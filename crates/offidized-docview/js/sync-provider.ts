// WebSocket sync provider for collaborative document editing.
//
// Wraps the DocEdit CRDT sync methods and the relay server binary
// protocol to enable real-time collaboration between multiple editors.

import type { DocEdit } from "../pkg/offidized_docview.js";

/** Message type prefixes for the relay server binary protocol. */
const MSG_SYNC = 0x00;
const MSG_AWARENESS = 0x01;
const MSG_REPLACE = 0x02;
const MSG_REQUEST_STATE = 0x03;
const FRAMED_GENERATION_BYTES = 4;
const RESYNC_INTERVAL_MS = 1200;
const RESYNC_BURST_COUNT = 3;
const AWARENESS_HEARTBEAT_INTERVAL_MS = 2000;
const AWARENESS_STALE_AFTER_MS = 8000;
const DISCONNECT_BROADCAST_REPEAT_MS = [24, 72] as const;
const DISCONNECT_BROADCAST_CLOSE_DELAY_MS = 120;

type TransportMode = "offline" | "broadcast" | "websocket" | "hybrid";

interface SyncStatus {
  wsConnected: boolean;
  bcConnected: boolean;
  mode: TransportMode;
  resyncing: boolean;
  divergenceSuspected: boolean;
  repairRequests: number;
}

export type AwarenessState = Record<string, unknown> | null;

export interface RemoteAwarenessPeer {
  senderId: string;
  state: Record<string, unknown>;
  lastUpdatedAt: number;
}

interface AwarenessEnvelope {
  seq: number;
  state: AwarenessState;
}

type BroadcastEnvelope =
  | {
      t: "hello" | "request_state" | "bye";
      roomId: string;
      senderId: string;
      generation?: number;
    }
  | {
      t: "sync";
      roomId: string;
      senderId: string;
      generation: number;
      payload: ArrayBuffer;
    }
  | {
      t: "replace";
      roomId: string;
      senderId: string;
      generation: number;
      payload: ArrayBuffer;
    }
  | {
      t: "awareness";
      roomId: string;
      senderId: string;
      awareness: AwarenessEnvelope;
    };

interface RemoteAwarenessEntry {
  seq: number;
  state: Record<string, unknown>;
  lastUpdatedAt: number;
}

interface AwarenessWireMessage {
  roomId: string;
  senderId: string;
  awareness: AwarenessEnvelope;
}

interface SyncDebugDetail {
  kind:
    | "awareness-clear"
    | "awareness-bye-recv"
    | "awareness-bye-send"
    | "awareness-expire"
    | "awareness-recv"
    | "awareness-send"
    | "bc-connect"
    | "bc-send"
    | "bc-recv"
    | "divergence-cleared"
    | "divergence-detected"
    | "resync-send"
    | "state-request-send"
    | "ws-open"
    | "ws-close"
    | "ws-error"
    | "ws-send"
    | "ws-recv";
  roomId: string;
  bytes?: number;
}

/**
 * WebSocket-based sync provider for collaborative editing.
 *
 * Connects to a relay server, exchanges CRDT state vectors and updates,
 * and calls back into the editor when remote changes arrive.
 */
export class SyncProvider {
  private ws: WebSocket | null = null;
  private wsConnected = false;
  private bc: BroadcastChannel | null = null;
  private bcConnected = false;
  private docEdit: DocEdit;
  private roomId: string;
  private wsUrl: string;
  private senderId: string;
  private onRemoteUpdate: (() => void) | null = null;
  private onReplaceSnapshot:
    | ((snapshot: Uint8Array, generation: number) => void)
    | null = null;
  private onStatusUpdate: ((status: SyncStatus) => void) | null = null;
  private onAwarenessUpdate:
    | ((peers: ReadonlyArray<RemoteAwarenessPeer>) => void)
    | null = null;

  /** Last-sent state vector, used for diff-based broadcasting. */
  private lastSentSv: Uint8Array | null = null;

  /** Guard against broadcasting while applying a remote update. */
  private applyingRemote = false;
  private needsResync = false;
  private resyncBurstsRemaining = 0;
  private resyncTimer: number | null = null;
  private awarenessTimer: number | null = null;
  private localAwareness: AwarenessState = null;
  private localAwarenessSeq = 0;
  private awarenessPausedForTests = false;
  private remoteAwareness = new Map<string, RemoteAwarenessEntry>();
  private remoteAwarenessSeq = new Map<string, number>();
  private textEncoder = new TextEncoder();
  private textDecoder = new TextDecoder();
  private generation = 0;
  private lastAppliedReplaceGeneration = -1;
  private pendingReplaceSnapshot: Uint8Array | null = null;
  private transportPausedForTests = false;
  private resyncing = false;
  private divergenceSuspected = false;
  private repairRequests = 0;

  constructor(docEdit: DocEdit, roomId: string, opts?: { wsUrl?: string }) {
    this.docEdit = docEdit;
    this.roomId = roomId;
    this.wsUrl = opts?.wsUrl ?? "ws://localhost:4567";
    this.senderId =
      globalThis.crypto?.randomUUID?.() ??
      `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  /** Connect to the relay server. On connect, send our state vector + full state. */
  connect(): void {
    if (this.ws || this.bc) return;

    this.lastSentSv = this.docEdit.encodeStateVector();
    this.startResyncTicker();
    this.startAwarenessTicker();
    this.connectBroadcastChannel();
    this.connectWebSocket();
    this.emitStatus();
  }

  /** Disconnect from the relay server. */
  disconnect(): void {
    const bcToClose = this.bc;
    const disconnectAwareness =
      this.localAwarenessSeq > 0
        ? {
            seq: this.localAwarenessSeq + 1,
            state: null,
          }
        : null;
    this.localAwareness = null;
    if (disconnectAwareness) {
      this.localAwarenessSeq = disconnectAwareness.seq;
    }

    if (this.ws) {
      this.ws.close();
      this.ws = null;
    }
    this.wsConnected = false;

    if (bcToClose && this.bcConnected) {
      this.scheduleBroadcastDisconnectNotices(bcToClose, disconnectAwareness);
    } else if (bcToClose) {
      bcToClose.close();
    }
    this.bc = null;
    this.bcConnected = false;
    this.stopResyncTicker();
    this.stopAwarenessTicker();
    this.needsResync = false;
    this.resyncBurstsRemaining = 0;
    this.clearRepairState();
    this.awarenessPausedForTests = false;
    this.clearRemoteAwareness();
    this.emitStatus();
  }

  /**
   * Send local changes to remote peers.
   *
   * Uses diff-based encoding for efficiency: compares against the
   * last-sent state vector and only sends operations the remote
   * hasn't seen.
   */
  broadcastUpdate(): void {
    if (this.applyingRemote) return;
    if (!this.wsConnected && !this.bcConnected) return;

    let update: Uint8Array;
    if (this.lastSentSv) {
      // Diff-based: only send what changed since last broadcast.
      update = this.docEdit.encodeDiff(this.lastSentSv);
    } else {
      // No previous state vector: send full state.
      update = this.docEdit.encodeStateAsUpdate();
    }

    let sent = false;
    if (this.wsConnected) {
      this.sendWsSync(update);
      sent = true;
    }
    if (this.bcConnected) {
      this.sendBroadcastSync(update);
      sent = true;
    }

    if (sent) {
      // Update the stored state vector for the next diff.
      this.lastSentSv = this.docEdit.encodeStateVector();
      this.queueResync();
    }
  }

  /** Register callback for when a remote update is applied. */
  onRemote(cb: () => void): void {
    this.onRemoteUpdate = cb;
  }

  /** Register callback for a remote document replacement snapshot. */
  onReplace(
    cb: ((snapshot: Uint8Array, generation: number) => void) | null,
  ): void {
    this.onReplaceSnapshot = cb;
  }

  /** Register callback for sync transport status changes. */
  onStatus(cb: (status: SyncStatus) => void): void {
    this.onStatusUpdate = cb;
    this.emitStatus();
  }

  /** Register callback for remote awareness changes. */
  onAwareness(
    cb: ((peers: ReadonlyArray<RemoteAwarenessPeer>) => void) | null,
  ): void {
    this.onAwarenessUpdate = cb;
    this.emitAwareness();
  }

  /** Update the local awareness state and broadcast it independently. */
  setLocalAwareness(state: AwarenessState): void {
    const normalized = this.normalizeAwarenessState(state);
    if (this.awarenessStatesEqual(this.localAwareness, normalized)) return;
    this.localAwareness = normalized;
    this.localAwarenessSeq += 1;
    this.broadcastAwareness();
  }

  /** Snapshot of all currently-active remote awareness peers. */
  getRemoteAwareness(): ReadonlyArray<RemoteAwarenessPeer> {
    return this.snapshotRemoteAwareness();
  }

  /** Whether the provider is currently connected. */
  isConnected(): boolean {
    return this.wsConnected || this.bcConnected;
  }

  /** Swap the active DocEdit instance without changing room identity. */
  attachDocEdit(docEdit: DocEdit, generation = this.generation): void {
    this.docEdit = docEdit;
    this.generation = generation;
    this.lastSentSv = this.docEdit.encodeStateVector();
    this.needsResync = false;
    this.resyncBurstsRemaining = 0;
    this.clearRepairState();
  }

  /** Broadcast an out-of-band room-level document replacement snapshot. */
  broadcastReplace(snapshot: Uint8Array): void {
    const copy = snapshot.slice();
    this.generation += 1;
    this.lastSentSv = this.docEdit.encodeStateVector();
    this.pendingReplaceSnapshot = copy;
    if (this.wsConnected) this.sendWsReplace(copy);
    if (this.bcConnected) this.sendBroadcastReplace(copy);
    this.pendingReplaceSnapshot = null;
    this.queueResync(2);
  }

  /**
   * Force an immediate full-state broadcast.
   *
   * Diff-based sync is the default, but richer structural edits like tables
   * and inline images benefit from an eager anti-entropy flush so peers
   * converge without waiting for the periodic resync ticker.
   */
  flushFullState(): void {
    if (!this.wsConnected && !this.bcConnected) return;
    this.sendFullState();
    this.queueResync(1);
  }

  /** Destroy the provider. */
  destroy(): void {
    this.disconnect();
    this.onRemoteUpdate = null;
    this.onReplaceSnapshot = null;
    this.onStatusUpdate = null;
    this.onAwarenessUpdate = null;
    this.lastSentSv = null;
    this.localAwareness = null;
    this.localAwarenessSeq = 0;
    this.awarenessPausedForTests = false;
    this.remoteAwareness.clear();
    this.remoteAwarenessSeq.clear();
  }

  /**
   * Test-only helper: pause or resume awareness transport without sending an
   * explicit awareness clear/bye. This allows browser harnesses to prove
   * stale-peer expiry on the surviving client.
   */
  setAwarenessPausedForTests(paused: boolean): void {
    this.awarenessPausedForTests = paused;
  }

  /**
   * Test-only helper: pause sync transport delivery in both directions
   * without destroying the provider. Resuming triggers an explicit state
   * request so stale peers repair immediately.
   */
  setTransportPausedForTests(paused: boolean): void {
    const wasPaused = this.transportPausedForTests;
    this.transportPausedForTests = paused;
    if (wasPaused && !paused && (this.wsConnected || this.bcConnected)) {
      this.requestCurrentState();
    }
  }

  // -- Private helpers --

  private connectWebSocket(): void {
    const url = `${this.wsUrl}/doc/${this.roomId}`;
    const ws = new WebSocket(url);
    ws.binaryType = "arraybuffer";
    this.ws = ws;

    ws.addEventListener("open", () => {
      this.wsConnected = true;
      this.emitStatus();
      this.emitDebug({ kind: "ws-open", roomId: this.roomId });

      // Send our full state so late joiners get everything.
      // (Don't send the state vector — it's not a valid update and
      // would crash peers that try to applyUpdate() on it.)
      if (this.pendingReplaceSnapshot) {
        this.sendWsReplace(this.pendingReplaceSnapshot);
        this.pendingReplaceSnapshot = null;
      } else {
        const fullUpdate = this.docEdit.encodeStateAsUpdate();
        this.sendWsSync(fullUpdate);
      }
      this.sendWsAwareness();
      this.queueResync(2);
    });

    ws.addEventListener("message", (event: MessageEvent) => {
      if (this.transportPausedForTests) return;
      const data = new Uint8Array(event.data as ArrayBuffer);
      if (data.length < 2) return;

      const msgType = data[0];

      if (msgType === MSG_SYNC) {
        const framed = this.decodeFramedPayload(data);
        if (!framed) return;
        this.emitDebug({
          kind: "ws-recv",
          roomId: this.roomId,
          bytes: framed.payload.length,
        });
        this.handleSyncMessage(framed.payload, framed.generation);
        return;
      }

      if (msgType === MSG_AWARENESS) {
        this.handleWsAwarenessMessage(data.subarray(1));
        return;
      }

      if (msgType === MSG_REPLACE) {
        const framed = this.decodeFramedPayload(data);
        if (!framed) return;
        this.handleReplaceMessage(framed.payload, framed.generation);
        return;
      }

      if (msgType === MSG_REQUEST_STATE) {
        const framed = this.decodeFramedPayload(data);
        if (!framed) return;
        this.handleStateRequest(framed.generation);
      }
    });

    ws.addEventListener("close", () => {
      this.wsConnected = false;
      this.ws = null;
      this.emitDebug({ kind: "ws-close", roomId: this.roomId });
      this.emitStatus();
    });

    ws.addEventListener("error", () => {
      this.wsConnected = false;
      this.emitDebug({ kind: "ws-error", roomId: this.roomId });
      this.emitStatus();
    });
  }

  private connectBroadcastChannel(): void {
    if (typeof BroadcastChannel === "undefined") return;
    const channel = new BroadcastChannel(
      `offidized-docview-sync:${this.roomId}`,
    );
    this.bc = channel;
    this.bcConnected = true;
    this.emitStatus();
    this.emitDebug({ kind: "bc-connect", roomId: this.roomId });

    channel.addEventListener("message", (event: MessageEvent<unknown>) => {
      this.handleBroadcastMessage(event.data);
    });

    // Signal readiness and request a full snapshot from any existing peers.
    this.postBroadcast({
      t: "hello",
      roomId: this.roomId,
      senderId: this.senderId,
      generation: this.generation,
    });
    this.postBroadcast({
      t: "request_state",
      roomId: this.roomId,
      senderId: this.senderId,
      generation: this.generation,
    });
    this.sendBroadcastAwareness();
  }

  private handleBroadcastMessage(raw: unknown): void {
    if (this.transportPausedForTests) return;
    const msg = this.parseBroadcastMessage(raw);
    if (!msg) return;
    if (msg.roomId !== this.roomId) return;
    if (msg.senderId === this.senderId) return;

    if (msg.t === "bye") {
      this.emitDebug({
        kind: "awareness-bye-recv",
        roomId: this.roomId,
      });
      this.remoteAwarenessSeq.set(msg.senderId, Number.POSITIVE_INFINITY);
      if (this.remoteAwareness.delete(msg.senderId)) {
        this.emitDebug({
          kind: "awareness-clear",
          roomId: this.roomId,
        });
        this.emitAwareness();
      }
      return;
    }

    if (msg.t === "hello" || msg.t === "request_state") {
      if ((msg.generation ?? 0) < this.generation) {
        this.sendBroadcastReplace(this.snapshotBytes());
        this.sendBroadcastAwareness();
        this.queueResync(1);
        return;
      }
      if ((msg.generation ?? 0) > this.generation) {
        this.requestCurrentState();
        return;
      }
      // A peer joined or asked for state: send a full update snapshot.
      this.sendBroadcastSync(this.docEdit.encodeStateAsUpdate());
      this.sendBroadcastAwareness();
      this.queueResync(1);
      return;
    }

    if (msg.t === "sync" && msg.payload) {
      this.emitDebug({
        kind: "bc-recv",
        roomId: this.roomId,
        bytes: msg.payload.byteLength,
      });
      this.handleSyncMessage(new Uint8Array(msg.payload), msg.generation);
      return;
    }

    if (msg.t === "replace" && msg.payload) {
      this.handleReplaceMessage(new Uint8Array(msg.payload), msg.generation);
      return;
    }

    if (msg.t === "awareness") {
      this.emitDebug({
        kind: "awareness-recv",
        roomId: this.roomId,
      });
      this.applyRemoteAwareness(msg.senderId, msg.awareness);
    }
  }

  private parseBroadcastMessage(raw: unknown): BroadcastEnvelope | null {
    if (!raw || typeof raw !== "object") return null;
    const msg = raw as Partial<BroadcastEnvelope>;
    if (msg.roomId !== this.roomId) return null;
    if (typeof msg.senderId !== "string") return null;
    if (
      msg.t !== "hello" &&
      msg.t !== "request_state" &&
      msg.t !== "bye" &&
      msg.t !== "sync" &&
      msg.t !== "replace" &&
      msg.t !== "awareness"
    ) {
      return null;
    }
    if (
      (msg.t === "sync" || msg.t === "replace") &&
      typeof (msg as { generation?: unknown }).generation !== "number"
    ) {
      return null;
    }
    if (msg.t === "sync" && !(msg.payload instanceof ArrayBuffer)) return null;
    if (msg.t === "replace" && !(msg.payload instanceof ArrayBuffer)) {
      return null;
    }
    if (
      msg.t === "awareness" &&
      !this.isAwarenessEnvelope((msg as { awareness?: unknown }).awareness)
    ) {
      return null;
    }
    return msg as BroadcastEnvelope;
  }

  /** Handle an incoming sync message (remote CRDT update). */
  private handleSyncMessage(payload: Uint8Array, generation: number): void {
    if (generation < this.generation) return;
    if (generation > this.generation) {
      this.requestCurrentState();
      return;
    }
    this.applyingRemote = true;
    try {
      this.docEdit.applyUpdate(payload);
      // Refresh diff baseline so future local broadcasts don't resend remote ops.
      this.lastSentSv = this.docEdit.encodeStateVector();
      this.clearRepairState();
      this.onRemoteUpdate?.();
    } catch (err) {
      console.error("SyncProvider: failed to apply remote update:", err);
    } finally {
      this.applyingRemote = false;
    }
  }

  private handleReplaceMessage(payload: Uint8Array, generation: number): void {
    if (generation < this.generation) return;
    if (generation <= this.lastAppliedReplaceGeneration) return;
    if (generation > this.generation) {
      this.beginRepair();
    }
    this.lastAppliedReplaceGeneration = generation;
    this.generation = generation;
    this.applyingRemote = true;
    try {
      this.onReplaceSnapshot?.(payload.slice(), generation);
    } finally {
      this.applyingRemote = false;
    }
  }

  private handleStateRequest(requestGeneration: number): void {
    if (requestGeneration < this.generation) {
      this.sendWsReplace(this.snapshotBytes());
      return;
    }
    if (requestGeneration === this.generation) {
      this.sendWsSync(this.docEdit.encodeStateAsUpdate());
    }
  }

  private postBroadcast(message: BroadcastEnvelope): void {
    if (this.transportPausedForTests) return;
    if (!this.bcConnected || !this.bc) return;
    this.bc.postMessage(message);
  }

  private sendBroadcastAwareness(): void {
    if (!this.bcConnected || !this.bc) return;
    const awareness = this.currentAwarenessEnvelope();
    if (!awareness) return;
    this.emitDebug({
      kind: "awareness-send",
      roomId: this.roomId,
    });
    this.postBroadcast({
      t: "awareness",
      roomId: this.roomId,
      senderId: this.senderId,
      awareness,
    });
  }

  private sendBroadcastSync(payload: Uint8Array): void {
    const copy = payload.slice();
    this.emitDebug({
      kind: "bc-send",
      roomId: this.roomId,
      bytes: copy.length,
    });
    this.postBroadcast({
      t: "sync",
      roomId: this.roomId,
      senderId: this.senderId,
      generation: this.generation,
      payload: copy.buffer,
    });
  }

  private sendBroadcastReplace(payload: Uint8Array): void {
    const copy = payload.slice();
    this.emitDebug({
      kind: "bc-send",
      roomId: this.roomId,
      bytes: copy.length,
    });
    this.postBroadcast({
      t: "replace",
      roomId: this.roomId,
      senderId: this.senderId,
      generation: this.generation,
      payload: copy.buffer,
    });
  }

  /** Send a sync message to relay server (0x00 prefix + payload). */
  private sendWsSync(payload: Uint8Array): void {
    if (this.transportPausedForTests) return;
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    const msg = this.frameBinaryMessage(MSG_SYNC, payload);
    this.emitDebug({
      kind: "ws-send",
      roomId: this.roomId,
      bytes: payload.length,
    });
    this.ws.send(msg.buffer);
  }

  private sendWsReplace(payload: Uint8Array): void {
    if (this.transportPausedForTests) return;
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    const msg = this.frameBinaryMessage(MSG_REPLACE, payload);
    this.emitDebug({
      kind: "ws-send",
      roomId: this.roomId,
      bytes: payload.length,
    });
    this.ws.send(msg.buffer);
  }

  private sendWsStateRequest(): void {
    if (this.transportPausedForTests) return;
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    const msg = this.frameBinaryMessage(MSG_REQUEST_STATE, new Uint8Array(0));
    this.ws.send(msg.buffer);
  }

  private sendWsAwareness(): void {
    if (this.transportPausedForTests) return;
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    const awareness = this.currentAwarenessEnvelope();
    if (!awareness) return;
    const body = this.textEncoder.encode(
      JSON.stringify({
        roomId: this.roomId,
        senderId: this.senderId,
        awareness,
      } satisfies AwarenessWireMessage),
    );
    const msg = new Uint8Array(1 + body.length);
    msg[0] = MSG_AWARENESS;
    msg.set(body, 1);
    this.emitDebug({
      kind: "awareness-send",
      roomId: this.roomId,
      bytes: body.length,
    });
    this.ws.send(msg.buffer);
  }

  private startResyncTicker(): void {
    if (this.resyncTimer != null) return;
    this.resyncTimer = window.setInterval(() => {
      if (!this.needsResync) return;
      if (!this.wsConnected && !this.bcConnected) return;
      this.sendFullState();
      this.resyncBurstsRemaining = Math.max(0, this.resyncBurstsRemaining - 1);
      this.needsResync = this.resyncBurstsRemaining > 0;
    }, RESYNC_INTERVAL_MS);
  }

  private stopResyncTicker(): void {
    if (this.resyncTimer == null) return;
    window.clearInterval(this.resyncTimer);
    this.resyncTimer = null;
  }

  private startAwarenessTicker(): void {
    if (this.awarenessTimer != null) return;
    this.awarenessTimer = window.setInterval(() => {
      this.expireStaleAwareness();
      if (this.awarenessPausedForTests || this.transportPausedForTests) return;
      if (!this.localAwareness) return;
      if (!this.wsConnected && !this.bcConnected) return;
      this.localAwarenessSeq += 1;
      this.broadcastAwareness();
    }, AWARENESS_HEARTBEAT_INTERVAL_MS);
  }

  private stopAwarenessTicker(): void {
    if (this.awarenessTimer == null) return;
    window.clearInterval(this.awarenessTimer);
    this.awarenessTimer = null;
  }

  private sendFullState(): void {
    const fullUpdate = this.docEdit.encodeStateAsUpdate();
    if (this.wsConnected) this.sendWsSync(fullUpdate);
    if (this.bcConnected) this.sendBroadcastSync(fullUpdate);
    this.lastSentSv = this.docEdit.encodeStateVector();
    this.emitDebug({
      kind: "resync-send",
      roomId: this.roomId,
      bytes: fullUpdate.length,
    });
  }

  private queueResync(burstCount = RESYNC_BURST_COUNT): void {
    this.needsResync = true;
    this.resyncBurstsRemaining = Math.max(
      this.resyncBurstsRemaining,
      burstCount,
    );
  }

  private broadcastAwareness(): void {
    if (this.awarenessPausedForTests) return;
    if (this.wsConnected) this.sendWsAwareness();
    if (this.bcConnected) this.sendBroadcastAwareness();
  }

  private handleWsAwarenessMessage(payload: Uint8Array): void {
    let raw: unknown;
    try {
      raw = JSON.parse(this.textDecoder.decode(payload));
    } catch (err) {
      console.error("SyncProvider: failed to parse awareness message:", err);
      return;
    }
    if (!this.isAwarenessWireMessage(raw)) return;
    if (raw.senderId === this.senderId) return;
    this.emitDebug({
      kind: "awareness-recv",
      roomId: this.roomId,
      bytes: payload.length,
    });
    this.applyRemoteAwareness(raw.senderId, raw.awareness);
  }

  private applyRemoteAwareness(
    senderId: string,
    awareness: AwarenessEnvelope,
  ): void {
    const lastSeenSeq = this.remoteAwarenessSeq.get(senderId);
    if (lastSeenSeq != null && awareness.seq <= lastSeenSeq) return;
    this.remoteAwarenessSeq.set(senderId, awareness.seq);

    if (awareness.state == null) {
      if (this.remoteAwareness.delete(senderId)) {
        this.emitDebug({
          kind: "awareness-clear",
          roomId: this.roomId,
        });
        this.emitAwareness();
      }
      return;
    }

    this.remoteAwareness.set(senderId, {
      seq: awareness.seq,
      state: awareness.state,
      lastUpdatedAt: Date.now(),
    });
    this.emitAwareness();
  }

  private expireStaleAwareness(): void {
    const cutoff = Date.now() - AWARENESS_STALE_AFTER_MS;
    let changed = false;
    for (const [senderId, entry] of this.remoteAwareness.entries()) {
      if (entry.lastUpdatedAt >= cutoff) continue;
      this.remoteAwareness.delete(senderId);
      changed = true;
      this.emitDebug({
        kind: "awareness-expire",
        roomId: this.roomId,
      });
    }
    if (changed) {
      this.emitAwareness();
    }
  }

  private clearRemoteAwareness(): void {
    if (this.remoteAwareness.size === 0) return;
    this.remoteAwareness.clear();
    this.remoteAwarenessSeq.clear();
    this.emitDebug({
      kind: "awareness-clear",
      roomId: this.roomId,
    });
    this.emitAwareness();
  }

  private scheduleBroadcastDisconnectNotices(
    channel: BroadcastChannel,
    awareness: AwarenessEnvelope | null,
  ): void {
    const send = () => {
      if (awareness) {
        channel.postMessage({
          t: "awareness",
          roomId: this.roomId,
          senderId: this.senderId,
          awareness,
        } satisfies BroadcastEnvelope);
      }
      this.emitDebug({
        kind: "awareness-bye-send",
        roomId: this.roomId,
      });
      channel.postMessage({
        t: "bye",
        roomId: this.roomId,
        senderId: this.senderId,
      } satisfies BroadcastEnvelope);
    };

    send();
    if (typeof window === "undefined") {
      channel.close();
      return;
    }

    for (const delay of DISCONNECT_BROADCAST_REPEAT_MS) {
      window.setTimeout(send, delay);
    }
    window.setTimeout(() => {
      channel.close();
    }, DISCONNECT_BROADCAST_CLOSE_DELAY_MS);
  }

  private emitAwareness(): void {
    this.onAwarenessUpdate?.(this.snapshotRemoteAwareness());
  }

  private frameBinaryMessage(prefix: number, payload: Uint8Array): Uint8Array {
    const msg = new Uint8Array(1 + FRAMED_GENERATION_BYTES + payload.length);
    msg[0] = prefix;
    new DataView(msg.buffer).setUint32(1, this.generation);
    msg.set(payload, 1 + FRAMED_GENERATION_BYTES);
    return msg;
  }

  private decodeFramedPayload(
    data: Uint8Array,
  ): { generation: number; payload: Uint8Array } | null {
    if (data.length < 1 + FRAMED_GENERATION_BYTES) return null;
    const generation = new DataView(
      data.buffer,
      data.byteOffset + 1,
      FRAMED_GENERATION_BYTES,
    ).getUint32(0);
    return {
      generation,
      payload: data.subarray(1 + FRAMED_GENERATION_BYTES),
    };
  }

  private requestCurrentState(): void {
    this.beginRepair();
    this.emitDebug({
      kind: "state-request-send",
      roomId: this.roomId,
    });
    if (this.bcConnected) {
      this.postBroadcast({
        t: "request_state",
        roomId: this.roomId,
        senderId: this.senderId,
        generation: this.generation,
      });
    }
    if (this.wsConnected) {
      this.sendWsStateRequest();
    }
    this.emitStatus();
  }

  private beginRepair(): void {
    this.repairRequests += 1;
    this.resyncing = true;
    if (!this.divergenceSuspected) {
      this.divergenceSuspected = true;
      this.emitDebug({
        kind: "divergence-detected",
        roomId: this.roomId,
      });
    }
    this.emitStatus();
  }

  private clearRepairState(): void {
    const hadRepairState = this.resyncing || this.divergenceSuspected;
    const hadDivergence = this.divergenceSuspected;
    this.resyncing = false;
    this.divergenceSuspected = false;
    if (!hadRepairState) return;
    if (hadDivergence) {
      this.emitDebug({
        kind: "divergence-cleared",
        roomId: this.roomId,
      });
    }
    this.emitStatus();
  }

  private snapshotBytes(): Uint8Array {
    return this.docEdit.save().slice();
  }

  private snapshotRemoteAwareness(): ReadonlyArray<RemoteAwarenessPeer> {
    return [...this.remoteAwareness.entries()]
      .map(([senderId, entry]) => ({
        senderId,
        state: entry.state,
        lastUpdatedAt: entry.lastUpdatedAt,
      }))
      .sort((a, b) => a.senderId.localeCompare(b.senderId));
  }

  private currentAwarenessEnvelope(): AwarenessEnvelope | null {
    if (this.localAwareness === null && this.localAwarenessSeq === 0) {
      return null;
    }
    return {
      seq: this.localAwarenessSeq,
      state: this.localAwareness,
    };
  }

  private isAwarenessWireMessage(raw: unknown): raw is AwarenessWireMessage {
    if (!raw || typeof raw !== "object") return false;
    const message = raw as Partial<AwarenessWireMessage>;
    return (
      message.roomId === this.roomId &&
      typeof message.senderId === "string" &&
      this.isAwarenessEnvelope(message.awareness)
    );
  }

  private isAwarenessEnvelope(raw: unknown): raw is AwarenessEnvelope {
    if (!raw || typeof raw !== "object") return false;
    const awareness = raw as Partial<AwarenessEnvelope>;
    if (typeof awareness.seq !== "number" || !Number.isFinite(awareness.seq)) {
      return false;
    }
    return (
      awareness.state === null || this.isAwarenessStateRecord(awareness.state)
    );
  }

  private isAwarenessStateRecord(raw: unknown): raw is Record<string, unknown> {
    return !!raw && typeof raw === "object" && !Array.isArray(raw);
  }

  private normalizeAwarenessState(state: AwarenessState): AwarenessState {
    return this.isAwarenessStateRecord(state) ? state : null;
  }

  private awarenessStatesEqual(
    left: AwarenessState,
    right: AwarenessState,
  ): boolean {
    if (left === right) return true;
    return JSON.stringify(left) === JSON.stringify(right);
  }

  private currentMode(): TransportMode {
    if (this.wsConnected && this.bcConnected) return "hybrid";
    if (this.wsConnected) return "websocket";
    if (this.bcConnected) return "broadcast";
    return "offline";
  }

  private emitStatus(): void {
    this.onStatusUpdate?.({
      wsConnected: this.wsConnected,
      bcConnected: this.bcConnected,
      mode: this.currentMode(),
      resyncing: this.resyncing,
      divergenceSuspected: this.divergenceSuspected,
      repairRequests: this.repairRequests,
    });
  }

  private emitDebug(detail: SyncDebugDetail): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(
      new CustomEvent("docedit-sync-debug", {
        detail,
      }),
    );
  }
}
