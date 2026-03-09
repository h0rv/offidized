import type { SectionModel } from "../types.ts";
import {
  layoutDocument,
  type LayoutConfig,
  type LayoutResult,
  type MeasuredBlock,
} from "./layout-engine.ts";

export const LAYOUT_WORKER_PROTOCOL_NAME = "offidized-docview.layout" as const;
export const LAYOUT_WORKER_PROTOCOL_VERSION = 1 as const;

export interface LayoutJobRequest {
  blocks: MeasuredBlock[];
  sections: SectionModel[];
  docVersion: number;
  config?: Partial<LayoutConfig>;
}

export interface LayoutWorkerRequest {
  kind: "layout.request";
  protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
  version: typeof LAYOUT_WORKER_PROTOCOL_VERSION;
  requestId: number;
  payload: LayoutJobRequest;
}

export interface LayoutWorkerSuccess {
  kind: "layout.result";
  protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
  version: typeof LAYOUT_WORKER_PROTOCOL_VERSION;
  requestId: number;
  result: LayoutResult;
}

export interface LayoutWorkerError {
  kind: "layout.error";
  protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
  version: typeof LAYOUT_WORKER_PROTOCOL_VERSION;
  requestId: number;
  error: string;
}

export type LayoutWorkerResponse = LayoutWorkerSuccess | LayoutWorkerError;

interface PendingRequest {
  req: LayoutJobRequest;
  resolve: (result: LayoutResult) => void;
  reject: (err: Error) => void;
}

function runLayoutSync(req: LayoutJobRequest): LayoutResult {
  return layoutDocument(req.blocks, req.sections, req.docVersion, req.config);
}

function isProtocolMessage(value: unknown): value is {
  protocol: string;
  version: number;
  requestId: number;
  kind: string;
} {
  if (!value || typeof value !== "object") return false;
  const msg = value as Record<string, unknown>;
  return (
    typeof msg.protocol === "string" &&
    typeof msg.version === "number" &&
    typeof msg.requestId === "number" &&
    typeof msg.kind === "string"
  );
}

function resolveLayoutWorkerUrl(): URL {
  const moduleUrl = import.meta.url;
  if (
    moduleUrl.startsWith("http://") ||
    moduleUrl.startsWith("https://") ||
    moduleUrl.startsWith("blob:")
  ) {
    return new URL("./layout-worker.ts", moduleUrl);
  }
  return new URL("js/canvas/layout-worker.ts", document.baseURI);
}

export class LayoutClient {
  private worker: Worker | null = null;
  private nextRequestId = 1;
  private pending = new Map<number, PendingRequest>();

  constructor() {
    this.startWorker();
  }

  hasWorker(): boolean {
    return this.worker !== null;
  }

  layoutSync(req: LayoutJobRequest): LayoutResult {
    return runLayoutSync(req);
  }

  layout(req: LayoutJobRequest): Promise<LayoutResult> {
    const worker = this.worker;
    if (!worker) {
      return Promise.resolve(runLayoutSync(req));
    }

    const requestId = this.nextRequestId++;
    const msg: LayoutWorkerRequest = {
      kind: "layout.request",
      protocol: LAYOUT_WORKER_PROTOCOL_NAME,
      version: LAYOUT_WORKER_PROTOCOL_VERSION,
      requestId,
      payload: req,
    };

    return new Promise<LayoutResult>((resolve, reject) => {
      this.pending.set(requestId, { req, resolve, reject });
      try {
        worker.postMessage(msg);
      } catch (err) {
        this.pending.delete(requestId);
        this.disableWorker(err, true);
        try {
          resolve(runLayoutSync(req));
        } catch (syncErr) {
          reject(
            syncErr instanceof Error
              ? syncErr
              : new Error(String(syncErr ?? "layout failed")),
          );
        }
      }
    });
  }

  destroy(): void {
    const worker = this.worker;
    if (worker) {
      worker.onmessage = null;
      worker.onerror = null;
      worker.onmessageerror = null;
      worker.terminate();
    }
    this.worker = null;

    for (const pending of this.pending.values()) {
      pending.reject(new Error("layout client destroyed"));
    }
    this.pending.clear();
  }

  private startWorker(): void {
    if (typeof Worker === "undefined") {
      return;
    }

    try {
      const worker = new Worker(resolveLayoutWorkerUrl(), {
        type: "module",
        name: "offidized-layout-worker",
      });
      worker.onmessage = (event: MessageEvent<unknown>) => {
        this.handleMessage(event.data);
      };
      worker.onerror = (event: ErrorEvent) => {
        const message = event.message || "layout worker error";
        this.disableWorker(new Error(message), true);
      };
      worker.onmessageerror = () => {
        this.disableWorker(new Error("layout worker message error"), true);
      };
      this.worker = worker;
    } catch (err) {
      this.disableWorker(err, false);
    }
  }

  private handleMessage(raw: unknown): void {
    if (!isProtocolMessage(raw)) return;
    if (raw.protocol !== LAYOUT_WORKER_PROTOCOL_NAME) return;

    const pending = this.pending.get(raw.requestId);
    if (!pending) return;

    this.pending.delete(raw.requestId);

    if (raw.version !== LAYOUT_WORKER_PROTOCOL_VERSION) {
      try {
        pending.resolve(runLayoutSync(pending.req));
      } catch (err) {
        pending.reject(
          err instanceof Error
            ? err
            : new Error(String(err ?? "layout version mismatch")),
        );
      }
      return;
    }

    const msg = raw as LayoutWorkerResponse;

    if (msg.kind === "layout.result") {
      pending.resolve(msg.result);
      return;
    }

    if (msg.kind === "layout.error") {
      try {
        pending.resolve(runLayoutSync(pending.req));
      } catch (err) {
        pending.reject(
          err instanceof Error ? err : new Error(msg.error || "layout failed"),
        );
      }
      return;
    }

    pending.reject(new Error("unknown layout worker response"));
  }

  private disableWorker(reason: unknown, warn: boolean): void {
    const worker = this.worker;
    if (worker) {
      worker.onmessage = null;
      worker.onerror = null;
      worker.onmessageerror = null;
      worker.terminate();
    }
    this.worker = null;

    if (warn) {
      const text =
        reason instanceof Error
          ? reason.message
          : typeof reason === "string"
            ? reason
            : "layout worker unavailable";
      console.warn("layout worker disabled; using sync fallback:", text);
    }

    if (this.pending.size === 0) {
      return;
    }

    const pendingEntries = Array.from(this.pending.values());
    this.pending.clear();

    for (const pending of pendingEntries) {
      try {
        pending.resolve(runLayoutSync(pending.req));
      } catch (err) {
        pending.reject(
          err instanceof Error
            ? err
            : new Error(String(err ?? "layout fallback failed")),
        );
      }
    }
  }
}
