/// <reference types="bun-types" />

// Collaborative editing relay server.
// Broadcasts CRDT updates between all clients in a room.
// Usage: bun run demo/relay-server.ts

const PORT = 4567;

interface RelaySocket {
  data?: { room?: string };
  send(data: Uint8Array | ArrayBuffer): void;
}

// Room state: accumulated updates per document
const rooms = new Map<
  string,
  {
    clients: Set<RelaySocket>;
    updates: Uint8Array[]; // accumulated updates for late joiners
  }
>();

function toUint8Array(
  message: string | ArrayBuffer | ArrayBufferView,
): Uint8Array | null {
  if (typeof message === "string") return null;
  if (message instanceof ArrayBuffer) {
    return new Uint8Array(message);
  }
  return new Uint8Array(message.buffer, message.byteOffset, message.byteLength);
}

Bun.serve<{ room: string }>({
  port: PORT,
  fetch(req, server) {
    const url = new URL(req.url);
    // Room ID from URL path: /doc/my-document-id
    const match = url.pathname.match(/^\/doc\/(.+)$/);
    const room = match?.[1];
    if (room && server.upgrade(req, { data: { room } })) {
      return;
    }
    return new Response("WebSocket relay server. Connect to /doc/<room-id>", {
      status: 200,
    });
  },
  websocket: {
    open(ws) {
      const room = ws.data?.room;
      if (!room) return;
      if (!rooms.has(room)) {
        rooms.set(room, { clients: new Set(), updates: [] });
      }
      const r = rooms.get(room)!;
      r.clients.add(ws);

      // Send accumulated state to the new client.
      // Protocol: first message type byte, then payload.
      // 0x00 = sync update, 0x01 = awareness
      for (const update of r.updates) {
        const msg = new Uint8Array(1 + update.length);
        msg[0] = 0x00; // sync update
        msg.set(update, 1);
        ws.send(msg);
      }

      console.log(`[${room}] Client connected (${r.clients.size} total)`);
    },
    message(ws, message) {
      const room = ws.data?.room;
      if (!room) return;
      const r = rooms.get(room);
      if (!r) return;

      const data = toUint8Array(message);
      if (!data || data.length === 0) return;
      const type = data[0];
      const payload = data.slice(1);

      if (type === 0x00) {
        // Sync update -- store and broadcast to others
        r.updates.push(payload);
        for (const client of r.clients) {
          if (client !== ws) {
            client.send(data);
          }
        }
      } else if (type === 0x01) {
        // Awareness -- broadcast to others (don't store)
        for (const client of r.clients) {
          if (client !== ws) {
            client.send(data);
          }
        }
      }
    },
    close(ws) {
      const room = ws.data?.room;
      if (!room) return;
      const r = rooms.get(room);
      if (r) {
        r.clients.delete(ws);
        console.log(`[${room}] Client disconnected (${r.clients.size} total)`);
        if (r.clients.size === 0) {
          // Keep room state for a while in case clients reconnect
          setTimeout(() => {
            const current = rooms.get(room);
            if (current && current.clients.size === 0) {
              rooms.delete(room);
              console.log(`[${room}] Room cleaned up`);
            }
          }, 60000); // 1 minute
        }
      }
    },
  },
});

console.log(`Relay server running on ws://localhost:${PORT}`);
console.log(`Connect to ws://localhost:${PORT}/doc/<room-id>`);
