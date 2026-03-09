export interface Message {
  role: "user" | "assistant";
  content: string;
}

export interface FileVersion {
  n: number;
  ts: number;
  path: string;
  isOriginal?: boolean;
}

export interface Session {
  id: string;
  title: string;
  conversation: Message[];
  currentFile: string;
  currentFilePath: string;
  workspaceDir: string;
  createdAt: number;
  lastActive: number;
  versions: FileVersion[];
}

export interface SessionSummary {
  id: string;
  title: string;
  lastActive: number;
  hasFile: boolean;
  messageCount: number;
}

const sessions = new Map<string, Session>();

export function createSession(): Session {
  const id = crypto.randomUUID();
  const sess: Session = {
    id,
    title: "",
    conversation: [],
    currentFile: "",
    currentFilePath: "",
    workspaceDir: "",
    createdAt: Date.now(),
    lastActive: Date.now(),
    versions: [],
  };
  sessions.set(id, sess);
  return sess;
}

export function getSession(id: string): Session | null {
  const sess = sessions.get(id);
  if (!sess) return null;
  sess.lastActive = Date.now();
  return sess;
}

export function getOrCreateSession(id: string | null): Session {
  if (id) {
    const existing = sessions.get(id);
    if (existing) {
      existing.lastActive = Date.now();
      return existing;
    }
  }
  return createSession();
}

export function deleteSession(id: string): void {
  sessions.delete(id);
}

export function hasSession(id: string): boolean {
  return sessions.has(id);
}

export function listSessions(): SessionSummary[] {
  return Array.from(sessions.values())
    .sort((a, b) => b.lastActive - a.lastActive)
    .map((s) => ({
      id: s.id,
      title: s.title || "New chat",
      lastActive: s.lastActive,
      hasFile: !!s.currentFile,
      messageCount: s.conversation.filter((m) => m.role === "user").length,
    }));
}

export function parseSid(cookieHeader: string | null): string | null {
  if (!cookieHeader) return null;
  for (const part of cookieHeader.split(";")) {
    const [k, ...rest] = part.trim().split("=");
    if (k?.trim() === "sid") return rest.join("=").trim();
  }
  return null;
}

export function makeSetCookie(sid: string): string {
  return `sid=${sid}; Path=/; Max-Age=7200; SameSite=Lax; HttpOnly`;
}

setInterval(
  async () => {
    const TWO_HOURS = 2 * 60 * 60 * 1000;
    const now = Date.now();
    for (const [id, sess] of sessions) {
      if (now - sess.lastActive > TWO_HOURS) {
        sessions.delete(id);
        if (sess.workspaceDir) {
          try {
            await Bun.$`rm -rf ${sess.workspaceDir}`.quiet();
          } catch {
            // ignore cleanup errors
          }
        }
      }
    }
  },
  15 * 60 * 1000,
);
