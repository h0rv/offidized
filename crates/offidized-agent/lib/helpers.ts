import { resolve } from "path";

export async function createWorkspace(): Promise<string> {
  const dir = resolve(`/tmp/offidized-${crypto.randomUUID().slice(0, 8)}`);
  await Bun.$`mkdir -p ${dir}`;
  return dir;
}

export function esc(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
