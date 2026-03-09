import { resolve } from "path";

const OFX = resolve(import.meta.dir, "../../../target/release/ofx");

export let ofxHelp: string;
try {
  const p = Bun.spawnSync([OFX, "--help"]);
  ofxHelp = new TextDecoder().decode(p.stdout).trim();
  for (const sub of [
    "info",
    "read",
    "set",
    "patch",
    "create",
    "replace",
    "charts",
    "pivots",
    "eval",
    "derive",
    "apply",
  ]) {
    const s = Bun.spawnSync([OFX, sub, "--help"]);
    ofxHelp +=
      `\n\n--- ofx ${sub} --help ---\n` +
      new TextDecoder().decode(s.stdout).trim();
  }
} catch {
  ofxHelp = "(ofx CLI not found — tools will not work)";
}
