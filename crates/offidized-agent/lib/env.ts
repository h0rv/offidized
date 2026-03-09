import { resolve } from "path";

const ENV_PATH = resolve(import.meta.dir, "..", ".env");
try {
  const envFile = await Bun.file(ENV_PATH).text();
  for (const line of envFile.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const val = trimmed
      .slice(eq + 1)
      .trim()
      .replace(/^["']|["']$/g, "");
    if (!process.env[key]) process.env[key] = val;
  }
  console.log("Loaded .env file");
} catch {
  console.log("No .env file found, using environment variables");
}

console.log("API keys available:", {
  openrouter: process.env.OPENROUTER_API_KEY ? "yes" : "no",
  google: process.env.GOOGLE_GENERATIVE_AI_API_KEY ? "yes" : "no",
  anthropic: process.env.ANTHROPIC_API_KEY ? "yes" : "no",
});
