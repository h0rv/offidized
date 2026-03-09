import { streamText, stepCountIs } from "ai";
import type { StreamTextResult } from "ai";
import { createAnthropic } from "@ai-sdk/anthropic";
import { createGoogleGenerativeAI } from "@ai-sdk/google";
import { createOpenRouter } from "@openrouter/ai-sdk-provider";
import type { Session } from "./session";
import { BASE_PROMPT } from "./prompt";
import { createTools } from "./tools";

export type StreamResult = StreamTextResult<any, any>;

export async function createStream(
  sid: string,
  sess: Session,
  keys: { openrouter?: string; google?: string; claude?: string },
): Promise<StreamResult> {
  const logPrefix = `[chat ${sid.slice(0, 8)}]`;

  let system = BASE_PROMPT;
  if (sess.currentFile) {
    system += `\n\nThe user has loaded "${sess.currentFile}". Use this exact filename in tool calls.`;
  }

  let model: any;
  let modelInfo: string;

  if (keys.openrouter) {
    const openrouter = createOpenRouter({ apiKey: keys.openrouter });
    const modelId =
      process.env.OPENROUTER_MODEL || "arcee-ai/trinity-large-preview:free";
    model = openrouter(modelId);
    modelInfo = `openrouter/${modelId}`;
  } else if (keys.google) {
    const google = createGoogleGenerativeAI({ apiKey: keys.google });
    model = google("gemini-2.5-flash");
    modelInfo = "google/gemini-2.5-flash";
  } else if (keys.claude) {
    const anthropic = createAnthropic({ apiKey: keys.claude });
    model = anthropic("claude-sonnet-4-6");
    modelInfo = "anthropic/claude-sonnet-4.6";
  } else {
    throw new Error("No API key provided");
  }

  console.log(`${logPrefix} using model: ${modelInfo}`);

  const tools = await createTools(sess.workspaceDir, logPrefix);

  return streamText({
    model,
    system,
    messages: sess.conversation.map((m) => ({
      role: m.role,
      content: m.content,
    })),
    tools,
    stopWhen: stepCountIs(100),
    onStepFinish({ toolCalls, toolResults, finishReason, usage }) {
      const calls = toolCalls
        .map((c) => {
          const input =
            "input" in c ? (c.input as Record<string, unknown>) : {};
          const summary =
            c.toolName === "bash"
              ? String(input.command ?? "").slice(0, 80)
              : c.toolName === "readFile"
                ? String(input.path ?? "").slice(0, 80)
                : c.toolName === "writeFile"
                  ? String(input.path ?? "").slice(0, 80)
                  : "";
          return `${c.toolName}(${summary})`;
        })
        .join(", ");
      const results = toolResults
        .map((r) => String((r as any).output ?? "").slice(0, 80))
        .join(" | ");
      console.log(
        `${logPrefix} step finish=${finishReason}`,
        calls ? `tools: ${calls}` : "",
        results ? `→ ${results}` : "",
        `tokens: ${usage?.totalTokens ?? "?"}`,
      );
    },
    onError({ error }) {
      console.error(`${logPrefix} stream error:`, error);
    },
  });
}
