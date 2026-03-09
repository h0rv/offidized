import { resolve } from "path";
import { createBashTool } from "bash-tool";
import type { BashToolkit, Sandbox } from "bash-tool";

export async function createTools(
  workspaceDir: string,
  logPrefix: string,
): Promise<BashToolkit["tools"]> {
  const ofxPath = resolve(import.meta.dir, "../../../target/release");

  const sandbox: Sandbox = {
    executeCommand: async (command: string) => {
      const proc = Bun.spawn(["sh", "-c", command], {
        cwd: workspaceDir,
        stdout: "pipe",
        stderr: "pipe",
        env: {
          ...process.env,
          PATH: `${ofxPath}:${process.env.PATH ?? ""}`,
        },
      });
      const stdout = await new Response(proc.stdout).text();
      const stderr = await new Response(proc.stderr).text();
      await proc.exited;
      return { stdout, stderr, exitCode: proc.exitCode ?? 1 };
    },
    readFile: async (path: string) => {
      return Bun.file(path).text();
    },
    writeFiles: async (
      files: Array<{ path: string; content: string | Buffer }>,
    ) => {
      for (const { path, content } of files) {
        await Bun.write(path, content);
      }
    },
  };

  const { tools } = await createBashTool({
    destination: workspaceDir,
    sandbox,
    onBeforeBashCall: ({ command }) => {
      console.log(`${logPrefix} tool:bash $ ${command.slice(0, 100)}`);
      return undefined;
    },
  });
  return tools;
}
