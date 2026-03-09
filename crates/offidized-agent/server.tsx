import "./lib/env";
import { Elysia } from "elysia";
import { registerRoutes } from "./lib/routes";
import { previewReady } from "./lib/preview";

const app = new Elysia();
registerRoutes(app);
app.listen(Number(process.env.PORT ?? 3000));

console.log(`http://localhost:${process.env.PORT ?? 3000}`);
if (!previewReady)
  console.log("  (preview disabled — viewer assets not built)");
