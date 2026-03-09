import js from "@eslint/js";
import ts from "typescript-eslint";

export default ts.config(
  { ignores: ["pkg/", "node_modules/"] },
  js.configs.recommended,
  ...ts.configs.recommended,
  {
    languageOptions: {
      globals: {
        document: "readonly",
        HTMLElement: "readonly",
        HTMLDivElement: "readonly",
        HTMLInputElement: "readonly",
        HTMLButtonElement: "readonly",
        HTMLSpanElement: "readonly",
        HTMLTableElement: "readonly",
        DocumentFragment: "readonly",
        ShadowRoot: "readonly",
        CustomEvent: "readonly",
        Event: "readonly",
        File: "readonly",
        Uint8Array: "readonly",
        performance: "readonly",
        fetch: "readonly",
        console: "readonly",
        customElements: "readonly",
      },
    },
    rules: {
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-unused-vars": [
        "warn",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
    },
  },
);
