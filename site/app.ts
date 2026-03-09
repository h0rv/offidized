import {
  init as initXlView,
  mount as mountXlView,
} from "../crates/offidized-xlview/js/xl-view.ts";
import {
  init as initXlEdit,
  mountEditor as mountXlEdit,
} from "../crates/offidized-xlview/js/xl-edit.ts";
import {
  init as initDocView,
  mount as mountDocView,
} from "../crates/offidized-docview/js/doc-view.ts";
import {
  init as initDocEdit,
  mountEditor as mountDocEdit,
} from "../crates/offidized-docview/js/doc-edit.ts";
import {
  init as initPptView,
  mount as mountPptView,
} from "../crates/offidized-pptview/js/ppt-view.ts";
import initCoreWasm, {
  Workbook as WasmWorkbook,
} from "../crates/offidized-wasm/pkg/offidized_wasm.js";

type DemoKind = "xlsx" | "docx" | "pptx";
type DemoMode = "view" | "edit";
type DocRendererKind = "canvas" | "html";

type Sample = {
  label: string;
  filename: string;
  url: string;
};

type DemoController = {
  load(data: Uint8Array): number | void;
  destroy(): void;
  save?: () => Uint8Array;
  download?: (filename?: string) => void;
  newBlank?: () => Uint8Array | void;
};

const assetUrl = (path: string) =>
  new URL(path, window.location.href).toString();

const samples: Record<DemoKind, Record<string, Sample>> = {
  xlsx: {
    portfolio: {
      label: "Portfolio",
      filename: "meridian_portfolio.xlsx",
      url: assetUrl("./assets/samples/meridian_portfolio.xlsx"),
    },
  },
  docx: {
    contract: {
      label: "Contract",
      filename: "07_contract.docx",
      url: assetUrl("./assets/samples/07_contract.docx"),
    },
    tables: {
      label: "Tables",
      filename: "02_tables.docx",
      url: assetUrl("./assets/samples/02_tables.docx"),
    },
  },
  pptx: {
    pitch: {
      label: "Pitch deck",
      filename: "07_pitch_deck.pptx",
      url: assetUrl("./assets/samples/07_pitch_deck.pptx"),
    },
    richtext: {
      label: "Rich text",
      filename: "04_rich_text_formatting.pptx",
      url: assetUrl("./assets/samples/04_rich_text_formatting.pptx"),
    },
  },
};

let coreInitPromise: Promise<void> | null = null;

async function initCore(): Promise<void> {
  if (!coreInitPromise) {
    coreInitPromise = initCoreWasm(
      assetUrl("./assets/wasm/offidized_wasm_bg.wasm"),
    ).then(() => {});
  }
  return coreInitPromise;
}

async function createBlankSpreadsheet(): Promise<Uint8Array> {
  await initCore();
  const workbook = new WasmWorkbook();
  workbook.addSheet("Sheet1");
  return workbook.toBytes();
}

async function mountSpreadsheet(
  container: HTMLElement,
  mode: DemoMode,
): Promise<DemoController> {
  const wasmUrl = assetUrl("./assets/wasm/offidized_xlview_bg.wasm");
  if (mode === "edit") {
    await initXlEdit(wasmUrl);
    return await mountXlEdit(container);
  }
  await initXlView(wasmUrl);
  return await mountXlView(container);
}

async function mountPresentation(
  container: HTMLElement,
): Promise<DemoController> {
  await initPptView(assetUrl("./assets/wasm/offidized_pptview_bg.wasm"));
  return await mountPptView(container);
}

function downloadBytes(
  bytes: Uint8Array,
  filename: string,
  mimeType: string,
): void {
  const blob = new Blob([bytes], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

async function readBytes(file: File): Promise<Uint8Array> {
  return new Uint8Array(await file.arrayBuffer());
}

async function fetchBytes(url: string): Promise<Uint8Array> {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}

function mimeTypeFor(kind: DemoKind): string {
  if (kind === "xlsx") {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (kind === "docx") {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  }
  return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

class Panel {
  private readonly kind: DemoKind;
  private readonly stage: HTMLElement;
  private readonly status: HTMLElement;
  private readonly fileInput: HTMLInputElement;
  private readonly newButton: HTMLButtonElement | null;
  private readonly saveButton: HTMLButtonElement | null;
  private readonly sampleButtons: HTMLButtonElement[];
  private readonly modeButtons: HTMLButtonElement[];
  private readonly rendererButtons: HTMLButtonElement[];
  private mode: DemoMode = "view";
  private renderer: DocRendererKind = "canvas";
  private controller: DemoController | null = null;
  private currentData: Uint8Array | null = null;
  private currentFilename = "";
  private initialized = false;

  constructor(root: HTMLElement) {
    this.kind = root.dataset.kind as DemoKind;
    this.stage = root.querySelector(".stage") as HTMLElement;
    this.status = root.querySelector(".panel-status") as HTMLElement;
    this.fileInput = root.querySelector(".file-input") as HTMLInputElement;
    this.newButton = root.querySelector(".new-button");
    this.saveButton = root.querySelector(".save-button");
    this.sampleButtons = Array.from(root.querySelectorAll(".sample-chip"));
    this.modeButtons = Array.from(
      root.querySelectorAll(".mode-chip[data-mode]"),
    );
    this.rendererButtons = Array.from(
      root.querySelectorAll(".mode-chip[data-renderer]"),
    );

    this.fileInput.addEventListener("change", async () => {
      const file = this.fileInput.files?.[0];
      if (!file) return;
      await this.loadBytes(await readBytes(file), file.name);
    });

    this.newButton?.addEventListener("click", async () => {
      await this.createBlank();
    });

    for (const button of this.sampleButtons) {
      button.addEventListener("click", async () => {
        const key = button.dataset.sample!;
        this.setActiveSample(key);
        await this.loadSample(samples[this.kind][key]);
      });
    }

    for (const button of this.modeButtons) {
      button.addEventListener("click", async () => {
        const nextMode = button.dataset.mode as DemoMode;
        if (nextMode === this.mode) return;
        await this.setMode(nextMode);
      });
    }

    for (const button of this.rendererButtons) {
      button.addEventListener("click", async () => {
        const nextRenderer = button.dataset.renderer as DocRendererKind;
        if (nextRenderer === this.renderer) return;
        this.snapshotCurrentState();
        this.renderer = nextRenderer;
        for (const item of this.rendererButtons) {
          item.classList.toggle(
            "active",
            item.dataset.renderer === nextRenderer,
          );
        }
        await this.mountController();
        if (this.currentData) {
          await this.loadBytes(this.currentData, this.currentFilename, false);
        }
      });
    }

    this.saveButton?.addEventListener("click", () => {
      if (!this.controller || this.kind === "pptx") return;
      if (this.controller.save) {
        this.currentData = this.controller.save();
      }
      if (this.kind === "xlsx" && this.controller.download) {
        this.controller.download(this.currentFilename || "edited.xlsx");
        return;
      }
      if (this.currentData) {
        downloadBytes(
          this.currentData,
          this.currentFilename || `edited.${this.kind}`,
          mimeTypeFor(this.kind),
        );
      }
    });
  }

  async init(): Promise<void> {
    if (this.initialized) return;
    this.initialized = true;
    const defaultSampleKey = this.sampleButtons[0]?.dataset.sample;
    if (!defaultSampleKey) return;
    this.setActiveSample(defaultSampleKey);
    await this.loadSample(samples[this.kind][defaultSampleKey]);
  }

  private setActiveSample(sampleKey: string): void {
    for (const button of this.sampleButtons) {
      button.classList.toggle("active", button.dataset.sample === sampleKey);
    }
  }

  private snapshotCurrentState(): void {
    if (!this.controller?.save) return;
    try {
      this.currentData = this.controller.save();
    } catch {
      // Fall back to the last successful snapshot.
    }
  }

  private async setMode(nextMode: DemoMode): Promise<void> {
    this.snapshotCurrentState();
    this.mode = nextMode;
    for (const button of this.modeButtons) {
      button.classList.toggle("active", button.dataset.mode === nextMode);
    }
    if (this.newButton) {
      this.newButton.hidden = nextMode !== "edit";
    }
    if (this.saveButton) {
      this.saveButton.hidden = nextMode !== "edit";
    }
    await this.mountController();
    if (this.currentData) {
      await this.loadBytes(this.currentData, this.currentFilename, false);
    }
  }

  private async mountController(): Promise<void> {
    this.controller?.destroy();
    this.controller = null;
    this.stage.innerHTML = "";

    if (this.kind === "docx") {
      this.controller =
        this.mode === "edit"
          ? await this.mountDocEditor()
          : await this.mountDocViewer();
      return;
    }

    if (this.kind === "xlsx") {
      this.controller = await mountSpreadsheet(this.stage, this.mode);
      return;
    }

    this.controller = await mountPresentation(this.stage);
  }

  private async mountDocViewer(): Promise<DemoController> {
    await initDocView(assetUrl("./assets/wasm/offidized_docview_bg.wasm"));
    const viewer = await mountDocView(this.stage, { renderer: this.renderer });
    return {
      load(data: Uint8Array): number {
        return viewer.load(data);
      },
      destroy(): void {
        viewer.destroy();
      },
    };
  }

  private async mountDocEditor(): Promise<DemoController> {
    await initDocEdit(assetUrl("./assets/wasm/offidized_docview_bg.wasm"));
    const editor = await mountDocEdit(this.stage, { renderer: this.renderer });
    return {
      load(data: Uint8Array): number {
        return editor.load(data);
      },
      destroy(): void {
        editor.destroy();
      },
      save(): Uint8Array {
        return editor.save();
      },
      newBlank(): Uint8Array {
        editor.loadBlank();
        return editor.save();
      },
    };
  }

  private async createBlank(): Promise<void> {
    if (this.mode !== "edit") return;

    try {
      if (!this.controller) {
        await this.mountController();
      }
      if (!this.controller) return;

      this.setActiveSample("__blank__");

      if (this.kind === "docx") {
        const bytes = this.controller.newBlank?.();
        if (bytes instanceof Uint8Array) {
          this.currentData = bytes;
        }
        this.currentFilename = "untitled.docx";
        this.status.textContent = "untitled.docx · edit mode";
        return;
      }

      if (this.kind === "xlsx") {
        const bytes = await createBlankSpreadsheet();
        await this.loadBytes(bytes, "untitled.xlsx");
      }
    } catch (error) {
      const filename = this.kind === "docx" ? "untitled.docx" : "untitled.xlsx";
      this.status.textContent = `Could not create ${filename}: ${(error as Error).message}`;
    }
  }

  private async loadSample(sample: Sample): Promise<void> {
    this.status.textContent = `Loading ${sample.label.toLowerCase()}…`;
    const bytes = await fetchBytes(sample.url);
    await this.loadBytes(bytes, sample.filename, false);
  }

  private async loadBytes(
    bytes: Uint8Array,
    filename: string,
    cacheData = true,
  ): Promise<void> {
    try {
      if (!this.controller) {
        await this.mountController();
      }
      if (!this.controller) return;

      this.currentData = cacheData ? bytes : new Uint8Array(bytes);
      this.currentFilename = filename;

      const elapsed = this.controller.load(bytes);
      const modeLabel =
        this.kind === "pptx"
          ? "viewer"
          : `${this.mode === "edit" ? "edit" : "view"} mode`;
      this.status.textContent = `${filename} · ${modeLabel}${typeof elapsed === "number" ? ` · ${elapsed}ms` : ""}`;
    } catch (error) {
      this.status.textContent = `Could not load ${filename}: ${(error as Error).message}`;
    }
  }
}

const panelElements = Array.from(
  document.querySelectorAll<HTMLElement>(".demo-panel"),
);
const panels = panelElements.map((panel) => new Panel(panel));
const panelsByKind = new Map<DemoKind, Panel>();
const tabButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>(".file-tab"),
);

for (let index = 0; index < panels.length; index += 1) {
  const kind = panelElements[index]?.dataset.kind as DemoKind | undefined;
  if (kind) {
    panelsByKind.set(kind, panels[index]);
  }
}

async function setActiveKind(kind: DemoKind): Promise<void> {
  for (const tab of tabButtons) {
    tab.classList.toggle("active", tab.dataset.target === kind);
  }
  for (const panel of panelElements) {
    panel.classList.toggle("active", panel.dataset.kind === kind);
  }
  await panelsByKind.get(kind)?.init();
}

for (const tab of tabButtons) {
  tab.addEventListener("click", () => {
    void setActiveKind(tab.dataset.target as DemoKind);
  });
}

await setActiveKind("xlsx");
