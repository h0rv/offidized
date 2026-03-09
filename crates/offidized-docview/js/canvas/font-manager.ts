// Font manager for the CanvasKit renderer.
//
// Handles loading fonts and mapping DocViewModel font family names
// to CanvasKit typefaces with a fallback chain.

import { getCanvasKit } from "./canvaskit-loader.ts";

// We use `unknown` for CanvasKit types since we lazy-load the module
// and don't want a hard dependency on canvaskit-wasm types at compile time.
type Typeface = unknown;

export interface FontManifestEntry {
  /** Family name used when registering with CanvasKit FontCollection. */
  family: string;
  /** Optional Google Fonts CSS endpoint or equivalent stylesheet URL. */
  cssUrl?: string;
  /** Optional URL for a font file (ttf/otf/woff2 supported by browser fetch). */
  url?: string;
  /** Optional already-loaded bytes to avoid any network fetch. */
  data?: ArrayBuffer;
}

export type LocalFontManifestProvider = () =>
  | ReadonlyArray<FontManifestEntry>
  | Promise<ReadonlyArray<FontManifestEntry>>;

const BUNDLED_FONT_SOURCES = [
  {
    family: "Noto Sans",
    url: "https://fonts.gstatic.com/s/notosans/v36/o-0mIpQlx3QUlC5A4PNB6Ryti20_6n1iPHjcz6L1SoM-jCpoiyD9A-9a6Vc.ttf",
  },
  {
    family: "Noto Serif",
    cssUrl:
      "https://fonts.googleapis.com/css2?family=Noto+Serif:wght@400&display=swap",
  },
  {
    family: "Noto Sans Mono",
    cssUrl:
      "https://fonts.googleapis.com/css2?family=Noto+Sans+Mono:wght@400&display=swap",
  },
  {
    family: "Comic Neue",
    cssUrl:
      "https://fonts.googleapis.com/css2?family=Comic+Neue:wght@400&display=swap",
  },
] as const;

/** Cached typeface instances keyed by font family name. */
const typefaceCache = new Map<string, Typeface>();
const bundledFontData = new Map<string, ArrayBuffer>();
let defaultTypeface: Typeface | null = null;
let localFontManifestProvider: LocalFontManifestProvider | null = null;
let allowRemoteBundledFonts = true;

type FontBucket = "sans" | "serif" | "mono";

const BUCKET_TO_FALLBACK_FAMILY: Record<FontBucket, string> = {
  sans: "Noto Sans",
  serif: "Noto Serif",
  mono: "Noto Sans Mono",
};

const ALIAS_TO_BUCKET = new Map<string, FontBucket>([
  // Sans (Word + Google Docs common defaults)
  ["aptos", "sans"],
  ["calibri", "sans"],
  ["arial", "sans"],
  ["arialnarrow", "sans"],
  ["helvetica", "sans"],
  ["segoeui", "sans"],
  ["tahoma", "sans"],
  ["verdana", "sans"],
  ["trebuchetms", "sans"],
  ["candara", "sans"],
  ["corbel", "sans"],
  ["bahnschrift", "sans"],
  ["franklingothic", "sans"],
  ["gillsans", "sans"],
  ["roboto", "sans"],
  ["opensans", "sans"],
  ["lato", "sans"],
  ["montserrat", "sans"],
  ["nunito", "sans"],
  ["poppins", "sans"],
  ["sourcesanspro", "sans"],
  // Serif
  ["cambria", "serif"],
  ["timesnewroman", "serif"],
  ["times", "serif"],
  ["georgia", "serif"],
  ["garamond", "serif"],
  ["palatino", "serif"],
  ["bookantiqua", "serif"],
  ["baskerville", "serif"],
  ["constantia", "serif"],
  ["merriweather", "serif"],
  ["sourceserifpro", "serif"],
  // Mono
  ["consolas", "mono"],
  ["couriernew", "mono"],
  ["courier", "mono"],
  ["lucidaconsole", "mono"],
  ["monaco", "mono"],
  ["menlo", "mono"],
  ["robotomono", "mono"],
  ["sourcecodepro", "mono"],
  ["inconsolata", "mono"],
  ["ubuntumono", "mono"],
  ["droidsansmono", "mono"],
  ["comicneue", "sans"],
  ["comicsansms", "sans"],
]);

const GENERIC_TO_BUCKET = new Map<string, FontBucket>([
  ["serif", "serif"],
  ["sans-serif", "sans"],
  ["sans", "sans"],
  ["system-ui", "sans"],
  ["ui-sans-serif", "sans"],
  ["monospace", "mono"],
  ["ui-monospace", "mono"],
]);

/**
 * Optional config hook for environments that provide a local/offline font
 * manifest. Existing behavior is preserved by default:
 * - remote bundled fonts are still enabled unless explicitly disabled
 * - without a provider, fonts are fetched from the built-in URLs
 */
export function configureFontManifestFallback(opts: {
  localManifestProvider?: LocalFontManifestProvider | null;
  allowRemoteFallback?: boolean;
}): void {
  if (opts.localManifestProvider !== undefined) {
    localFontManifestProvider = opts.localManifestProvider;
  }
  if (opts.allowRemoteFallback !== undefined) {
    allowRemoteBundledFonts = opts.allowRemoteFallback;
  }
}

function normalizeFamilyToken(token: string): string {
  const trimmed = token.trim().replace(/^["']+|["']+$/g, "");
  return trimmed.toLowerCase();
}

function findLoadedFamily(token: string): string | null {
  for (const family of bundledFontData.keys()) {
    if (family.toLowerCase() === token) {
      return family;
    }
  }
  return null;
}

function compactFamilyToken(token: string): string {
  return token.replace(/[^a-z0-9]/g, "");
}

function getBucketFromToken(token: string): FontBucket | null {
  const genericBucket = GENERIC_TO_BUCKET.get(token);
  if (genericBucket) return genericBucket;

  const aliasBucket = ALIAS_TO_BUCKET.get(compactFamilyToken(token));
  if (aliasBucket) return aliasBucket;

  if (
    token.includes("mono") ||
    token.includes("code") ||
    token.includes("typewriter")
  ) {
    return "mono";
  }
  if (
    token.includes("serif") ||
    token.includes("times") ||
    token.includes("garamond") ||
    token.includes("cambria")
  ) {
    return "serif";
  }
  return null;
}

async function loadTypefaceSource(
  ck: { Typeface: { MakeFreeTypeFaceFromData(data: ArrayBuffer): Typeface } },
  source: FontManifestEntry,
): Promise<ArrayBuffer> {
  let data = source.data;
  if (!data && !source.url && !source.cssUrl) {
    throw new Error(
      `Font source '${source.family}' is missing 'data', 'url', and 'cssUrl'`,
    );
  }
  if (!data) {
    const url =
      source.url ?? (await resolveCssFontUrl(source.cssUrl as string));
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(
        `Failed to load font '${source.family}' from '${url}': HTTP ${response.status}`,
      );
    }
    data = await response.arrayBuffer();
  }
  const tf = ck.Typeface.MakeFreeTypeFaceFromData(data);
  if (tf) {
    typefaceCache.set(source.family, tf);
    if (!defaultTypeface) defaultTypeface = tf;
  }
  return data;
}

async function resolveCssFontUrl(cssUrl: string): Promise<string> {
  const response = await fetch(cssUrl);
  if (!response.ok) {
    throw new Error(
      `Failed to load font stylesheet '${cssUrl}': HTTP ${response.status}`,
    );
  }
  const cssText = await response.text();
  const match = cssText.match(/url\(([^)]+)\)/i);
  if (!match?.[1]) {
    throw new Error(`No font URL found in stylesheet '${cssUrl}'`);
  }
  return match[1].trim().replace(/^['"]|['"]$/g, "");
}

async function getLocalManifest(): Promise<ReadonlyArray<FontManifestEntry>> {
  if (!localFontManifestProvider) return [];
  const localManifest = await localFontManifestProvider();
  const deduped = new Map<string, FontManifestEntry>();
  for (const source of localManifest) {
    const family = source.family.trim();
    if (!family) continue;
    if (!source.data && !source.url) continue;
    if (!source.data && !source.url && !source.cssUrl) continue;
    deduped.set(family.toLowerCase(), {
      family,
      data: source.data,
      url: source.url,
      cssUrl: source.cssUrl,
    });
  }
  return Array.from(deduped.values());
}

/**
 * Load bundled fallback fonts. Called once during renderer initialization.
 */
export async function loadBundledFonts(): Promise<
  Array<{ family: string; data: ArrayBuffer }>
> {
  if (bundledFontData.size > 0) {
    return Array.from(bundledFontData.entries()).map(([family, data]) => ({
      family,
      data,
    }));
  }

  const ck = (await getCanvasKit()) as {
    Typeface: { MakeFreeTypeFaceFromData(data: ArrayBuffer): Typeface };
  };

  // Prefer caller-provided local/offline manifest when configured.
  const localManifest = await getLocalManifest();
  for (const source of localManifest) {
    try {
      const data = await loadTypefaceSource(ck, source);
      bundledFontData.set(source.family, data);
    } catch {
      // Local sources are best-effort; remote fallback can still fill gaps.
    }
  }

  if (allowRemoteBundledFonts) {
    for (const source of BUNDLED_FONT_SOURCES) {
      if (bundledFontData.has(source.family)) continue;
      try {
        const data = await loadTypefaceSource(ck, source);
        bundledFontData.set(source.family, data);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        console.warn(
          `Skipping unavailable bundled font '${source.family}': ${message}`,
        );
      }
    }
  }

  if (!bundledFontData.has("Noto Sans")) {
    const localNames = localManifest.map((s) => s.family).join(", ");
    throw new Error(
      allowRemoteBundledFonts
        ? "Failed to initialize default bundled font 'Noto Sans'."
        : `No usable default font loaded. Local manifest families: [${localNames}]`,
    );
  }
  if (!defaultTypeface) {
    const sansData = bundledFontData.get("Noto Sans");
    if (sansData) {
      const tf = ck.Typeface.MakeFreeTypeFaceFromData(sansData);
      if (tf) {
        typefaceCache.set("Noto Sans", tf);
        defaultTypeface = tf;
      }
    }
  }

  return Array.from(bundledFontData.entries()).map(([family, data]) => ({
    family,
    data,
  }));
}

/**
 * Get a typeface for the given font family name.
 * Returns the default typeface if the requested family isn't loaded.
 */
export function getTypeface(family: string): Typeface | null {
  return typefaceCache.get(family) ?? defaultTypeface;
}

/**
 * Get the default font data buffer (for FontCollection registration).
 */
export function getBundledFontData(family = "Noto Sans"): ArrayBuffer | null {
  return bundledFontData.get(family) ?? null;
}

/**
 * Map a DocViewModel font family name to the family names
 * that CanvasKit's FontCollection knows about.
 *
 * Uses a deterministic mapping so pagination doesn't depend on browser
 * fallback behavior.
 */
export function resolveFontFamily(_family?: string): string {
  const family = (_family ?? "").trim();
  if (!family) return BUCKET_TO_FALLBACK_FAMILY.sans;

  // Try family list in order (e.g. `Calibri, Arial, sans-serif`).
  const tokens = family
    .split(",")
    .map((token) => normalizeFamilyToken(token))
    .filter((token) => token.length > 0);

  for (const token of tokens) {
    const loaded = findLoadedFamily(token);
    if (loaded) return loaded;

    const bucket = getBucketFromToken(token);
    if (!bucket) continue;
    const resolved = BUCKET_TO_FALLBACK_FAMILY[bucket];
    if (bundledFontData.has(resolved)) return resolved;
    return BUCKET_TO_FALLBACK_FAMILY.sans;
  }

  return BUCKET_TO_FALLBACK_FAMILY.sans;
}
