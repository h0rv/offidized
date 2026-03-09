import * as esbuild from 'esbuild';
import { copyFileSync } from 'fs';

const watch = process.argv.includes('--watch');

/** @type {esbuild.BuildOptions} */
const shared = {
  bundle: false,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  sourcemap: true,
};

const entryPoints = [
  // Demo viewer + editor
  { in: 'src/demo/viewer.ts', out: 'src/demo/viewer' },
  { in: 'src/demo/editor.ts', out: 'src/demo/editor' },
  // Bench
  { in: 'bench/bench.ts', out: 'bench/bench' },
  { in: 'bench/providers/index.ts', out: 'bench/providers/index' },
  { in: 'bench/providers/xlview.ts', out: 'bench/providers/xlview' },
  { in: 'bench/providers/sheetjs.ts', out: 'bench/providers/sheetjs' },
  // Browser tests
  { in: 'tests/browser/run_scroll_test.ts', out: 'tests/browser/run_scroll_test' },
  // Public API wrappers
  { in: 'js/xl-view.ts', out: 'pkg/xl-view' },
  { in: 'js/xl-edit.ts', out: 'pkg/xl-edit' },
];

async function build() {
  for (const entry of entryPoints) {
    const opts = {
      ...shared,
      entryPoints: [entry.in],
      outfile: entry.out + '.js',
    };

    // Node scripts need node platform
    if (
      entry.in === 'tests/browser/run_scroll_test.ts'
    ) {
      opts.platform = 'node';
    }

    if (watch) {
      const ctx = await esbuild.context(opts);
      await ctx.watch();
      console.log(`Watching ${entry.in}...`);
    } else {
      await esbuild.build(opts);
      console.log(`Built ${entry.out}.js`);
    }
  }

  // Copy type declarations that esbuild doesn't handle
  copyFileSync('js/xl-view.d.ts', 'pkg/xl-view.d.ts');
  console.log('Copied pkg/xl-view.d.ts');
  copyFileSync('js/xl-edit.d.ts', 'pkg/xl-edit.d.ts');
  console.log('Copied pkg/xl-edit.d.ts');

  if (!watch) {
    console.log('\nAll TypeScript files compiled successfully.');
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
