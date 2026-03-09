export interface EditorPerfSample {
  op: string;
  applyMs: number;
  viewModelMs: number;
  renderMs: number;
  totalMs: number;
  atMs: number;
}

export interface EditorPerfSummary {
  samples: number;
  avgTotalMs: number;
  p95TotalMs: number;
  maxTotalMs: number;
  avgApplyMs: number;
  avgViewModelMs: number;
  avgRenderMs: number;
}

export class EditorPerfTracker {
  private readonly maxSamples: number;
  private readonly samples: EditorPerfSample[] = [];

  constructor(maxSamples: number = 120) {
    this.maxSamples = Math.max(10, maxSamples);
  }

  record(sample: EditorPerfSample): EditorPerfSummary {
    this.samples.push(sample);
    if (this.samples.length > this.maxSamples) {
      this.samples.splice(0, this.samples.length - this.maxSamples);
    }
    return this.summary();
  }

  summary(): EditorPerfSummary {
    if (this.samples.length === 0) {
      return {
        samples: 0,
        avgTotalMs: 0,
        p95TotalMs: 0,
        maxTotalMs: 0,
        avgApplyMs: 0,
        avgViewModelMs: 0,
        avgRenderMs: 0,
      };
    }

    let total = 0;
    let apply = 0;
    let viewModel = 0;
    let render = 0;
    let max = 0;
    const totals: number[] = [];

    for (const sample of this.samples) {
      total += sample.totalMs;
      apply += sample.applyMs;
      viewModel += sample.viewModelMs;
      render += sample.renderMs;
      max = Math.max(max, sample.totalMs);
      totals.push(sample.totalMs);
    }

    totals.sort((a, b) => a - b);
    const p95Idx = Math.min(
      totals.length - 1,
      Math.max(0, Math.floor(totals.length * 0.95) - 1),
    );
    const count = this.samples.length;

    return {
      samples: count,
      avgTotalMs: total / count,
      p95TotalMs: totals[p95Idx] ?? 0,
      maxTotalMs: max,
      avgApplyMs: apply / count,
      avgViewModelMs: viewModel / count,
      avgRenderMs: render / count,
    };
  }
}
