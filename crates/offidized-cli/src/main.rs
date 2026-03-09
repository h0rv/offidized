mod cmd;
mod format;
mod output;
mod range_parse;

use std::path::PathBuf;

use clap::{Parser, Subcommand};

/// CLI for reading, writing, and manipulating OOXML files.
#[derive(Parser)]
#[command(name = "offidized", version, about)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Show file metadata (sheets, paragraph count, etc.).
    Info {
        /// Path to the OOXML file.
        file: PathBuf,
    },

    /// Read cell values or paragraphs.
    Read {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Cell range for xlsx, e.g. "Sheet1!A1:D10".
        range: Option<String>,
        /// Output format: json (default) or csv.
        #[arg(long, default_value = "json")]
        format: String,
        /// Paragraph index range for docx, e.g. "0-4".
        #[arg(long)]
        paragraphs: Option<String>,
    },

    /// Set a single cell value or paragraph text.
    Set {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Target: "Sheet1!A1" for xlsx, paragraph index for docx.
        target: String,
        /// The value to set.
        value: String,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// Bulk edit via stdin JSON.
    Patch {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// Find and replace text in all cells or runs.
    Replace {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Text to find.
        find: String,
        /// Replacement text.
        replace_with: String,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// Access raw OPC package parts.
    Part {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Part URI to extract (e.g. "/xl/workbook.xml").
        uri: Option<String>,
        /// List all parts as JSON.
        #[arg(long)]
        list: bool,
    },

    /// Create an empty OOXML file.
    Create {
        /// Path for the new file.
        file: PathBuf,
        /// Sheet names to pre-create (xlsx only; default: Sheet1).
        ///
        /// Example: ofx create report.xlsx Dashboard Data "Monthly Trend"
        #[arg(trailing_var_arg = true)]
        sheets: Vec<String>,
    },

    /// Copy a cell range to another location.
    CopyRange {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Source range, e.g. "Sheet1!A1:B10".
        source_range: String,
        /// Destination cell, e.g. "Sheet1!D1".
        dest_cell: String,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// Move a cell range to another location.
    MoveRange {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Source range, e.g. "Sheet1!A1:B10".
        source_range: String,
        /// Destination cell, e.g. "Sheet1!D1".
        dest_cell: String,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// List pivot tables with details.
    Pivots {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Optional sheet name to filter pivot tables.
        #[arg(long)]
        sheet: Option<String>,
    },

    /// List charts with details.
    Charts {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Optional sheet name to filter charts.
        #[arg(long)]
        sheet: Option<String>,
    },

    /// Add a chart to a worksheet.
    AddChart {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Sheet name to add the chart to.
        #[arg(long)]
        sheet: String,
        /// Chart type: bar, line, pie, area, scatter, doughnut, radar.
        #[arg(long = "type")]
        chart_type: String,
        /// Chart title.
        #[arg(long)]
        title: Option<String>,
        /// Series spec: "Name | [cats_range] | vals_range" or "Name | vals_range".
        /// Repeat for multiple series.
        #[arg(long = "series", action = clap::ArgAction::Append)]
        series: Vec<String>,
        /// Two-cell anchor: "from_col,from_row,to_col,to_row" (zero-based).
        #[arg(long, default_value = "0,0,9,14")]
        anchor: String,
        /// Bar direction: col (vertical) or bar (horizontal).
        #[arg(long = "bar-direction")]
        bar_direction: Option<String>,
        /// Grouping: clustered, stacked, percent-stacked, standard.
        #[arg(long)]
        grouping: Option<String>,
        /// Legend position: b, t, l, r, tr.
        #[arg(long = "legend-pos")]
        legend_pos: Option<String>,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
    },

    /// Evaluate a formula.
    Eval {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Formula to evaluate, e.g. "=SUM(A1:A10)".
        formula: String,
        /// Sheet name (defaults to first sheet).
        #[arg(long)]
        sheet: Option<String>,
    },

    /// Lint workbook structure and formulas.
    Lint {
        /// Path to the OOXML file.
        file: PathBuf,
    },

    /// List unified editable content nodes for any supported Office format.
    Nodes {
        /// Path to the OOXML file.
        file: PathBuf,
    },

    /// Show unified edit capabilities for a file.
    Capabilities {
        /// Path to the OOXML file.
        file: PathBuf,
    },

    /// Apply unified content edits across xlsx/docx/pptx.
    Edit {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Edit spec in the form `<id>=<text>`. Repeat for multiple edits.
        #[arg(long = "edit", action = clap::ArgAction::Append)]
        edits: Vec<String>,
        /// Path to JSON array of edits. Supports optional `file`, `group`, and typed `payload`.
        #[arg(long = "edits-json")]
        edits_json: Option<PathBuf>,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite input file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
        /// Skip checksum validation.
        #[arg(long)]
        force: bool,
        /// Fail if any apply diagnostics or skipped edits remain (lint diagnostics excluded).
        #[arg(long)]
        strict: bool,
        /// Run unified edit lint checks before apply (missing targets, ambiguous anchors, table coords).
        #[arg(long)]
        lint: bool,
    },

    /// Derive a text IR from an Office file.
    Derive {
        /// Path to the OOXML file.
        file: PathBuf,
        /// Output file path (writes to stdout if not provided).
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// IR mode: content, style, or full.
        #[arg(long, default_value = "content")]
        mode: String,
        /// Single sheet name (xlsx only).
        #[arg(long)]
        sheet: Option<String>,
    },

    /// Apply changes from an IR file to an Office document.
    Apply {
        /// Path to the IR file (reads from stdin if not provided).
        file: Option<PathBuf>,
        /// Output file path.
        #[arg(short = 'o', long = "output")]
        output: Option<PathBuf>,
        /// Overwrite source file in place.
        #[arg(short = 'i', long = "in-place")]
        in_place: bool,
        /// Override source file path (default: from IR header).
        #[arg(long = "source")]
        source: Option<PathBuf>,
        /// Skip checksum validation.
        #[arg(long)]
        force: bool,
        /// Show changes without saving.
        #[arg(long)]
        dry_run: bool,
    },
}

fn main() {
    let cli = Cli::parse();
    if let Err(e) = run(cli) {
        eprintln!("error: {e:#}");
        std::process::exit(1);
    }
}

fn run(cli: Cli) -> anyhow::Result<()> {
    match cli.command {
        Command::Info { file } => cmd::info::run(&file),
        Command::Read {
            file,
            range,
            format,
            paragraphs,
        } => cmd::read::run(&file, range.as_deref(), &format, paragraphs.as_deref()),
        Command::Set {
            file,
            target,
            value,
            output,
            in_place,
        } => cmd::set::run(&file, &target, &value, output.as_deref(), in_place),
        Command::Patch {
            file,
            output,
            in_place,
        } => cmd::patch::run(&file, output.as_deref(), in_place),
        Command::Replace {
            file,
            find,
            replace_with,
            output,
            in_place,
        } => cmd::replace::run(&file, &find, &replace_with, output.as_deref(), in_place),
        Command::Part { file, uri, list } => cmd::part::run(&file, uri.as_deref(), list),
        Command::Create { file, sheets } => cmd::create::run(&file, &sheets),
        Command::CopyRange {
            file,
            source_range,
            dest_cell,
            output,
            in_place,
        } => cmd::copy_range::run(
            &file,
            &source_range,
            &dest_cell,
            output.as_deref(),
            in_place,
        ),
        Command::MoveRange {
            file,
            source_range,
            dest_cell,
            output,
            in_place,
        } => cmd::move_range::run(
            &file,
            &source_range,
            &dest_cell,
            output.as_deref(),
            in_place,
        ),
        Command::Pivots { file, sheet } => cmd::pivots::run(&file, sheet.as_deref()),
        Command::Charts { file, sheet } => cmd::charts::run(&file, sheet.as_deref()),
        Command::AddChart {
            file,
            sheet,
            chart_type,
            title,
            series,
            anchor,
            bar_direction,
            grouping,
            legend_pos,
            output,
            in_place,
        } => cmd::add_chart::run(
            &file,
            &sheet,
            &chart_type,
            title.as_deref(),
            &series,
            &anchor,
            bar_direction.as_deref(),
            grouping.as_deref(),
            legend_pos.as_deref(),
            output.as_deref(),
            in_place,
        ),
        Command::Eval {
            file,
            formula,
            sheet,
        } => cmd::eval::run(&file, &formula, sheet.as_deref()),
        Command::Lint { file } => cmd::lint::run(&file),
        Command::Nodes { file } => cmd::nodes::run(&file),
        Command::Capabilities { file } => cmd::capabilities::run(&file),
        Command::Edit {
            file,
            edits,
            edits_json,
            output,
            in_place,
            force,
            strict,
            lint,
        } => cmd::edit::run(
            &file,
            &edits,
            edits_json.as_deref(),
            output.as_deref(),
            in_place,
            force,
            strict,
            lint,
        ),
        Command::Derive {
            file,
            output,
            mode,
            sheet,
        } => cmd::derive::run(&file, output.as_deref(), &mode, sheet.as_deref()),
        Command::Apply {
            file,
            output,
            in_place,
            source,
            force,
            dry_run,
        } => cmd::apply::run(
            file.as_deref(),
            output.as_deref(),
            in_place,
            source.as_deref(),
            force,
            dry_run,
        ),
    }
}
