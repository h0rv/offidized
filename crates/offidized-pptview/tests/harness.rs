#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
//! Test harness for offidized-pptview.
//!
//! - `diagnose_*`: Load real pptx files and report conversion quality.
//! - `render_*`: Generate standalone HTML files for visual inspection.
//! - `regression_*`: Snapshot-based regression tests.
//!
//! Run with: cargo test -p offidized-pptview --test harness -- --nocapture

use offidized_pptview::convert::convert_presentation;
use offidized_pptview::model::PresentationViewModel;
use offidized_pptx::Presentation;
use std::path::{Path, PathBuf};

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

fn project_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .map(PathBuf::from)
        .expect("could not resolve project root")
}

fn shapecrawler_dir() -> Option<PathBuf> {
    let d = project_root().join("references/ShapeCrawler/tests/ShapeCrawler.DevTests/Assets");
    d.is_dir().then_some(d)
}

fn openxml_dir() -> Option<PathBuf> {
    let d = project_root()
        .join("references/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles");
    d.is_dir().then_some(d)
}

fn output_dir() -> PathBuf {
    let d = project_root().join("artifacts/pptview-harness");
    std::fs::create_dir_all(&d).ok();
    d
}

fn collect_pptx(dir: &Path) -> Vec<PathBuf> {
    let mut result = Vec::new();
    walk(dir, &mut result);
    result.sort();
    result
}

fn walk(dir: &Path, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = std::fs::read_dir(dir) {
        for entry in entries.flatten() {
            let p = entry.path();
            if p.is_dir() {
                walk(&p, out);
            } else if p.extension().is_some_and(|e| e == "pptx") {
                out.push(p);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Diagnostic: summarize conversion quality
// ---------------------------------------------------------------------------

struct DiagResult {
    file: String,
    slides: usize,
    raw_shapes: usize,
    geom_ok: usize,
    #[allow(dead_code)]
    geom_missing: usize,
    converted_shapes: usize,
    shapes_with_text: usize,
    open_error: bool,
}

fn diagnose(path: &Path) -> DiagResult {
    let file = path
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();

    let pres = match Presentation::open(path) {
        Ok(p) => p,
        Err(_) => {
            return DiagResult {
                file,
                slides: 0,
                raw_shapes: 0,
                geom_ok: 0,
                geom_missing: 0,
                converted_shapes: 0,
                shapes_with_text: 0,
                open_error: true,
            };
        }
    };

    let slides = pres.slides().len();
    let raw_shapes: usize = pres.slides().iter().map(|s| s.shapes().len()).sum();
    let geom_ok: usize = pres
        .slides()
        .iter()
        .flat_map(|s| s.shapes())
        .filter(|sh| sh.geometry().is_some())
        .count();
    let geom_missing = raw_shapes - geom_ok;

    let model = convert_presentation(&pres).expect("convert failed");
    let converted_shapes: usize = model.slides.iter().map(|s| s.shapes.len()).sum();
    let shapes_with_text: usize = model
        .slides
        .iter()
        .flat_map(|s| &s.shapes)
        .filter(|sh| sh.text.is_some())
        .count();

    DiagResult {
        file,
        slides,
        raw_shapes,
        geom_ok,
        geom_missing,
        converted_shapes,
        shapes_with_text,
        open_error: false,
    }
}

#[test]
fn diagnose_all_files() {
    let mut all_files = Vec::new();
    if let Some(d) = shapecrawler_dir() {
        all_files.extend(collect_pptx(&d));
    }
    if let Some(d) = openxml_dir() {
        all_files.extend(collect_pptx(&d));
    }

    if all_files.is_empty() {
        eprintln!("No reference pptx files found, skipping diagnostic.");
        return;
    }

    let mut total = 0;
    let mut errors = 0;
    let mut empty = 0;
    let mut with_text = 0;

    eprintln!(
        "\n{:10} {:50} {:>6} {:>6} {:>8} {:>8} {:>6}",
        "STATUS", "FILE", "SLIDES", "RAW", "GEOM_OK", "CONVERT", "TEXT"
    );
    eprintln!("{}", "-".repeat(100));

    for file in &all_files {
        let d = diagnose(file);
        total += 1;

        if d.open_error {
            errors += 1;
            eprintln!("{:10} {:50} OPEN_ERROR", "ERR", d.file);
            continue;
        }

        let status = if d.converted_shapes == 0 && d.raw_shapes > 0 {
            empty += 1;
            "EMPTY"
        } else if d.shapes_with_text > 0 {
            with_text += 1;
            "OK"
        } else if d.converted_shapes > 0 {
            "NO_TEXT"
        } else {
            "BLANK"
        };

        eprintln!(
            "{:10} {:50} {:>6} {:>6} {:>8} {:>8} {:>6}",
            status,
            d.file,
            d.slides,
            d.raw_shapes,
            d.geom_ok,
            d.converted_shapes,
            d.shapes_with_text
        );
    }

    eprintln!("\n  Total: {total} files, {errors} errors, {empty} empty, {with_text} with text");

    // This is the key regression assertion:
    // No file with raw shapes should produce zero converted shapes.
    let empty_files: Vec<_> = all_files
        .iter()
        .filter_map(|f| {
            let d = diagnose(f);
            if !d.open_error && d.raw_shapes > 0 && d.converted_shapes == 0 {
                Some(d.file)
            } else {
                None
            }
        })
        .collect();

    assert!(
        empty_files.is_empty(),
        "Files with shapes that produced 0 converted shapes: {empty_files:?}"
    );
}

// ---------------------------------------------------------------------------
// HTML snapshot renderer — for visual inspection without WASM
// ---------------------------------------------------------------------------

fn render_slide_html(model: &PresentationViewModel, slide_index: usize) -> String {
    let slide = &model.slides[slide_index];
    let mut html = String::new();

    html.push_str(&format!(
        "<div class=\"slide\" style=\"width:{}pt;height:{}pt;position:relative;background:#fff;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.3);\">",
        model.slide_width_pt, model.slide_height_pt
    ));

    // Background
    if let Some(bg) = &slide.background {
        match bg {
            offidized_pptview::model::BackgroundModel::Solid { color } => {
                html.push_str(&format!(
                    "<div style=\"position:absolute;inset:0;background:#{color};\"></div>"
                ));
            }
            offidized_pptview::model::BackgroundModel::Gradient { css } => {
                html.push_str(&format!(
                    "<div style=\"position:absolute;inset:0;background:{css};\"></div>"
                ));
            }
        }
    }

    // Shapes
    for shape in &slide.shapes {
        if shape.hidden {
            continue;
        }

        let mut style = format!(
            "position:absolute;left:{}pt;top:{}pt;width:{}pt;height:{}pt;box-sizing:border-box;overflow:hidden;",
            shape.x_pt, shape.y_pt, shape.width_pt, shape.height_pt
        );

        if let Some(rot) = shape.rotation {
            style.push_str(&format!("transform:rotate({rot}deg);"));
        }

        if let Some(fill) = &shape.fill {
            match fill {
                offidized_pptview::model::ShapeFillModel::Solid { color } => {
                    style.push_str(&format!("background-color:#{color};"));
                }
                offidized_pptview::model::ShapeFillModel::Gradient { css } => {
                    style.push_str(&format!("background:{css};"));
                }
                offidized_pptview::model::ShapeFillModel::None => {}
            }
        }

        if let Some(outline) = &shape.outline {
            let w = outline.width_pt.unwrap_or(1.0);
            let ds = outline.dash_style.as_deref().unwrap_or("solid");
            let c = outline.color.as_deref().unwrap_or("000");
            style.push_str(&format!("border:{w}pt {ds} #{c};"));
        }

        if shape.preset_geometry.as_deref() == Some("ellipse") {
            style.push_str("border-radius:50%;");
        }

        html.push_str(&format!("<div style=\"{style}\">"));

        // Image
        if let Some(idx) = shape.image_index {
            if let Some(img) = model.images.get(idx) {
                html.push_str(&format!(
                    "<img src=\"{}\" style=\"width:100%;height:100%;object-fit:contain;\">",
                    img.data_uri
                ));
            }
        }

        // Text
        if let Some(text) = &shape.text {
            let mut text_style = String::from("width:100%;height:100%;");

            if let Some(anchor) = &text.anchor {
                text_style.push_str("display:flex;flex-direction:column;");
                match anchor.as_str() {
                    "middle" => text_style.push_str("justify-content:center;"),
                    "bottom" => text_style.push_str("justify-content:flex-end;"),
                    _ => {}
                }
            }

            if let Some(insets) = &text.insets {
                text_style.push_str(&format!(
                    "padding:{}pt {}pt {}pt {}pt;",
                    insets.top_pt, insets.right_pt, insets.bottom_pt, insets.left_pt
                ));
            }

            html.push_str(&format!("<div style=\"{text_style}\">"));

            for para in &text.paragraphs {
                let mut pstyle = String::from("margin:0;");
                if let Some(a) = &para.alignment {
                    pstyle.push_str(&format!("text-align:{a};"));
                }
                if let Some(sb) = para.spacing_before_pt {
                    pstyle.push_str(&format!("margin-top:{sb}pt;"));
                }
                if let Some(sa) = para.spacing_after_pt {
                    pstyle.push_str(&format!("margin-bottom:{sa}pt;"));
                }
                if let Some(ls) = para.line_spacing {
                    pstyle.push_str(&format!("line-height:{ls};"));
                }

                html.push_str(&format!("<p style=\"{pstyle}\">"));

                // Bullet
                if let Some(bullet) = &para.bullet {
                    let bchar = bullet.char.as_deref().unwrap_or("\u{2022}");
                    html.push_str(&format!("<span style=\"margin-right:4pt;\">{bchar}</span>"));
                }

                for run in &para.runs {
                    let mut rstyle = String::new();
                    if let Some(ff) = &run.font_family {
                        rstyle.push_str(&format!("font-family:\"{ff}\",Calibri,sans-serif;"));
                    }
                    if let Some(fs) = run.font_size_pt {
                        rstyle.push_str(&format!("font-size:{fs}pt;"));
                    }
                    if run.bold {
                        rstyle.push_str("font-weight:bold;");
                    }
                    if run.italic {
                        rstyle.push_str("font-style:italic;");
                    }
                    if run.underline {
                        rstyle.push_str("text-decoration:underline;");
                    }
                    if let Some(c) = &run.color {
                        rstyle.push_str(&format!("color:#{c};"));
                    }

                    let escaped = html_escape(&run.text);
                    html.push_str(&format!("<span style=\"{rstyle}\">{escaped}</span>"));
                }

                html.push_str("</p>");
            }

            html.push_str("</div>");
        }

        // Table
        if let Some(table) = &shape.table {
            html.push_str("<table style=\"width:100%;height:100%;border-collapse:collapse;\">");
            for row in &table.rows {
                html.push_str("<tr>");
                for cell in &row.cells {
                    if cell.v_merge {
                        continue;
                    }
                    let mut tdstyle =
                        String::from("border:1px solid #ccc;padding:2pt 4pt;font-size:10pt;");
                    if let Some(fc) = &cell.fill_color {
                        tdstyle.push_str(&format!("background-color:#{fc};"));
                    }
                    let mut attrs = String::new();
                    if cell.grid_span > 1 {
                        attrs.push_str(&format!(" colspan=\"{}\"", cell.grid_span));
                    }
                    if cell.row_span > 1 {
                        attrs.push_str(&format!(" rowspan=\"{}\"", cell.row_span));
                    }
                    let escaped = html_escape(&cell.text);
                    html.push_str(&format!("<td style=\"{tdstyle}\"{attrs}>{escaped}</td>"));
                }
                html.push_str("</tr>");
            }
            html.push_str("</table>");
        }

        html.push_str("</div>"); // shape
    }

    html.push_str("</div>"); // slide
    html
}

fn render_presentation_html(model: &PresentationViewModel, title: &str) -> String {
    let mut html = String::new();
    html.push_str(&format!(
        r#"<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<style>
  body {{ background:#333; font-family:Calibri,Arial,sans-serif; padding:20px; }}
  h1 {{ color:#fff; font-size:16px; margin:0 0 10px; }}
  .slide-wrapper {{ margin:20px auto; text-align:center; }}
  .slide-label {{ color:#aaa; font-size:12px; margin-bottom:4px; }}
  .slide {{ display:inline-block; background:#fff; }}
</style>
</head>
<body>
<h1>{title} &mdash; {n} slides, {w}x{h}pt</h1>
"#,
        n = model.slides.len(),
        w = model.slide_width_pt,
        h = model.slide_height_pt,
    ));

    for i in 0..model.slides.len() {
        let shapes = model.slides[i].shapes.len();
        let text_shapes = model.slides[i]
            .shapes
            .iter()
            .filter(|s| s.text.is_some())
            .count();
        html.push_str(&format!(
            "<div class=\"slide-wrapper\"><div class=\"slide-label\">Slide {} \u{2014} {} shapes, {} with text</div>",
            i + 1, shapes, text_shapes
        ));
        html.push_str(&render_slide_html(model, i));
        html.push_str("</div>\n");
    }

    html.push_str("</body></html>");
    html
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

#[test]
fn render_selected_files_to_html() {
    let out = output_dir();

    // Render a curated set of interesting files
    let candidates = [
        (
            "ShapeCrawler",
            shapecrawler_dir(),
            vec![
                "002.pptx",
                "007_2 slides.pptx",
                "008.pptx",
                "020.pptx",
                "031.pptx",
                "058_bg-fill.pptx",
                "061_font-color.pptx",
                "065 table.pptx",
                "077 grouped shape.pptx",
                "078 textbox.pptx",
            ],
        ),
        (
            "OpenXML",
            openxml_dir(),
            vec![
                "Presentation.pptx",
                "animation.pptx",
                "o09_Performance_typical.pptx",
            ],
        ),
    ];

    let mut rendered = 0;

    for (source, dir, names) in &candidates {
        let Some(dir) = dir else { continue };
        let all_files = collect_pptx(dir);

        for name in names {
            let file = all_files
                .iter()
                .find(|p| p.file_name().is_some_and(|n| n.to_string_lossy() == *name));

            let Some(file) = file else {
                eprintln!("  SKIP: {source}/{name} — not found");
                continue;
            };

            let pres = match Presentation::open(file) {
                Ok(p) => p,
                Err(e) => {
                    eprintln!("  SKIP: {source}/{name} — {e}");
                    continue;
                }
            };

            let model = convert_presentation(&pres).expect("convert failed");
            let title = format!("{source}/{name}");
            let html = render_presentation_html(&model, &title);

            let safe_name = name.replace(' ', "_").replace(".pptx", ".html");
            let out_path = out.join(format!("{source}_{safe_name}"));
            std::fs::write(&out_path, html).expect("write html");
            eprintln!("  RENDERED: {}", out_path.display());
            rendered += 1;
        }
    }

    eprintln!("\n  Rendered {rendered} files to {}", out.display());
}

// ---------------------------------------------------------------------------
// JSON snapshot regression
// ---------------------------------------------------------------------------

#[test]
fn snapshot_model_json() {
    // Pick a simple, stable file and snapshot its model as JSON.
    // If conversion changes, the snapshot changes → easy to review.
    let dir = match shapecrawler_dir() {
        Some(d) => d,
        None => {
            eprintln!("ShapeCrawler assets not found, skipping snapshot.");
            return;
        }
    };

    let file = dir.join("002.pptx");
    if !file.exists() {
        eprintln!("002.pptx not found, skipping snapshot.");
        return;
    }

    let pres = Presentation::open(&file).expect("open 002.pptx");
    let model = convert_presentation(&pres).expect("convert 002.pptx");

    let json = serde_json::to_string_pretty(&model).expect("serialize model");

    let out = output_dir().join("snapshot_002.json");
    let prev = std::fs::read_to_string(&out).ok();

    std::fs::write(&out, &json).expect("write snapshot");

    if let Some(prev) = prev {
        if prev != json {
            eprintln!("SNAPSHOT CHANGED: {}", out.display());
            eprintln!("  Review the diff to verify the change is expected.");
            // Don't fail — just warn. Use git diff on the snapshot file.
        } else {
            eprintln!("  Snapshot unchanged: {}", out.display());
        }
    } else {
        eprintln!("  New snapshot created: {}", out.display());
    }

    // Basic sanity: 002.pptx should have 3 slides with shapes
    assert_eq!(model.slides.len(), 3, "002.pptx should have 3 slides");
    assert!(
        model.slides.iter().any(|s| !s.shapes.is_empty()),
        "at least one slide should have shapes"
    );
}
