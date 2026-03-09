use std::collections::{BTreeMap, BTreeSet, VecDeque};
use std::fs::{self, File};
use std::io::{BufReader, Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use anyhow::{anyhow, bail, Context, Result};
use offidized_docx::Document;
use offidized_opc::relationship::TargetMode;
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, PartData};
use offidized_pptx::Presentation;
use offidized_xlsx::Workbook;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use serde::Serialize;
use zip::ZipArchive;

#[derive(Debug, Clone)]
struct Args {
    references_root: PathBuf,
    output_dir: Option<PathBuf>,
    extensions: String,
    jobs: usize,
    max_files: Option<usize>,
    ignore_list: Option<PathBuf>,
    skip_csharp: bool,
    dotnet_runner_dll: Option<PathBuf>,
    dotnet_cli_home: PathBuf,
    keep_all_outputs: bool,
    compare_input: bool,
    rust_engine: RustEngine,
}

#[derive(Debug, Clone)]
struct RunConfig {
    references_root: PathBuf,
    output_dir: PathBuf,
    rust_output_root: PathBuf,
    csharp_output_root: PathBuf,
    skip_csharp: bool,
    dotnet_runner_dll: Option<PathBuf>,
    dotnet_cli_home: PathBuf,
    keep_all_outputs: bool,
    compare_input: bool,
    rust_engine: RustEngine,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RustEngine {
    FormatApis,
    OpcCore,
}

impl RustEngine {
    fn from_arg(value: &str) -> Result<Self> {
        match value {
            "format" => Ok(Self::FormatApis),
            "opc" => Ok(Self::OpcCore),
            _ => bail!("invalid --rust-engine value `{value}`; expected `format` or `opc`"),
        }
    }

    fn as_str(&self) -> &'static str {
        match self {
            Self::FormatApis => "format",
            Self::OpcCore => "opc",
        }
    }
}

#[derive(Debug, Clone, Serialize)]
struct StepResult {
    ok: bool,
    error: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
struct ContentDigest {
    size_bytes: u64,
    checksum: String,
}

#[derive(Debug, Clone, Serialize)]
struct EntryComparison {
    equal: bool,
    only_in_rust: usize,
    only_in_csharp: usize,
    differing_entries: usize,
    sample_differences: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Serialize)]
struct RelationshipSnapshot {
    rel_type: String,
    target: String,
    target_mode: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
struct PartSnapshot {
    uri: String,
    content_type: Option<String>,
    relationships: Vec<RelationshipSnapshot>,
    payload: ContentDigest,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
struct PackageSnapshot {
    package_relationships: Vec<RelationshipSnapshot>,
    parts: Vec<PartSnapshot>,
}

#[derive(Debug, Clone, Serialize)]
struct OpcComparison {
    equal: bool,
    missing_in_rust: usize,
    missing_in_csharp: usize,
    content_type_mismatches: usize,
    relationship_mismatches: usize,
    payload_mismatches: usize,
    package_relationship_mismatch: bool,
    sample_differences: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
struct FileResult {
    input_relative_path: String,
    extension: String,
    rust_roundtrip: StepResult,
    csharp_roundtrip: Option<StepResult>,
    raw_zip_equal: Option<bool>,
    entry_comparison: Option<EntryComparison>,
    opc_comparison: Option<OpcComparison>,
    passed: bool,
    failure_reason: Option<String>,
    duration_ms: u128,
    rust_output: Option<String>,
    csharp_output: Option<String>,
}

#[derive(Debug, Serialize)]
struct Report {
    generated_unix_timestamp: u64,
    references_root: String,
    output_dir: String,
    extensions: Vec<String>,
    jobs: usize,
    max_files: Option<usize>,
    ignore_list: Option<String>,
    ignored_files: usize,
    skip_csharp: bool,
    compare_input: bool,
    rust_engine: String,
    dotnet_runner_dll: Option<String>,
    dotnet_cli_home: Option<String>,
    total_files: usize,
    passed_files: usize,
    failed_files: usize,
    rust_roundtrip_failures: usize,
    csharp_roundtrip_failures: usize,
    comparison_failures: usize,
    raw_zip_exact_matches: usize,
    entry_level_matches: usize,
    opc_snapshot_matches: usize,
    duration_ms: u128,
    results: Vec<FileResult>,
}

fn main() -> Result<()> {
    let args = parse_args()?;

    if args.jobs == 0 {
        bail!("--jobs must be >= 1");
    }
    if !args.compare_input && !args.skip_csharp && args.dotnet_runner_dll.is_none() {
        bail!("--dotnet-runner-dll is required unless --skip-csharp is provided");
    }

    let references_root = absolute_path(&args.references_root)?;
    if !references_root.is_dir() {
        bail!(
            "references root does not exist or is not a directory: {}",
            references_root.display()
        );
    }

    let extension_filter = parse_extensions(args.extensions.as_str())?;
    let mut files = discover_files(&references_root, &extension_filter)?;
    let ignore_entries = load_ignore_entries(args.ignore_list.as_deref(), &references_root)
        .with_context(|| "load ignore-list entries (--ignore-list)".to_string())?;
    let mut ignored_files = 0_usize;
    if !ignore_entries.is_empty() {
        files.retain(|path| {
            let relative = path.strip_prefix(&references_root).unwrap_or(path);
            let key = normalize_path_for_report(relative);
            let ignored = ignore_entries.contains(key.as_str());
            if ignored {
                ignored_files = ignored_files.saturating_add(1);
            }
            !ignored
        });
    }
    if let Some(max_files) = args.max_files {
        if max_files < files.len() {
            files.truncate(max_files);
        }
    }
    if files.is_empty() {
        bail!(
            "no files discovered under {} for extensions {:?}",
            references_root.display(),
            extension_filter
        );
    }

    let output_dir = args
        .output_dir
        .as_ref()
        .map(|path| absolute_path(path))
        .transpose()?
        .unwrap_or_else(default_output_dir);

    let rust_output_root = output_dir.join("outputs").join("rust");
    let csharp_output_root = output_dir.join("outputs").join("csharp");
    fs::create_dir_all(&rust_output_root)?;
    if !args.skip_csharp && !args.compare_input {
        fs::create_dir_all(&csharp_output_root)?;
    }

    let config = RunConfig {
        references_root: references_root.clone(),
        output_dir: output_dir.clone(),
        rust_output_root,
        csharp_output_root,
        skip_csharp: args.skip_csharp,
        dotnet_runner_dll: args.dotnet_runner_dll.clone(),
        dotnet_cli_home: absolute_path(&args.dotnet_cli_home)?,
        keep_all_outputs: args.keep_all_outputs,
        compare_input: args.compare_input,
        rust_engine: args.rust_engine,
    };

    println!(
        "Differential scan: {} files under {}",
        files.len(),
        references_root.display()
    );
    println!(
        "Output directory: {} (keep_all_outputs={})",
        output_dir.display(),
        config.keep_all_outputs
    );
    println!("Workers: {}", args.jobs);
    if args.ignore_list.is_some() || !ignore_entries.is_empty() {
        println!("Ignored files: {}", ignored_files);
    }
    println!("Rust engine: {}", config.rust_engine.as_str());
    if config.compare_input {
        println!("Comparison baseline: input package (--compare-input)");
    } else if config.skip_csharp {
        println!("C# runner: skipped (--skip-csharp)");
    } else if let Some(dll) = config.dotnet_runner_dll.as_ref() {
        println!("C# runner dll: {}", dll.display());
    }

    let started = Instant::now();
    let results = run_files(files.as_slice(), &config, args.jobs);
    let duration_ms = started.elapsed().as_millis();

    let report = build_report(
        &args,
        &config,
        files.len(),
        ignored_files,
        duration_ms,
        results,
    );
    write_report_files(&output_dir, &report)?;

    println!(
        "Differential summary: passed={} failed={} total={}",
        report.passed_files, report.failed_files, report.total_files
    );
    println!(
        "Signals: raw_zip_exact={} entry_level={} opc_snapshot={}",
        report.raw_zip_exact_matches, report.entry_level_matches, report.opc_snapshot_matches
    );
    println!(
        "Reports: {} and {}",
        output_dir.join("report.md").display(),
        output_dir.join("report.json").display()
    );

    if report.failed_files > 0 {
        bail!("differential failures detected: {}", report.failed_files);
    }

    Ok(())
}

fn usage_text() -> &'static str {
    "Usage: offidized-differential [options]\n\n
Options:
  --references-root PATH       Root directory to scan (default: references)
  --output-dir PATH            Artifact output directory
  --extensions LIST            Comma-separated extensions (default: docx,xlsx,pptx)
  --jobs N                     Parallel workers (default: 1)
  --max-files N                Process at most N discovered files
  --ignore-list PATH           Relative-path ignore list file (one fixture path per line)
  --rust-engine MODE           Rust roundtrip engine: format|opc (default: format)
  --compare-input              Compare Rust output against original input (no C# runner)
  --skip-csharp                Skip C# differential side and run Rust-only checks
  --dotnet-runner-dll PATH     Path to built C# runner DLL (required unless --skip-csharp)
  --dotnet-cli-home PATH       DOTNET_CLI_HOME (default: .dotnet_home)
  --keep-all-outputs           Keep output files for passing and failing cases
  --help                       Show this help text"
}

fn parse_args() -> Result<Args> {
    let mut references_root = PathBuf::from("references");
    let mut output_dir = None;
    let mut extensions = "docx,xlsx,pptx".to_string();
    let mut jobs: usize = 1;
    let mut max_files = None;
    let mut ignore_list = None;
    let mut skip_csharp = false;
    let mut dotnet_runner_dll = None;
    let mut dotnet_cli_home = PathBuf::from(".dotnet_home");
    let mut keep_all_outputs = false;
    let mut compare_input = false;
    let mut rust_engine = RustEngine::FormatApis;

    let mut args = std::env::args().skip(1);
    while let Some(flag) = args.next() {
        match flag.as_str() {
            "--references-root" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --references-root"))?;
                references_root = PathBuf::from(value);
            }
            "--output-dir" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --output-dir"))?;
                output_dir = Some(PathBuf::from(value));
            }
            "--extensions" => {
                extensions = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --extensions"))?;
            }
            "--jobs" => {
                let raw = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --jobs"))?;
                jobs = raw
                    .parse::<usize>()
                    .context("--jobs must be a positive integer")?;
            }
            "--max-files" => {
                let raw = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --max-files"))?;
                max_files = Some(
                    raw.parse::<usize>()
                        .context("--max-files must be a positive integer")?,
                );
            }
            "--ignore-list" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --ignore-list"))?;
                ignore_list = Some(PathBuf::from(value));
            }
            "--rust-engine" => {
                let raw = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --rust-engine"))?;
                rust_engine = RustEngine::from_arg(raw.as_str())?;
            }
            "--compare-input" => {
                compare_input = true;
            }
            "--skip-csharp" => {
                skip_csharp = true;
            }
            "--dotnet-runner-dll" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --dotnet-runner-dll"))?;
                dotnet_runner_dll = Some(PathBuf::from(value));
            }
            "--dotnet-cli-home" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("missing value for --dotnet-cli-home"))?;
                dotnet_cli_home = PathBuf::from(value);
            }
            "--keep-all-outputs" => {
                keep_all_outputs = true;
            }
            "--help" | "-h" => {
                println!("{}", usage_text());
                std::process::exit(0);
            }
            unknown => {
                bail!("unknown option: {unknown}\n\n{}", usage_text());
            }
        }
    }

    Ok(Args {
        references_root,
        output_dir,
        extensions,
        jobs,
        max_files,
        ignore_list,
        skip_csharp,
        dotnet_runner_dll,
        dotnet_cli_home,
        keep_all_outputs,
        compare_input,
        rust_engine,
    })
}

fn run_files(files: &[PathBuf], config: &RunConfig, jobs: usize) -> Vec<FileResult> {
    if jobs <= 1 {
        return files
            .iter()
            .enumerate()
            .map(|(index, path)| process_file(index, path, config))
            .collect::<Vec<_>>();
    }

    let queue = Arc::new(Mutex::new((0..files.len()).collect::<VecDeque<_>>()));
    let results = Arc::new(Mutex::new(Vec::with_capacity(files.len())));
    let files = Arc::new(files.to_vec());
    let mut handles = Vec::new();

    for _ in 0..jobs {
        let queue = Arc::clone(&queue);
        let results = Arc::clone(&results);
        let files = Arc::clone(&files);
        let config = config.clone();

        handles.push(thread::spawn(move || loop {
            let index = match queue.lock() {
                Ok(mut guard) => guard.pop_front(),
                Err(_) => None,
            };
            let Some(index) = index else {
                break;
            };

            let path = &files[index];
            let result = process_file(index, path, &config);

            if let Ok(mut guard) = results.lock() {
                guard.push(result);
            }
        }));
    }

    for handle in handles {
        let _ = handle.join();
    }

    let mut collected = match results.lock() {
        Ok(guard) => guard.clone(),
        Err(_) => Vec::new(),
    };
    collected.sort_by(|left, right| left.input_relative_path.cmp(&right.input_relative_path));
    collected
}

fn process_file(index: usize, input_path: &Path, config: &RunConfig) -> FileResult {
    let started = Instant::now();
    let relative_path = input_path
        .strip_prefix(&config.references_root)
        .unwrap_or(input_path);
    let input_relative_path = normalize_path_for_report(relative_path);
    let extension = input_path
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .unwrap_or_default();

    let rust_output_path = config.rust_output_root.join(format!(
        "{:08}_{}",
        index,
        normalize_path_for_fs(relative_path)
    ));
    let csharp_output_path = config.csharp_output_root.join(format!(
        "{:08}_{}",
        index,
        normalize_path_for_fs(relative_path)
    ));

    let mut rust_output_report_path = Some(rust_output_path.display().to_string());
    let mut csharp_output_report_path = if config.skip_csharp || config.compare_input {
        None
    } else {
        Some(csharp_output_path.display().to_string())
    };

    let rust_roundtrip = match roundtrip_rust(input_path, &rust_output_path, config.rust_engine) {
        Ok(()) => StepResult {
            ok: true,
            error: None,
        },
        Err(error) => StepResult {
            ok: false,
            error: Some(format!("{error:#}")),
        },
    };

    let csharp_roundtrip = if config.skip_csharp || config.compare_input {
        None
    } else {
        Some(
            match roundtrip_csharp(input_path, &csharp_output_path, config) {
                Ok(()) => StepResult {
                    ok: true,
                    error: None,
                },
                Err(error) => StepResult {
                    ok: false,
                    error: Some(format!("{error:#}")),
                },
            },
        )
    };

    let mut raw_zip_equal = None;
    let mut entry_comparison = None;
    let mut opc_comparison = None;
    let mut passed = false;
    let mut failure_reason = None;

    if !rust_roundtrip.ok {
        failure_reason = Some("rust roundtrip failed".to_string());
    } else if config.compare_input {
        match compare_outputs(&rust_output_path, input_path) {
            Ok((raw_equal, entry, opc)) => {
                raw_zip_equal = Some(raw_equal);
                entry_comparison = Some(entry.clone());
                opc_comparison = Some(opc.clone());
                if entry.equal && opc.equal {
                    passed = true;
                } else {
                    failure_reason = Some(format!(
                        "comparison mismatch against input: entry_equal={} opc_equal={}",
                        entry.equal, opc.equal
                    ));
                }
            }
            Err(error) => {
                failure_reason = Some(format!("comparison error against input: {error}"));
            }
        }
    } else if let Some(csharp_step) = csharp_roundtrip.as_ref() {
        if !csharp_step.ok {
            failure_reason = Some("csharp roundtrip failed".to_string());
        } else {
            match compare_outputs(&rust_output_path, &csharp_output_path) {
                Ok((raw_equal, entry, opc)) => {
                    raw_zip_equal = Some(raw_equal);
                    entry_comparison = Some(entry.clone());
                    opc_comparison = Some(opc.clone());
                    if entry.equal && opc.equal {
                        passed = true;
                    } else {
                        failure_reason = Some(format!(
                            "comparison mismatch: entry_equal={} opc_equal={}",
                            entry.equal, opc.equal
                        ));
                    }
                }
                Err(error) => {
                    failure_reason = Some(format!("comparison error: {error}"));
                }
            }
        }
    } else {
        passed = true;
    }

    if passed && !config.keep_all_outputs {
        let _ = fs::remove_file(&rust_output_path);
        rust_output_report_path = None;

        if !config.skip_csharp && !config.compare_input {
            let _ = fs::remove_file(&csharp_output_path);
            csharp_output_report_path = None;
        }
    }

    FileResult {
        input_relative_path,
        extension,
        rust_roundtrip,
        csharp_roundtrip,
        raw_zip_equal,
        entry_comparison,
        opc_comparison,
        passed,
        failure_reason,
        duration_ms: started.elapsed().as_millis(),
        rust_output: rust_output_report_path,
        csharp_output: csharp_output_report_path,
    }
}

fn roundtrip_rust(input: &Path, output: &Path, engine: RustEngine) -> Result<()> {
    ensure_parent_dir(output)?;

    if engine == RustEngine::OpcCore {
        let package = Package::open(input)
            .with_context(|| format!("open OPC package `{}`", input.display()))?;
        package
            .save(output)
            .with_context(|| format!("save OPC package `{}`", output.display()))?;
        return Ok(());
    }

    let extension = input
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .unwrap_or_default();

    match extension.as_str() {
        "docx" => {
            let document = Document::open(input)
                .with_context(|| format!("open DOCX `{}`", input.display()))?;
            document
                .save(output)
                .with_context(|| format!("save DOCX `{}`", output.display()))?;
        }
        "xlsx" => {
            let workbook = Workbook::open(input)
                .with_context(|| format!("open XLSX `{}`", input.display()))?;
            workbook
                .save(output)
                .with_context(|| format!("save XLSX `{}`", output.display()))?;
        }
        "pptx" => {
            let mut presentation = Presentation::open(input)
                .with_context(|| format!("open PPTX `{}`", input.display()))?;
            presentation
                .save(output)
                .with_context(|| format!("save PPTX `{}`", output.display()))?;
        }
        _ => bail!("unsupported extension for rust roundtrip: `{extension}`"),
    }

    Ok(())
}

fn roundtrip_csharp(input: &Path, output: &Path, config: &RunConfig) -> Result<()> {
    ensure_parent_dir(output)?;
    let runner = config
        .dotnet_runner_dll
        .as_ref()
        .ok_or_else(|| anyhow!("missing dotnet runner dll path"))?;

    let output_result = Command::new("dotnet")
        .arg(runner)
        .arg(input)
        .arg(output)
        .env("DOTNET_CLI_HOME", &config.dotnet_cli_home)
        .env("DOTNET_SKIP_FIRST_TIME_EXPERIENCE", "1")
        .env("DOTNET_NOLOGO", "1")
        .output()
        .with_context(|| format!("execute dotnet runner `{}`", runner.display()))?;

    if output_result.status.success() {
        Ok(())
    } else {
        let stdout = String::from_utf8_lossy(&output_result.stdout);
        let stderr = String::from_utf8_lossy(&output_result.stderr);
        bail!(
            "dotnet runner failed for `{}` -> `{}` (code {:?})\nstdout:\n{}\nstderr:\n{}",
            input.display(),
            output.display(),
            output_result.status.code(),
            stdout.trim(),
            stderr.trim()
        );
    }
}

fn compare_outputs(
    rust_output: &Path,
    csharp_output: &Path,
) -> Result<(bool, EntryComparison, OpcComparison)> {
    let raw_equal = stream_bytes_equal(rust_output, csharp_output)?;

    let rust_entries = zip_entry_digest_map(rust_output)
        .with_context(|| format!("read zip entries from `{}`", rust_output.display()))?;
    let csharp_entries = zip_entry_digest_map(csharp_output)
        .with_context(|| format!("read zip entries from `{}`", csharp_output.display()))?;
    let entry_comparison = compare_entry_digests(&rust_entries, &csharp_entries);

    let rust_snapshot = package_snapshot(rust_output)
        .with_context(|| format!("read OPC snapshot from `{}`", rust_output.display()))?;
    let csharp_snapshot = package_snapshot(csharp_output)
        .with_context(|| format!("read OPC snapshot from `{}`", csharp_output.display()))?;
    let opc_comparison = compare_opc_snapshots(&rust_snapshot, &csharp_snapshot);

    Ok((raw_equal, entry_comparison, opc_comparison))
}

fn stream_bytes_equal(left: &Path, right: &Path) -> Result<bool> {
    let left_meta = fs::metadata(left)?;
    let right_meta = fs::metadata(right)?;
    if left_meta.len() != right_meta.len() {
        return Ok(false);
    }

    let mut left_reader = BufReader::new(File::open(left)?);
    let mut right_reader = BufReader::new(File::open(right)?);
    let mut left_buffer = [0_u8; 8192];
    let mut right_buffer = [0_u8; 8192];

    loop {
        let left_read = left_reader.read(&mut left_buffer)?;
        let right_read = right_reader.read(&mut right_buffer)?;
        if left_read != right_read {
            return Ok(false);
        }
        if left_read == 0 {
            return Ok(true);
        }
        if left_buffer[..left_read] != right_buffer[..right_read] {
            return Ok(false);
        }
    }
}

fn zip_entry_digest_map(path: &Path) -> Result<BTreeMap<String, ContentDigest>> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut map = BTreeMap::new();

    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        if entry.is_dir() {
            continue;
        }

        let name = canonicalize_zip_entry_name(entry.name());
        let mut payload = Vec::new();
        entry.read_to_end(&mut payload)?;
        let digest = digest_entry_payload(name.as_str(), &payload);
        map.insert(name, digest);
    }

    Ok(map)
}

fn digest_bytes(bytes: &[u8]) -> ContentDigest {
    let mut checksum = 1469598103934665603_u64;
    for byte in bytes {
        checksum ^= u64::from(*byte);
        checksum = checksum.wrapping_mul(1099511628211_u64);
    }

    ContentDigest {
        size_bytes: bytes.len() as u64,
        checksum: format!("{:016x}", checksum),
    }
}

fn digest_entry_payload(name: &str, bytes: &[u8]) -> ContentDigest {
    let should_try_xml = is_xml_name(name) || looks_like_xml(bytes);
    digest_payload(bytes, should_try_xml)
}

fn digest_part_payload(
    part_uri: &str,
    content_type: Option<&str>,
    data: &PartData,
) -> ContentDigest {
    let should_try_xml = matches!(data, PartData::Xml(_))
        || content_type.is_some_and(is_xml_content_type)
        || is_xml_name(part_uri)
        || looks_like_xml(data.as_bytes());
    digest_payload(data.as_bytes(), should_try_xml)
}

fn digest_payload(bytes: &[u8], should_try_xml: bool) -> ContentDigest {
    if should_try_xml {
        if let Some(canonical) = canonicalize_xml_payload(bytes) {
            return digest_bytes(&canonical);
        }
    }
    digest_bytes(bytes)
}

fn canonicalize_zip_entry_name(name: &str) -> String {
    normalize_path(name, false)
}

fn canonicalize_package_uri(uri: &str) -> String {
    normalize_path(uri, true)
}

fn normalize_path(path: &str, absolute: bool) -> String {
    let normalized = path.replace('\\', "/");
    let mut segments = Vec::new();

    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                let _ = segments.pop();
            }
            value => segments.push(value),
        }
    }

    if absolute {
        if segments.is_empty() {
            "/".to_string()
        } else {
            format!("/{}", segments.join("/"))
        }
    } else {
        segments.join("/")
    }
}

fn is_xml_name(name: &str) -> bool {
    let normalized = name.replace('\\', "/");
    normalized
        .rsplit_once('.')
        .map(|(_, extension)| {
            let extension = extension.to_ascii_lowercase();
            extension == "xml" || extension == "rels"
        })
        .unwrap_or(false)
}

fn is_xml_content_type(content_type: &str) -> bool {
    let content_type = content_type.to_ascii_lowercase();
    content_type.contains("xml") || content_type.contains("+xml")
}

fn looks_like_xml(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }

    let mut offset = 0;
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        offset = 3;
    }

    while offset < bytes.len() && bytes[offset].is_ascii_whitespace() {
        offset += 1;
    }

    bytes.get(offset).copied() == Some(b'<')
}

fn canonicalize_xml_payload(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len()));
    let mut buf = Vec::new();
    let mut saw_element = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) => {
                saw_element = true;
                write_canonical_element(&mut writer, event, false).ok()?;
            }
            Ok(Event::Empty(event)) => {
                saw_element = true;
                write_canonical_element(&mut writer, event, true).ok()?;
            }
            Ok(Event::End(event)) => {
                let name = String::from_utf8_lossy(event.name().as_ref()).into_owned();
                writer
                    .write_event(Event::End(BytesEnd::new(name.as_str())))
                    .ok()?;
            }
            Ok(Event::Text(event)) => {
                let text = event.xml_content().ok()?.into_owned();
                if !text.trim().is_empty() {
                    writer
                        .write_event(Event::Text(BytesText::new(text.as_str())))
                        .ok()?;
                }
            }
            Ok(Event::CData(event)) => {
                let text = String::from_utf8_lossy(event.as_ref()).into_owned();
                if !text.trim().is_empty() {
                    writer
                        .write_event(Event::Text(BytesText::new(text.as_str())))
                        .ok()?;
                }
            }
            Ok(Event::Decl(_)) | Ok(Event::Comment(_)) | Ok(Event::PI(_)) => {}
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => return None,
        }
        buf.clear();
    }

    saw_element.then(|| writer.into_inner())
}

fn write_canonical_element(
    writer: &mut Writer<Vec<u8>>,
    event: BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    let name = String::from_utf8_lossy(event.name().as_ref()).into_owned();
    let mut attributes = Vec::new();
    for attribute in event.attributes().with_checks(false) {
        let attribute = attribute?;
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        let value = attribute.unescape_value()?.into_owned();
        attributes.push((key, value));
    }
    attributes.sort();

    let mut canonical = BytesStart::new(name.as_str());
    for (key, value) in &attributes {
        canonical.push_attribute((key.as_str(), value.as_str()));
    }

    if empty {
        writer.write_event(Event::Empty(canonical))?;
    } else {
        writer.write_event(Event::Start(canonical))?;
    }
    Ok(())
}

fn split_uri_suffix(value: &str) -> (&str, &str) {
    let query = value.find('?');
    let fragment = value.find('#');
    let boundary = match (query, fragment) {
        (Some(left), Some(right)) => Some(left.min(right)),
        (Some(index), None) | (None, Some(index)) => Some(index),
        (None, None) => None,
    };
    boundary
        .map(|index| value.split_at(index))
        .unwrap_or((value, ""))
}

fn normalize_internal_relationship_target(
    source_part_uri: Option<&PartUri>,
    target: &str,
) -> String {
    let (raw_path, suffix) = split_uri_suffix(target);
    let normalized_input = raw_path.replace('\\', "/");

    let normalized_path = if normalized_input.starts_with('/') {
        canonicalize_package_uri(normalized_input.as_str())
    } else if normalized_input.is_empty() {
        source_part_uri
            .map(|part_uri| canonicalize_package_uri(part_uri.as_str()))
            .unwrap_or_else(|| "/".to_string())
    } else {
        let base = match source_part_uri.cloned() {
            Some(part_uri) => part_uri,
            None => match PartUri::new("/") {
                Ok(root) => root,
                Err(_) => {
                    return format!(
                        "{}{}",
                        canonicalize_package_uri(normalized_input.as_str()),
                        suffix
                    )
                }
            },
        };
        match base.resolve_relative(normalized_input.as_str()) {
            Ok(resolved) => canonicalize_package_uri(resolved.as_str()),
            Err(_) => canonicalize_package_uri(normalized_input.as_str()),
        }
    };

    format!("{normalized_path}{suffix}")
}

fn compare_entry_digests(
    rust_entries: &BTreeMap<String, ContentDigest>,
    csharp_entries: &BTreeMap<String, ContentDigest>,
) -> EntryComparison {
    let mut only_in_rust = 0_usize;
    let mut only_in_csharp = 0_usize;
    let mut differing_entries = 0_usize;
    let mut sample_differences = Vec::new();
    let mut keys = BTreeSet::new();

    keys.extend(rust_entries.keys().cloned());
    keys.extend(csharp_entries.keys().cloned());

    for key in keys {
        match (rust_entries.get(&key), csharp_entries.get(&key)) {
            (Some(_), None) => {
                only_in_rust = only_in_rust.saturating_add(1);
                if sample_differences.len() < 20 {
                    sample_differences.push(format!("only in rust: {key}"));
                }
            }
            (None, Some(_)) => {
                only_in_csharp = only_in_csharp.saturating_add(1);
                if sample_differences.len() < 20 {
                    sample_differences.push(format!("only in csharp: {key}"));
                }
            }
            (Some(left), Some(right)) => {
                if left != right {
                    differing_entries = differing_entries.saturating_add(1);
                    if sample_differences.len() < 20 {
                        sample_differences.push(format!(
                            "entry differs: {key} (rust size={} checksum={}, csharp size={} checksum={})",
                            left.size_bytes, left.checksum, right.size_bytes, right.checksum
                        ));
                    }
                }
            }
            (None, None) => {}
        }
    }

    EntryComparison {
        equal: only_in_rust == 0 && only_in_csharp == 0 && differing_entries == 0,
        only_in_rust,
        only_in_csharp,
        differing_entries,
        sample_differences,
    }
}

fn package_snapshot(path: &Path) -> Result<PackageSnapshot> {
    let package = Package::open(path)?;

    let mut package_relationships = snapshot_relationships(package.relationships().iter(), None);
    package_relationships.sort();

    let mut parts = package
        .parts()
        .map(|part| {
            let mut relationships =
                snapshot_relationships(part.relationships.iter(), Some(&part.uri));
            relationships.sort();
            PartSnapshot {
                uri: canonicalize_package_uri(part.uri.as_str()),
                content_type: part.content_type.clone(),
                relationships,
                payload: digest_part_payload(
                    part.uri.as_str(),
                    part.content_type.as_deref(),
                    &part.data,
                ),
            }
        })
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.uri.cmp(&right.uri));

    Ok(PackageSnapshot {
        package_relationships,
        parts,
    })
}

fn snapshot_relationships<'a>(
    relationships: impl Iterator<Item = &'a offidized_opc::Relationship>,
    source_part_uri: Option<&PartUri>,
) -> Vec<RelationshipSnapshot> {
    relationships
        .map(|relationship| RelationshipSnapshot {
            rel_type: relationship.rel_type.clone(),
            target: match relationship.target_mode {
                TargetMode::Internal => {
                    normalize_internal_relationship_target(source_part_uri, &relationship.target)
                }
                TargetMode::External => relationship.target.clone(),
            },
            target_mode: match relationship.target_mode {
                TargetMode::Internal => "Internal".to_string(),
                TargetMode::External => "External".to_string(),
            },
        })
        .collect::<Vec<_>>()
}

fn compare_opc_snapshots(
    rust_snapshot: &PackageSnapshot,
    csharp_snapshot: &PackageSnapshot,
) -> OpcComparison {
    let mut missing_in_rust = 0_usize;
    let mut missing_in_csharp = 0_usize;
    let mut content_type_mismatches = 0_usize;
    let mut relationship_mismatches = 0_usize;
    let mut payload_mismatches = 0_usize;
    let package_relationship_mismatch =
        rust_snapshot.package_relationships != csharp_snapshot.package_relationships;
    let mut sample_differences = Vec::new();

    let rust_parts = rust_snapshot
        .parts
        .iter()
        .map(|part| (part.uri.as_str(), part))
        .collect::<BTreeMap<_, _>>();
    let csharp_parts = csharp_snapshot
        .parts
        .iter()
        .map(|part| (part.uri.as_str(), part))
        .collect::<BTreeMap<_, _>>();

    let mut uris = BTreeSet::new();
    uris.extend(rust_parts.keys().copied());
    uris.extend(csharp_parts.keys().copied());

    for uri in uris {
        match (rust_parts.get(uri), csharp_parts.get(uri)) {
            (Some(_), None) => {
                missing_in_csharp = missing_in_csharp.saturating_add(1);
                if sample_differences.len() < 20 {
                    sample_differences.push(format!("part missing in csharp output: {uri}"));
                }
            }
            (None, Some(_)) => {
                missing_in_rust = missing_in_rust.saturating_add(1);
                if sample_differences.len() < 20 {
                    sample_differences.push(format!("part missing in rust output: {uri}"));
                }
            }
            (Some(left), Some(right)) => {
                if left.content_type != right.content_type {
                    content_type_mismatches = content_type_mismatches.saturating_add(1);
                    if sample_differences.len() < 20 {
                        sample_differences.push(format!(
                            "content type mismatch for {uri}: rust={:?} csharp={:?}",
                            left.content_type, right.content_type
                        ));
                    }
                }
                if left.relationships != right.relationships {
                    relationship_mismatches = relationship_mismatches.saturating_add(1);
                    if sample_differences.len() < 20 {
                        sample_differences.push(format!("relationship mismatch for part {uri}"));
                    }
                }
                if left.payload != right.payload {
                    payload_mismatches = payload_mismatches.saturating_add(1);
                    if sample_differences.len() < 20 {
                        sample_differences.push(format!(
                            "payload mismatch for {uri}: rust(size={},checksum={}) csharp(size={},checksum={})",
                            left.payload.size_bytes,
                            left.payload.checksum,
                            right.payload.size_bytes,
                            right.payload.checksum
                        ));
                    }
                }
            }
            (None, None) => {}
        }
    }

    let equal = missing_in_rust == 0
        && missing_in_csharp == 0
        && content_type_mismatches == 0
        && relationship_mismatches == 0
        && payload_mismatches == 0
        && !package_relationship_mismatch;

    OpcComparison {
        equal,
        missing_in_rust,
        missing_in_csharp,
        content_type_mismatches,
        relationship_mismatches,
        payload_mismatches,
        package_relationship_mismatch,
        sample_differences,
    }
}

fn build_report(
    args: &Args,
    config: &RunConfig,
    total_files: usize,
    ignored_files: usize,
    duration_ms: u128,
    results: Vec<FileResult>,
) -> Report {
    let passed_files = results.iter().filter(|result| result.passed).count();
    let failed_files = total_files.saturating_sub(passed_files);
    let rust_roundtrip_failures = results
        .iter()
        .filter(|result| !result.rust_roundtrip.ok)
        .count();
    let csharp_roundtrip_failures = results
        .iter()
        .filter(|result| {
            result
                .csharp_roundtrip
                .as_ref()
                .is_some_and(|step| !step.ok)
        })
        .count();
    let comparison_failures = results
        .iter()
        .filter(|result| {
            let has_comparison =
                result.entry_comparison.is_some() || result.opc_comparison.is_some();
            result.rust_roundtrip.ok && has_comparison && !result.passed
        })
        .count();
    let raw_zip_exact_matches = results
        .iter()
        .filter(|result| result.raw_zip_equal == Some(true))
        .count();
    let entry_level_matches = results
        .iter()
        .filter(|result| {
            result
                .entry_comparison
                .as_ref()
                .is_some_and(|comparison| comparison.equal)
        })
        .count();
    let opc_snapshot_matches = results
        .iter()
        .filter(|result| {
            result
                .opc_comparison
                .as_ref()
                .is_some_and(|comparison| comparison.equal)
        })
        .count();

    let generated_unix_timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs())
        .unwrap_or(0);

    Report {
        generated_unix_timestamp,
        references_root: config.references_root.display().to_string(),
        output_dir: config.output_dir.display().to_string(),
        extensions: parse_extensions(args.extensions.as_str())
            .unwrap_or_default()
            .into_iter()
            .collect::<Vec<_>>(),
        jobs: args.jobs,
        max_files: args.max_files,
        ignore_list: args
            .ignore_list
            .as_ref()
            .map(|path| path.display().to_string()),
        ignored_files,
        skip_csharp: args.skip_csharp || args.compare_input,
        compare_input: args.compare_input,
        rust_engine: args.rust_engine.as_str().to_string(),
        dotnet_runner_dll: config
            .dotnet_runner_dll
            .as_ref()
            .map(|path| path.display().to_string()),
        dotnet_cli_home: if args.skip_csharp || args.compare_input {
            None
        } else {
            Some(config.dotnet_cli_home.display().to_string())
        },
        total_files,
        passed_files,
        failed_files,
        rust_roundtrip_failures,
        csharp_roundtrip_failures,
        comparison_failures,
        raw_zip_exact_matches,
        entry_level_matches,
        opc_snapshot_matches,
        duration_ms,
        results,
    }
}

fn write_report_files(output_dir: &Path, report: &Report) -> Result<()> {
    fs::create_dir_all(output_dir)?;
    let json_path = output_dir.join("report.json");
    let markdown_path = output_dir.join("report.md");

    let json = serde_json::to_vec_pretty(report)?;
    fs::write(&json_path, json)?;

    let mut markdown = String::new();
    markdown.push_str("# Differential Regression Report\n\n");
    markdown.push_str(&format!(
        "- references_root: `{}`\n",
        report.references_root
    ));
    markdown.push_str(&format!("- output_dir: `{}`\n", report.output_dir));
    markdown.push_str(&format!(
        "- generated_unix_timestamp: `{}`\n",
        report.generated_unix_timestamp
    ));
    markdown.push_str(&format!(
        "- extensions: `{}`\n",
        report.extensions.join(",")
    ));
    markdown.push_str(&format!("- skip_csharp: `{}`\n", report.skip_csharp));
    markdown.push_str(&format!("- compare_input: `{}`\n", report.compare_input));
    markdown.push_str(&format!("- rust_engine: `{}`\n", report.rust_engine));
    markdown.push_str(&format!("- jobs: `{}`\n", report.jobs));
    markdown.push_str(&format!("- max_files: `{:?}`\n", report.max_files));
    markdown.push_str(&format!("- ignore_list: `{:?}`\n", report.ignore_list));
    markdown.push_str(&format!("- ignored_files: `{}`\n", report.ignored_files));
    markdown.push_str(&format!("- total_files: `{}`\n", report.total_files));
    markdown.push_str(&format!("- passed_files: `{}`\n", report.passed_files));
    markdown.push_str(&format!("- failed_files: `{}`\n", report.failed_files));
    markdown.push_str(&format!(
        "- rust_roundtrip_failures: `{}`\n",
        report.rust_roundtrip_failures
    ));
    markdown.push_str(&format!(
        "- csharp_roundtrip_failures: `{}`\n",
        report.csharp_roundtrip_failures
    ));
    markdown.push_str(&format!(
        "- comparison_failures: `{}`\n",
        report.comparison_failures
    ));
    markdown.push_str(&format!(
        "- raw_zip_exact_matches: `{}`\n",
        report.raw_zip_exact_matches
    ));
    markdown.push_str(&format!(
        "- entry_level_matches: `{}`\n",
        report.entry_level_matches
    ));
    markdown.push_str(&format!(
        "- opc_snapshot_matches: `{}`\n",
        report.opc_snapshot_matches
    ));
    markdown.push_str(&format!("- duration_ms: `{}`\n", report.duration_ms));
    markdown.push_str("\n## Failure Criteria\n\n");
    markdown.push_str(
        "A file is marked failed if Rust roundtrip fails, C# roundtrip fails, or either of these checks fails:\n",
    );
    markdown.push_str("- ZIP entry-level canonical comparison (entry names + payload digests)\n");
    markdown.push_str(
        "- OPC snapshot comparison (package/part relationships, content types, and payload digests)\n",
    );
    markdown.push_str(
        "Raw ZIP byte equality is reported as an additional strict signal but is not the pass/fail gate.\n",
    );

    if report.failed_files > 0 {
        markdown.push_str("\n## Failures\n\n");
        markdown.push_str("| Input | Reason |\n");
        markdown.push_str("| --- | --- |\n");
        for result in report
            .results
            .iter()
            .filter(|result| !result.passed)
            .take(300)
        {
            let reason = result
                .failure_reason
                .clone()
                .unwrap_or_else(|| "unknown failure".to_string())
                .replace('|', "\\|");
            markdown.push_str(&format!(
                "| `{}` | {} |\n",
                result.input_relative_path, reason
            ));
        }
        let hidden = report.failed_files.saturating_sub(300);
        if hidden > 0 {
            markdown.push_str(&format!("\n... {} additional failures omitted.\n", hidden));
        }
    }

    let mut file = File::create(&markdown_path)?;
    file.write_all(markdown.as_bytes())?;

    Ok(())
}

fn discover_files(root: &Path, extension_filter: &BTreeSet<String>) -> Result<Vec<PathBuf>> {
    let mut discovered = Vec::new();
    discover_files_recursive(root, extension_filter, &mut discovered)?;
    discovered.sort();
    Ok(discovered)
}

fn load_ignore_entries(
    ignore_list_path: Option<&Path>,
    references_root: &Path,
) -> Result<BTreeSet<String>> {
    let Some(ignore_list_path) = ignore_list_path else {
        return Ok(BTreeSet::new());
    };

    let ignore_path = if ignore_list_path.is_absolute() {
        ignore_list_path.to_path_buf()
    } else {
        absolute_path(ignore_list_path)?
    };

    if !ignore_path.is_file() {
        bail!(
            "ignore list does not exist or is not a file: {}",
            ignore_path.display()
        );
    }

    let contents = fs::read_to_string(&ignore_path)
        .with_context(|| format!("read ignore list file `{}`", ignore_path.display()))?;
    let mut entries = BTreeSet::new();
    for raw_line in contents.lines() {
        let line = raw_line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        let normalized = line.replace('\\', "/");
        if normalized.starts_with('/') {
            let as_path = Path::new(normalized.as_str());
            if let Ok(relative) = as_path.strip_prefix(references_root) {
                entries.insert(normalize_path_for_report(relative));
                continue;
            }
        }
        entries.insert(normalized);
    }

    Ok(entries)
}

fn discover_files_recursive(
    directory: &Path,
    extension_filter: &BTreeSet<String>,
    discovered: &mut Vec<PathBuf>,
) -> Result<()> {
    for entry in fs::read_dir(directory)? {
        let entry = entry?;
        let file_type = entry.file_type()?;
        let path = entry.path();

        if file_type.is_symlink() {
            continue;
        }
        if file_type.is_dir() {
            discover_files_recursive(path.as_path(), extension_filter, discovered)?;
            continue;
        }
        if !file_type.is_file() {
            continue;
        }

        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .map(|value| value.to_ascii_lowercase());
        if extension
            .as_ref()
            .is_some_and(|ext| extension_filter.contains(ext.as_str()))
        {
            discovered.push(path);
        }
    }

    Ok(())
}

fn parse_extensions(raw: &str) -> Result<BTreeSet<String>> {
    let parsed = raw
        .split(',')
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(|value| value.trim_start_matches('.').to_ascii_lowercase())
        .filter(|value| !value.is_empty())
        .collect::<BTreeSet<_>>();
    if parsed.is_empty() {
        bail!("no extensions provided");
    }
    Ok(parsed)
}

fn ensure_parent_dir(path: &Path) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    Ok(())
}

fn normalize_path_for_report(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn normalize_path_for_fs(path: &Path) -> String {
    normalize_path_for_report(path).replace('/', "__")
}

fn absolute_path(path: &Path) -> Result<PathBuf> {
    if path.is_absolute() {
        Ok(path.to_path_buf())
    } else {
        Ok(std::env::current_dir()?.join(path))
    }
}

fn default_output_dir() -> PathBuf {
    let unix_seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs())
        .unwrap_or(0);
    PathBuf::from("artifacts")
        .join("differential-regression")
        .join(unix_seconds.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    struct TempZip {
        path: PathBuf,
    }

    impl TempZip {
        fn new(label: &str, entries: &[(&str, &[u8])]) -> Self {
            let path = unique_temp_zip_path(label);
            write_zip(path.as_path(), entries);
            Self { path }
        }

        fn path(&self) -> &Path {
            self.path.as_path()
        }
    }

    impl Drop for TempZip {
        fn drop(&mut self) {
            let _ = fs::remove_file(self.path.as_path());
        }
    }

    fn unique_temp_zip_path(label: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time should be after unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "offidized-differential-{label}-{}-{nanos}.zip",
            std::process::id()
        ))
    }

    fn write_zip(path: &Path, entries: &[(&str, &[u8])]) {
        let file = File::create(path).expect("temporary zip file should be created");
        let mut writer = zip::ZipWriter::new(file);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            writer
                .start_file(*name, options)
                .expect("zip entry should start");
            writer
                .write_all(bytes)
                .expect("zip entry bytes should be written");
        }

        writer.finish().expect("zip file should finish");
    }

    fn relationship_snapshot(target: &str) -> RelationshipSnapshot {
        let source = PartUri::new("/word/document.xml").expect("source part URI should be valid");
        RelationshipSnapshot {
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                .to_string(),
            target: normalize_internal_relationship_target(Some(&source), target),
            target_mode: "Internal".to_string(),
        }
    }

    fn package_snapshot_with_part_relationship(
        relationship: RelationshipSnapshot,
    ) -> PackageSnapshot {
        PackageSnapshot {
            package_relationships: Vec::new(),
            parts: vec![PartSnapshot {
                uri: "/word/document.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                        .to_string(),
                ),
                relationships: vec![relationship],
                payload: digest_bytes(b"<w:document/>"),
            }],
        }
    }

    fn package_snapshot_from_raw_relationship(
        relationship: offidized_opc::Relationship,
    ) -> PackageSnapshot {
        let source = PartUri::new("/word/document.xml").expect("source part URI should be valid");
        let mut relationships =
            snapshot_relationships(std::iter::once(&relationship), Some(&source));
        relationships.sort();

        PackageSnapshot {
            package_relationships: Vec::new(),
            parts: vec![PartSnapshot {
                uri: "/word/document.xml".to_string(),
                content_type: Some(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                        .to_string(),
                ),
                relationships,
                payload: digest_bytes(b"<w:document/>"),
            }],
        }
    }

    #[test]
    fn zip_path_normalization_keeps_opc_snapshot_equivalent() {
        let left = TempZip::new("zip-path-left", &[("word/document.xml", b"<w:document/>")]);
        let right = TempZip::new(
            "zip-path-right",
            &[("word\\.\\sections\\..\\document.xml", b"<w:document/>")],
        );

        let (_, entry, opc) = compare_outputs(left.path(), right.path())
            .expect("zip fixtures should be readable and comparable");

        assert!(
            entry.equal,
            "entry comparison should normalize slash and backslash path separators"
        );
        assert!(
            opc.equal,
            "opc snapshot comparison should treat equivalent zip paths as the same part"
        );
    }

    #[test]
    fn relationship_target_normalization_treats_equivalent_paths_as_equal() {
        let left =
            package_snapshot_with_part_relationship(relationship_snapshot("media/image1.png"));
        let right = package_snapshot_with_part_relationship(relationship_snapshot(
            "./media/../media/image1.png",
        ));

        let comparison = compare_opc_snapshots(&left, &right);

        assert!(
            comparison.equal,
            "relationship targets with equivalent paths should compare equal"
        );
        assert_eq!(
            comparison.relationship_mismatches, 0,
            "no relationship mismatches are expected for normalized-equivalent targets"
        );
    }

    #[test]
    fn relationship_comparison_ignores_ids_when_other_fields_match() {
        let left = package_snapshot_from_raw_relationship(offidized_opc::Relationship {
            id: "rId1".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                .to_string(),
            target: "media/image1.png".to_string(),
            target_mode: TargetMode::Internal,
        });
        let right = package_snapshot_from_raw_relationship(offidized_opc::Relationship {
            id: "generatedId42".to_string(),
            rel_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                .to_string(),
            target: "media/image1.png".to_string(),
            target_mode: TargetMode::Internal,
        });

        let comparison = compare_opc_snapshots(&left, &right);

        assert!(
            comparison.equal,
            "relationship IDs should not affect comparison if type/target/mode are equivalent"
        );
        assert_eq!(
            comparison.relationship_mismatches, 0,
            "no relationship mismatches are expected when only relationship IDs differ"
        );
    }

    #[test]
    fn package_relationship_target_normalization_uses_package_root() {
        let source = PartUri::new("/word/document.xml").expect("source part URI should be valid");
        let relationship = offidized_opc::Relationship {
            id: "rId1".to_string(),
            rel_type: "type/x".to_string(),
            target: "./sections/../document.xml".to_string(),
            target_mode: TargetMode::Internal,
        };

        let part_relative = snapshot_relationships(std::iter::once(&relationship), Some(&source));
        let package_relative = snapshot_relationships(std::iter::once(&relationship), None);

        assert_eq!(
            part_relative[0].target, "/word/document.xml",
            "part-level relationships should resolve targets relative to the source part URI"
        );
        assert_eq!(
            package_relative[0].target, "/document.xml",
            "package-level relationships should resolve relative targets against package root"
        );
    }

    #[test]
    fn canonical_xml_digest_equates_format_and_attribute_order_variants() {
        let left = TempZip::new(
            "canonical-xml-left",
            &[(
                "doc.xml",
                br#"<root b="2" a="1"><child key="v">text</child></root>"#,
            )],
        );
        let right = TempZip::new(
            "canonical-xml-right",
            &[(
                "doc.xml",
                br#"<?xml version="1.0" encoding="UTF-8"?>
<!--ignored-->
<?processing ignored?>
<root a="1" b="2">
    <child key="v">text</child>
</root>"#,
            )],
        );

        let (_, entry, opc) = compare_outputs(left.path(), right.path())
            .expect("xml fixtures should be readable and comparable");

        assert!(
            entry.equal,
            "canonical XML digesting should ignore formatting and attribute-order differences"
        );
        assert!(
            opc.equal,
            "part payload digests should match for canonical-equivalent XML payloads"
        );
    }
}
