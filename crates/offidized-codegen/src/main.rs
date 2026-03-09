use std::env;
use std::ffi::OsString;
use std::path::PathBuf;

use anyhow::{bail, Context, Result};

fn main() {
    if let Err(error) = run() {
        eprintln!("error: {error:#}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let mut args = env::args_os();
    let program = args
        .next()
        .unwrap_or_else(|| OsString::from("offidized-codegen"));

    let data_root = next_required_arg(&mut args, usage(program.to_string_lossy().as_ref()))?;
    let output_dir = next_required_arg(&mut args, usage(program.to_string_lossy().as_ref()))?;

    if args.next().is_some() {
        bail!("{}", usage(program.to_string_lossy().as_ref()));
    }

    let generated = offidized_codegen::generate_from_data_root(
        PathBuf::from(data_root),
        PathBuf::from(output_dir),
    )
    .context("generation failed")?;

    for file in generated.files {
        println!("{}", file.display());
    }

    Ok(())
}

fn next_required_arg(args: &mut impl Iterator<Item = OsString>, usage: String) -> Result<OsString> {
    args.next().with_context(|| usage)
}

fn usage(program: &str) -> String {
    format!("usage: {program} <openxml-sdk-data-root> <output-dir>")
}
