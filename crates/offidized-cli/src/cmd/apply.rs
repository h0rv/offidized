use std::path::{Path, PathBuf};

use anyhow::{bail, Result};
use offidized_ir::ApplyOptions;

/// Run the `apply` subcommand.
pub fn run(
    ir_file: Option<&Path>,
    output: Option<&Path>,
    in_place: bool,
    source_override: Option<&Path>,
    force: bool,
    dry_run: bool,
) -> Result<()> {
    // Read IR from file or stdin
    let ir = if let Some(path) = ir_file {
        std::fs::read_to_string(path)?
    } else {
        use std::io::Read;
        let mut buf = String::new();
        std::io::stdin().read_to_string(&mut buf)?;
        buf
    };

    // Parse header to get source path for in-place resolution
    let (header, _) = offidized_ir::IrHeader::parse(&ir).map_err(|e| anyhow::anyhow!("{e}"))?;

    let source = source_override
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from(&header.source));

    // Resolve output path
    let dest = match (output, in_place) {
        (Some(_), true) => bail!("cannot specify both -o and -i"),
        (Some(o), false) => o.to_path_buf(),
        (None, true) => source,
        (None, false) => bail!("must specify -o <path> or -i for in-place edit"),
    };

    let options = ApplyOptions {
        source_override: source_override.map(PathBuf::from),
        force,
    };

    if dry_run {
        eprintln!("dry-run: would apply changes to {}", dest.display());
        eprintln!("(full dry-run diff not yet implemented)");
        return Ok(());
    }

    let result = offidized_ir::apply(&ir, &dest, &options).map_err(|e| anyhow::anyhow!("{e}"))?;

    // Report results to stderr (stdout may be piped)
    eprintln!(
        "applied: {} updated, {} created, {} cleared",
        result.cells_updated, result.cells_created, result.cells_cleared,
    );
    for warning in &result.warnings {
        eprintln!("warning: {warning}");
    }

    Ok(())
}
