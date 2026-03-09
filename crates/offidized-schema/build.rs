use std::env;
use std::path::PathBuf;

fn main() -> anyhow::Result<()> {
    run().map_err(|err| err.context("failed to generate schema catalogs"))
}

fn run() -> anyhow::Result<()> {
    let manifest_dir = PathBuf::from(env::var("CARGO_MANIFEST_DIR")?);
    let default_data_root = manifest_dir.join("../../references/Open-XML-SDK/data");
    let data_root = env::var_os("OFFIDIZED_OPENXML_DATA_ROOT")
        .map(PathBuf::from)
        .unwrap_or(default_data_root);

    if !data_root.exists() {
        anyhow::bail!(
            "OpenXML data root does not exist: {} (set OFFIDIZED_OPENXML_DATA_ROOT)",
            data_root.display()
        );
    }

    let out_dir = PathBuf::from(env::var("OUT_DIR")?);
    offidized_codegen::generate_from_data_root(&data_root, &out_dir)?;

    println!("cargo:rerun-if-env-changed=OFFIDIZED_OPENXML_DATA_ROOT");
    println!("cargo:rerun-if-changed={}", data_root.display());

    Ok(())
}
