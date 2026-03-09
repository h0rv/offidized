use std::io::Write as _;
use std::path::Path;

use anyhow::{bail, Result};
use offidized_opc::Package;

use crate::output::PartInfo;

pub fn run(path: &Path, uri: Option<&str>, list: bool) -> Result<()> {
    let package = Package::open(path)?;

    if list {
        let parts: Vec<PartInfo> = package
            .parts()
            .map(|p| PartInfo {
                uri: p.uri.as_str().to_string(),
                content_type: p.content_type.clone(),
                size_bytes: p.data.len(),
            })
            .collect();
        println!("{}", serde_json::to_string(&parts)?);
        return Ok(());
    }

    let Some(uri) = uri else {
        bail!("specify a part URI or use --list to list all parts");
    };

    let part = package
        .get_part(uri)
        .ok_or_else(|| anyhow::anyhow!("part not found: {uri:?}"))?;

    std::io::stdout().write_all(part.data.as_bytes())?;
    Ok(())
}
