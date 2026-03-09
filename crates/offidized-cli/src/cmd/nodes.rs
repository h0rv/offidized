use std::path::Path;

use anyhow::Result;
use offidized_ir::derive_content_nodes;
use serde::Serialize;

#[derive(Debug, Serialize)]
struct NodeOut {
    id: String,
    kind: String,
    text: String,
}

pub fn run(file: &Path) -> Result<()> {
    let nodes = derive_content_nodes(file)?;
    let out: Vec<NodeOut> = nodes
        .into_iter()
        .map(|node| NodeOut {
            id: node.id_string(),
            kind: format!("{:?}", node.kind),
            text: node.text,
        })
        .collect();
    println!("{}", serde_json::to_string_pretty(&out)?);
    Ok(())
}
