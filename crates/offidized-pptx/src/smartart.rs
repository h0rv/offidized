// ── SmartArt / Diagram support ──
//
// SmartArt graphics in OOXML are represented as `p:graphicFrame` elements
// whose `a:graphicData` child has URI `http://schemas.openxmlformats.org/drawingml/2006/diagram`.
// Inside, a `dgm:relIds` element references four external parts:
//   r:dm  – diagram data     (dgm+xml)
//   r:lo  – diagram layout   (dgm+xml)
//   r:qs  – diagram style    (dgm+xml)
//   r:cs  – diagram colors   (dgm+xml)
// Plus an optional extended part for the diagram drawing itself.
//
// This module provides domain types for SmartArt metadata.  The primary goal
// is **roundtrip preservation**: we detect SmartArt graphic frames during parse,
// store the relationship IDs so the underlying parts survive save, and expose
// enough structure for callers to inspect (and eventually edit) the diagram.

/// Namespace URI used in `a:graphicData` to identify SmartArt/diagram content.
pub const DIAGRAM_GRAPHIC_DATA_URI: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/diagram";

/// Relationship type for the diagram drawing extended part.
pub const DIAGRAM_DRAWING_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";

/// Content type for the diagram drawing part.
pub const DIAGRAM_DRAWING_CONTENT_TYPE: &str =
    "application/vnd.ms-office.drawingml.diagramDrawing+xml";

// ── SmartArt type classification ──

/// Known SmartArt layout families.
///
/// OOXML defines dozens of diagram layouts.  We enumerate the most common ones
/// and fall back to `Other` for anything we don't explicitly recognise.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SmartArtType {
    /// Basic block list (`urn:microsoft.com/office/officeart/2005/8/layout/vList1` etc.).
    BasicBlockList,
    /// Basic process (linear left-to-right steps).
    BasicProcess,
    /// Continuous block process.
    ContinuousBlockProcess,
    /// Organization chart / hierarchy.
    OrgChart,
    /// Hierarchy layout.
    Hierarchy,
    /// Radial / hub-and-spoke.
    Radial,
    /// Venn diagram.
    Venn,
    /// Pyramid.
    Pyramid,
    /// Cycle diagram.
    Cycle,
    /// Matrix / grid.
    Matrix,
    /// Unrecognised layout – stores the raw layout URI or name.
    Other(String),
}

// ── Relationship IDs ──

/// The four (optionally five) relationship IDs that link a SmartArt graphic
/// frame to its external diagram parts.
///
/// These are stored as opaque `String` values and are written back verbatim
/// during serialization to preserve roundtrip fidelity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartArtRelIds {
    /// `r:dm` – relationship ID for the diagram **data** part.
    pub data_rel_id: String,
    /// `r:lo` – relationship ID for the diagram **layout** part.
    pub layout_rel_id: String,
    /// `r:qs` – relationship ID for the diagram **quick style** part.
    pub style_rel_id: String,
    /// `r:cs` – relationship ID for the diagram **colors** part.
    pub colors_rel_id: String,
}

impl SmartArtRelIds {
    /// Create a new set of SmartArt relationship IDs.
    pub fn new(
        data_rel_id: impl Into<String>,
        layout_rel_id: impl Into<String>,
        style_rel_id: impl Into<String>,
        colors_rel_id: impl Into<String>,
    ) -> Self {
        Self {
            data_rel_id: data_rel_id.into(),
            layout_rel_id: layout_rel_id.into(),
            style_rel_id: style_rel_id.into(),
            colors_rel_id: colors_rel_id.into(),
        }
    }
}

// ── SmartArt node ──

/// A single node within a SmartArt diagram.
///
/// Nodes form a tree: each node may have children, producing hierarchical
/// diagrams (org charts, pyramids, etc.).
///
/// Modelled after ShapeCrawler's `SmartArtNode` – each node carries a stable
/// `model_id` used to correlate entries across the data, layout, and drawing
/// parts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartArtNode {
    /// Stable identifier within the diagram data part (e.g. `"1"`, `"p1"`).
    pub model_id: String,
    /// Display text of the node.
    pub text: String,
    /// Nesting depth (0 = top-level).
    pub level: u32,
    /// Child nodes.
    pub children: Vec<SmartArtNode>,
}

impl SmartArtNode {
    /// Create a new top-level node with the given text.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            model_id: String::new(),
            text: text.into(),
            level: 0,
            children: Vec::new(),
        }
    }

    /// Create a node with explicit model ID, text, and level.
    pub fn with_id(model_id: impl Into<String>, text: impl Into<String>, level: u32) -> Self {
        Self {
            model_id: model_id.into(),
            text: text.into(),
            level,
            children: Vec::new(),
        }
    }

    /// Add a child node and return `&mut Self` for chaining.
    pub fn add_child(&mut self, child: SmartArtNode) -> &mut Self {
        self.children.push(child);
        self
    }

    /// Recursively count this node plus all descendants.
    pub fn descendant_count(&self) -> usize {
        1 + self
            .children
            .iter()
            .map(SmartArtNode::descendant_count)
            .sum::<usize>()
    }
}

// ── SmartArt (top-level) ──

/// High-level representation of a SmartArt graphic embedded in a slide.
///
/// Stores the diagram metadata, relationship IDs required for roundtrip, and
/// the logical node tree extracted from the diagram data part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SmartArt {
    /// Classified diagram type (best-effort; falls back to `Other`).
    pub diagram_type: SmartArtType,
    /// Logical nodes of the diagram.
    pub nodes: Vec<SmartArtNode>,
    /// Relationship IDs linking back to the four diagram parts.
    pub rel_ids: SmartArtRelIds,
    /// Optional style name / quick-style identifier.
    pub style: Option<String>,
    /// Optional color scheme name.
    pub color_scheme: Option<String>,
}

impl SmartArt {
    /// Create a new `SmartArt` with the given relationship IDs and no nodes.
    pub fn new(rel_ids: SmartArtRelIds) -> Self {
        Self {
            diagram_type: SmartArtType::Other(String::new()),
            nodes: Vec::new(),
            rel_ids,
            style: None,
            color_scheme: None,
        }
    }

    /// Total number of top-level nodes.
    pub fn node_count(&self) -> usize {
        self.nodes.len()
    }

    /// Total number of nodes including all descendants.
    pub fn total_node_count(&self) -> usize {
        self.nodes.iter().map(SmartArtNode::descendant_count).sum()
    }

    /// Add a top-level node.
    pub fn add_node(&mut self, node: SmartArtNode) -> &mut Self {
        self.nodes.push(node);
        self
    }

    /// Iterate over the top-level nodes.
    pub fn iter_nodes(&self) -> impl Iterator<Item = &SmartArtNode> {
        self.nodes.iter()
    }
}

// ── Detection helper ──

/// Returns `true` if the given `a:graphicData` URI indicates SmartArt content.
///
/// This is the primary detection mechanism used during shape parsing: when a
/// `p:graphicFrame` contains `a:graphicData` with the diagram URI, the shape
/// should be treated as SmartArt.
pub fn is_smartart_graphic_data_uri(uri: &str) -> bool {
    uri == DIAGRAM_GRAPHIC_DATA_URI
}

// ── Tests ──

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn smartart_node_new() {
        let node = SmartArtNode::new("Hello");
        assert_eq!(node.text, "Hello");
        assert_eq!(node.level, 0);
        assert!(node.children.is_empty());
        assert!(node.model_id.is_empty());
    }

    #[test]
    fn smartart_node_with_id() {
        let node = SmartArtNode::with_id("p1", "First", 0);
        assert_eq!(node.model_id, "p1");
        assert_eq!(node.text, "First");
        assert_eq!(node.level, 0);
    }

    #[test]
    fn smartart_node_children() {
        let mut parent = SmartArtNode::with_id("p1", "Parent", 0);
        let child = SmartArtNode::with_id("p2", "Child", 1);
        parent.add_child(child);

        assert_eq!(parent.children.len(), 1);
        assert_eq!(parent.children[0].text, "Child");
        assert_eq!(parent.children[0].level, 1);
    }

    #[test]
    fn smartart_node_descendant_count() {
        let mut root = SmartArtNode::new("Root");
        let mut child1 = SmartArtNode::new("Child1");
        child1.add_child(SmartArtNode::new("Grandchild"));
        root.add_child(child1);
        root.add_child(SmartArtNode::new("Child2"));

        // root(1) + child1(1) + grandchild(1) + child2(1) = 4
        assert_eq!(root.descendant_count(), 4);
    }

    #[test]
    fn smartart_rel_ids() {
        let rel_ids = SmartArtRelIds::new("rId1", "rId2", "rId3", "rId4");
        assert_eq!(rel_ids.data_rel_id, "rId1");
        assert_eq!(rel_ids.layout_rel_id, "rId2");
        assert_eq!(rel_ids.style_rel_id, "rId3");
        assert_eq!(rel_ids.colors_rel_id, "rId4");
    }

    #[test]
    fn smartart_new_and_add_nodes() {
        let rel_ids = SmartArtRelIds::new("rId1", "rId2", "rId3", "rId4");
        let mut sa = SmartArt::new(rel_ids);

        assert_eq!(sa.node_count(), 0);
        assert_eq!(sa.total_node_count(), 0);

        sa.add_node(SmartArtNode::new("A"));
        sa.add_node(SmartArtNode::new("B"));

        assert_eq!(sa.node_count(), 2);
        assert_eq!(sa.total_node_count(), 2);
    }

    #[test]
    fn smartart_total_node_count_with_hierarchy() {
        let rel_ids = SmartArtRelIds::new("rId1", "rId2", "rId3", "rId4");
        let mut sa = SmartArt::new(rel_ids);

        let mut parent = SmartArtNode::with_id("p1", "CEO", 0);
        parent.add_child(SmartArtNode::with_id("p2", "VP1", 1));
        parent.add_child(SmartArtNode::with_id("p3", "VP2", 1));
        sa.add_node(parent);
        sa.add_node(SmartArtNode::new("Independent"));

        // CEO(1) + VP1(1) + VP2(1) + Independent(1) = 4
        assert_eq!(sa.total_node_count(), 4);
        assert_eq!(sa.node_count(), 2);
    }

    #[test]
    fn smartart_iter_nodes() {
        let rel_ids = SmartArtRelIds::new("rId1", "rId2", "rId3", "rId4");
        let mut sa = SmartArt::new(rel_ids);
        sa.add_node(SmartArtNode::new("X"));
        sa.add_node(SmartArtNode::new("Y"));

        let texts: Vec<&str> = sa.iter_nodes().map(|n| n.text.as_str()).collect();
        assert_eq!(texts, vec!["X", "Y"]);
    }

    #[test]
    fn smartart_type_variants() {
        let t = SmartArtType::BasicBlockList;
        assert_eq!(t, SmartArtType::BasicBlockList);

        let other = SmartArtType::Other("customLayout".into());
        assert_eq!(other, SmartArtType::Other("customLayout".into()));
        assert_ne!(other, SmartArtType::Venn);
    }

    #[test]
    fn is_smartart_uri_detection() {
        assert!(is_smartart_graphic_data_uri(DIAGRAM_GRAPHIC_DATA_URI));
        assert!(!is_smartart_graphic_data_uri(
            "http://schemas.openxmlformats.org/drawingml/2006/chart"
        ));
        assert!(!is_smartart_graphic_data_uri(
            "http://schemas.openxmlformats.org/drawingml/2006/table"
        ));
        assert!(!is_smartart_graphic_data_uri(""));
    }

    #[test]
    fn smartart_style_and_color_scheme() {
        let rel_ids = SmartArtRelIds::new("rId1", "rId2", "rId3", "rId4");
        let mut sa = SmartArt::new(rel_ids);
        sa.style = Some("SimpleFill".into());
        sa.color_scheme = Some("Colorful".into());

        assert_eq!(sa.style.as_deref(), Some("SimpleFill"));
        assert_eq!(sa.color_scheme.as_deref(), Some("Colorful"));
    }
}
