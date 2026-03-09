#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SharedElementDescriptor {
    pub schema_path: &'static str,
    pub class_name: &'static str,
    pub qualified_name: &'static str,
}

pub const ELEMENTS: &[SharedElementDescriptor] = &[
    SharedElementDescriptor {
        schema_path: "/a:txBody",
        class_name: "TextBody",
        qualified_name: "a:txBody",
    },
    SharedElementDescriptor {
        schema_path: "/a:txBody/a:p",
        class_name: "DrawingParagraph",
        qualified_name: "a:p",
    },
    SharedElementDescriptor {
        schema_path: "/a:txBody/a:p/a:r",
        class_name: "DrawingRun",
        qualified_name: "a:r",
    },
    SharedElementDescriptor {
        schema_path: "/a:txBody/a:p/a:r/a:t",
        class_name: "DrawingText",
        qualified_name: "a:t",
    },
    SharedElementDescriptor {
        schema_path: "/a:blip",
        class_name: "Blip",
        qualified_name: "a:blip",
    },
    SharedElementDescriptor {
        schema_path: "/a:stretch",
        class_name: "Stretch",
        qualified_name: "a:stretch",
    },
    SharedElementDescriptor {
        schema_path: "/a:xfrm",
        class_name: "Transform2D",
        qualified_name: "a:xfrm",
    },
    SharedElementDescriptor {
        schema_path: "/a:prstGeom",
        class_name: "PresetGeometry",
        qualified_name: "a:prstGeom",
    },
    SharedElementDescriptor {
        schema_path: "/mc:AlternateContent",
        class_name: "AlternateContent",
        qualified_name: "mc:AlternateContent",
    },
    SharedElementDescriptor {
        schema_path: "/v:shape",
        class_name: "VmlShape",
        qualified_name: "v:shape",
    },
];
