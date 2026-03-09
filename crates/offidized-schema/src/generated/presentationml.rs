#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PresentationmlElementDescriptor {
    pub schema_path: &'static str,
    pub class_name: &'static str,
    pub qualified_name: &'static str,
}

pub const ELEMENTS: &[PresentationmlElementDescriptor] = &[
    PresentationmlElementDescriptor {
        schema_path: "/p:presentation",
        class_name: "Presentation",
        qualified_name: "p:presentation",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:presentation/p:sldMasterIdLst",
        class_name: "SlideMasterIdList",
        qualified_name: "p:sldMasterIdLst",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:presentation/p:sldIdLst",
        class_name: "SlideIdList",
        qualified_name: "p:sldIdLst",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:presentation/p:sldIdLst/p:sldId",
        class_name: "SlideId",
        qualified_name: "p:sldId",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:sld",
        class_name: "Slide",
        qualified_name: "p:sld",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:sld/p:cSld",
        class_name: "CommonSlideData",
        qualified_name: "p:cSld",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:sld/p:cSld/p:spTree",
        class_name: "ShapeTree",
        qualified_name: "p:spTree",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:sld/p:cSld/p:spTree/p:sp",
        class_name: "Shape",
        qualified_name: "p:sp",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:sld/p:cSld/p:spTree/p:graphicFrame",
        class_name: "GraphicFrame",
        qualified_name: "p:graphicFrame",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:notes",
        class_name: "NotesSlide",
        qualified_name: "p:notes",
    },
    PresentationmlElementDescriptor {
        schema_path: "/p:notes/p:cSld",
        class_name: "NotesCommonSlideData",
        qualified_name: "p:cSld",
    },
];
