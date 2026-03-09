//! Tests for SlideLayout and SlideMaster write API.
//!
//! These tests demonstrate creating masters and layouts from scratch,
//! modifying them, and saving presentations with custom masters.

use offidized_pptx::{Presentation, SlideLayout, SlideMaster};

#[test]
fn create_master_with_default_layout() {
    let master = SlideMaster::with_default_layout(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    assert_eq!(master.relationship_id(), "rId1");
    assert_eq!(master.part_uri(), "/ppt/slideMasters/slideMaster1.xml");
    assert_eq!(master.layouts().len(), 1);
    assert_eq!(master.layouts()[0].name(), "Title Slide");
}

#[test]
fn create_layout_with_title_placeholder() {
    let layout = SlideLayout::with_title_placeholder(
        "Title Slide",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    );

    assert_eq!(layout.name(), "Title Slide");
    assert_eq!(layout.layout_type(), Some("title"));
    assert_eq!(layout.shapes().len(), 1);
    assert_eq!(layout.shapes()[0].name(), "Title 1");
}

#[test]
fn create_layout_with_title_and_content() {
    let layout = SlideLayout::with_title_and_content(
        "Title and Content",
        "rId1",
        "/ppt/slideLayouts/slideLayout2.xml",
        "rId11",
    );

    assert_eq!(layout.name(), "Title and Content");
    assert_eq!(layout.layout_type(), Some("obj"));
    assert_eq!(layout.shapes().len(), 2);
    assert_eq!(layout.shapes()[0].name(), "Title 1");
    assert_eq!(layout.shapes()[1].name(), "Content Placeholder 2");
}

#[test]
fn add_placeholder_to_layout() {
    use offidized_pptx::PlaceholderType;

    let mut layout = SlideLayout::new(
        "Custom Layout",
        "rId1",
        "/ppt/slideLayouts/slideLayout3.xml",
        "rId12",
    );

    // Add a title placeholder
    let title_idx = layout.add_placeholder(PlaceholderType::Title, 100, 200, 8000, 1000);
    assert_eq!(title_idx, 0);

    // Add a content placeholder
    let content_idx = layout.add_placeholder(PlaceholderType::Object, 100, 1500, 8000, 4000);
    assert_eq!(content_idx, 1);

    assert_eq!(layout.shapes().len(), 2);
    assert_eq!(layout.shapes()[0].name(), "Placeholder 1");
    assert_eq!(layout.shapes()[1].name(), "Placeholder 2");
}

#[test]
fn modify_layout_properties() {
    let mut layout = SlideLayout::new(
        "Blank",
        "rId1",
        "/ppt/slideLayouts/slideLayout4.xml",
        "rId13",
    );

    // Initially not dirty (just created, hasn't been loaded from a file)
    assert!(!layout.is_dirty());

    // Modify name
    layout.set_name("Custom Blank");
    assert_eq!(layout.name(), "Custom Blank");
    assert!(layout.is_dirty());

    // Modify layout type
    layout.set_layout_type("blank");
    assert_eq!(layout.layout_type(), Some("blank"));

    // Modify preserve flag
    layout.set_preserve(true);
    assert!(layout.preserve());
}

#[test]
fn add_multiple_layouts_to_master() {
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    // Add title layout
    let title_layout = SlideLayout::with_title_placeholder(
        "Title Slide",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    );
    master.add_layout(title_layout);

    // Add title and content layout
    let content_layout = SlideLayout::with_title_and_content(
        "Title and Content",
        "rId2",
        "/ppt/slideLayouts/slideLayout2.xml",
        "rId11",
    );
    master.add_layout(content_layout);

    // Add blank layout
    let blank_layout = SlideLayout::new(
        "Blank",
        "rId3",
        "/ppt/slideLayouts/slideLayout3.xml",
        "rId12",
    );
    master.add_layout(blank_layout);

    assert_eq!(master.layouts().len(), 3);
    assert_eq!(master.layouts()[0].name(), "Title Slide");
    assert_eq!(master.layouts()[1].name(), "Title and Content");
    assert_eq!(master.layouts()[2].name(), "Blank");
}

#[test]
fn create_presentation_with_custom_master() {
    let mut prs = Presentation::new();

    // Create a master with two layouts
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    let title_layout = SlideLayout::with_title_placeholder(
        "Title Slide",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    );
    master.add_layout(title_layout);

    let content_layout = SlideLayout::with_title_and_content(
        "Title and Content",
        "rId2",
        "/ppt/slideLayouts/slideLayout2.xml",
        "rId11",
    );
    master.add_layout(content_layout);

    // Add master to presentation
    prs.add_slide_master(master);

    // Verify it was added
    assert_eq!(prs.slide_masters_v2().len(), 1);
    assert_eq!(prs.slide_masters_v2()[0].layouts().len(), 2);

    // Add a slide
    prs.add_slide_with_title("Test Slide");
    assert_eq!(prs.slide_count(), 1);
}

#[test]
fn save_presentation_with_custom_master() {
    let mut prs = Presentation::new();

    // Create a master with a default layout
    let master = SlideMaster::with_default_layout(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );
    prs.add_slide_master(master);

    // Add slides
    prs.add_slide_with_title("Slide 1");
    prs.add_slide_with_title("Slide 2");

    // Save to a temp file
    let temp_file = std::env::temp_dir().join("test_master_write.pptx");
    prs.save(&temp_file).expect("should save");

    // Verify file exists
    assert!(temp_file.exists());

    // Reopen and verify structure
    let prs2 = Presentation::open(&temp_file).expect("should reopen");
    assert_eq!(prs2.slide_count(), 2);
    // Note: slide_masters_v2 won't be populated on load yet (that's a future enhancement)
    // but slide_masters (legacy) should have the master
    assert!(!prs2.slide_masters().is_empty());

    // Clean up
    let _ = std::fs::remove_file(&temp_file);
}

#[test]
fn modify_existing_master() {
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    // Add initial layout
    let layout = SlideLayout::with_title_placeholder(
        "Title Slide",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    );
    master.add_layout(layout);

    // Modify the layout
    if let Some(layout) = master.layout_mut(0) {
        layout.set_name("Modified Title");
        layout.set_preserve(true);
    }

    // Verify changes
    assert_eq!(master.layouts()[0].name(), "Modified Title");
    assert!(master.layouts()[0].preserve());
    assert!(master.is_dirty());
}

#[test]
fn remove_layout_from_master() {
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    // Add three layouts
    for i in 1..=3 {
        let layout = SlideLayout::new(
            format!("Layout {i}"),
            "rId1",
            format!("/ppt/slideLayouts/slideLayout{i}.xml"),
            format!("rId1{i}"),
        );
        master.add_layout(layout);
    }

    assert_eq!(master.layouts().len(), 3);

    // Remove middle layout
    let removed = master.remove_layout(1);
    assert!(removed.is_some());
    assert_eq!(removed.unwrap().name(), "Layout 2");

    assert_eq!(master.layouts().len(), 2);
    assert_eq!(master.layouts()[0].name(), "Layout 1");
    assert_eq!(master.layouts()[1].name(), "Layout 3");
}

#[test]
fn presentation_layout_access() {
    let mut prs = Presentation::new();

    // Create a master with two layouts
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    master.add_layout(SlideLayout::new(
        "Layout 1",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    ));

    master.add_layout(SlideLayout::new(
        "Layout 2",
        "rId2",
        "/ppt/slideLayouts/slideLayout2.xml",
        "rId11",
    ));

    prs.add_slide_master(master);

    // Access layout through presentation
    let layout = prs.layout(0, 0).expect("should have layout");
    assert_eq!(layout.name(), "Layout 1");

    let layout2 = prs.layout(0, 1).expect("should have layout 2");
    assert_eq!(layout2.name(), "Layout 2");

    // Modify layout through presentation
    if let Some(layout_mut) = prs.layout_mut(0, 0) {
        layout_mut.set_name("Modified Layout");
    }

    assert_eq!(prs.layout(0, 0).unwrap().name(), "Modified Layout");
}

#[test]
fn apply_layout_to_slide() {
    let mut prs = Presentation::new();

    // Create a master with three different layouts
    let mut master = SlideMaster::new(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );

    master.add_layout(SlideLayout::with_title_placeholder(
        "Title Slide",
        "rId1",
        "/ppt/slideLayouts/slideLayout1.xml",
        "rId10",
    ));

    master.add_layout(SlideLayout::with_title_and_content(
        "Title and Content",
        "rId2",
        "/ppt/slideLayouts/slideLayout2.xml",
        "rId11",
    ));

    master.add_layout(SlideLayout::new(
        "Blank",
        "rId3",
        "/ppt/slideLayouts/slideLayout3.xml",
        "rId12",
    ));

    prs.add_slide_master(master);

    // Add slides and apply different layouts
    let slide1 = prs.add_slide_with_title("Title Slide");
    slide1.apply_layout(0, 0); // Use "Title Slide" layout

    let slide2 = prs.add_slide_with_title("Content Slide");
    slide2.apply_layout(0, 1); // Use "Title and Content" layout

    let slide3 = prs.add_slide_with_title("Blank Slide");
    slide3.apply_layout(0, 2); // Use "Blank" layout

    // Verify layout references
    assert_eq!(prs.slides()[0].layout_reference(), Some((0, 0)));
    assert_eq!(prs.slides()[1].layout_reference(), Some((0, 1)));
    assert_eq!(prs.slides()[2].layout_reference(), Some((0, 2)));

    // Save and verify
    let temp_file = std::env::temp_dir().join("test_apply_layout.pptx");
    prs.save(&temp_file).expect("should save");
    assert!(temp_file.exists());

    // Clean up
    let _ = std::fs::remove_file(&temp_file);
}

#[test]
fn clear_layout_from_slide() {
    let mut prs = Presentation::new();

    // Create a master with a layout
    let master = SlideMaster::with_default_layout(
        "rId1".to_string(),
        "/ppt/slideMasters/slideMaster1.xml".to_string(),
    );
    prs.add_slide_master(master);

    // Add slide and apply layout
    let slide = prs.add_slide_with_title("Test");
    slide.apply_layout(0, 0);
    assert_eq!(slide.layout_reference(), Some((0, 0)));

    // Clear layout
    slide.clear_layout();
    assert_eq!(slide.layout_reference(), None);
}
