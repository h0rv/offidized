//! # offidized-pptx
//!
//! High-level PowerPoint API inspired by python-pptx.
//!
//! ```ignore
//! use offidized_pptx::Presentation;
//!
//! let mut prs = Presentation::new();
//! prs.add_slide_with_title("Q4 Results");
//! prs.add_slide_with_title("Summary");
//! prs.save("results.pptx")?;
//! ```

pub mod actions;
pub mod animation_advanced;
pub mod chart;
pub mod chart_template;
pub mod chart_types;
pub mod color;
pub mod color_resolver;
pub mod comment;
pub mod custom_show;
pub mod error;
pub mod image;
pub mod layout_inheritance;
pub mod media_controls;
pub mod placeholder_editor;
pub mod presentation;
pub mod presentation_properties;
pub mod shape;
pub mod slide;
pub mod slide_layout;
pub mod slide_layout_io;
pub mod slide_master;
pub mod slide_master_io;
pub mod smartart;
pub mod table;
pub mod table_style;
pub mod text;
pub mod theme;
pub mod timing;
pub mod transition;
pub mod vba_preserve;

pub use actions::{ActionButtonType, ActionType, EmbeddedObject, ShapeAction};
pub use animation_advanced::{
    AnimationEffect, AnimationEffectType, AnimationRestart, AnimationSequence, AnimationTiming,
    AnimationTrigger,
};
pub use chart::{
    Chart, ChartAxis, ChartDataLabel, ChartSeries, ChartType, LegendPosition, LineStyle,
    MarkerShape, ScatterStyle, SeriesBorder, SeriesFill, SeriesMarker,
};
pub use chart_template::{ChartStyle, ChartTemplate, ChartTemplateRegistry};
pub use chart_types::{
    Area3DChart, Bar3DChart, BubbleChart, Column3DChart, CombinationChart, Line3DChart, Pie3DChart,
    StockChart, SurfaceChart,
};
pub use color::{ColorTransform, ShapeColor};
pub use color_resolver::resolve_scheme_color;
pub use comment::SlideComment;
pub use custom_show::{CustomShow, SlideShowSettings, SlideShowType};
pub use error::{PptxError, Result};
pub use image::{Image, ImageCrop};
pub use layout_inheritance::{
    find_placeholder_in_shapes, parse_placeholder_from_xml, InheritanceResolver,
    ReferencedPlaceholder, ResolvedBackground, ResolvedFont, ResolvedTransform,
};
pub use media_controls::{
    AudioStartMode, AudioType, MediaContent, MediaPlaybackControls, VideoType,
};
pub use placeholder_editor::{
    PlaceholderContent, PlaceholderEditor, PlaceholderFactory, PlaceholderInheritance,
};
pub use presentation::Presentation;
pub use presentation_properties::PresentationProperties;
pub use shape::{
    ArrowSize, ArrowType, AutoFitType, BulletProperties, BulletStyle, ConnectionInfo, GradientFill,
    GradientFillType, GradientStop, LineArrow, LineCompoundStyle, LineDashStyle, LineSpacing,
    LineSpacingUnit, MediaType, ParagraphProperties, PatternFill, PatternFillType, PictureFill,
    PlaceholderType, Shape, ShapeFill, ShapeGeometry, ShapeGlow, ShapeOutline, ShapeParagraph,
    ShapeReflection, ShapeShadow, ShapeType, SpacingUnit, SpacingValue, TextAlignment, TextAnchor,
};
pub use slide::{PresentationSection, ShapeGroup, Slide, SlideBackground, SlideHeaderFooter};
pub use slide_layout::SlideLayout;
pub use slide_master::SlideMaster;
pub use smartart::{SmartArt, SmartArtNode, SmartArtType};
pub use table::{CellBorder, CellBorders, CellTextAnchor, Table, TableCell, TextDirection};
pub use table_style::{
    get_builtin_style, get_style_by_guid, parse_table_style_id, parse_table_style_options,
    write_table_style_id, write_table_style_options, TableStyle, TableStyleOptions, TableStyleType,
};
pub use text::{RunProperties, StrikethroughStyle, TextRun, UnderlineStyle};
pub use theme::{ThemeColorRef, ThemeColorScheme, ThemeFontScheme};
pub use timing::{SlideAnimationNode, SlideTiming};
pub use transition::{SlideTransition, SlideTransitionKind, TransitionSound, TransitionSpeed};
pub use vba_preserve::{VbaMacroContainer, VbaProject, VbaSignature};
