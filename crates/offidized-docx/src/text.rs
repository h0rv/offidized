/// Plain text value used by higher-level content objects.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Text {
    value: String,
}

impl Text {
    pub fn new(value: impl Into<String>) -> Self {
        Self {
            value: value.into(),
        }
    }

    pub fn as_str(&self) -> &str {
        &self.value
    }
}
