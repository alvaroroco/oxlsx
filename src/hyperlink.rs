#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Hyperlink {
    target: String,
}

impl Hyperlink {
    pub fn new(target: impl Into<String>) -> Self {
        Self {
            target: target.into(),
        }
    }

    pub fn target(&self) -> &str {
        &self.target
    }
}

impl From<&str> for Hyperlink {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

impl From<String> for Hyperlink {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

#[cfg(test)]
mod tests {
    use super::Hyperlink;

    #[test]
    fn hyperlink_model_round_trip() {
        let source = "https://example.com/docs";
        let hyperlink = Hyperlink::new(source);

        assert_eq!(hyperlink.target(), source);
        assert_eq!(hyperlink, Hyperlink::from(source));
        assert_eq!(hyperlink, Hyperlink::from(source.to_string()));
    }
}
