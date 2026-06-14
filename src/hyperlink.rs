//! Cell hyperlinks (`<hyperlinks>` worksheet element + external relationships).
//!
//! Mirrors openpyxl's `openpyxl.worksheet.hyperlink.Hyperlink`. A hyperlink is
//! attached to a cell and serializes to a `<hyperlink>` child of the worksheet
//! `<hyperlinks>` collection; external targets additionally require a sheet
//! relationship with `TargetMode="External"` (see [`crate::sheet_rels`]).
//!
//! This Cut targets external-URL hyperlinks. Internal `location` links and the
//! `tooltip`/`display` attributes are preserved on read-modify-write but are
//! not yet ergonomically authored — they are kept verbatim on the model so a
//! round-trip never loses them.

/// A cell hyperlink.
///
/// PyO3-ready: no borrowed fields, no lifetime parameters, all fields owned so
/// the struct is `'static` and can back a `#[pyclass]` directly.
///
/// `target` is the hyperlink destination (an external URL for external links).
/// `location`, `tooltip`, and `display` are optional and preserved verbatim
/// across read-modify-write.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Hyperlink {
    /// The hyperlink destination — an external URL for external links.
    pub target: String,
    /// Optional in-document location (e.g. `"Sheet1!A1"`); preserved verbatim.
    pub location: Option<String>,
    /// Optional hover tooltip text; preserved verbatim.
    pub tooltip: Option<String>,
    /// Optional display text override; preserved verbatim.
    pub display: Option<String>,
}

impl Hyperlink {
    /// Build an external-URL hyperlink with no `location`/`tooltip`/`display`.
    pub fn new_external(target: &str) -> Self {
        Self {
            target: target.to_string(),
            location: None,
            tooltip: None,
            display: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Hyperlink;

    #[test]
    fn new_external_sets_target_and_leaves_optional_fields_empty() {
        let link = Hyperlink::new_external("https://example.com");
        assert_eq!(link.target, "https://example.com");
        assert_eq!(link.location, None);
        assert_eq!(link.tooltip, None);
        assert_eq!(link.display, None);
    }

    #[test]
    fn preserves_optional_fields_when_set() {
        let link = Hyperlink {
            target: "https://rust-lang.org".to_string(),
            location: Some("Sheet1!A1".to_string()),
            tooltip: Some("Rust".to_string()),
            display: Some("The Rust site".to_string()),
        };
        assert_eq!(link.target, "https://rust-lang.org");
        assert_eq!(link.location.as_deref(), Some("Sheet1!A1"));
        assert_eq!(link.tooltip.as_deref(), Some("Rust"));
        assert_eq!(link.display.as_deref(), Some("The Rust site"));
    }
}
