#[derive(Debug, thiserror::Error)]
pub enum OxlsxError {
    #[error("io: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("invalid xlsx format: {0}")]
    InvalidFormat(String),
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("duplicate table displayName: {0}")]
    DuplicateTableName(String),
    #[error("operation intersects table range: {0}")]
    TableIntersection(String),
    #[error("move_range with translate=true is not supported: formula translation is out of scope")]
    UnsupportedTranslate,
}

#[cfg(test)]
mod tests {
    use super::OxlsxError;

    #[test]
    fn maps_from_io_error() {
        let err: OxlsxError = std::io::Error::other("boom").into();
        assert!(matches!(err, OxlsxError::Io(_)));
    }

    #[test]
    fn renders_missing_part_message() {
        let err = OxlsxError::MissingPart("xl/styles.xml".to_string());
        assert_eq!(err.to_string(), "missing part: xl/styles.xml");
    }

    #[test]
    fn renders_table_intersection_message() {
        let err = OxlsxError::TableIntersection("A1:C3".to_string());
        assert_eq!(err.to_string(), "operation intersects table range: A1:C3");
    }

    #[test]
    fn renders_unsupported_translate_message() {
        let err = OxlsxError::UnsupportedTranslate;
        assert_eq!(
            err.to_string(),
            "move_range with translate=true is not supported: formula translation is out of scope"
        );
    }
}
