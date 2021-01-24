use pulldown_cmark::Tag;

pub fn cmarktag_stringify(tag: &Tag<'_>) -> Option<String> {
    match tag {
        Tag::Heading(idx) => Some(format!("Heading{}", idx)),
        Tag::List(_) => Some("List".to_string()),
        _ => None,
    }
}
