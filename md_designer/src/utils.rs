use pulldown_cmark::Tag;

use crate::constant::CUSTOM_PREFIX_KEY;

pub fn cmarktag_stringify(tag: &Tag<'_>) -> Option<String> {
    match tag {
        Tag::Heading(idx) => Some(format!("Heading{}", idx)),
        Tag::List(_) => Some("List".to_string()),
        _ => None,
    }
}

pub fn get_custom_prefix_key(prefix: &str) -> String {
    format!("{} {}", CUSTOM_PREFIX_KEY.clone(), prefix)
}

pub fn get_custom_prefix_as_normal_list(prefix: &str) -> String {
    format!("* !!!{}{}", CUSTOM_PREFIX_KEY.clone(), prefix)
}

pub fn custom_prefix_to_key(text_with_custom_prefix: Option<&str>) -> Option<String> {
    if let Some(text) = text_with_custom_prefix {
        if let Some(stripped) = text.strip_prefix(&format!("!!!{}", CUSTOM_PREFIX_KEY.clone())) {
            if let Some(prefix) = stripped.chars().next() {
                return Some(get_custom_prefix_key(&prefix.to_string()));
            }
        }
    }
    None
}

pub fn get_custom_prefix_end_idx() -> usize {
    format!("!!!{}{}", CUSTOM_PREFIX_KEY.clone(), "a").len() + 1
}
