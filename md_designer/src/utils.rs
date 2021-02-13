use anyhow::{anyhow, Result};
use log::info;
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

pub fn get_output_filename(filename: &str) -> Result<&str> {
    if filename.is_empty() {
        Err(anyhow!("output filename is empty."))
    } else {
        let result = if let Some(stripped) = filename.strip_suffix(".xlsx") {
            stripped
        } else {
            filename
        };
        info!("output filename without extension: {}", result);
        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_output_filename() {
        assert_eq!("output", get_output_filename("output").unwrap());
        assert_eq!("output", get_output_filename("output.xlsx").unwrap());
    }

    #[test]
    fn test_get_output_filename_error() {
        assert!(get_output_filename("").is_err());
    }
}
