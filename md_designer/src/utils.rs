use anyhow::{anyhow, Result};
use pulldown_cmark::Tag;

pub fn string_to_cmarktag(input: &str) -> Result<Tag<'_>> {
    match input {
        "Heading1" => Ok(Tag::Heading(1)),
        "Heading2" => Ok(Tag::Heading(2)),
        "Heading3" => Ok(Tag::Heading(3)),
        "Heading4" => Ok(Tag::Heading(4)),
        "Heading5" => Ok(Tag::Heading(5)),
        "Heading6" => Ok(Tag::Heading(6)),
        "Heading7" => Ok(Tag::Heading(7)),
        "Heading8" => Ok(Tag::Heading(8)),
        "List" => Ok(Tag::List(None)),
        _ => Err(anyhow!("input string is malformed")),
    }
}
