use std::collections::HashMap;

use anyhow::Result;
use pulldown_cmark::Tag;

use crate::{
    constant::AUTO_INCREMENT_KEY,
    rule::Rule,
    utils::{cmarktag_stringify, custom_prefix_to_key, get_custom_prefix_key},
};

#[derive(Debug, PartialEq)]
pub struct Mapping {
    blocks: Vec<Block>,
}

impl Mapping {
    pub fn new(rule: &Rule) -> Result<Self> {
        let mut blocks = vec![];
        rule.doc.blocks.iter().for_each(|block| {
            let mut mapping = HashMap::new();
            let mut last_key = None;
            block.columns.iter().enumerate().for_each(|(idx, column)| {
                if column.auto_increment {
                    mapping.insert(AUTO_INCREMENT_KEY.clone(), idx);
                } else if let Some(prefix) = &column.custom_prefix {
                    mapping.insert(get_custom_prefix_key(prefix), idx);
                } else {
                    mapping.insert(column.cmark_tag.clone(), idx);
                }
                if column.is_last {
                    if let Some(prefix) = &column.custom_prefix {
                        last_key = Some(get_custom_prefix_key(prefix));
                    } else {
                        last_key = Some(column.cmark_tag.clone());
                    }
                }
            });
            blocks.push(Block {
                title: block.title.clone(),
                mapping,
                last_key,
            });
        });
        Ok(Mapping { blocks })
    }
}

impl Mapping {
    pub fn get_idx(
        &self,
        block_idx: usize,
        tag: Option<&Tag<'_>>,
        text_with_custom_prefix: Option<&str>,
    ) -> Option<&usize> {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.get_idx(tag, text_with_custom_prefix);
        }
        None
    }

    pub fn get_auto_increment_idx(&self, block_idx: usize) -> Option<&usize> {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.get_auto_increment_idx();
        }
        None
    }

    pub fn get_size(&self, block_idx: usize) -> Option<usize> {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.get_size();
        }
        None
    }

    pub fn is_last_key(
        &self,
        block_idx: usize,
        tag: Option<&Tag<'_>>,
        text_with_custom_prefix: Option<&str>,
    ) -> bool {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.is_last_key(tag, text_with_custom_prefix);
        }
        false
    }

    pub fn get_title(&self, block_idx: usize) -> Option<String> {
        if let Some(block) = self.blocks.get(block_idx) {
            return Some(block.title.clone());
        }
        None
    }
}

impl Default for Mapping {
    fn default() -> Self {
        Self { blocks: vec![] }
    }
}

#[derive(Debug, PartialEq)]
struct Block {
    title: String,
    mapping: HashMap<String, usize>,
    last_key: Option<String>,
}

impl Block {
    pub fn get_idx(
        &self,
        tag: Option<&Tag<'_>>,
        text_with_custom_prefix: Option<&str>,
    ) -> Option<&usize> {
        if let Some(key) = custom_prefix_to_key(text_with_custom_prefix) {
            return self.mapping.get(&key);
        } else if let Some(t) = tag {
            if let Some(tag_str) = cmarktag_stringify(t) {
                return self.mapping.get(&tag_str);
            }
        }
        None
    }

    pub fn get_auto_increment_idx(&self) -> Option<&usize> {
        self.mapping.get(&AUTO_INCREMENT_KEY.clone())
    }

    pub fn get_size(&self) -> Option<usize> {
        Some(self.mapping.len())
    }

    pub fn is_last_key(
        &self,
        tag: Option<&Tag<'_>>,
        text_with_custom_prefix: Option<&str>,
    ) -> bool {
        let k = if let Some(key) = custom_prefix_to_key(text_with_custom_prefix) {
            key
        } else if let Some(t) = tag {
            if let Some(tag_str) = cmarktag_stringify(t) {
                tag_str
            } else {
                return false;
            }
        } else {
            return false;
        };
        if let Some(last_key) = &self.last_key {
            return &k == last_key;
        }
        false
    }
}

impl Default for Block {
    fn default() -> Self {
        Self {
            title: String::default(),
            mapping: HashMap::new(),
            last_key: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::fs::read_to_string;

    use super::*;

    use crate::{
        constant::AUTO_INCREMENT_KEY,
        utils::{get_custom_prefix_as_normal_list, get_custom_prefix_key},
    };

    #[test]
    fn test_auto_increment_idx_empty() {
        let mapping = Mapping::default();
        assert!(mapping.get_auto_increment_idx(0).is_none());
    }

    #[test]
    fn test_get_size_empty() {
        let mapping = Mapping::default();
        assert!(mapping.get_size(0).is_none());
    }

    #[test]
    fn test_get_title() {
        let mut mapping = Mapping::default();
        mapping.blocks.push(Block {
            title: String::from("Block Title"),
            ..Default::default()
        });
        assert_eq!(Some(String::from("Block Title")), mapping.get_title(0));
        assert!(mapping.get_title(99).is_none());
    }

    #[test]
    fn test_mapping() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/default_rule.yml").unwrap()).unwrap();
        let mapping = Mapping::new(&rule).unwrap();
        let mut map = HashMap::new();
        map.insert(AUTO_INCREMENT_KEY.clone(), 0);
        map.insert("Heading2".to_string(), 1);
        map.insert("Heading3".to_string(), 2);
        map.insert("Heading4".to_string(), 3);
        map.insert("Heading5".to_string(), 4);
        map.insert("Heading6".to_string(), 5);
        map.insert("Heading7".to_string(), 6);
        map.insert("Heading8".to_string(), 7);
        map.insert("List".to_string(), 8);
        let expected = Mapping {
            blocks: vec![Block {
                title: String::from("Block Title"),
                mapping: map,
                last_key: Some(String::from("List")),
            }],
        };
        assert_eq!(expected, mapping);
        assert!(!mapping.is_last_key(0, Some(&Tag::Heading(8)), None));
        assert!(mapping.is_last_key(0, Some(&Tag::List(None)), None));
    }

    #[test]
    fn test_mapping_various_lists() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/various_list.yml").unwrap()).unwrap();
        let mapping = Mapping::new(&rule).unwrap();
        let mut map = HashMap::new();
        map.insert(AUTO_INCREMENT_KEY.clone(), 0);
        map.insert("Heading2".to_string(), 1);
        map.insert("Heading3".to_string(), 2);
        map.insert("Heading4".to_string(), 3);
        map.insert("Heading5".to_string(), 4);
        map.insert("Heading6".to_string(), 5);
        map.insert("Heading7".to_string(), 6);
        map.insert("Heading8".to_string(), 7);
        map.insert("List".to_string(), 8);
        map.insert(get_custom_prefix_key("+"), 9);
        map.insert(get_custom_prefix_key("$"), 10);
        let mut expected = Mapping {
            blocks: vec![Block {
                title: String::from("Block Title 1"),
                mapping: map,
                last_key: Some(get_custom_prefix_key("$")),
            }],
        };
        let mut map = HashMap::new();
        map.insert(AUTO_INCREMENT_KEY.clone(), 0);
        map.insert("Heading2".to_string(), 1);
        map.insert(get_custom_prefix_key("$"), 2);
        map.insert(get_custom_prefix_key("+"), 3);
        expected.blocks.push(Block {
            title: String::from("Block Title 2"),
            mapping: map,
            last_key: Some(get_custom_prefix_key("+")),
        });
        assert_eq!(expected, mapping);
        assert!(!mapping.is_last_key(0, Some(&Tag::Heading(8)), None));
        assert!(mapping.is_last_key(
            0,
            None,
            Some(
                &get_custom_prefix_as_normal_list("$")
                    .strip_prefix("* ")
                    .unwrap()
            )
        ));
    }
}
