use std::collections::HashMap;

use anyhow::Result;
use pulldown_cmark::Tag;

use crate::{constant::AUTO_INCREMENT_KEY, rule::Rule, utils::cmarktag_stringify};

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
                } else {
                    mapping.insert(column.cmark_tag.clone(), idx);
                }
                if column.is_last {
                    last_key = Some(column.cmark_tag.clone());
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
    pub fn get_idx(&self, block_idx: usize, tag: &Tag<'_>) -> Option<&usize> {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.get_idx(tag);
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

    pub fn is_last_key(&self, block_idx: usize, tag: &Tag<'_>) -> bool {
        if let Some(block) = self.blocks.get(block_idx) {
            return block.is_last_key(tag);
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
    pub fn get_idx(&self, tag: &Tag<'_>) -> Option<&usize> {
        if let Some(tag_str) = cmarktag_stringify(tag) {
            return self.mapping.get(&tag_str);
        }
        None
    }

    pub fn get_auto_increment_idx(&self) -> Option<&usize> {
        self.mapping.get(&AUTO_INCREMENT_KEY.clone())
    }

    pub fn get_size(&self) -> Option<usize> {
        Some(self.mapping.len())
    }

    pub fn is_last_key(&self, tag: &Tag<'_>) -> bool {
        if let Some(tag_str) = cmarktag_stringify(tag) {
            if let Some(k) = &self.last_key {
                return k == &tag_str;
            }
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
    use super::*;

    use crate::constant::AUTO_INCREMENT_KEY;

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
        let rule = Rule::marshal(
            r#"
doc:
  blocks:
    - title: Block Title
      content:
      - column: No
        isNum: true
      - group: Variation
        columns:
        - column: Variation 1
          md: Heading2
        - column: Variation 2
          md: Heading3
        - column: Variation 3
          md: Heading4
        - column: Variation 4
          md: Heading5
        - column: Variation 5
          md: Heading6
        - column: Variation 6
          md: Heading7
        - column: Variation 7
          md: Heading8
      - column: Description
        md: List
            "#,
        )
        .unwrap();
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
        assert!(!mapping.is_last_key(0, &Tag::Heading(8)));
        assert!(mapping.is_last_key(0, &Tag::List(None)));
    }
}
