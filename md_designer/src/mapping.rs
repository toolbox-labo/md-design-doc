use std::collections::HashMap;

use anyhow::Result;
use pulldown_cmark::Tag;

use crate::{rule::Rule, utils::cmarktag_stringify, constant::AUTO_INCREMENT_KEY};

#[derive(Debug, PartialEq)]
pub struct Mapping {
    pub mappings: Vec<HashMap<String, usize>>,
}

impl Mapping {
    pub fn new(rule: &Rule) -> Result<Self> {
        let mut mappings = vec![];
        rule.doc.blocks.iter().for_each(|block| {
            let mut data = HashMap::new();
            block.columns.iter().enumerate().for_each(|(idx, column)| {
                if column.auto_increment {
                    data.insert(AUTO_INCREMENT_KEY.clone(), idx);
                } else {
                    data.insert(column.cmark_tag.clone(), idx);
                }
            });
            mappings.push(data);
        });
        Ok(Mapping { mappings })
    }
}

impl Mapping {
    pub fn get_idx(&self, block_idx: usize, tag: &Tag<'_>) -> Option<&usize> {
        if let Some(map) = self.mappings.get(block_idx) {
            if let Some(tag_str) = cmarktag_stringify(tag) {
                return map.get(&tag_str);
            }
        }
        None
    }

    pub fn get_auto_increment_idx(&self, block_idx: usize) -> Option<&usize> {
        if let Some(map) = self.mappings.get(block_idx) {
            return map.get(&AUTO_INCREMENT_KEY.clone());
        }
        None
    }

    pub fn get_size(&self, block_idx: usize) -> Option<usize> {
        if let Some(map) = self.mappings.get(block_idx) {
            return Some(map.len());
        }
        None
    }
}

impl Default for Mapping {
    fn default() -> Self {
        Self { mappings: vec![] }
    }
}
