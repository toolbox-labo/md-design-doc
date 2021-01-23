use std::collections::HashMap;

use anyhow::Result;
use pulldown_cmark::Tag;

use crate::{rule::Rule, utils::cmarktag_stringify};

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
                data.insert(column.cmark_tag.clone(), idx);
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

    pub fn get_size(&self, block_idx: usize) -> Option<usize> {
        //println!("{:?}", self);
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
