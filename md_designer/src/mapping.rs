use std::collections::HashMap;

use anyhow::Result;

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
