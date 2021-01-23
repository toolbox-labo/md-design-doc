use std::rc::Rc;

use anyhow::Result;
use yaml_rust::{Yaml, YamlLoader};

pub struct Rule {
    doc: Doc,
}

impl Rule {
    pub fn marshal(input: &str) -> Result<Self> {
        let docs = YamlLoader::load_from_str(input).unwrap();
        let doc = &docs[0]["doc"];
        let mut blcs = vec![];
        // TODO: validation

        if let Some(blocks) = doc["blocks"].as_vec() {
            blocks.iter().for_each(|v| {
                let mut blc = Block::default();
                if let Some(block) = v["block"].as_vec() {
                    block.iter().for_each(|w| {
                        if let Some(col_or_grp) = w.as_hash() {
                            let (col_or_grp_list, group) = if col_or_grp
                                .contains_key(&Yaml::String("column".to_string()))
                            {
                                (vec![col_or_grp], None)
                            } else if col_or_grp.contains_key(&Yaml::String("group".to_string())) {
                                (
                                    col_or_grp
                                        .get(&Yaml::String("columns".to_string()))
                                        .unwrap()
                                        .as_vec()
                                        .unwrap()
                                        .iter()
                                        .map(|v| v.as_hash().unwrap())
                                        .collect(),
                                    Some(Rc::new(Group {
                                        title: String::from(
                                            col_or_grp
                                                .get(&Yaml::String("group".to_string()))
                                                .unwrap()
                                                .as_str()
                                                .unwrap_or(""),
                                        ),
                                    })),
                                )
                            } else {
                                (vec![], None)
                            };
                            col_or_grp_list.iter().for_each(|clm| {
                                blc.columns.push(Column {
                                    title: String::from(
                                        clm.get(&Yaml::String("column".to_string()))
                                            .unwrap()
                                            .as_str()
                                            .unwrap_or(""),
                                    ),
                                    auto_increment: clm
                                        .get(&Yaml::String("isNum".to_string()))
                                        .unwrap_or(&Yaml::Boolean(false))
                                        .as_bool()
                                        .unwrap_or(false),
                                    cmark_tag: String::from(
                                        clm.get(&Yaml::String("md".to_string()))
                                            .unwrap_or(&Yaml::String("".to_string()))
                                            .as_str()
                                            .unwrap(),
                                    ),
                                    group: if let Some(g) = &group {
                                        Some(g.clone())
                                    } else {
                                        None
                                    },
                                });
                            });
                        }
                    });
                }
                blcs.push(blc);
            });
        }

        Ok(Rule {
            doc: Doc { blocks: vec![] },
        })
    }
}

impl Default for Rule {
    fn default() -> Self {
        Rule {
            doc: Doc::default(),
        }
    }
}

struct Doc {
    blocks: Vec<Block>,
}

impl Default for Doc {
    fn default() -> Self {
        Doc { blocks: vec![] }
    }
}

#[derive(Debug)]
struct Block {
    columns: Vec<Column>,
}

impl Default for Block {
    fn default() -> Self {
        Block { columns: vec![] }
    }
}

#[derive(Debug)]
struct Column {
    title: String,
    auto_increment: bool,
    cmark_tag: String,
    group: Option<Rc<Group>>,
}

impl Default for Column {
    fn default() -> Self {
        Column {
            title: String::default(),
            auto_increment: false,
            cmark_tag: String::default(),
            group: None,
        }
    }
}

#[derive(Debug)]
struct Group {
    title: String,
}
