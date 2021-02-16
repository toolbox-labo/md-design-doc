use std::rc::Rc;

use anyhow::{Context, Result};
use log::{debug, info};
use regex::Regex;
use yaml_rust::{Yaml, YamlLoader};

use crate::utils::get_custom_prefix_as_normal_list;

#[derive(Debug, PartialEq, Clone)]
pub struct MergeInfo {
    pub title: String,
    pub from: u16,
    pub to: u16,
}

impl MergeInfo {
    pub fn new(title: &str, from: u16, to: u16) -> Self {
        MergeInfo {
            title: title.to_string(),
            from,
            to,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Rule {
    pub doc: Doc,
}

impl Rule {
    pub fn marshal(input: &str) -> Result<Self> {
        info!("parsing rules...");
        let docs = YamlLoader::load_from_str(input)?;
        let doc = &docs[0]["doc"];
        let mut blcs = vec![];
        // TODO: validation

        if let Some(blocks) = doc["blocks"].as_vec() {
            for v in blocks.iter() {
                let mut blc = Block::default();
                if let Some(title) = v["title"].as_str() {
                    blc.title = title.to_string();
                }
                if let Some(block) = v["content"].as_vec() {
                    let mut idx: usize = 0;
                    let mut group_from: Option<usize> = None;
                    let mut group_to: Option<usize> = None;
                    for (i, w) in block.iter().enumerate() {
                        if let Some(col_or_grp) = w.as_hash() {
                            let (col_or_grp_list, group) = if col_or_grp
                                .contains_key(&Yaml::String("column".to_string()))
                            {
                                idx += 1;
                                (vec![col_or_grp], None)
                            } else if col_or_grp.contains_key(&Yaml::String("group".to_string())) {
                                let grp_list = col_or_grp
                                    .get(&Yaml::String("columns".to_string()))
                                    .with_context(|| "columns key is required in group")?
                                    .as_vec()
                                    .with_context(|| "columns must be array")?;
                                group_from = Some(idx);
                                idx = idx.saturating_add(grp_list.len() - 1);
                                group_to = Some(idx);
                                (
                                    grp_list.iter().map(|v| v.as_hash().unwrap()).collect(),
                                    Some(Rc::new(Group {
                                        title: String::from(
                                            col_or_grp
                                                .get(&Yaml::String("group".to_string()))
                                                // It is clear that group key exists
                                                .unwrap()
                                                .as_str()
                                                // allows group value to be empty
                                                .unwrap_or(""),
                                        ),
                                    })),
                                )
                            } else {
                                return Err(anyhow::anyhow!("All values of 'block' key must have either keys 'column' or 'group'"));
                            };
                            for clm in col_or_grp_list.iter() {
                                blc.columns.push(Column {
                                    title: String::from(
                                        clm.get(&Yaml::String("column".to_string()))
                                            .with_context(|| "column key is required")?
                                            .as_str()
                                            // allows column value to be empty
                                            .unwrap_or(""),
                                    ),
                                    auto_increment: clm
                                        .get(&Yaml::String("isNum".to_string()))
                                        // allows key isNum to be undefined
                                        .unwrap_or(&Yaml::Boolean(false))
                                        .as_bool()
                                        .unwrap_or(false),
                                    cmark_tag: String::from(
                                        clm.get(&Yaml::String("md".to_string()))
                                            // allows key md to be undefined
                                            // this is for auto incremented column
                                            .unwrap_or(&Yaml::String("".to_string()))
                                            .as_str()
                                            .unwrap(),
                                    ),
                                    custom_prefix: {
                                        if let Some(prefix) =
                                            clm.get(&Yaml::String("customPrefix".to_string()))
                                        {
                                            //Some(prefix.as_str().unwrap_or("").to_string())
                                            let p: Result<&str> = if let Some(p) = prefix.as_str() {
                                                if p.len() != 1 {
                                                    return Err(anyhow::anyhow!("Custom prefix's length must be 1. Your input is {}", p.len()));
                                                }
                                                Ok(p)
                                            } else {
                                                return Err(anyhow::anyhow!("Custom prefix is malformed. It could not be converted into string: {:?}", prefix));
                                            };
                                            Some(p?.to_string())
                                        } else {
                                            None
                                        }
                                    },
                                    group: if let Some(g) = &group {
                                        Some(g.clone())
                                    } else {
                                        None
                                    },
                                    is_last: i == block.len().saturating_sub(1),
                                });
                            }
                            if let Some(g) = &group {
                                blc.merge_info.push(MergeInfo::new(
                                    g.title.as_str(),
                                    group_from.unwrap() as u16,
                                    group_to.unwrap() as u16,
                                ));
                            }
                        }
                    }
                }
                blcs.push(blc);
            }
        }
        let rule = Rule {
            doc: Doc { blocks: blcs },
        };

        info!("OK");
        debug!("parsed rule: \n{:?}", rule);
        Ok(rule)
    }

    /// This function filters the custom prefix lists into normal lists.
    /// In addition, it prepends `!!!CUSTOMPREFIX<prefix>` to be able to be checked if they're custom prefix lists or not.
    /// For example: `+ hogehoge` -> `* !!!CUSTOMPREFIX+ hogehoge`
    pub fn filter(&self, input: &str) -> String {
        let separator = Regex::new(r"(?m)^---(.*)").expect("Invalid regex");
        let mut result = vec![];
        for (idx, block) in separator.split(input).enumerate() {
            let mut block_replaced = block.to_owned();
            if let Some(b) = self.doc.blocks.get(idx) {
                for column in b.columns.iter() {
                    if let Some(prefix) = &column.custom_prefix {
                        let mut lines = vec![];
                        for line in block_replaced.lines() {
                            let replacement = get_custom_prefix_as_normal_list(&prefix);
                            if let Some(stripped) = line.trim().strip_prefix(prefix) {
                                // check if stripped text starts with ' '
                                // - 'D Description' -> repleace
                                // - '  Description' -> NOT replace
                                if stripped.strip_prefix(" ").is_some() {
                                    lines.push(format!("{}{}", &replacement, stripped));
                                } else {
                                    lines.push(line.to_string());
                                }
                            } else {
                                lines.push(line.to_string());
                            }
                        }
                        block_replaced = lines.join("\n");
                    }
                }
            }
            result.push(block_replaced);
        }
        format!("{}{}", result.join("\n---"), "\n")
    }
}

impl Default for Rule {
    fn default() -> Self {
        Rule {
            doc: Doc::default(),
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Doc {
    pub blocks: Vec<Block>,
}

impl Default for Doc {
    fn default() -> Self {
        Doc { blocks: vec![] }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Block {
    pub title: String,
    pub columns: Vec<Column>,
    pub merge_info: Vec<MergeInfo>,
}

impl Default for Block {
    fn default() -> Self {
        Block {
            title: String::default(),
            columns: vec![],
            merge_info: vec![],
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Column {
    pub title: String,
    pub auto_increment: bool,
    pub cmark_tag: String,
    pub custom_prefix: Option<String>,
    pub group: Option<Rc<Group>>,
    pub is_last: bool,
}

impl Default for Column {
    fn default() -> Self {
        Column {
            title: String::default(),
            auto_increment: false,
            cmark_tag: String::default(),
            custom_prefix: None,
            group: None,
            is_last: false,
        }
    }
}

#[derive(Debug, PartialEq, Clone)]
pub struct Group {
    pub title: String,
}

#[cfg(test)]
mod tests {
    use std::fs::read_to_string;

    use super::*;

    #[test]
    fn test_default_rule() {
        let rule = Rule::default();
        let expected = Rule {
            doc: Doc::default(),
        };
        assert_eq!(expected, rule);
    }

    #[test]
    fn test_marshal_invalid_key() {
        let rule = Rule::marshal(&read_to_string("test_case/rule/invalid_key.yml").unwrap());
        assert!(rule.is_err());
    }

    #[test]
    fn test_marshal() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/default_rule.yml").unwrap()).unwrap();
        let group = Rc::new(Group {
            title: String::from("Variation"),
        });
        let expected = Rule {
            doc: Doc {
                blocks: vec![Block {
                    title: String::from("Block Title"),
                    columns: vec![
                        Column {
                            title: String::from("No"),
                            auto_increment: true,
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 1"),
                            cmark_tag: String::from("Heading2"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 2"),
                            cmark_tag: String::from("Heading3"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 3"),
                            cmark_tag: String::from("Heading4"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 4"),
                            cmark_tag: String::from("Heading5"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 5"),
                            cmark_tag: String::from("Heading6"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 6"),
                            cmark_tag: String::from("Heading7"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Variation 7"),
                            cmark_tag: String::from("Heading8"),
                            group: Some(group.clone()),
                            ..Default::default()
                        },
                        Column {
                            title: String::from("Description"),
                            cmark_tag: String::from("List"),
                            is_last: true,
                            ..Default::default()
                        },
                    ],
                    merge_info: vec![MergeInfo {
                        title: String::from("Variation"),
                        from: 1,
                        to: 7,
                    }],
                }],
            },
        };
        assert_eq!(expected, rule);
    }

    #[test]
    fn test_marshal_various_list() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/various_list.yml").unwrap()).unwrap();
        let group = Rc::new(Group {
            title: String::from("Variation"),
        });
        let expected = Rule {
            doc: Doc {
                blocks: vec![
                    Block {
                        title: String::from("Block Title 1"),
                        columns: vec![
                            Column {
                                title: String::from("No"),
                                auto_increment: true,
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 1"),
                                cmark_tag: String::from("Heading2"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 2"),
                                cmark_tag: String::from("Heading3"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 3"),
                                cmark_tag: String::from("Heading4"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 4"),
                                cmark_tag: String::from("Heading5"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 5"),
                                cmark_tag: String::from("Heading6"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 6"),
                                cmark_tag: String::from("Heading7"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Variation 7"),
                                cmark_tag: String::from("Heading8"),
                                group: Some(group.clone()),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Description"),
                                cmark_tag: String::from("List"),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Procedure"),
                                cmark_tag: String::from("List"),
                                custom_prefix: Some(String::from("+")),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Date"),
                                cmark_tag: String::from("List"),
                                custom_prefix: Some(String::from("$")),
                                is_last: true,
                                ..Default::default()
                            },
                        ],
                        merge_info: vec![MergeInfo {
                            title: String::from("Variation"),
                            from: 1,
                            to: 7,
                        }],
                    },
                    Block {
                        title: String::from("Block Title 2"),
                        columns: vec![
                            Column {
                                title: String::from("No"),
                                auto_increment: true,
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Column 1"),
                                cmark_tag: String::from("Heading2"),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Description"),
                                cmark_tag: String::from("List"),
                                custom_prefix: Some(String::from("$")),
                                ..Default::default()
                            },
                            Column {
                                title: String::from("Result"),
                                cmark_tag: String::from("List"),
                                custom_prefix: Some(String::from("+")),
                                is_last: true,
                                ..Default::default()
                            },
                        ],
                        merge_info: vec![],
                    },
                ],
            },
        };
        assert_eq!(expected, rule);
    }

    #[test]
    fn test_marshal_various_list_prefix_too_long() {
        assert!(Rule::marshal(
            &read_to_string("test_case/rule/various_list_prefix_too_long.yml").unwrap(),
        )
        .is_err());
    }

    #[test]
    fn test_filter() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/various_list.yml").unwrap()).unwrap();
        let result = rule.filter(&read_to_string("test_case/input/various_list.md").unwrap());
        let expected = read_to_string("test_case/input/various_list_filtered.md").unwrap();
        assert_eq!(expected, result);
    }

    #[test]
    fn test_filter_confusing() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/list_confusing_prefix.yml").unwrap())
                .unwrap();
        let result =
            rule.filter(&read_to_string("test_case/input/list_confusing_prefix.md").unwrap());
        let expected = read_to_string("test_case/input/list_confusing_prefix_filtered.md").unwrap();
        assert_eq!(expected, result);
    }
}
