use std::rc::Rc;

use anyhow::{Context, Result};
use log::{debug, info};
use yaml_rust::{Yaml, YamlLoader};

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
    pub group: Option<Rc<Group>>,
    pub is_last: bool,
}

impl Default for Column {
    fn default() -> Self {
        Column {
            title: String::default(),
            auto_increment: false,
            cmark_tag: String::default(),
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
        let rule = Rule::marshal(
            r#"
doc:
  blocks:
    - title: Block Title
      content:
      - invalid_key: wrong!!
            "#,
        );
        assert!(rule.is_err());
    }

    #[test]
    fn test_marshal() {
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
}
