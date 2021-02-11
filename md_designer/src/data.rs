#[cfg(not(test))]
use log::{debug, info};

#[cfg(test)]
use std::{println as info, println as debug};

use anyhow::{anyhow, Result};
use pulldown_cmark::{CowStr, Event, Options, Parser, Tag};

use crate::{mapping::Mapping, rule::Rule};

#[cfg(feature = "excel")]
use xlsxwriter::*;

pub struct CellRange {
    from: u16,
    to: u16,
}

impl CellRange {
    pub fn new(from: u16, to: u16) -> Self {
        CellRange { from, to }
    }

    pub fn contain(&self, pos: u16) -> bool {
        self.from <= pos && pos <= self.to
    }
}

#[derive(Debug, PartialEq)]
pub struct Data {
    sheets: Vec<Sheet>,
    rule: Rule,
    mapping: Mapping,
}

impl Default for Data {
    fn default() -> Self {
        Self {
            sheets: vec![],
            rule: Rule::default(),
            mapping: Mapping::default(),
        }
    }
}

impl Data {
    pub fn marshal(input: &str, rule: Rule) -> Result<Self> {
        info!("parsing input with parsed rules...");
        // trim first empty lines
        let input = input.trim_start();

        // convert the rule into mapping
        let mapping = Mapping::new(&rule)?;

        // check is first line is Heading(1)
        // (sheet name is required)
        if !input.starts_with("# ") {
            return Err(anyhow!("input must start with '# ' (sheet name)."));
        }

        // marshal
        // expand parser to be able to handle 7th heading
        let mut options = Options::empty();
        options.insert(Options::ENABLE_TASKLISTS);
        //let input = Data::custom_filter(input);
        let parser = Parser::new_ext(&input, options);
        let mut parser_filtered = vec![];
        // pulldown_cmark does not support Heading(7) and Heading(8)
        // so they should be handled by hand
        parser.for_each(|event| match event {
            Event::Text(ref text) => {
                if let Some(content) = text.strip_prefix("####### ") {
                    parser_filtered.push(Event::Start(Tag::Heading(7)));
                    parser_filtered.push(Event::Text(CowStr::from(content.to_string())));
                } else if let Some(content) = text.strip_prefix("######## ") {
                    parser_filtered.push(Event::Start(Tag::Heading(8)));
                    parser_filtered.push(Event::Text(CowStr::from(content.to_string())));
                } else {
                    parser_filtered.push(event);
                }
            }
            _ => {
                parser_filtered.push(event);
            }
        });

        let mut current_block: usize = 0;
        let mut current_column: usize = 0;
        let mut sheet = Sheet::default();
        let mut block = Block::default();
        let mut row = Row::new(current_block, &mapping);
        let mut start_new_line = false;
        let mut is_sheet_name = false;
        let mut current_row = 1;
        let mut previous_idx: usize = 0;

        parser_filtered.iter().for_each(|event| {
            // if true, next text data is append to current column
            debug!("event: {:?}", event);
            match event {
                Event::Start(tag) => {
                    // check previous tag id
                    // if current tag id is smaller than previous one or equal, start new line
                    if let Some(current_idx) = mapping.get_idx(current_block, &tag) {
                        if current_idx <= &previous_idx {
                            start_new_line = true;
                        }
                    } else if mapping.is_last_key(current_block, tag) {
                        start_new_line = true;
                    }
                    if let Tag::Heading(1) = tag {
                        // Heading 1 is the sheet name
                        is_sheet_name = true;
                    } else {
                        if start_new_line {
                            // start a new row
                            // insert auto incremented id if rule exists
                            debug!("start a new line");
                            if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
                                row.columns[*id_idx] = format!("{}", current_row);
                            }
                            block.rows.push(row.clone());
                            row = Row::new(current_block, &mapping);
                            start_new_line = false;
                            previous_idx = 0;
                            current_row += 1;
                        }
                        if let Some(column_idx) = mapping.get_idx(current_block, &tag) {
                            current_column = *column_idx;
                        }
                    }
                }
                Event::Text(text) => {
                    if is_sheet_name {
                        sheet.sheet_name = Some(text.to_string());
                    } else {
                        row.columns[current_column] =
                            Data::concat(&row.columns.get(current_column), &text);
                    }
                }
                Event::End(tag) => {
                    is_sheet_name = false;
                    // store this tag idx as previous tag idx to be used by next loop
                    if let Some(idx) = mapping.get_idx(current_block, &tag) {
                        previous_idx = *idx;
                    }
                }
                Event::Rule => {
                    debug!("start a new block");
                    // push the last row and push block to blocks
                    if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
                        row.columns[*id_idx] = format!("{}", current_row);
                    }
                    block.rows.push(row.clone());
                    if let Some(title) = mapping.get_title(current_block) {
                        block.title = title;
                    }
                    sheet.blocks.push(block.clone());
                    // start a new block
                    block = Block::default();
                    current_block += 1;
                    current_row = 1;
                    current_column = 0;
                    previous_idx = 0;
                    //                    start_new_line = true;
                    row = Row::new(current_block, &mapping);
                }
                _ => {}
            }
        });
        // push the last row and block
        if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
            row.columns[*id_idx] = format!("{}", current_row);
        }
        if let Some(title) = mapping.get_title(current_block) {
            block.title = title;
        }
        block.rows.push(row);
        sheet.blocks.push(block);

        let data = Self {
            sheets: vec![sheet],
            rule,
            mapping,
        };

        info!("OK");
        debug!("parsed data: \n{:?}", data);
        Ok(data)
    }

    #[cfg(feature = "excel")]
    pub fn export_excel(&self, file_name: &str) -> Result<()> {
        info!("exporting excel file ({}.xlsx)...", file_name);
        // TODO: customizable start positions
        let (_start_x, _start_y) = (0, 0);
        let (block_start_x, mut block_start_y) = (0, 0);
        let workbook = Workbook::new(&format!("{}.xlsx", file_name));
        for sheet in self.sheets.iter() {
            let mut s = workbook.add_worksheet(sheet.sheet_name.as_deref())?;
            let wrap_format = workbook.add_format().set_text_wrap();
            for (idx, block) in sheet.blocks.iter().enumerate() {
                // render the block title
                s.write_string(block_start_y, block_start_x, &block.title, None)?;
                block_start_y += 1;
                let mut merged_posisitons: Vec<CellRange> = vec![];
                if let Some(b) = self.rule.doc.blocks.get(idx) {
                    // Header
                    // render the merged cells first
                    // and store the merged column indexes
                    for merge_info in b.merge_info.iter() {
                        s.merge_range(
                            block_start_y,
                            merge_info.from,
                            block_start_y,
                            merge_info.to,
                            &merge_info.title,
                            None,
                        )?;
                        merged_posisitons.push(CellRange::new(merge_info.from, merge_info.to));
                    }
                    // render the remaining headers
                    let header_merged = !merged_posisitons.is_empty();
                    for (pos_x, column) in b.columns.iter().enumerate() {
                        let pos_x = pos_x as u16;
                        // check if pos_x is within merged range
                        let mut in_merged_range = false;
                        for merged_pos in merged_posisitons.iter() {
                            if merged_pos.contain(pos_x) {
                                in_merged_range = true;
                                break;
                            }
                        }
                        if in_merged_range {
                            s.write_string(block_start_y + 1, pos_x, &column.title, None)?;
                        } else if header_merged {
                            s.merge_range(
                                block_start_y,
                                pos_x,
                                block_start_y + 1,
                                pos_x,
                                &column.title,
                                None,
                            )?;
                        } else {
                            s.write_string(block_start_y, pos_x, &column.title, None)?;
                        }
                    }
                    if header_merged {
                        block_start_y += 1;
                    }

                    // Body
                    let _body_start_x = block_start_x;
                    let body_start_y = block_start_y + 1;
                    let mut last_y = 0;
                    for (y_offset, row) in block.rows.iter().enumerate() {
                        for (x_offset, column) in row.columns.iter().enumerate() {
                            s.write_string(
                                body_start_y + (y_offset as u32),
                                block_start_x + x_offset as u16,
                                &column,
                                Some(&wrap_format),
                            )?;
                        }
                        last_y = y_offset;
                    }

                    // update block_start_y for the next block
                    block_start_y += (last_y + 3) as u32;
                }
            }
        }
        info!("OK");
        Ok(())
    }

    fn concat(target: &Option<&String>, input: &str) -> String {
        if let Some(str) = target {
            format!("{}\n{}", str, input)
        } else {
            input.to_string()
        }
    }
}

#[derive(Debug, PartialEq)]
struct Sheet {
    sheet_name: Option<String>,
    blocks: Vec<Block>,
}

impl Default for Sheet {
    fn default() -> Self {
        Self {
            sheet_name: None,
            blocks: vec![],
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Block {
    title: String,
    rows: Vec<Row>,
}

impl Default for Block {
    fn default() -> Self {
        Self {
            title: String::default(),
            rows: vec![],
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Row {
    columns: Vec<String>,
}

impl Row {
    fn new(block_idx: usize, mapping: &Mapping) -> Self {
        Row {
            columns: vec![String::default(); mapping.get_size(block_idx).unwrap_or(0)],
        }
    }
}

impl Default for Row {
    fn default() -> Self {
        Self { columns: vec![] }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_range_contain() {
        let cell_range = CellRange::new(0, 10);
        assert!(cell_range.contain(0));
        assert!(cell_range.contain(10));
        assert!(cell_range.contain(5));
        assert!(!cell_range.contain(11));
    }

    #[test]
    fn test_default_data() {
        let data = Data::default();
        let expected = Data {
            sheets: vec![],
            rule: Rule::default(),
            mapping: Mapping::default(),
        };
        assert_eq!(expected, data);
    }

    fn get_default_rule() -> Rule {
        Rule::marshal(
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
        .unwrap()
    }

    #[test]
    fn test_marshal_error() {
        let rule = get_default_rule();
        let data = Data::marshal(
            r#"
## Test Variation 1
### Test Variation 1-1
#### Test Variation 1-1-1
* Test Description
  more lines...
            "#,
            rule,
        );
        assert!(data.is_err());
    }

    #[test]
    fn test_marshal() {
        let rule = get_default_rule();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            r#"
# Sheet Name
## Test Variation 1
### Test Variation 1-1
#### Test Variation 1-1-1
##### Test Variation 1-1-1-1
###### Test Variation 1-1-1-1-1
####### Test Variation 1-1-1-1-1-1
######## Test Variation 1-1-1-1-1-1-1
* Test Description
  more lines...
## Test Variation 2
### Test Variation 2-1
#### Test Variation 2-1-1
##### Test Variation 2-1-1-1
* Test Description
  more lines...
##### Test Variation 2-1-1-2
* Test Description
  more lines...
            "#,
            rule,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![Sheet {
                sheet_name: Some(String::from("Sheet Name")),
                blocks: vec![Block {
                    title: String::from("Block Title"),
                    rows: vec![
                        Row {
                            columns: vec![
                                String::from("1"),
                                String::from("\nTest Variation 1"),
                                String::from("\nTest Variation 1-1"),
                                String::from("\nTest Variation 1-1-1"),
                                String::from("\nTest Variation 1-1-1-1"),
                                String::from("\nTest Variation 1-1-1-1-1"),
                                String::from("\nTest Variation 1-1-1-1-1-1"),
                                String::from("\nTest Variation 1-1-1-1-1-1-1"),
                                String::from("\nTest Description\nmore lines..."),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("2"),
                                String::from("\nTest Variation 2"),
                                String::from("\nTest Variation 2-1"),
                                String::from("\nTest Variation 2-1-1"),
                                String::from("\nTest Variation 2-1-1-1"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("\nTest Description\nmore lines..."),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("3"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("\nTest Variation 2-1-1-2"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("\nTest Description\nmore lines..."),
                            ],
                        },
                    ],
                }],
            }],
            mapping,
            rule: rule_clone,
        };
        assert_eq!(expected, data);
    }

    #[test]
    fn test_marshal_multiple_blocks() {
        let rule = Rule::marshal(
            r#"
doc:
  blocks:
    - title: Block Title 1
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
    - title: Block Title 2
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
    - title: Block Title 3
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
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            r#"
# Sheet Name
## Test Variation A 1
### Test Variation A 1-1
#### Test Variation A 1-1-1
##### Test Variation A 1-1-1-1
###### Test Variation A 1-1-1-1-1
####### Test Variation A 1-1-1-1-1-1
######## Test Variation A 1-1-1-1-1-1-1
* Test Description
  more lines...
## Test Variation A 2
### Test Variation A 2-1
#### Test Variation A 2-1-1
##### Test Variation A 2-1-1-1
* Test Description
  more lines...
##### Test Variation A 2-1-1-2
* Test Description
  more lines...
---
## Test Variation B 1
### Test Variation B 1-1
#### Test Variation B 1-1-1
##### Test Variation B 1-1-1-1
###### Test Variation B 1-1-1-1-1
####### Test Variation B 1-1-1-1-1-1
######## Test Variation B 1-1-1-1-1-1-1
* Test Description
  more lines...
## Test Variation B 2
### Test Variation B 2-1
#### Test Variation B 2-1-1
##### Test Variation B 2-1-1-1
* Test Description
  more lines...
##### Test Variation B 2-1-1-2
* Test Description
  more lines...
---
## Test Variation C 1
### Test Variation C 1-1
#### Test Variation C 1-1-1
##### Test Variation C 1-1-1-1
###### Test Variation C 1-1-1-1-1
####### Test Variation C 1-1-1-1-1-1
######## Test Variation C 1-1-1-1-1-1-1
* Test Description
  more lines...
## Test Variation C 2
### Test Variation C 2-1
#### Test Variation C 2-1-1
##### Test Variation C 2-1-1-1
* Test Description
  more lines...
##### Test Variation C 2-1-1-2
* Test Description
  more lines...
            "#,
            rule,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![Sheet {
                sheet_name: Some(String::from("Sheet Name")),
                blocks: vec![
                    Block {
                        title: String::from("Block Title 1"),
                        rows: vec![
                            Row {
                                columns: vec![
                                    String::from("1"),
                                    String::from("\nTest Variation A 1"),
                                    String::from("\nTest Variation A 1-1"),
                                    String::from("\nTest Variation A 1-1-1"),
                                    String::from("\nTest Variation A 1-1-1-1"),
                                    String::from("\nTest Variation A 1-1-1-1-1"),
                                    String::from("\nTest Variation A 1-1-1-1-1-1"),
                                    String::from("\nTest Variation A 1-1-1-1-1-1-1"),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("\nTest Variation A 2"),
                                    String::from("\nTest Variation A 2-1"),
                                    String::from("\nTest Variation A 2-1-1"),
                                    String::from("\nTest Variation A 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Variation A 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                        ],
                    },
                    Block {
                        title: String::from("Block Title 2"),
                        rows: vec![
                            Row {
                                columns: vec![
                                    String::from("1"),
                                    String::from("\nTest Variation B 1"),
                                    String::from("\nTest Variation B 1-1"),
                                    String::from("\nTest Variation B 1-1-1"),
                                    String::from("\nTest Variation B 1-1-1-1"),
                                    String::from("\nTest Variation B 1-1-1-1-1"),
                                    String::from("\nTest Variation B 1-1-1-1-1-1"),
                                    String::from("\nTest Variation B 1-1-1-1-1-1-1"),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("\nTest Variation B 2"),
                                    String::from("\nTest Variation B 2-1"),
                                    String::from("\nTest Variation B 2-1-1"),
                                    String::from("\nTest Variation B 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Variation B 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                        ],
                    },
                    Block {
                        title: String::from("Block Title 3"),
                        rows: vec![
                            Row {
                                columns: vec![
                                    String::from("1"),
                                    String::from("\nTest Variation C 1"),
                                    String::from("\nTest Variation C 1-1"),
                                    String::from("\nTest Variation C 1-1-1"),
                                    String::from("\nTest Variation C 1-1-1-1"),
                                    String::from("\nTest Variation C 1-1-1-1-1"),
                                    String::from("\nTest Variation C 1-1-1-1-1-1"),
                                    String::from("\nTest Variation C 1-1-1-1-1-1-1"),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("\nTest Variation C 2"),
                                    String::from("\nTest Variation C 2-1"),
                                    String::from("\nTest Variation C 2-1-1"),
                                    String::from("\nTest Variation C 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Variation C 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("\nTest Description\nmore lines..."),
                                ],
                            },
                        ],
                    },
                ],
            }],
            mapping,
            rule: rule_clone,
        };
        assert_eq!(expected, data);
    }

    #[test]
    fn test_marshal_without_list() {
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
            "#,
        )
        .unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            r#"
# Sheet Name
## Test Variation 1
### Test Variation 1-1
### Test Variation 1-2
## Test Variation 2
## Test Variation 3
### Test Variation 3-1
### Test Variation 3-2
            "#,
            rule,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![Sheet {
                sheet_name: Some(String::from("Sheet Name")),
                blocks: vec![Block {
                    title: String::from("Block Title"),
                    rows: vec![
                        Row {
                            columns: vec![
                                String::from("1"),
                                String::from("\nTest Variation 1"),
                                String::from("\nTest Variation 1-1"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("2"),
                                String::default(),
                                String::from("\nTest Variation 1-2"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("3"),
                                String::from("\nTest Variation 2"),
                                String::default(),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("4"),
                                String::from("\nTest Variation 3"),
                                String::from("\nTest Variation 3-1"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("5"),
                                String::default(),
                                String::from("\nTest Variation 3-2"),
                            ],
                        },
                    ],
                }],
            }],
            mapping,
            rule: rule_clone,
        };
        assert_eq!(expected, data);
    }

    #[test]
    fn test_export_excel() {
        let rule = get_default_rule();
        let data = Data::marshal(
            r#"
# Sheet Name
## Test Variation 1
### Test Variation 1-1
#### Test Variation 1-1-1
##### Test Variation 1-1-1-1
###### Test Variation 1-1-1-1-1
####### Test Variation 1-1-1-1-1-1
######## Test Variation 1-1-1-1-1-1-1
* Test Description
  more lines...
## Test Variation 2
### Test Variation 2-1
#### Test Variation 2-1-1
##### Test Variation 2-1-1-1
* Test Description
  more lines...
##### Test Variation 2-1-1-2
* Test Description
  more lines...
            "#,
            rule,
        )
        .unwrap();
        let file_name = "unit_test";
        assert!(data.export_excel("unit_test").is_ok());
        std::fs::remove_file(format!("{}.xlsx", file_name)).unwrap();
    }

    #[test]
    fn test_concat() {
        // None
        let target = None;
        let input = "input";
        let result = Data::concat(&target, input);
        assert_eq!(String::from("input"), result);
        // Some
        let data = String::from("target");
        let target = Some(&data);
        let input = "input";
        let result = Data::concat(&target, input);
        assert_eq!(String::from("target\ninput"), result);
    }
}
