use anyhow::{anyhow, Result};
use pulldown_cmark::{Event, Options, Parser, Tag};

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
        // trim first empty lines
        let input = input.trim_start();

        // convert the rule into mapping
        let mapping = Mapping::new(&rule).unwrap();

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
        let parser = parser.map(|event| match event {
            Event::Text(ref text) => {
                if text.starts_with("####### ") {
                    return Event::Start(Tag::Heading(7));
                } else if text.starts_with("######## ") {
                    return Event::Start(Tag::Heading(8));
                } else {
                    event
                }
            }
            _ => event,
        });

        let current_block: usize = 0;
        let mut current_column: usize = 0;
        let mut sheet = Sheet::default();
        let mut block = Block::default();
        let mut row = Row::new(current_block, &mapping);
        let mut last_is_list = false;
        let mut is_sheet_name = false;
        let mut current_row = 1;

        parser.for_each(|event| {
            // is last line soft break
            // if true, next text data is append to current column
            match event {
                Event::Start(tag) => {
                    if let Tag::Heading(1) = tag {
                        // Heading 1 is the sheet name
                        is_sheet_name = true;
                    } else {
                        if last_is_list {
                            // start a new row
                            // insert auto incremented id if rule exists
                            if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
                                row.columns[*id_idx] = format!("{}", current_row);
                            }
                            block.rows.push(row.clone());
                            row = Row::new(current_block, &mapping);
                            last_is_list = false;
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
                    if let Tag::List(_) = tag {
                        last_is_list = true;
                    }
                    is_sheet_name = false;
                }
                _ => {}
            }
        });
        // push the last row
        if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
            row.columns[*id_idx] = format!("{}", current_row);
        }
        block.rows.push(row);
        sheet.blocks.push(block);

        Ok(Self {
            sheets: vec![sheet],
            rule,
            mapping,
        })
    }

    #[cfg(feature = "excel")]
    pub fn export_excel(&self) -> Result<()> {
        // TODO: customizable start positions
        let (_start_x, _start_y) = (0, 0);
        let (block_start_x, mut block_start_y) = (0, 0);
        let workbook = Workbook::new("test.xlsx");
        self.sheets.iter().for_each(|sheet| {
            let mut s = workbook.add_worksheet(sheet.sheet_name.as_deref()).unwrap();
            let wrap_format = workbook.add_format().set_text_wrap();
            sheet.blocks.iter().enumerate().for_each(|(idx, block)| {
                let mut merged_posisitons: Vec<CellRange> = vec![];
                if let Some(b) = self.rule.doc.blocks.get(idx) {
                    // Header
                    // render the merged cells first
                    // and store the merged column indexes
                    b.merge_info.iter().for_each(|merge_info| {
                        s.merge_range(
                            block_start_y,
                            merge_info.from,
                            block_start_y,
                            merge_info.to,
                            &merge_info.title,
                            None,
                        )
                        .unwrap();
                        merged_posisitons.push(CellRange::new(merge_info.from, merge_info.to));
                    });
                    // render the remaining headers
                    let header_merged = !merged_posisitons.is_empty();
                    b.columns.iter().enumerate().for_each(|(pos_x, column)| {
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
                            s.write_string(block_start_y + 1, pos_x, &column.title, None)
                                .unwrap();
                        } else if header_merged {
                            s.merge_range(
                                block_start_y,
                                pos_x,
                                block_start_y + 1,
                                pos_x,
                                &column.title,
                                None,
                            )
                            .unwrap();
                        } else {
                            s.write_string(block_start_y, pos_x, &column.title, None)
                                .unwrap();
                        }
                    });
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
                            )
                            .unwrap();
                        }
                        last_y = y_offset;
                    }

                    // update block_start_y for the next block
                    block_start_y += (last_y + 1) as u32;
                }
            });
        });
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
    rows: Vec<Row>,
}

impl Default for Block {
    fn default() -> Self {
        Self { rows: vec![] }
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
    fn test_marshal() {
        let rule = Rule::marshal(
            r#"
doc:
  blocks:
    - block:
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
## Test Variation 1
### Test Variation 1-1
#### Test Variation 1-1-1
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
                    rows: vec![Row {
                        columns: vec![
                            String::from("1"),
                            String::from("\nTest Variation 1"),
                            String::from("\nTest Variation 1-1"),
                            String::from("\nTest Variation 1-1-1"),
                            String::default(),
                            String::default(),
                            String::default(),
                            String::default(),
                            String::from("\nTest Description\nmore lines..."),
                        ],
                    }],
                }],
            }],
            mapping,
            rule: rule_clone,
        };
        assert_eq!(expected, data);
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
