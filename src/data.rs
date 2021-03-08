#[cfg(not(test))]
use log::{debug, info};

#[cfg(test)]
use std::{println as info, println as debug};

use anyhow::{anyhow, Result};
use pulldown_cmark::{CowStr, Event, Options, Parser, Tag};

use crate::{
    mapping::Mapping,
    rule::Rule,
    utils::{custom_prefix_to_key, get_custom_prefix_end_idx},
};

#[cfg(feature = "excel")]
use xlsxwriter::*;

#[derive(Debug)]
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
        // escape md notation without beginnig of line
        info!("escape input");
        let input = Data::escape_notation(input);

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

        let input = rule.filter(input);

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

        let mut current_sheet: usize = 0;
        let mut current_block: usize = 0;
        let mut current_column: usize = 0;
        let mut current_row = 1;
        // Column idx in the previous loop.
        // It's used for checking if the new line should be started.
        let mut previous_idx: usize = 0;
        // used for checking if the new line should be started.
        let mut previous_is_list = false;
        let mut sheets = vec![];
        let mut sheet = Sheet::default();
        let mut block = Block::default();
        let mut row = Row::new(current_block, &mapping);
        // new line should be started in current loop?
        let mut start_new_line = false;
        // is started tag has the sheet name?
        let mut is_sheet_name = false;
        // is the first row since the new block started?
        let mut block_start = false;

        parser_filtered.iter().for_each(|event| {
            // if true, next text data is append to current column
            debug!("event: {:?}", event);
            match event {
                Event::Start(tag) => {
                    // check previous tag id
                    // if current tag id is smaller than previous one or equal, start new line
                    if let Some(current_idx) = mapping.get_idx(current_block, Some(&tag), None) {
                        if current_idx <= &previous_idx {
                            // if Tag::List starts, check the previous tag and if it's also tag,
                            // skip starting a new line.
                            if let Tag::List(_) = tag {
                                if !previous_is_list {
                                    start_new_line = true;
                                }
                            } else {
                                start_new_line = true;
                            }
                        }
                    } else if mapping.is_last_key(current_block, Some(tag), None) {
                        start_new_line = true;
                    }
                    if let Tag::Heading(1) = tag {
                        // Heading 1 is the sheet name
                        is_sheet_name = true;
                    } else {
                        if start_new_line {
                            // start a new row
                            // insert auto incremented id if rule exists
                            debug!("start a new line (Event::Start)");
                            if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
                                row.columns[*id_idx] = format!("{}", current_row);
                            }
                            block.rows.push(row.clone());
                            row = Row::new(current_block, &mapping);
                            start_new_line = false;
                            previous_idx = 0;
                            current_row += 1;
                        }
                        if let Some(column_idx) = mapping.get_idx(current_block, Some(&tag), None) {
                            current_column = *column_idx;
                        }
                    }
                }
                Event::Text(text) => {
                    if is_sheet_name {
                        current_sheet += 1;
                        if current_sheet > 1 {
                            debug!("start a new sheet");
                            // start a new sheet
                            // push the last row and block
                            if let Some(id_idx) = mapping.get_auto_increment_idx(current_block) {
                                row.columns[*id_idx] = format!("{}", current_row);
                            }
                            if let Some(title) = mapping.get_title(current_block) {
                                block.title = title;
                            }
                            block.rows.push(row.clone());
                            sheet.blocks.push(block.clone());
                            sheets.push(sheet.clone());
                            // reset variables
                            current_block = 0;
                            current_column = 0;
                            current_row = 1;
                            previous_idx = 0;
                            previous_is_list = false;
                            sheet = Sheet::default();
                            block = Block::default();
                            row = Row::new(current_block, &mapping);
                            start_new_line = false;
                            is_sheet_name = false;
                            block_start = false;
                        }
                        sheet.sheet_name = Some(Data::reverse_escape_notation(&text.to_string()));
                        debug!("sheet name pushed: {:?}", sheet.sheet_name);
                    } else if custom_prefix_to_key(Some(text)).is_some() {
                        if let Some(column_idx) = mapping.get_idx(current_block, None, Some(text)) {
                            if column_idx < &current_column && !block_start {
                                // start a new row
                                debug!("start a new line (Event::Text)");
                                if let Some(id_idx) = mapping.get_auto_increment_idx(current_block)
                                {
                                    row.columns[*id_idx] = format!("{}", current_row);
                                }
                                block.rows.push(row.clone());
                                row = Row::new(current_block, &mapping);
                                start_new_line = false;
                                previous_idx = 0;
                                current_row += 1;
                            }
                            row.columns[*column_idx] = Data::concat(
                                &row.columns.get(*column_idx),
                                &Data::reverse_escape_notation(
                                    &text[get_custom_prefix_end_idx()..],
                                ),
                            );
                            current_column = *column_idx;
                            debug!(
                                "cell pushed => sheet: {}, block: {}, row: {}, column: {}",
                                current_sheet, current_block, current_row, current_column
                            );
                        }
                        block_start = false;
                    } else {
                        row.columns[current_column] = Data::concat(
                            &row.columns.get(current_column),
                            &Data::reverse_escape_notation(&text),
                        );
                        debug!(
                            "cell pushed => sheet: {}, block: {}, row: {}, column: {}",
                            current_sheet, current_block, current_row, current_column
                        );
                    }
                }
                Event::End(tag) => {
                    is_sheet_name = false;
                    // store this tag idx as previous tag idx to be used by next loop
                    if let Some(idx) = mapping.get_idx(current_block, Some(&tag), None) {
                        previous_idx = *idx;
                    }
                    // store if current tag is list to be used by next loop
                    if let Tag::List(_) = tag {
                        previous_is_list = true;
                    } else {
                        previous_is_list = false;
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
                    row = Row::new(current_block, &mapping);
                    block_start = true;
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
        sheets.push(sheet);

        let data = Self {
            sheets,
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
        let workbook = Workbook::new(&format!("{}.xlsx", file_name));
        let title_format = workbook.add_format().set_font_size(16.0).set_bold();
        let head_row_format = workbook
            .add_format()
            .set_text_wrap()
            .set_align(FormatAlignment::CenterAcross)
            .set_align(FormatAlignment::VerticalCenter)
            .set_border(FormatBorder::Thin)
            .set_bg_color(FormatColor::Cyan);
        let data_row_format = workbook
            .add_format()
            .set_text_wrap()
            .set_align(FormatAlignment::Left)
            .set_align(FormatAlignment::VerticalTop)
            .set_border(FormatBorder::Thin);
        for sheet in self.sheets.iter() {
            let (_start_x, _start_y) = (0, 0);
            let (block_start_x, mut block_start_y) = (0, 0);
            let mut s = workbook.add_worksheet(sheet.sheet_name.as_deref())?;
            for (idx, block) in sheet.blocks.iter().enumerate() {
                // render the block title
                s.write_string(
                    block_start_y,
                    block_start_x,
                    &block.title,
                    Some(&title_format),
                )?;
                block_start_y += 1;
                let mut merged_positions: Vec<CellRange> = vec![];
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
                            Some(&head_row_format),
                        )?;
                        debug!("(header)merge_range -> start_y: {:?}, start_x: {:?}, end_y: {:?}, end_x: {:?}, text: {:?}", block_start_y, merge_info.from, block_start_y, merge_info.to, &merge_info.title);
                        merged_positions.push(CellRange::new(merge_info.from, merge_info.to));
                    }
                    debug!("merged_positions: {:?}", merged_positions);
                    // render the remaining headers
                    let header_merged = !merged_positions.is_empty();
                    for (pos_x, column) in b.columns.iter().enumerate() {
                        let pos_x = pos_x as u16;
                        // check if pos_x is within merged range
                        let mut in_merged_range = false;
                        for merged_pos in merged_positions.iter() {
                            if merged_pos.contain(pos_x) {
                                in_merged_range = true;
                                break;
                            }
                        }
                        if in_merged_range {
                            s.write_string(
                                block_start_y + 1,
                                pos_x,
                                &column.title,
                                Some(&head_row_format),
                            )?;
                            debug!(
                                "(header)write_string -> y: {:?}, x: {:?}, text: {:?}",
                                block_start_y + 1,
                                pos_x,
                                &column.title
                            );
                        } else if header_merged {
                            s.merge_range(
                                block_start_y,
                                pos_x,
                                block_start_y + 1,
                                pos_x,
                                &column.title,
                                Some(&head_row_format),
                            )?;
                            debug!("(header)merge_range -> start_y: {:?}, start_x: {:?}, end_y: {:?}, end_x: {:?}, text: {:?}", block_start_y, pos_x, block_start_y + 1, pos_x, &column.title);
                        } else {
                            s.write_string(
                                block_start_y,
                                pos_x,
                                &column.title,
                                Some(&head_row_format),
                            )?;
                            debug!(
                                "(header)write_string -> y: {:?}, x: {:?}, text: {:?}",
                                block_start_y, pos_x, &column.title
                            );
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
                                Some(&data_row_format),
                            )?;
                            debug!(
                                "(content)write_string -> y: {:?}, x: {:?}, text: {:?}",
                                body_start_y + (y_offset as u32),
                                block_start_x + x_offset as u16,
                                &column
                            );
                        }
                        last_y = y_offset;
                    }

                    // update block_start_y for the next block
                    block_start_y += (last_y + 3) as u32;
                }
            }
        }
        workbook.close()?;
        info!("OK");
        Ok(())
    }

    fn concat(target: &Option<&String>, input: &str) -> String {
        if let Some(str) = target {
            if !str.is_empty() {
                return format!("{}\n{}", str, input);
            }
        }
        input.to_string()
    }

    fn escape_notation(input: &str) -> String {
        let mut result: String = "".to_string();
        // split \n and convert by one line
        let splits: Vec<&str> = input.split('\n').collect();
        for split in splits {
            // skip first of line
            let mut split = split.to_string();
            let mut head: String = "".to_string();
            if !split.is_empty() {
                head = split.remove(0).to_string();
            }
            let replaced = split.replace("*", "--asterisk--");
            result.push_str(&format!("\n{}{}", head, replaced));
        }
        result
    }

    fn reverse_escape_notation(input: &str) -> String {
        input.replace("--asterisk--", "*")
    }
}

#[derive(Debug, PartialEq, Clone)]
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
    use std::fs::read_to_string;

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
        Rule::marshal(&read_to_string("test_case/rule/default_rule.yml").unwrap()).unwrap()
    }

    #[test]
    fn test_marshal_error() {
        let rule = get_default_rule();
        let data = Data::marshal(
            &read_to_string("test_case/input/error_input.md").unwrap(),
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
            &read_to_string("test_case/input/single_block_multi_row.md").unwrap(),
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
                                String::from("Test Variation 1"),
                                String::from("Test Variation 1-1"),
                                String::from("Test Variation 1-1-1"),
                                String::from("Test Variation 1-1-1-1"),
                                String::from("Test Variation 1-1-1-1-1"),
                                String::from("Test Variation 1-1-1-1-1-1"),
                                String::from("Test Variation 1-1-1-1-1-1-1"),
                                String::from("Test Description\nmore lines..."),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("2"),
                                String::from("Test Variation 2"),
                                String::from("Test Variation 2-1"),
                                String::from("Test Variation 2-1-1"),
                                String::from("Test Variation 2-1-1-1"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("Test Description\nmore lines..."),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("3"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("Test Variation 2-1-1-2"),
                                String::default(),
                                String::default(),
                                String::default(),
                                String::from("Test Description\nmore lines..."),
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
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/multi_block.yml").unwrap()).unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/multi_block_multi_row.md").unwrap(),
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
                                    String::from("Test Variation A 1"),
                                    String::from("Test Variation A 1-1"),
                                    String::from("Test Variation A 1-1-1"),
                                    String::from("Test Variation A 1-1-1-1"),
                                    String::from("Test Variation A 1-1-1-1-1"),
                                    String::from("Test Variation A 1-1-1-1-1-1"),
                                    String::from("Test Variation A 1-1-1-1-1-1-1"),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("Test Variation A 2"),
                                    String::from("Test Variation A 2-1"),
                                    String::from("Test Variation A 2-1-1"),
                                    String::from("Test Variation A 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Variation A 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
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
                                    String::from("Test Variation B 1"),
                                    String::from("Test Variation B 1-1"),
                                    String::from("Test Variation B 1-1-1"),
                                    String::from("Test Variation B 1-1-1-1"),
                                    String::from("Test Variation B 1-1-1-1-1"),
                                    String::from("Test Variation B 1-1-1-1-1-1"),
                                    String::from("Test Variation B 1-1-1-1-1-1-1"),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("Test Variation B 2"),
                                    String::from("Test Variation B 2-1"),
                                    String::from("Test Variation B 2-1-1"),
                                    String::from("Test Variation B 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Variation B 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
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
                                    String::from("Test Variation C 1"),
                                    String::from("Test Variation C 1-1"),
                                    String::from("Test Variation C 1-1-1"),
                                    String::from("Test Variation C 1-1-1-1"),
                                    String::from("Test Variation C 1-1-1-1-1"),
                                    String::from("Test Variation C 1-1-1-1-1-1"),
                                    String::from("Test Variation C 1-1-1-1-1-1-1"),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("Test Variation C 2"),
                                    String::from("Test Variation C 2-1"),
                                    String::from("Test Variation C 2-1-1"),
                                    String::from("Test Variation C 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Variation C 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
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
    fn test_marshal_various_list() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/various_list.yml").unwrap()).unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/various_list.md").unwrap(),
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
                                    String::from("Test Variation 1"),
                                    String::from("Test Variation 1-1"),
                                    String::from("Test Variation 1-1-1"),
                                    String::from("Test Variation 1-1-1-1"),
                                    String::from("Test Variation 1-1-1-1-1"),
                                    String::from("Test Variation 1-1-1-1-1-1"),
                                    String::from("Test Variation 1-1-1-1-1-1-1"),
                                    String::from("Test Description\nmore lines..."),
                                    String::from("Procedure A-A\nProcedure A-B\nProcedure A-C"),
                                    String::from("2021/01/01"),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("Test Variation 2"),
                                    String::from("Test Variation 2-1"),
                                    String::from("Test Variation 2-1-1"),
                                    String::from("Test Variation 2-1-1-1"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
                                    String::from("Procedure B-A\nProcedure B-B"),
                                    String::from("2021/01/01"),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Variation 2-1-1-2"),
                                    String::default(),
                                    String::default(),
                                    String::default(),
                                    String::from("Test Description\nmore lines..."),
                                    String::from("Procedure"),
                                    String::from("2021/01/02"),
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
                                    String::from("cell 1"),
                                    String::default(),
                                    String::from("OK"),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("cell 2"),
                                    String::from("Description\nmore lines..."),
                                    String::from("NG"),
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
    fn test_marshal_various_list_multiple_sheet() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/various_list.yml").unwrap()).unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/various_list_multiple_sheet.md").unwrap(),
            rule,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![
                Sheet {
                    sheet_name: Some(String::from("Sheet Name 1")),
                    blocks: vec![
                        Block {
                            title: String::from("Block Title 1"),
                            rows: vec![
                                Row {
                                    columns: vec![
                                        String::from("1"),
                                        String::from("Test Variation 1"),
                                        String::from("Test Variation 1-1"),
                                        String::from("Test Variation 1-1-1"),
                                        String::from("Test Variation 1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1-1-1"),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure A-A\nProcedure A-B\nProcedure A-C"),
                                        String::from("2021/01/01"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("2"),
                                        String::from("Test Variation 2"),
                                        String::from("Test Variation 2-1"),
                                        String::from("Test Variation 2-1-1"),
                                        String::from("Test Variation 2-1-1-1"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure B-A\nProcedure B-B"),
                                        String::from("2021/01/01"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("3"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Variation 2-1-1-2"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure"),
                                        String::from("2021/01/02"),
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
                                        String::from("cell 1"),
                                        String::default(),
                                        String::from("OK"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("2"),
                                        String::from("cell 2"),
                                        String::from("Description\nmore lines..."),
                                        String::from("NG"),
                                    ],
                                },
                            ],
                        },
                    ],
                },
                Sheet {
                    sheet_name: Some(String::from("Sheet Name 2")),
                    blocks: vec![
                        Block {
                            title: String::from("Block Title 1"),
                            rows: vec![
                                Row {
                                    columns: vec![
                                        String::from("1"),
                                        String::from("Test Variation 1"),
                                        String::from("Test Variation 1-1"),
                                        String::from("Test Variation 1-1-1"),
                                        String::from("Test Variation 1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1-1"),
                                        String::from("Test Variation 1-1-1-1-1-1-1"),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure A-A\nProcedure A-B\nProcedure A-C"),
                                        String::from("2021/01/01"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("2"),
                                        String::from("Test Variation 2"),
                                        String::from("Test Variation 2-1"),
                                        String::from("Test Variation 2-1-1"),
                                        String::from("Test Variation 2-1-1-1"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure B-A\nProcedure B-B"),
                                        String::from("2021/01/01"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("3"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Variation 2-1-1-2"),
                                        String::default(),
                                        String::default(),
                                        String::default(),
                                        String::from("Test Description\nmore lines..."),
                                        String::from("Procedure"),
                                        String::from("2021/01/02"),
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
                                        String::from("cell 1"),
                                        String::default(),
                                        String::from("OK"),
                                    ],
                                },
                                Row {
                                    columns: vec![
                                        String::from("2"),
                                        String::from("cell 2"),
                                        String::from("Description\nmore lines..."),
                                        String::from("NG"),
                                    ],
                                },
                            ],
                        },
                    ],
                },
            ],
            mapping,
            rule: rule_clone,
        };
        assert_eq!(expected, data);
    }

    #[test]
    fn test_marshal_without_list() {
        let rule =
            Rule::marshal(&read_to_string("test_case/rule/without_list.yml").unwrap()).unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/without_list.md").unwrap(),
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
                                String::from("Test Variation 1"),
                                String::from("Test Variation 1-1"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("2"),
                                String::default(),
                                String::from("Test Variation 1-2"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("3"),
                                String::from("Test Variation 2"),
                                String::default(),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("4"),
                                String::from("Test Variation 3"),
                                String::from("Test Variation 3-1"),
                            ],
                        },
                        Row {
                            columns: vec![
                                String::from("5"),
                                String::default(),
                                String::from("Test Variation 3-2"),
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
    fn test_marshal_only_list() {
        let rule = Rule::marshal(&read_to_string("test_case/rule/only_list.yml").unwrap()).unwrap();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/only_list.md").unwrap(),
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
                                    String::from("cell A 1\nmore lines..."),
                                    String::from("cell B 1\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("cell A 2"),
                                    String::from("cell B 2"),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("3"),
                                    String::from("cell A 3\nmore lines..."),
                                    String::default(),
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
                                    String::from("another cell A 1\nmore lines..."),
                                    String::from("another cell B 1\nmore lines...\nmore lines..."),
                                ],
                            },
                            Row {
                                columns: vec![
                                    String::from("2"),
                                    String::from("another cell A 2"),
                                    String::default(),
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
    fn test_marshal_escape_asterisk() {
        let rule = get_default_rule();
        let rule_clone = rule.clone();
        let mapping = Mapping::new(&rule).unwrap();
        let data = Data::marshal(
            &read_to_string("test_case/input/escape_asterisk.md").unwrap(),
            rule,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![Sheet {
                sheet_name: Some(String::from("Sheet Name")),
                blocks: vec![Block {
                    title: String::from("Block Title"),
                    rows: vec![Row {
                        columns: vec![
                            String::from("1"),
                            String::from("Test Variation 1"),
                            String::from("Test Variation 1-1"),
                            String::from("Test Variation 1-1-1"),
                            String::default(),
                            String::default(),
                            String::default(),
                            String::default(),
                            String::from(
                                "Test Description_astarisk\nsingle *\ndouble **\nwith space * * *",
                            ),
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
    fn test_export_excel() {
        let rule = get_default_rule();
        let data = Data::marshal(
            &read_to_string("test_case/input/single_block_multi_row.md").unwrap(),
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
        // Some but target is empty
        let data = String::from("");
        let target = Some(&data);
        let input = "input";
        let result = Data::concat(&target, input);
        assert_eq!(String::from("input"), result);
    }
}
