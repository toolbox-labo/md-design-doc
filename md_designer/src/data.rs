use anyhow::{anyhow, Result};
use pulldown_cmark::{Event, Options, Parser, Tag};
use regex::Regex;

use crate::{mapping::Mapping, rule::Rule};

#[cfg(feature = "excel")]
use xlsxwriter::*;

enum State {
    Nothing,
    Heading(u32),
    List,
    //Check,
}

enum List {
    Nothing,
    Description,
    Checks,
    Procedure,
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
        let mapping = Mapping::new(rule).unwrap();

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

        parser.for_each(|event| {
            // is last line soft break
            // if true, next text data is append to current column
            //let mut is_fb = false;
            match event {
                Event::SoftBreak => {
                    //is_fb = true;
                }
                Event::Start(tag) => {
                    if last_is_list {
                        // start a new row
                        block.rows.push(row.clone());
                        row = Row::new(current_block, &mapping);
                        last_is_list = false;
                    }
                    if let Some(column_idx) = mapping.get_idx(current_block, &tag) {
                        current_column = *column_idx;
                    }
                }
                Event::Text(text) => {
                    row.columns[current_column] =
                        Data::concat(&row.columns.get(current_column), &text);
                }
                Event::End(Tag::List(_)) => {
                    last_is_list = true;
                }
                _ => {}
            }
        });
        // push the last row
        block.rows.push(row);
        sheet.blocks.push(block);

        println!("{:?}", sheet);

        Ok(Self {
            sheets: vec![sheet],
            rule,
            mapping,
        })
    }

    #[cfg(feature = "excel")]
    pub fn export_excel(&self) -> Result<()> {
        let workbook = Workbook::new("test.xlsx");
        self.sheets.iter().for_each(|sheet| {
            let mut s = workbook.add_worksheet(sheet.sheet_name.as_deref()).unwrap();
        })
    }

    /*
    pub fn marshal(input: &str, rule: &Rule) -> Result<Self> {
        // trim first empty lines
        let input = input.trim_start();

        // check is first line is Heading(1)
        // (sheet name is required)
        if !input.starts_with("# ") {
            return Err(anyhow!("input must start with '# ' (sheet name)."));
        }

        // marshal
        let mut options = Options::empty();
        options.insert(Options::ENABLE_TASKLISTS);
        let input = Data::custom_filter(input);
        let parser = Parser::new_ext(&input, options);
        let mut sheet = Sheet::default();
        let mut row = Row::default();
        let mut last_is_list = false;
        let mut state = State::Nothing;
        let mut current_list = List::Nothing;
        let mut soft_break = false;
        // expand parser to be able to handle 7th heading
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
        parser.for_each(|event| {
            let mut is_fb = false;
            match event {
                Event::SoftBreak => {
                    is_fb = true;
                }
                Event::Start(tag) => {
                    if last_is_list {
                        // start a new row
                        sheet.rows.push(row.clone());
                        row = Row::default();
                        last_is_list = false;
                    }
                    match tag {
                        Tag::Heading(num) => state = State::Heading(num),
                        Tag::List(_) => state = State::List,
                        _ => {}
                    }
                }
                Event::Text(text) => {
                    if soft_break {
                        match current_list {
                            List::Description => {
                                row.description = Some(Data::concat(&row.description, &text));
                            }
                            List::Checks => {
                                row.checks = Some(Data::concat(&row.checks, &text));
                            }
                            List::Procedure => {
                                row.procedure = Some(Data::concat(&row.procedure, &text));
                            }
                            _ => {}
                        }
                    } else {
                        match state {
                            State::Heading(num) => match num {
                                1 => sheet.sheet_name = Some(text.to_string()),
                                2 => {
                                    row.variation_1 = Some(Data::concat(&row.variation_1, &text));
                                }
                                3 => {
                                    row.variation_2 = Some(Data::concat(&row.variation_2, &text));
                                }
                                4 => {
                                    row.variation_3 = Some(Data::concat(&row.variation_3, &text));
                                }
                                5 => {
                                    row.variation_4 = Some(Data::concat(&row.variation_4, &text));
                                }
                                6 => {
                                    row.variation_5 = Some(Data::concat(&row.variation_5, &text));
                                }
                                7 => {
                                    row.variation_6 = Some(Data::concat(&row.variation_6, &text));
                                }
                                8 => {
                                    row.variation_7 = Some(Data::concat(&row.variation_7, &text));
                                }
                                _ => {}
                            },
                            State::List => {
                                if let Some(txt) = text.strip_prefix("!!DSC!!") {
                                    // description
                                    row.description = Some(Data::concat(&row.description, &txt));
                                    current_list = List::Description;
                                } else if let Some(txt) = text.strip_prefix("!!CHK!!") {
                                    // checks
                                    row.checks = Some(Data::concat(&row.checks, &txt));
                                    current_list = List::Checks;
                                } else {
                                    // procedure
                                    row.procedure = Some(Data::concat(&row.procedure, &text));
                                    current_list = List::Procedure;
                                }
                            }
                            _ => {}
                        }
                    }
                }
                Event::End(Tag::List(_)) => {
                    last_is_list = true;
                }
                _ => {}
            }
            soft_break = is_fb;
        });
        // push the last row
        sheet.rows.push(row);

        Ok(Self {
            sheets: vec![sheet],
        })
    }
    */

    /*
    #[cfg(feature = "excel")]
    pub fn export_excel(&self) -> Result<()> {
        let workbook = Workbook::new("test.xlsx");
        self.sheets.iter().for_each(|sheet| {
            let mut s = workbook.add_worksheet(sheet.sheet_name.as_deref()).unwrap();
            let wrap_format = workbook.add_format().set_text_wrap();
            // header
            s.merge_range(0, 0, 1, 0, "試験項番", None).unwrap();
            s.merge_range(0, 1, 0, 7, "試験バリエーション", None)
                .unwrap();
            s.write_string(1, 1, "項目1", None).unwrap();
            s.write_string(1, 2, "項目2", None).unwrap();
            s.write_string(1, 3, "項目3", None).unwrap();
            s.write_string(1, 4, "項目4", None).unwrap();
            s.write_string(1, 5, "項目5", None).unwrap();
            s.write_string(1, 6, "項目6", None).unwrap();
            s.write_string(1, 7, "項目7", None).unwrap();
            s.merge_range(0, 8, 1, 8, "試験概要", None).unwrap();
            s.merge_range(0, 9, 1, 9, "試験手順", None).unwrap();
            s.merge_range(0, 10, 1, 10, "確認内容", None).unwrap();
            s.merge_range(0, 11, 1, 11, "優先度", None).unwrap();
            for i in 0..=2 {
                let title = format!("{}回目", i + 1);
                let offset = 4 * i;
                s.merge_range(0, 12 + offset, 0, 15 + offset, title.as_str(), None)
                    .unwrap();
                s.write_string(1, 12 + offset, "試験予定日", None).unwrap();
                s.write_string(1, 13 + offset, "試験実施日", None).unwrap();
                s.write_string(1, 14 + offset, "試験者", None).unwrap();
                s.write_string(1, 15 + offset, "試験結果", None).unwrap();
            }
            // body
            sheet.rows.iter().enumerate().for_each(|(i, row)| {
                let current_row_idx = i + 2;
                s.write_string(
                    current_row_idx as u32,
                    0,
                    (i + 1).to_string().as_str(),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    1,
                    &row.variation_1.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    2,
                    &row.variation_2.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    3,
                    &row.variation_3.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    4,
                    &row.variation_4.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    5,
                    &row.variation_5.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    6,
                    &row.variation_6.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    7,
                    &row.variation_7.as_ref().unwrap_or(&"".to_string()),
                    None,
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    8,
                    &row.description.as_ref().unwrap_or(&"".to_string()),
                    Some(&wrap_format),
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    9,
                    &row.procedure.as_ref().unwrap_or(&"".to_string()),
                    Some(&wrap_format),
                )
                .unwrap();
                s.write_string(
                    current_row_idx as u32,
                    10,
                    &row.checks.as_ref().unwrap_or(&"".to_string()),
                    Some(&wrap_format),
                )
                .unwrap();
            });
        });
        Ok(())
    }
    */

    fn custom_filter(input: &str) -> String {
        let list_1 = Regex::new(r" *(\* )").unwrap();
        let input = list_1.replace_all(&input, "- !!DSC!!");
        let list_1 = Regex::new(r" *(- \[[ |\*]\])").unwrap();
        let input = list_1.replace_all(&input, "- !!CHK!!");
        input.to_string()
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

/*
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_marshal() {
        let data = Data::marshal(
            r#"
# Sheet Name
## Test Variation 1
### Test Variation 1-1
#### Test Variation 1-1-1
* Test Description
  more lines...
- Test Procedure(1)
- Test Procedure(2)
- [ ] Confirmation item(1)
- [ ] Confirmation item(2)
"#,
        )
        .unwrap();
        let expected = Data {
            sheets: vec![Sheet {
                sheet_name: Some(String::from("Sheet Name")),
                rows: vec![Row {
                    variation_1: Some(String::from("Test Variation 1")),
                    variation_2: Some(String::from("Test Variation 1-1")),
                    variation_3: Some(String::from("Test Variation 1-1-1")),
                    description: Some(String::from("Test Description\nmore lines...")),
                    procedure: Some(String::from("Test Procedure(1)\nTest Procedure(2)")),
                    checks: Some(String::from(" Confirmation item(1)\n Confirmation item(2)")),
                    ..Default::default()
                }],
            }],
        };
        assert_eq!(expected, data);
    }

    #[test]
    fn test_custom_filter() {
        let input = r#"
# Sheet1
## variation 1
### variation 2
#### variation 3
* Description
  lines...
  lines...
- Procedure 1
- Procedure 2
- [ ] check 1
- [*] check 2
"#;
        let expected = r#"
# Sheet1
## variation 1
### variation 2
#### variation 3
- !!DSC!!Description
  lines...
  lines...
- Procedure 1
- Procedure 2
- !!CHK!! check 1
- !!CHK!! check 2
"#;

        assert_eq!(expected, Data::custom_filter(input));
    }

    #[test]
    fn test_concat() {
        // None
        let target = None;
        let input = "input";
        let result = Data::concat(&target, input);
        assert_eq!(String::from("input"), result);
        // Some
        let target = Some("target".to_string());
        let input = "input";
        let result = Data::concat(&target, input);
        assert_eq!(String::from("target\ninput"), result);
    }
}
*/
