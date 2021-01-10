use anyhow::{anyhow, Result};
use pulldown_cmark::{html, Event, Options, Parser, Tag};
use regex::Regex;

enum State {
    Nothing,
    Heading(u32),
    List,
    Check,
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
}

impl Default for Data {
    fn default() -> Self {
        Self { sheets: vec![] }
    }
}

impl Data {
    pub fn marshal(input: &str) -> Result<(Self)> {
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
                    return event;
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
                    }
                    match tag {
                        Tag::Heading(num) => { state = State::Heading(num) },
                        Tag::List(_) => state = State::List,
                        _ => {}
                    }
                }
                Event::Text(text) => {
                    if soft_break {
                        match current_list {
                            List::Description => {
                                row.description = Some(format!(
                                    "{}\n{}",
                                    row.description.as_ref().unwrap_or(&"".to_string()),
                                    &text
                                ));
                            }
                            List::Checks => {
                                row.checks = Some(format!(
                                    "{}\n{}",
                                    row.checks.as_ref().unwrap_or(&"".to_string()),
                                    &text
                                ));
                            }
                            List::Procedure => {
                                row.procedure = Some(format!(
                                    "{}\n{}",
                                    row.procedure.as_ref().unwrap_or(&"".to_string()),
                                    &text
                                ));
                            }
                            _ => {}
                        }
                    } else {
                        match state {
                            State::Heading(num) => match num {
                                1 => {
                                    sheet.sheet_name = Some(text.to_string())
                                }
                                2 => {
                                    row.variation_1 = Some(format!(
                                        "{}\n{}",
                                        row.variation_1.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                3 => {
                                    row.variation_2 = Some(format!(
                                        "{}\n{}",
                                        row.variation_2.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                4 => {
                                    row.variation_3 = Some(format!(
                                        "{}\n{}",
                                        row.variation_3.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                5 => {
                                    row.variation_4 = Some(format!(
                                        "{}\n{}",
                                        row.variation_4.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                6 => {
                                    row.variation_5 = Some(format!(
                                        "{}\n{}",
                                        row.variation_5.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                7 => {
                                    row.variation_6 = Some(format!(
                                        "{}\n{}",
                                        row.variation_6.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                8 => {
                                    row.variation_7 = Some(format!(
                                        "{}\n{}",
                                        row.variation_7.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ))
                                }
                                _ => {}
                            },
                            State::List => {
                                if text.starts_with("!!DSC!!") {
                                    // description
                                    row.description = Some(format!(
                                        "{}\n{}",
                                        row.description.as_ref().unwrap_or(&"".to_string()),
                                        &text[7..]
                                    ));
                                    current_list = List::Description;
                                } else if text.starts_with("!!CHK!!") {
                                    // checks
                                    row.checks = Some(format!(
                                        "{}\n{}",
                                        row.checks.as_ref().unwrap_or(&"".to_string()),
                                        &text[7..]
                                    ));
                                    current_list = List::Checks;
                                } else {
                                    // procedure
                                    row.procedure = Some(format!(
                                        "{}\n{}",
                                        row.procedure.as_ref().unwrap_or(&"".to_string()),
                                        text
                                    ));
                                    current_list = List::Procedure;
                                }
                            }
                            _ => {}
                        }
                    }
                }
                Event::End(tag) => match tag {
                    Tag::List(_) => {
                        last_is_list = true;
                    }
                    _ => {}
                },
                _ => {}
            }
            soft_break = is_fb;
        });
        // push the last row
        sheet.rows.push(row.clone());

        Ok(Self { sheets: vec![sheet] })
    }

    fn custom_filter(input: &str) -> String {
        let list_1 = Regex::new(r" *(\* )").unwrap();
        let input = list_1.replace_all(&input, "- !!DSC!!");
        let list_1 = Regex::new(r" *(- \[[ |\*]\])").unwrap();
        let input = list_1.replace_all(&input, "- !!CHK!!");
        input.to_string()
    }
}

#[derive(Debug, PartialEq)]
struct Sheet {
    sheet_name: Option<String>,
    rows: Vec<Row>,
}

impl Default for Sheet {
    fn default() -> Self {
        Self { sheet_name: None, rows: vec![] }
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Row {
    variation_1: Option<String>,
    variation_2: Option<String>,
    variation_3: Option<String>,
    variation_4: Option<String>,
    variation_5: Option<String>,
    variation_6: Option<String>,
    variation_7: Option<String>,
    description: Option<String>,
    procedure: Option<String>,
    checks: Option<String>,
    is_priority_high: bool,
}

impl Default for Row {
    fn default() -> Self {
        Self {
            variation_1: None,
            variation_2: None,
            variation_3: None,
            variation_4: None,
            variation_5: None,
            variation_6: None,
            variation_7: None,
            description: None,
            procedure: None,
            checks: None,
            is_priority_high: false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_marshal() {
        let data = Data::marshal(r#"
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
"#
        ).unwrap();
        let expected = Data {
            sheets: vec![
                Sheet {
                    sheet_name: Some(String::from("Sheet Name")),
                    rows: vec![
                        Row {
                            variation_1: Some(String::from("Test Variation 1")),
                            variation_2: Some(String::from("Test Variation 1-1")),
                            variation_3: Some(String::from("Test Variation 1-1-1")),
                            description: Some(String::from("Test Description\nmore lines...")),
                            procedure: Some(String::from("Test Procedure(1)\nTest Procedure(2)")),
                            checks: Some(String::from("Confirmation item(1)\nConfirmation item(2)")),
                            ..Default::default()
                        }
                    ]
                }
            ]
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
}
