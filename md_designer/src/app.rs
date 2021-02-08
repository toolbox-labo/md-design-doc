use anyhow::Result;

use crate::{data::Data, rule::Rule};

pub struct App {
    pub data: Data,
    file_name: String,
}

impl App {
    pub fn new(file_name: &str, input: &str, rule: Rule) -> Result<Self> {
        Ok(App {
            data: Data::marshal(input, rule)?,
            file_name: file_name.to_string(),
        })
    }

    #[cfg(feature = "excel")]
    pub fn export_excel(&self) -> Result<()> {
        self.data.export_excel(&self.file_name)?;
        Ok(())
    }
}
