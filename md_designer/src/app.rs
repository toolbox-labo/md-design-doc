use anyhow::Result;

use crate::{data::Data, rule::Rule};

pub struct App {
    pub data: Data,
}

impl App {
    pub fn new(input: &str, rule: Rule) -> Result<Self> {
        Ok(App {
            data: Data::marshal(input, rule)?,
        })
    }

    #[cfg(feature = "excel")]
    pub fn export_excel(&self) -> Result<()> {
        self.data.export_excel()?;
        Ok(())
    }
}
