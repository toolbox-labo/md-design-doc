use anyhow::Result;

use crate::data::Data;

pub struct App {
    pub data: Data,
}

impl App {
    pub fn new(input: &str) -> Result<Self> {
        Ok(App {
            data: Data::marshal(input)?,
        })
    }

    #[cfg(feature="excel")]
    pub fn export_excel(&self) -> Result<()> {
        self.data.export_excel()?;
        Ok(())
    }
}
