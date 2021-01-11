use anyhow::Result;
use pulldown_cmark::{html, Options, Parser};

use crate::data::Data;

pub struct App {
    data: Data,
}

impl App {
    pub fn new(input: &str) -> Result<Self> {
        Ok(App {
            data: Data::marshal(input)?,
        })
    }

    pub fn export_excel(&self) -> Result<()> {
        self.data.export_excel()?;
        Ok(())
    }
}
