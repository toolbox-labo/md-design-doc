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

    pub fn parse(input: &str) -> Result<()> {
        Ok(())
    }
}
