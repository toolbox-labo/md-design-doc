#![warn(rust_2018_idioms)]

use std::fs;

use anyhow::Result;
use clap::{crate_authors, crate_description, crate_name, crate_version, App as ClapApp, Arg};

use md_designer::app::App;

fn main() -> Result<()> {
    // setup clap
    let clap = ClapApp::new(crate_name!())
        .author(crate_authors!())
        .version(crate_version!())
        .about(crate_description!())
        .arg(Arg::with_name("path").required(true).help("input file path (.md)"))
        .get_matches();

    // check user input
    let input_text = if let Some(path) = clap.value_of("path") {
        fs::read_to_string(path)?
    } else {
        "".to_string()
    };

    let app = App::new(&input_text)?;

    app.export_excel()?;

    Ok(())
}
