#![warn(rust_2018_idioms)]

use std::{env, fs, io::Write, path::Path};

use anyhow::{Context, Result};
use chrono::Local;
use clap::{crate_authors, crate_description, crate_name, crate_version, App as ClapApp, Arg};
use log::{debug, info};

use md_designer::{app::App, rule::Rule, utils::get_output_filename};

fn main() -> Result<()> {
    // setup clap
    let clap = ClapApp::new(crate_name!())
        .author(crate_authors!())
        .version(crate_version!())
        .about(crate_description!())
        .arg(
            Arg::with_name("path")
                .required(true)
                .help("input file path (.md)"),
        )
        .arg(
            Arg::with_name("conf_path")
                .required(true)
                .help("config file path (.yml)"),
        )
        .arg(
            Arg::with_name("output_filename")
                .short("o")
                .takes_value(true)
                .help("output file name. '.xlsx' is optional."),
        )
        .arg(
            Arg::with_name("verbose")
                .short("v")
                .help("verbose (print errors/warnings/info logs)"),
        )
        .arg(
            Arg::with_name("very_verbose")
                .short("d")
                .help("very verbose (also print debug logs)"),
        )
        .get_matches();

    // setup logging
    let log_level = {
        if clap.is_present("verbose") {
            Some("info")
        } else if clap.is_present("very_verbose") {
            Some("debug")
        } else {
            None
        }
    };
    if let Some(level) = log_level {
        env::set_var("RUST_LOG", level);
        env_logger::Builder::from_default_env()
            .format(|buf, record| {
                writeln!(
                    buf,
                    "{} [{}] - {}",
                    Local::now().format("%Y-%m-%dT%H:%M:%S"),
                    record.level(),
                    record.args(),
                )
            })
            .init();
    }

    let path = Path::new(clap.value_of("path").unwrap());
    info!("input file: {:?}", &path);
    let input_text = fs::read_to_string(&path)?;
    debug!("input file content: \n{}", &input_text);
    let cfg_path = Path::new(clap.value_of("conf_path").unwrap());
    info!("rule file: {:?}", &cfg_path);
    let cfg_text = fs::read_to_string(&cfg_path)?;
    debug!("rule file content: \n{}", &cfg_text);

    let rule = Rule::marshal(&cfg_text)?;

    let app = App::new(
        get_output_filename(
            clap.value_of("output_filename").unwrap_or(
                path.file_stem()
                    .with_context(|| "Input file path is malformed")?
                    .to_str()
                    .unwrap(),
            ),
        )?,
        &input_text,
        rule,
    )?;

    app.export_excel()?;

    info!("DONE");
    Ok(())
}
