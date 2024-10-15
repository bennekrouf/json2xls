
mod generate_excel_from_json;
mod models;

use crate::generate_excel_from_json::generate_excel_from_json;
use crate::models::ExcelConfig;
use std::fs::File;
use std::io::BufReader;
use serde_yaml::from_reader;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load the configuration from the YAML file
    let file = File::open("configuration.yml")?;
    let reader = BufReader::new(file);
    let config: ExcelConfig = from_reader(reader)?;

    // Generate Excel from the JSON input
    generate_excel_from_json(config)?;

    Ok(())
}

