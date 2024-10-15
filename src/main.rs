
mod models;
mod utils;

use crate::utils::generate_excel_from_json::generate_excel_from_json;
use crate::models::ExcelConfig;
use std::fs::File;
use std::io::BufReader;
use serde_yaml::from_reader;
use tracing_subscriber;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    tracing_subscriber::fmt()
        .with_max_level(tracing::Level::INFO)  // Set the maximum level to TRACE
        .init();
    // Load the configuration from the YAML file
    let file = File::open("configuration.yml")?;
    let reader = BufReader::new(file);
    let config: ExcelConfig = from_reader(reader)?;

    // Generate Excel from the JSON input
    generate_excel_from_json(config)?;

    Ok(())
}

