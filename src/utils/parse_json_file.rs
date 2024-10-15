
use serde_json::Value;
use std::fs::File;
use std::io::BufReader;

// Function to parse the JSON file
pub fn parse_json_file(json_file_path: &str) -> Result<Value, Box<dyn std::error::Error>> {
    let file = File::open(json_file_path)?;
    let reader = BufReader::new(file);
    let data: Value = serde_json::from_reader(reader)?;
    Ok(data)
}
