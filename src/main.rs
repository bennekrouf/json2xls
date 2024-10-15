
mod generate_excel_from_json;
mod models;

use crate::generate_excel_from_json::generate_excel_from_json;
use crate::models::ExcelConfig;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create the Excel configuration
    let config = ExcelConfig {
        json_file_path: "all_projects.json".to_string(),
        output_excel_file: "all_output.xlsx".to_string(),
        title_bg_color: 0x565E73,        // Blue background for title
        title_font_color: 0xFFFFFF,      // White font for title using hex code
        row_alt_color: 0xADD8E6,         // Light blue background for alternating rows
        repository_name_key: "application".to_string(),  // JSON key for repository name
        object_array_key: "debt".to_string(),            // JSON key for debt
    };

    // Generate Excel from the JSON input
    generate_excel_from_json(config)?;

    Ok(())
}

