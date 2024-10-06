mod generate_excel_from_json;

use crate::generate_excel_from_json::generate_excel_from_json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Specify the JSON file path and the output Excel file path
    let json_file_path = "all_projects.json";
    let output_excel_file = "all_output.xlsx";

    // Generate Excel from the JSON input
    generate_excel_from_json(json_file_path, output_excel_file)?;

    Ok(())
}

