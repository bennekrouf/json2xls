
use serde_json::Value;
use std::fs::File;
use std::io::BufReader;
use xlsxwriter::{Workbook, Format, format::FormatAlignment, format::FormatColor, Worksheet};
use std::collections::HashMap;

// Externalized color constants
const TITLE_BG_COLOR: u32 = 0x565E73;  // Blue background for title
const TITLE_FONT_COLOR: FormatColor = FormatColor::White;  // White font for title
const ROW_ALT_COLOR: u32 = 0xADD8E6;   // Light blue background for alternating rows

// Function to merge cells and write headers
fn write_merged_header(
    worksheet: &mut Worksheet,
    start_row: u32,
    start_col: u16,
    end_col: u16,
    text: &str,
    format: &Format,
) -> Result<(), Box<dyn std::error::Error>> {
    worksheet.merge_range(start_row, start_col, start_row, end_col, text, Some(format))?;
    Ok(())
}

// Function to generate an Excel file from JSON input
fn generate_excel_from_json(json_file_path: &str, output_excel_file: &str) -> Result<(), Box<dyn std::error::Error>> {
    // Parse the JSON file
    let file = File::open(json_file_path)?;
    let reader = BufReader::new(file);
    
    // Deserialize into a dynamic `serde_json::Value`
    let data: Value = serde_json::from_reader(reader)?;

    // Create a new Excel workbook
    let workbook = Workbook::new(output_excel_file)?;

    // Create format for title headers (white font, blue background)
    let mut title_format = Format::new();
    title_format
        .set_bold()
        .set_bg_color(FormatColor::Custom(TITLE_BG_COLOR))   // Externalized title background color
        .set_font_color(TITLE_FONT_COLOR)                    // Externalized font color
        .set_align(FormatAlignment::Center);

    // Create format for alternating rows (custom light blue background)
    let mut light_blue_format = Format::new();
    light_blue_format.set_bg_color(FormatColor::Custom(ROW_ALT_COLOR));  // Externalized alternating row color

    // Create a format for regular rows (no background color)
    let no_color_format = Format::new();

    // Start processing each top-level key (e.g., ACM, BPM)
    if let Value::Object(top_level) = data {
        for (key, repositories_value) in top_level {
            // Add a worksheet for each top-level key
            let mut worksheet = workbook.add_worksheet(Some(&key))?;

            // Ensure repositories_value is an array
            if let Value::Array(repositories) = repositories_value {
                // Collect all unique keys from the "debt" array for headers
                let mut debt_headers = vec!["repository_name".to_string()];
                if let Some(Value::Object(repo)) = repositories.get(0) {
                    if let Some(Value::Array(debt_array)) = repo.get("debt") {
                        if let Some(Value::Object(debt)) = debt_array.get(0) {
                            for debt_key in debt.keys() {
                                debt_headers.push(debt_key.clone());
                            }
                        }
                    }
                }

                // Write the headers for the worksheet
                for (col, header) in debt_headers.iter().enumerate() {
                    worksheet.write_string(0, col as u16, header, Some(&title_format))?;
                }

                // Write the data into the worksheet starting from row 1 (index 0)
                let mut row = 1;
                for repository_value in repositories {
                    if let Value::Object(repository) = repository_value {
                        if let Some(repo_name) = repository.get("repository_name") {
                            if let Value::String(ref repo_name) = repo_name {
                                if let Some(Value::Array(debts)) = repository.get("debt") {
                                    for (i, debt_value) in debts.iter().enumerate() {
                                        if let Value::Object(debt) = debt_value {
                                            // Alternate between light blue and no color for the rows
                                            let format = if i % 2 == 0 {
                                                &light_blue_format
                                            } else {
                                                &no_color_format
                                            };

                                            // Write repository_name in the first column
                                            worksheet.write_string(row, 0, &repo_name, Some(format))?;

                                            // Write debt details in subsequent columns
                                            for (j, debt_key) in debt_headers.iter().skip(1).enumerate() {
                                                if let Some(Value::String(debt_value)) = debt.get(debt_key) {
                                                    worksheet.write_string(row, (j + 1) as u16, debt_value, Some(format))?;
                                                }
                                            }
                                            row += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Close the workbook to finalize the file
    workbook.close()?;

    println!("Excel file '{}' created successfully!", output_excel_file);

    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Specify the JSON file path and the output Excel file path
    let json_file_path = "all_projects.json";
    let output_excel_file = "all_output.xlsx";

    // Generate Excel from the JSON input
    generate_excel_from_json(json_file_path, output_excel_file)?;

    Ok(())
}

