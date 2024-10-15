use serde_json::Value;
use std::fs::File;
use std::io::BufReader;
use xlsxwriter::{Workbook, Format, format::FormatAlignment, format::FormatColor};

use crate::models::ExcelConfig;

pub fn generate_excel_from_json(
    config: ExcelConfig
) -> Result<(), Box<dyn std::error::Error>> {
    // Parse the JSON file
    let json_file_path = &config.json_file_path;
    let output_excel_file = &config.output_excel_file;
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
        .set_bg_color(FormatColor::Custom(config.title_bg_color))   // Use config title background color
        .set_font_color(config.title_font_color)                    // Use config font color
        .set_align(FormatAlignment::Center);

    // Create format for alternating rows (custom light blue background)
    let mut light_blue_format = Format::new();
    light_blue_format.set_bg_color(FormatColor::Custom(config.row_alt_color));  // Use config alternating row color

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
                let mut debt_headers = vec![config.repository_name_key.clone()];
                if let Some(Value::Object(repo)) = repositories.get(0) {
                    if let Some(Value::Array(debt_array)) = repo.get(&config.object_array_key) {
                        if let Some(Value::Object(debt)) = debt_array.get(0) {
                            for object_array_key in debt.keys() {
                                debt_headers.push(object_array_key.clone());
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
                        if let Some(repo_name) = repository.get(&config.repository_name_key) {
                            if let Value::String(ref repo_name) = repo_name {
                                if let Some(Value::Array(debts)) = repository.get(&config.object_array_key) {
                                    for (i, debt_value) in debts.iter().enumerate() {
                                        if let Value::Object(debt) = debt_value {
                                            // Alternate between light blue and no color for the rows
                                            let format = if i % 2 == 0 {
                                                &light_blue_format
                                            } else {
                                                &no_color_format
                                            };

                                            // Write repository_name in the first column
                                            worksheet.write_string(row, 0, repo_name, Some(format))?;

                                            // Write debt details in subsequent columns
                                            for (j, object_array_key) in debt_headers.iter().skip(1).enumerate() {
                                                if let Some(Value::String(debt_value)) = debt.get(object_array_key) {
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
