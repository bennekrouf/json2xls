
use std::fs::File;
use std::io::BufReader;
use serde::Deserialize;
use std::collections::HashMap;
use xlsxwriter::{Workbook, Format, format::FormatAlignment, format::FormatColor, Worksheet};

// Externalized color constants
const TITLE_BG_COLOR: u32 = 0x565E73;  // Blue background for title
const TITLE_FONT_COLOR: FormatColor = FormatColor::White;  // White font for title
const ROW_ALT_COLOR: u32 = 0xADD8E6;   // Light blue background for alternating rows

// Struct to represent debt inside each repository
#[derive(Deserialize)]
struct Debt {
    cycle: String,
    product: String,
    reason: String,
    status: String,
}

// Struct to represent each repository and its debts
#[derive(Deserialize)]
struct Repository {
    repository_name: String,
    debt: Vec<Debt>,
}

// Type alias for the entire data structure: top-level keys map to a list of repositories
type RepoMap = HashMap<String, Vec<Repository>>;

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
    
    // Deserialize into the RepoMap struct that matches the JSON structure
    let data: RepoMap = serde_json::from_reader(reader)?;

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
    let mut row = 0;
    for (key, repositories) in data {
        // Add a worksheet for each top-level key
        let mut worksheet = workbook.add_worksheet(Some(&key))?;

        // Write the first column header "repository_name"
        worksheet.write_string(0, 0, "repository_name", Some(&title_format))?;

        // Merge columns 1 to 4 for "debt" header with blue background and white text
        write_merged_header(&mut worksheet, 0, 1, 4, "debt", &title_format)?;

        // Write sub-headers for the "debt" section: cycle, product, reason, status with blue background
        worksheet.write_string(1, 1, "cycle", Some(&title_format))?;
        worksheet.write_string(1, 2, "product", Some(&title_format))?;
        worksheet.write_string(1, 3, "reason", Some(&title_format))?;
        worksheet.write_string(1, 4, "status", Some(&title_format))?;

        // Write the data into the worksheet starting from row 2 (index 1)
        row = 2;
        for repository in repositories {
            for (i, debt) in repository.debt.into_iter().enumerate() {
                // Alternate between light blue and no color for the rows
                let format = if i % 2 == 0 {
                    &light_blue_format
                } else {
                    &no_color_format
                };

                worksheet.write_string(row, 0, &repository.repository_name, Some(format))?;
                worksheet.write_string(row, 1, &debt.cycle, Some(format))?;
                worksheet.write_string(row, 2, &debt.product, Some(format))?;
                worksheet.write_string(row, 3, &debt.reason, Some(format))?;
                worksheet.write_string(row, 4, &debt.status, Some(format))?;
                row += 1;
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
    let output_excel_file = "acm_output.xlsx";

    // Generate Excel from the JSON input
    generate_excel_from_json(json_file_path, output_excel_file)?;

    Ok(())
}

