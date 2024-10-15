
use xlsxwriter::Workbook;

use crate::models::ExcelConfig;
use crate::utils::write_summary_tab::write_summary_tab;
use crate::utils::process_top_level_keys_for_summary::process_top_level_keys_for_summary;
use crate::utils::create_format::create_format;
use crate::utils::parse_json_file::parse_json_file;

pub fn generate_excel_from_json(
    config: ExcelConfig
) -> Result<(), Box<dyn std::error::Error>> {
    // Load and parse the JSON file
    let data = parse_json_file(&config.json_file_path)?;

    // Create a new Excel workbook
    let workbook = Workbook::new(&config.output_excel_file)?;

    // Prepare formats
    let title_format = create_format(&workbook, true, config.title_bg_color, config.title_font_color, true);
    let light_blue_format = create_format(&workbook, false, config.row_alt_color, 0, false);
    let no_color_format = create_format(&workbook, false, 0, 0, false);

    // **NEW**: Create the Summary worksheet first
    let mut summary_worksheet = workbook.add_worksheet(Some("Summary"))?;

    // Collections for accumulating all the rows and headers for the summary
    let mut all_data_for_summary: Vec<(String, Vec<String>, Vec<String>)> = vec![];
    let mut all_headers_for_summary: Vec<String> = vec![];

    // Process each top-level key (e.g., ACM, BPM) and create corresponding worksheets
    process_top_level_keys_for_summary(
        &data,
        &workbook,
        &config,
        &title_format,
        &light_blue_format,
        &no_color_format,
        &mut all_data_for_summary,
        &mut all_headers_for_summary // Add this missing vector
    )?;

    // **NEW**: Now write all accumulated data to the Summary worksheet
    write_summary_tab(
        &mut summary_worksheet,
        &all_data_for_summary,
        &all_headers_for_summary, // Pass the headers to the summary tab
        &title_format,
        &light_blue_format
    )?;

    // Finalize the workbook
    workbook.close()?;
    println!("Excel file '{}' created successfully!", config.output_excel_file);

    Ok(())
}

