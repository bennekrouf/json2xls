
use serde_json::Value;
use xlsxwriter::{Workbook, Format};
use crate::models::ExcelConfig;
use crate::utils::extract_repository_row::extract_repository_row;
use crate::utils::write_repository_data::write_repository_data;
use crate::utils::write_headers::write_headers;
use crate::utils::extract_headers::extract_headers;
use tracing::{info, trace}; // Import trace

pub fn process_top_level_keys_for_summary(
    data: &Value,
    workbook: &Workbook,
    config: &ExcelConfig,
    title_format: &Format,
    light_blue_format: &Format,
    no_color_format: &Format,
    all_data_for_summary: &mut Vec<(String, Vec<String>, Vec<String>)>, // Collect all data for summary, with key and headers
    all_headers_for_summary: &mut Vec<String> // Collect all unique headers for the summary
) -> Result<(), Box<dyn std::error::Error>> {
    if let Value::Object(top_level) = data {
        for (key, repositories_value) in top_level.iter() {
            // Add a new worksheet for each key
            let mut worksheet = workbook.add_worksheet(Some(key))?;

            // Extract headers from the repositories
            let debt_headers = extract_headers(repositories_value.as_array().unwrap(), &config)?;
            info!("Result headers : {:?}", debt_headers);

            // Update the summary headers to include all unique headers
            for header in debt_headers.iter() {
                if !all_headers_for_summary.contains(header) { // Change to &header
                    trace!(header = %header, "Adding unique header to summary headers");
                    all_headers_for_summary.push(header.clone());
                }
            }

            // Write the headers to the worksheet
            write_headers(&mut worksheet, &debt_headers, title_format)?;

            // Write the repository data to each specific worksheet and collect for summary
            repositories_value.as_array().unwrap().iter().enumerate().for_each(|(i, repo)| {
                write_repository_data(repo, &mut worksheet, i + 1, &debt_headers, config, light_blue_format, no_color_format);

                // info!("")

                // Collect data for summary
                let repo_data = extract_repository_row(repo, &debt_headers, config);

                // info!(
                //     key = %key, 
                //     repo_data = ?repo_data, 
                //     debt_headers = ?debt_headers, 
                //     "Collecting data for summary"
                // );

                all_data_for_summary.push((key.clone(), debt_headers.clone(), repo_data)); // Collect key, headers, and row data
            });
        }
    }
    Ok(())
}

