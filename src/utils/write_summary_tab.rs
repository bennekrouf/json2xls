
use xlsxwriter::Format;
use crate::utils::write_headers::write_headers;
use tracing::trace;
use tracing::info;

pub fn write_summary_tab(
    summary_worksheet: &mut xlsxwriter::Worksheet,
    all_data_for_summary: &[(String, Vec<String>, Vec<String>)], // (Tab name, headers, row data)
    all_headers_for_summary: &[String], // All unique headers for the summary
    title_format: &Format,
    row_format: &Format
) -> Result<(), Box<dyn std::error::Error>> {
    if !all_data_for_summary.is_empty() {
        // First, write the headers, including the "Group" column
        let mut full_headers = vec!["Group".to_string()];
        full_headers.extend(all_headers_for_summary.iter().cloned());

        info!("Writing headers: {:?}", full_headers);
        write_headers(summary_worksheet, &full_headers, title_format)?;

        // Write each row, aligning the data with the headers
        for (row, (group_name, repo_headers, repo_data)) in all_data_for_summary.iter().enumerate() {
            // trace!(row = row + 1, group_name = %group_name, "Writing group name to row");

            summary_worksheet.write_string((row + 1) as u32, 0, group_name, Some(row_format))?; // Write the group (tab) name
// info!("all headers for summary : {:?} repo data: {:?}", all_headers_for_summary, repo_data);
            // For each summary header, find the corresponding data in the repo_data based on repo_headers
            for (col, summary_header) in all_headers_for_summary.iter().enumerate() {
                let col_index = repo_headers.iter().position(|h| h == summary_header);
                let cell_value = col_index.map_or("", |index| &repo_data[index]);

                // info!(
                //     row = row + 1,
                //     column = col + 1,
                //     header = %summary_header,
                //     col_index = ?col_index,
                //     value = %cell_value,
                //     "Writing data to worksheet"
                // );

                summary_worksheet.write_string((row + 1) as u32, (col + 1) as u16, cell_value, Some(row_format))?;
            }
        }
    } else {
        trace!("No data available for summary tab");
    }

    Ok(())
}

