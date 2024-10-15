
use serde_json::Value;
use xlsxwriter::Format;

use crate::models::ExcelConfig;

// Function to write repository data into the worksheet
use tracing::{trace, warn};
pub fn write_repository_data(
    repository_value: &Value,
    worksheet: &mut xlsxwriter::Worksheet,
    mut row: usize,  // make row mutable to increment it
    debt_headers: &[String],
    config: &ExcelConfig,
    light_blue_format: &Format,
    no_color_format: &Format
) {
    // Check if the repository is a JSON object
    if let Value::Object(repository) = repository_value {
        // Get repository name
        if let Some(Value::String(repo_name)) = repository.get(&config.repository_name_key) {
            trace!(repository_name = %repo_name, "Processing repository");

            // Get the debts array from the repository object
            if let Some(Value::Array(debts)) = repository.get(&config.object_array_key) {
                // trace!(?debts, "Debts array found");

                // Process each debt object within the debts array
                debts.iter().enumerate().for_each(|(i, debt)| {
                    let format = if i % 2 == 0 { light_blue_format } else { no_color_format };

                    // Write repository name to the first column
                    if let Err(e) = worksheet.write_string(row as u32, 0, repo_name, Some(format)) {
                        warn!(error = %e, "Failed to write repository name");
                    }

                    // Log the current debt being processed
                    // trace!(debt_index = i, ?debt, "Processing debt");

                    // Iterate through each debt header, skipping the repository name header
                    debt_headers.iter().skip(1).enumerate().for_each(|(j, header)| {
                        // Try to fetch the value associated with the current header in the debt object
                        if let Some(Value::String(debt_value)) = debt.get(header) {

                            // Write the debt value to the respective column
                            if let Err(e) = worksheet.write_string(row as u32, (j + 1) as u16, debt_value, Some(format)) {
                                warn!(header = %header, debt_index = i, error = %e, "Failed to write debt value");
                            } else {
                                // Log that the debt value has been written successfully
                                // trace!(header = %header, value = %debt_value, "Debt value written to worksheet");
                            }
                        } else {
                            // Log a warning if the header is missing in the current debt object
                            warn!(header = %header, debt_index = i, "Missing value for header in debt");
                        }
                    });

                    // Increment the row for the next debt
                    row += 1;
                });
            } else {
                // Log a warning if the debts array is not found
                warn!(repository_name = %repo_name, "No debts array found for repository");
            }
        } else {
            // Log a warning if the repository name is missing
            warn!("Repository name not found in the provided object");
        }
    } else {
        // Log a warning if the repository is not a valid JSON object
        warn!(?repository_value, "Expected JSON object, but found invalid repository object");
    }
}
