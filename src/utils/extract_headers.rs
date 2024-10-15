
use serde_json::Value;
use crate::models::ExcelConfig;
use tracing::info; // Import info and trace

// Function to extract headers from the repositories' debt array
pub fn extract_headers(repositories: &[Value], config: &ExcelConfig) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let mut debt_headers = vec![config.repository_name_key.clone()];
    let mut found_headers = false; // Flag to track if any headers are found

    repositories.iter().filter_map(|repo| repo.get(&config.object_array_key))
        .filter_map(|debts| debts.as_array())
        .flat_map(|debts| debts.iter().filter_map(|debt| debt.as_object()))
        .for_each(|debt| {
            for object_array_key in debt.keys() {
                if !debt_headers.contains(object_array_key) {
                    debt_headers.push(object_array_key.clone());
                    found_headers = true; // Set flag if at least one header is found
                    // Log the found header at INFO level
                    // info!(header = %object_array_key, "Found header");
                }
            }
        });

    // Check if no headers were found and return an error if so
    if !found_headers {
        return Err("No headers found in the provided repositories".into());
    }

    Ok(debt_headers)
}

