
use serde_json::Value;
use crate::models::ExcelConfig;

// Function to extract a row of data for the repository
pub fn extract_repository_row(
    repository_value: &Value,
    debt_headers: &[String],
    config: &ExcelConfig
) -> Vec<String> {
    let mut row_data = vec![];

    // Log start of processing
    tracing::info!("Processing repository repository={:?}", repository_value);

    if let Value::Object(repository) = repository_value {
        // Push the repository name if it exists
        if let Some(Value::String(repo_name)) = repository.get(&config.repository_name_key) {
            tracing::trace!("Found repository name repo_name={}", repo_name);
            row_data.push(repo_name.clone());
        } else {
            tracing::warn!("Repository name not found or is not a string");
            row_data.push(String::new()); // Add empty value for repo_name if not found
        }

        // Get the debt array
        if let Some(Value::Array(debt_array)) = repository.get("debt") {
            // Loop through the debt objects
            for debt in debt_array {
                if let Value::Object(debt_object) = debt {
                    // Add the other debt values based on headers
                    for header in debt_headers.iter().skip(1) {
                        match debt_object.get(header) {
                            Some(Value::String(debt_value)) => {
                                row_data.push(debt_value.clone());
                            },
                            Some(_) => {
                                // If the value is found but is not a String
                                tracing::warn!("Header value is not a string, adding empty value header={}", header);
                                row_data.push(String::new());
                            },
                            None => {
                                // Log if the header is not found in the debt object
                                tracing::warn!("Header not found in repository, adding empty value header={}", header);
                                row_data.push(String::new());
                            }
                        }
                    }
                }
            }
        } else {
            tracing::warn!("Debt array not found in repository");
        }
    } else {
        tracing::warn!("Invalid repository format: expected Object, got {:?}", repository_value);
    }

    // Log the completed extraction
    tracing::trace!("Completed extraction for repository row_data={:?}", row_data);

    row_data
}

