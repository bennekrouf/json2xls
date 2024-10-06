
# JSON to Excel Converter

This Rust program provides a function to generate an Excel file from a JSON input file. The function reads a JSON file containing repository information and converts it into a structured Excel spreadsheet, organizing data into worksheets based on top-level keys.

## Function: `generate_excel_from_json`

### Description
The `generate_excel_from_json` function takes the path to a JSON file as input and generates an Excel file at the specified output path. Each top-level key in the JSON data corresponds to a separate worksheet in the Excel file. The function formats the headers and alternating rows for better readability.

### Input File Format
The input JSON file should follow the structure outlined below:

```json
{
    "top_level_key_1": [
        {
            "repository_name": "Repo 1",
            "debt": [
                {
                    "debt_key_1": "value1",
                    "debt_key_2": "value2"
                },
                {
                    "debt_key_1": "value3",
                    "debt_key_2": "value4"
                }
            ]
        },
        {
            "repository_name": "Repo 2",
            "debt": [
                {
                    "debt_key_1": "value5",
                    "debt_key_2": "value6"
                }
            ]
        }
    ],
    "top_level_key_2": [
        {
            "repository_name": "Repo 3",
            "debt": [
                {
                    "debt_key_1": "value7"
                }
            ]
        }
    ]
}
```

## Example

Given the above JSON structure, the function will create an Excel file with separate worksheets for top_level_key_1 and top_level_key_2. Each worksheet will contain columns for repository_name and the keys found within the debt objects.

## Usage

To use the function, call it with the appropriate file paths:

```rust
generate_excel_from_json("path/to/input.json", "path/to/output.xlsx")?;
```

## Requirements

- Rust (1.50 or later)
- Dependencies: serde_json, xlsxwriter

## License

This project is licensed under the MIT License.
