
use serde::Deserialize;

#[derive(Debug, Deserialize)]
pub struct ExcelConfig {
    pub json_file_path: String,
    pub output_excel_file: String,
    #[serde(deserialize_with = "hex_color_to_u32")]
    pub title_bg_color: u32,
    #[serde(deserialize_with = "hex_color_to_u32")]
    pub title_font_color: u32,
    #[serde(deserialize_with = "hex_color_to_u32")]
    pub row_alt_color: u32,
    pub repository_name_key: String,
    pub object_array_key: String,
}

// Helper function to convert a hex string in YAML to u32
fn hex_color_to_u32<'de, D>(deserializer: D) -> Result<u32, D::Error>
where
    D: serde::Deserializer<'de>,
{
    let hex_str: String = Deserialize::deserialize(deserializer)?;
    // Remove leading `#` if present, then parse the hex string
    u32::from_str_radix(hex_str.trim_start_matches('#'), 16).map_err(serde::de::Error::custom)
}

