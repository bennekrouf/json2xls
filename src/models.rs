use xlsxwriter::format::FormatColor;

pub struct ExcelConfig {
    pub json_file_path: String,
    pub output_excel_file: String,
    pub title_bg_color: u32,
    pub title_font_color: FormatColor,
    pub row_alt_color: u32,
    pub repository_name_key: String,
    pub object_array_key: String,
}
