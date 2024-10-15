
use xlsxwriter::{Workbook, Format, format::FormatAlignment, format::FormatColor};

// Function to create formats for the Excel cells
pub fn create_format(_: &Workbook, is_title: bool, bg_color: u32, font_color: u32, bold: bool) -> Format {
    let mut format = Format::new();
    if bold {
        format.set_bold();
    }
    if is_title {
        format.set_align(FormatAlignment::Center);
    }
    if bg_color != 0 {
        format.set_bg_color(FormatColor::Custom(bg_color));
    }
    if font_color != 0 {
        format.set_font_color(FormatColor::Custom(font_color));
    }
    format
}
