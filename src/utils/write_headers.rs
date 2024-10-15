use xlsxwriter::Format;

// Function to write headers to the worksheet
pub fn write_headers(
    worksheet: &mut xlsxwriter::Worksheet,
    headers: &[String],
    title_format: &Format
) -> Result<(), Box<dyn std::error::Error>> {
    headers.iter().enumerate().for_each(|(col, header)| {
        worksheet.write_string(0, col as u16, header, Some(title_format)).unwrap();
    });
    Ok(())
}
