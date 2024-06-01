
#[cfg(feature = "xlsxwriter")]
use std::collections::HashMap;
#[cfg(feature = "xlsxwriter")]
use xlsx_batch_reader::write::XlsxWriter;

#[cfg(feature = "xlsxwriter")]
fn main() -> Result<(), Box<dyn std::error::Error>> {

    let mut writer = XlsxWriter::new();
    writer.with_columns("Sheet1".to_string(), vec!["A".to_string(), "B".to_string(), "C".to_string(), "D".to_string()], true);

    let row: HashMap<String, i32> = vec![("A".to_string(), 1), ("C".to_string(), 3)].into_iter().collect();
    writer.append_row_by_name("Sheet1", row)?;
    let row1: HashMap<String, &str> = vec![("A".to_string(), "A3"), ("B".to_string(), "B3"), ("D".to_string(), "D3")].into_iter().collect();
    let row2: HashMap<String, &str> = vec![("B".to_string(), "B4"), ("C".to_string(), "C4")].into_iter().collect();
    writer.append_rows_by_name("Sheet1", vec![row1, row2])?;

    writer.save_as("xlsx/out.xlsx")?;
    Ok(())
}

#[cfg(not(feature = "xlsxwriter"))]
fn main() {
    println!("Please enable the feature 'rust_xlsxwriter' to run this example.");
}