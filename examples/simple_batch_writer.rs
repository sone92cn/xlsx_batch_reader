
#[cfg(feature = "xlsxwriter")]
use xlsx_batch_reader::{get_num_from_ord, read::XlsxBook, write::XlsxWriter};

#[cfg(feature = "xlsxwriter")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = XlsxWriter::new();
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // each row will have 3 cells.
        let mut sheet = book.get_sheet_by_name(&shname, 100, 0, 1, get_num_from_ord("C".as_bytes())?, true)?;

        // the sheet name will be write at the begin of each row
        let pre_cells = vec![shname];
        if let Some((rows_nums, rows_data)) = sheet.get_remaining_cells()? {
            writer.append_rows("sheet", rows_nums, rows_data, &pre_cells)?;
            // if you don't want row numbers to be writed before data, set nrows = vec![];
        }; 
    };
    writer.save_as("xlsx/out.xlsx")?;
    Ok(())
}

#[cfg(not(feature = "xlsxwriter"))]
fn main() {
    println!("Please enable the feature 'rust_xlsxwriter' to run this example.");
}

// cargo run --example merged_range --features xlsxwriter