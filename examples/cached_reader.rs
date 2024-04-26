
#[cfg(feature = "cached")]
use xlsx_batch_reader::{read::XlsxBook, MAX_COL_NUM};

#[cfg(feature = "cached")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // iter_batch will be supported in the future
        // the tail empty cells will be ignored, if you want the length of cells in each row is fixed, you can set right_ncol to a number not MAX_COL_NUM
        let sheet = book.get_cached_sheet_by_name(&shname, 100, 1, 1, MAX_COL_NUM, false)?;

        println!("sheet: {}, row_ranges: {:?}, col_ranges: {:?}", sheet.sheet_name(), sheet.row_range(), sheet.column_range());

        let (_, merge_info) = sheet.get_cell_value_with_merge_info("B2")?;

        match merge_info {
            (true, None) => {
                println!("B2 is a merged cell(not top left cell)");
            },
            (true, Some((nrow, ncol))) => {
                println!("B2 is a merged cell(top left cell), taking {nrow} row(s) and {ncol} column(s)");
            },
            _ => {
                println!("B2 is not a merged cell");
            }
        };

        let a4 = sheet.get_cell_value("A4")?;
        println!("A4={:?}", a4);
    }
    Ok(())
}


#[cfg(not(feature = "cached"))]
fn main() {
    println!("Please enable the feature 'cached' to run this example.");
}