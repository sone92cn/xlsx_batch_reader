use xlsx_batch_reader::{get_num_from_ord, is_merged_cell, read::XlsxBook};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // each row will have 3 cells.
        let mut sheet = book.get_sheet_by_name(&shname, 100, 0, 1, get_num_from_ord("C".as_bytes())?, true)?;

        // this is not necessary, if you don't care about the headers.
        let (_, _header) = sheet.get_header_row()?;
        if let Some((_rows_nums, _rows_data)) = sheet.get_remaining_cells()? {
            //  some code
        }; 

        // should be called when all data have been scaned.
        let merged_rngs = sheet.get_merged_ranges()?;
        match is_merged_cell(merged_rngs, 2, get_num_from_ord("A".as_bytes())?) {
            (true, None) => {
                println!("a merged cell(not top left cell)");
            },
            (true, Some((nrow, ncol))) => {
                println!("a merged cell(top left cell), taking {nrow} row(s) and {ncol} column(s)");
            },
            _ => {
                println!("not a merged cell");
            }
        }
    }
    Ok(())
}