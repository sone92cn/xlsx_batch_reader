
#[cfg(feature = "cached")]
use xlsx_batch_reader::{get_ord_from_tuple, read::XlsxBook, MAX_COL_NUM};

#[cfg(feature = "cached")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // the tail empty cells will be ignored, if you want the length of cells in each row is fixed, you can set right_ncol to a number not MAX_COL_NUM
        let sheet = book.get_cached_sheet_by_name(&shname, 100, 0, 1, MAX_COL_NUM, false)?;

        for (rows_nums, rows_data) in sheet {
            // empty rows will be skiped
            for (row, cells) in rows_nums.into_iter().zip(rows_data) {
                for (col, cel) in cells.into_iter().enumerate() {
                    let val: String = cel.get()?.unwrap();   // supprted types: String, i64, f64, bool, NaiveDate
                    println!("the value of {} is {val}; raw cell is {:?}", get_ord_from_tuple(row, (col+1) as u16)?, cel);  
                }
            }
        };
    }
    Ok(())
}


#[cfg(not(feature = "cached"))]
fn main() {
    println!("Please enable the feature 'cached' to run this example.");
}