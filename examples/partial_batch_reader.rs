
use std::collections::HashMap;

use xlsx_batch_reader::{get_ord_from_tuple, read::XlsxBook, MAX_COL_NUM};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // the tail empty cells will be ignored, if you want the length of cells in each row is fixed, you can set right_ncol to a number not MAX_COL_NUM
        let mut sheet = book.get_sheet_by_name(&shname, 100, 0, 1, MAX_COL_NUM, true)?;

        let mut skip_until = HashMap::new();
        skip_until.insert("A".into(), "col1".into());
        skip_until.insert("C".into(), "col3".into());
        sheet.with_skip_until(&skip_until);
        let mut read_before = HashMap::new();
        read_before.insert("B".into(), "sum".into());
        sheet.with_read_before(&read_before);
        // only rows after skip_until-row(included) and before read_before(excluded) will be returned

        for batch in sheet {
            let (rows_nums, rows_data) = batch?;
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