use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use xlsx_batch_reader::{read::XlsxBook, MAX_COL_NUM};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // the tail empty cells will be ignored, if you want the length of cells in each row is fixed, you can set right_ncol to a number not MAX_COL_NUM
        let mut sheet = book.get_sheet_by_name(&shname, 100, 3, 1, MAX_COL_NUM, false)?;

        if let Some((_, rows_data)) = sheet.get_remaining_cells()? {
            let row = &rows_data[0];
            let val_dt: NaiveDate = row[0].get()?.unwrap();
            let val_tm: NaiveTime = row[0].get()?.unwrap();
            let val_dttm: NaiveDateTime = row[0].get()?.unwrap();
            println!("date:{}\ntime:{}\ndatetime:{}", val_dt, val_tm, val_dttm);
        }; 
    }
    Ok(())
}