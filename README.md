# XlsxBatchReader
An Excel/OpenDocument Spreadsheets file batch reader, in pure Rust. This crate supports Office 2007 or newer file formats(xlsx, xlsm, etc). The most obvious difference from other Excel file reading crates is that it does not read the whole file into memory, but read in batches. So that it can maintain low memory usage, especially when reading large files.
This crate supports date and time recognition, as well as obtaining merged cell ranges. For faster speed, it only supports reading data, not support formulas and other styles.

# Examples
1. simple reader
```rust
use xlsx_batch_reader::{get_ord_from_tuple, read::XlsxBook, MAX_COL_NUM};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut book = XlsxBook::new("xlsx/test.xlsx", true)?;
    for shname in book.get_visible_sheets().clone() {
        // left_ncol should not be 0
        // the tail empty rows will be ignored, if you want the length of cells in each row is fixed, you can set right_ncol to a number not MAX_COL_NUM
        let sheet = book.get_sheet_by_name(&shname, 100, 0, 1, MAX_COL_NUM, false)?;

        for batch in sheet {
            let (rows_nums, rows_data) = batch?;
            // empty rows will be skiped
            for (row, cells) in rows_nums.into_iter().zip(rows_data) {
                for (col, cel) in cells.into_iter().enumerate() {
                    // supprted types: String, i64, f64, bool, Date32, Timestamp(v0.1.4), NaiveDate, NaiveDateTime(v0.1.2), NaiveTime(v0.1.2)
                    let val: String = cel.get()?.unwrap();   
                    println!("the value of {} is {val}; raw cell is {:?}", get_ord_from_tuple(row, (col+1) as u16)?, cel);   
                }
            }
        };
    }
    Ok(())
}
```
possible output:
```text
the value of A1 is a; raw cell is Shared("a")
the value of B1 is ; raw cell is Blank
the value of C1 is c; raw cell is Shared("c")
the value of D1 is d; raw cell is Shared("d")
the value of A2 is 1; raw cell is Number(1.0)
the value of B2 is ; raw cell is Blank
the value of C2 is s; raw cell is Shared("s")
the value of A4 is 2024-01-04; raw cell is Date(45295.58405092593)
the value of B4 is ; raw cell is Blank
the value of C4 is 4; raw cell is Number(4.0)  
```

2. merged ranges
```rust
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
```
possible output:
```text
a merged cell(top left cell), taking 2 row(s) and 2 column(s)
```

3. read date and time
```rust
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use xlsx_batch_reader::{read::XlsxBook, Timestamp, MAX_COL_NUM};

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
            let val_stamp: Timestamp = row[0].get()?.unwrap();   // since v0.1.4
            println!("date:{}\ntime:{}\ndatetime:{}\ntimestamp:{}", val_dt, val_tm, val_dttm, val_stamp.utc());
        }; 
    }
    Ok(())
}
```
possible output:
```text
date:2024-01-04
time:14:01:02
datetime:2024-01-04 14:01:02
timestamp:1704376862
```

4. cached reader(feature cached should be enabled)
```rust
use xlsx_batch_reader::{read::XlsxBook, MAX_COL_NUM};

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
```
possible output:
```shell
sheet: Sheet1, row_ranges: (2, 4), col_ranges: (1, 4)
B2 is a merged cell(not top left cell)
A4=Date(45295.58405092593)
```


5. simple batch writer (feature xlsxwriter should be enabled)
```rust
use xlsx_batch_reader::{get_num_from_ord, read::XlsxBook, write::XlsxWriter};

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
```

# Features
| Features | Description |
| --- | --- |
| `cached` | Enable cached xlsx sheet (read all sheet data into memory when created) |
| `xlsxwriter` | Enable simple xlsx writer |


# Todos
- [ ] support to iter cached sheet by batch


# Changelogs
## [0.1.6] - 2024.4.26
### Fixed
* project description


## [0.1.5] - 2024.4.26
### Added
* support read all data into memory when sheet created(fearure `cached` should be enabled)

### Fixed
* unable to read the size of sheet 


## [0.1.4] - 2024.4.15
### Added
* get cell value as timestamp

### Changed
* Optimaze date&time recognition algorithm for better performance


## [0.1.3] - 2024.4.14
### Fixed
* unable to use feature xlsxwrite

## [0.1.2] - 2024.4.13
### Added
* get cell value as datetime and time

### Changed
* output error message in English


## [0.1.1] - 2024.4.13
### Added
* simple writer example


## [0.1.0] - 2023.4.13
### Added
* first release