//! An Excel/OpenDocument Spreadsheets file batch reader, in pure Rust. This crate supports Office 2007 or newer file formats(xlsx, xlsm, etc). The most obvious difference from other Excel file reading crates is that it does not read the whole file into memory, but read in batches. So that it can maintain low memory usage, especially when reading large files.
use chrono::Local;
use anyhow::{anyhow, Result};
use lazy_static::lazy_static;
use read::FromCellValue;

/// Excel file reader
pub mod read;
/// Excel file writer
#[cfg(feature = "xlsxwriter")]
pub mod write;


/// reexport chrono
pub use chrono;

/// days since UNIX epoch
pub type Date32 = i32;  
/// seconds since UNIX epoch
#[derive(Debug)]
pub struct Timestamp(i64);
/// seconds since midnight
#[derive(Debug)]
pub struct Timesecond(i32);
/// row number
pub type RowNum = u32;
/// column number
pub type ColNum = u16;
/// merged range
pub type MergedRange = ((RowNum, ColNum), (RowNum, ColNum));

/// max column number
pub static MAX_COL_NUM: u16 = std::u16::MAX;
/// max row number
pub static MAX_ROW_NUM: u32 = std::u32::MAX;

impl Timestamp {
    /// use cell value as UTC datatime and get timestamp
    pub fn utc(&self) -> i64 {
        self.0
    }
    /// use cell value as local datatime and get timestamp, the local time zone is determined by the time zone in which it runs
    pub fn local(&self) -> i64 {
        self.0 - *LOCAL_OFFSET
    }
}

// i64 into Timestamp
impl Into<Timestamp> for i64 {
    fn into(self) -> Timestamp {
        Timestamp(self)
    }
}

// f64 into Timestamp
impl Into<Timestamp> for f64 {
    fn into(self) -> Timestamp {
        Timestamp(self as i64)
    }
}

// i32 into Timesecond
impl Into<Timesecond> for i32 {
    fn into(self) -> Timesecond {
        Timesecond(self)
    }
}

// Timesecond into i32
impl From<Timesecond> for i32 {
    fn from(ts: Timesecond) -> i32 {
        ts.0
    }
}

/// Convert character based Excel cell column addresses to number. If you pass parameter D to this function, you will get 4
pub fn get_num_from_ord(addr: &[u8]) -> Result<ColNum>{
    let mut i: usize;
    let mut j: ColNum;
    let mut col: ColNum;
    let addr = addr.to_ascii_uppercase();
    (col, i, j) = (0, addr.len(), 1);
    while i > 0 {
        i -= 1;
        if addr[i] > b'@' {
            col += ((addr[i] - b'@') as ColNum) * j;
            j *= 26;
        }
    };
    Ok(col)
}

/// Convert number based Excel cell column addresses to character. If you pass parameter 4 to this function, you will get D
pub fn get_ord_from_num(num: ColNum) -> Result<String> {
    let mut col = num;
    let mut addr = Vec::with_capacity(2);
    while col > 26 {
        addr.push(((col % 26) + 64) as u8 as char);
        col = col / 26;
    }; 
    addr.push((col + 64) as u8 as char);
    addr.reverse();
    Ok(String::from_iter(addr))
}

/// Convert character based Excel cell addresses to numbers. If you pass parameter D2 to this function, you will get (2, 4)
pub fn get_tuple_from_ord(addr: &[u8]) -> Result<(RowNum, ColNum)> {
    let mut i: usize;
    let mut j: ColNum;
    let mut col: ColNum;
    let mut row: Option<RowNum> = None;
    let addr = addr.to_ascii_uppercase();
    (col, i, j) = (0, addr.len(), 1);
    while i > 0 {
        i -= 1;
        if addr[i] > b'@' {
            if row.is_none() {
                row = Some(String::from_utf8(addr[i+1..].to_vec())?.parse::<RowNum>()?);
            };
            col += ((addr[i] - b'@') as ColNum) * j;
            j *= 26;
        }
    };
    if let Some(row) = row {
        Ok((row, col))
    } else {
        return Err(anyhow!("invalid cell address: {:?}", addr))
    }
}

/// Convert numbers based Excel cell addresses to characters. If you pass parameter (2, 4) to this function, you will get D2.
pub fn get_ord_from_tuple(row: RowNum, col: ColNum) -> Result<String> {
    match get_ord_from_num(col) {
        Ok(col) => {
            Ok(format!("{}{}", col, row))
        },
        Err(e) => Err(e)
    }
}

/// check whether the cell is a merged cell. If it is the first cell in the merged area, return the size of the merged area. RowNum and ColNum start from 1.
pub fn is_merged_cell(mgs: &Vec<MergedRange>, row: RowNum, col: ColNum) -> (bool, Option<(RowNum, ColNum)>) {
    for (left_top, right_end) in mgs {
        if left_top.0 <= row && left_top.1 <= col && right_end.0 >= row && right_end.1 >= col {
            if left_top.0 == row && left_top.1 == col {
                return (true, Some((right_end.0-row+1, right_end.1-col+1)))
            } else {
                return (true, None);
            }
        };
    }
    return (false, None)
}

/// Cell Value Type
#[derive(Debug, Clone)]
pub enum CellValue<'a> {
    Blank,
    Bool(bool),
    Number(f64),
    Date(f64),
    Time(f64),
    Datetime(f64),
    Shared(&'a String),
    String(String),
    Error(String)
}

impl<'a> CellValue<'a> {
    /// Attention: as to blank cell, String will return String::new(), and other types will return None. 
    pub fn get<T: FromCellValue>(&'a self) -> Result<Option<T>> {
        T::try_from_cval(self)
    }
}


lazy_static! {
    /// local time zone offset
    static ref LOCAL_OFFSET: i64 = Local::now().offset().local_minus_utc() as i64;
}
