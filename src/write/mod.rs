use std::{collections::HashMap, path::Path};
use anyhow::{anyhow, Result};
use lazy_static::lazy_static;
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError, Format};

use crate::{CellValue, ColNum, RowNum};

pub use rust_xlsxwriter::IntoExcelData;

impl IntoExcelData for CellValue<'_> {
    fn write<'a>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        match self {
            CellValue::Blank => {},
            CellValue::Error(_) => {},
            CellValue::Bool(v) => {
                worksheet.write_boolean(row, col, v)?;
            },
            CellValue::Number(v) => {
                worksheet.write_number(row, col, v)?;
            },
            CellValue::Date(v) => {
                worksheet.write_number_with_format(row, col, v, &FMT_DEFAULT_DATE)?;
            },
            CellValue::Time(v) => {
                worksheet.write_number_with_format(row, col, v, &FMT_DEFAULT_TIME)?;
            },
            CellValue::Datetime(v) => {
                worksheet.write_number_with_format(row, col, v, &FMT_DEFAULT_DATETIME)?;
            },
            CellValue::Shared(v) => {
                worksheet.write_string(row, col, v)?;
            },
            CellValue::String(v) => {
                worksheet.write_string(row, col, v)?;
            }
        };
        Ok(worksheet)
    }

    fn write_with_format<'a, 'b>(
        self,
        worksheet: &'a mut Worksheet,
        row: RowNum,
        col: ColNum,
        format: &'b Format,
    ) -> Result<&'a mut Worksheet, XlsxError> {
        match self {
            CellValue::Blank => {},
            CellValue::Error(_) => {},
            CellValue::Bool(v) => {
                worksheet.write_boolean_with_format(row, col, v, format)?;
            },
            CellValue::Number(v) => {
                worksheet.write_number_with_format(row, col, v, format)?;
            },
            CellValue::Date(v) => {
                worksheet.write_number_with_format(row, col, v, format)?;
            },
            CellValue::Time(v) => {
                worksheet.write_number_with_format(row, col, v, format)?;
            },
            CellValue::Datetime(v) => {
                worksheet.write_number_with_format(row, col, v, format)?;
            },
            CellValue::Shared(v) => {
                worksheet.write_string_with_format(row, col, v, format)?;
            },
            CellValue::String(v) => {
                worksheet.write_string_with_format(row, col, v, format)?;
            }
        };
        Ok(worksheet)
    }
}

struct Sheet{
    sheet: Worksheet,
    nextrow: u32,
}

impl Sheet {
    pub fn write_rows<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, nrows: Vec<u32>, data: Vec<Vec<T>>, head: &Vec<H>) -> Result<()> {
        // 若nrows的长度为0，则不写行号
        let add_nrow = {nrows.len() > 0}; 
        let mut icol = head.len() as ColNum;
        if add_nrow {
            if nrows.len() != data.len() {
                return Err(anyhow!("the length of nrows is not equal to the length of data".to_string()));
            } else {
                icol += 1;
            }
        }
        for (i, row) in data.into_iter().enumerate() {
            if head.len() > 0 {
                self.sheet.write_row(self.nextrow, 0, head.clone())?;
            }
            if add_nrow {
                self.sheet.write_number(self.nextrow, icol - 1, nrows[i] as f64)?;
            }
            self.sheet.write_row(self.nextrow, icol, row)?;   // row-u32; col-u16
            self.nextrow = self.nextrow + 1;
        };
        Ok(())
    }
    pub fn close_sheet(self) -> Result<Worksheet> {
        Ok(self.sheet)
    }
}

// xlsx_writer
pub struct XlsxWriter {
    names: Vec<String>,
    sheets: HashMap<String, Sheet>,
    opened: bool
}

impl XlsxWriter {
    pub fn new() -> Self {
        Self {
            names: vec![],
            sheets: HashMap::new(),
            opened: true,
        }
    }
    /// check whether the sheet exists
    pub fn has_sheet(&self, shname: &String) -> bool {
        self.names.contains(shname)
    }
    /// The length of nrows and data should be equal.
    pub fn append_rows<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, name: &str, nrows: Vec<u32>, data: Vec<Vec<T>>, pre_cells: &Vec<H>) -> Result<()> {
        if self.opened {
            if !self.sheets.contains_key(name) {
                let mut sht = Sheet {
                    nextrow: 0,
                    sheet: Worksheet::new(),
                };
                sht.sheet.set_name(name)?;
                self.sheets.insert(name.to_owned(), sht);
                self.names.push(name.to_owned());
            };
            self.sheets.get_mut(name).ok_or(anyhow!("sheet-{} not exist", name))?.write_rows(nrows, data, pre_cells)?;
            Ok(())
        }else{
            Err(anyhow!("cannot write saved workbook".to_string()))
        }
    }
    /// save as file
    pub fn save_as<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        let mut book = Workbook::new();
        for name in self.names.iter() {
            let sht = self.sheets.remove(name).unwrap().close_sheet()?;
            book.push_worksheet(sht);
        };
        book.save(path)?;
        self.opened = false;
        Ok(())
    }
}

lazy_static! {
    static ref FMT_DEFAULT_DATE: Format = Format::new().set_num_format_index(14);
    static ref FMT_DEFAULT_TIME: Format = Format::new().set_num_format_index(21);
    static ref FMT_DEFAULT_DATETIME: Format = Format::new().set_num_format_index(22);
}