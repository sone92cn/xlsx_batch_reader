use std::{collections::HashMap, path::Path};
use anyhow::{anyhow, Result};
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
    /// append one row to sheet
    /// pre_cells: write at the head of the row
    /// row_num: number write after the pre_cells and before the row_data (usually get from XLsxSheet)，not the row number of the target row
    /// row_data: write after the head and the nrow
    #[inline]
    fn write_row<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, pre_cells: &Vec<H>, row_num: Option<&RowNum>, row_data: Vec<T>) -> Result<()> {
        let mut icol = pre_cells.len() as ColNum;
        if icol > 0 {
            self.sheet.write_row(self.nextrow, 0, pre_cells.clone())?;
        }
        if let Some(nrow) = row_num {
            self.sheet.write_number(self.nextrow, icol, *nrow as f64)?;
            icol += 1;
        }
        self.sheet.write_row(self.nextrow, icol, row_data)?;   // row-u32; col-u16
        self.nextrow += 1;
        Ok(())
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
    #[inline]
    pub fn has_sheet(&self, shname: &String) -> bool {
        self.names.contains(shname)
    }
    /// get mutable sheet
    #[inline]
    fn get_sheet_mut(&mut self, name: &str) -> Result<&mut Sheet> {
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
            self.sheets.get_mut(name).ok_or(anyhow!("sheet-{} not exist", name))
        } else {
            Err(anyhow!("cannot write saved workbook"))
        }
    }
    /// append one row to sheet
    /// name: sheet name, if not exists, create a new sheet
    /// nrow: number write after the pre_cells and before the row_data (usually get from XLsxSheet)，not the row number of the target row
    /// data: write after the head and the nrow
    /// pre_cells: write at the head of the row
    pub fn append_row<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, name: &str, nrow: Option<&RowNum>, data: Vec<T>, pre_cells: &Vec<H>)  -> Result<()> {
        let sheet = self.get_sheet_mut(name)?;
        sheet.write_row(pre_cells, nrow, data)?;
        Ok(())
    }
    /// append many rows to sheet
    /// name: sheet name, if not exists, create a new sheet
    /// nrows: number write after the pre_cells and before the row_data (usually get from XLsxSheet)，not the row number of the target row
    /// data: write after the head and the nrow
    /// pre_cells: write at the head of the row
    pub fn append_rows<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, name: &str, nrows: Vec<u32>, data: Vec<Vec<T>>, pre_cells: &Vec<H>) -> Result<()> {
        let sheet = self.get_sheet_mut(name)?;
        // 若nrows的长度为0，则不写行号
        if nrows.len() > 0 && nrows.len() != data.len() {
            return Err(anyhow!("the length of nrows is not equal to the length of data".to_string()));
        }
        for (i, rdata) in data.into_iter().enumerate() {
            sheet.write_row(pre_cells, nrows.get(i), rdata)?;
        };
        Ok(())
    }
    /// save as xlsx file, can only run once each writer
    pub fn save_as<P: AsRef<Path>>(&mut self, path: P) -> Result<()> {
        if self.opened {
            let mut book = Workbook::new();
            for name in &self.names {
                if let Some(sheet) = self.sheets.remove(name) {
                    book.push_worksheet(sheet.sheet)
                }
            }
            book.save(path)?;
            self.opened = false;
            Ok(())
        } else {
            Err(anyhow!("cannot save saved workbook"))
        }
    }
}

lazy_static::lazy_static! {
    static ref FMT_DEFAULT_DATE: Format = Format::new().set_num_format_index(14);
    static ref FMT_DEFAULT_TIME: Format = Format::new().set_num_format_index(21);
    static ref FMT_DEFAULT_DATETIME: Format = Format::new().set_num_format_index(22);
}