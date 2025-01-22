use std::collections::HashMap;

use anyhow::{anyhow, Result};
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError, Format, IntoExcelData};

use crate::{CellValue, ColNum, RowNum};

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

pub struct XlsxWriter {
    book: Workbook,
    rows: HashMap<String, RowNum>,
    open: bool,
    columns: HashMap<String, HashMap<String, ColNum>>,
    prepends: HashMap<String, bool>
}

impl XlsxWriter {
    /// new xlsx writer
    pub fn new() -> Self {
        Self {
            open: true,
            book: Workbook::new(),
            rows: HashMap::new(),
            columns: HashMap::new(),
            prepends: HashMap::new()
        }
    }
    /// check whether the sheet exists
    #[inline]
    pub fn has_sheet(&mut self, shname: &String) -> bool {
        match self.book.worksheet_from_name(&shname) {
            Ok(_) => true,
            Err(_) => false
        }
    }
    /// set columns, if you have many sheets call this for each sheet   
    /// shname: sheet name   
    /// columns: column names   
    /// add_to_top: if true, the column names will be added to the top of the sheet
    pub fn with_columns(&mut self, shname: String, columns: Vec<String>, add_to_top: bool) {
        let mut maps = HashMap::with_capacity(columns.len());
        for (i, colname) in columns.into_iter().enumerate() {
            maps.insert(colname, i as ColNum);
        }
        self.columns.insert(shname.clone(), maps);
        self.prepends.insert(shname, add_to_top);
    }

    /// get raw worksheet and total rows if yu want to do some actions on raw worksheet(such as styles).
    /// if you append data to the worksheet directly, the total rows will not change.
    /// do not append data to the worksheet directly. use `append_row/append_rows...` method instead.
    #[inline]
    pub fn get_sheet_mut<'a, 'b>(&'a mut self, shname: &'b str) -> Result<(&'a mut Worksheet, RowNum)> {
        if self.open {
            if !self.rows.contains_key(shname) {
                // self.book.add
                let sheet = self.book.add_worksheet();
                sheet.set_name(shname)?;
                if self.prepends.get(shname) == Some(&true) {
                    if let Some(columns) = self.columns.get(shname) {
                        for (colval, colnum) in columns {
                            sheet.write(0, *colnum, colval)?;
                        }
                        self.rows.insert(shname.to_owned(), 1);
                    } else {
                        self.rows.insert(shname.to_owned(), 0);
                    }
                } else {
                    self.rows.insert(shname.to_owned(), 0);
                }
            }
            Ok((self.book.worksheet_from_name(shname)?, self.rows.get(shname).unwrap_or(&0).to_owned()))
        } else {
            Err(anyhow!("cannot write saved workbook"))
        }
    }
    /// append one row to sheet   
    /// name: sheet name, if not exists, create a new sheet   
    /// nrow: number write after the pre_cells and before the row_data (usually get from XLsxSheet)，not the row number of the target row   
    /// data: write after the head and the nrow   
    /// pre_cells: write at the head of the row   
    /// if you will call this function many times, it is better to use append_rows
    pub fn append_row<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, shname: &str, nrow: Option<&RowNum>, data: Vec<T>, pre_cells: &Vec<H>)  -> Result<()> {
        let (sheet, mut irow) = self.get_sheet_mut(shname)?;

        let mut icol = 0;
        let npre = pre_cells.len() as ColNum;
        if npre > 0 {
            sheet.write_row(irow, icol, pre_cells.clone())?;
            icol += npre;
        }
        if let Some(nrow) = nrow {
            sheet.write_number(irow, icol, *nrow)?;
            icol += 1;
        }
        sheet.write_row(irow, icol, data)?;
        irow += 1;
        self.rows.insert(shname.to_owned(), irow);
        Ok(())
    }
    /// append many rows to sheet   
    /// name: sheet name, if not exists, create a new sheet   
    /// nrows: number write after the pre_cells and before the row_data (usually get from XLsxSheet)，not the row number of the target row   
    /// data: write after the head and the nrow   
    /// pre_cells: write at the head of the row   
    pub fn append_rows<T: IntoExcelData, H: IntoExcelData+Clone>(&mut self, shname: &str, nrows: Vec<u32>, data: Vec<Vec<T>>, pre_cells: &Vec<H>) -> Result<()> {
        let (sheet, mut irow) = self.get_sheet_mut(shname)?;
        // 若nrows的长度为0，则不写行号
        if nrows.len() > 0 && nrows.len() != data.len() {
            return Err(anyhow!("the length of nrows is not equal to the length of data".to_string()));
        }

        let mut icol;
        let npre = pre_cells.len() as ColNum;
        for (i, rdata) in data.into_iter().enumerate() {
            icol = 0;
            if npre > 0 {
                sheet.write_row(irow, icol, pre_cells.clone())?;
                icol += npre;
            }
            if let Some(nrow) = nrows.get(i) {
                sheet.write_number(irow, icol, *nrow)?;
                icol += 1;
            }
            sheet.write_row(irow, icol, rdata)?;
            irow += 1;
        };
        self.rows.insert(shname.to_owned(), irow);
        Ok(())
    }
    /// append one row to sheet by column name    
    /// name: sheet name, if not exists, create a new sheet    
    /// data: the data to write   
    pub fn append_row_by_name<T: IntoExcelData>(&mut self, shname: &str, data: HashMap<String, T>) -> Result<()> {
        if let Some(columns) = self.columns.get(shname) {
            let columns = columns.clone();
            let (sheet, mut irow) = self.get_sheet_mut(shname)?;
            for (colname, colval) in data.into_iter() {
                if let Some(colnum) = columns.get(&colname) {
                    sheet.write(irow, *colnum, colval)?;
                } else {
                    return Err(anyhow!("column name {} not found", colname));
                }
            }
            irow += 1;
            self.rows.insert(shname.to_owned(), irow);
            Ok(())
        } else {
            Err(anyhow!("columns not set"))
        }
    }
    
    /// append many rows to sheet by column name    
    /// name: sheet name, if not exists, create a new sheet    
    /// data: the data to write   
    pub fn append_rows_by_name<T: IntoExcelData>(&mut self, shname: &str, data: Vec<HashMap<String, T>>) -> Result<()> {
        if let Some(columns) = self.columns.get(shname) {
            let columns = columns.clone();
            let (sheet, mut irow) = self.get_sheet_mut(shname)?;
            for rdata in data.into_iter() {
                for (colname, colval) in rdata.into_iter() {
                    if let Some(colnum) = columns.get(&colname) {
                        sheet.write(irow, *colnum, colval)?;
                    } else {
                        return Err(anyhow!("column name {} not found", colname));
                    }
                }
                irow += 1;
            };
            self.rows.insert(shname.to_owned(), irow);
            Ok(())
        } else {
            Err(anyhow!("columns not set"))
        }
    }
    /// save as xlsx file, can only run once each writer
    pub fn save_as<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<()> {
        if self.open {
            self.book.save(path)?;
            self.open = false;
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
    static ref EMPTY_COLUMNS: HashMap<String, ColNum> = HashMap::new();
}