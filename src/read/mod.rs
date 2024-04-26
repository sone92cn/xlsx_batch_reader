use std::{cmp::max, collections::HashMap, fs::File, io::BufReader, path::Path};
use anyhow::{anyhow, Result};
use zip::{ZipArchive, read::ZipFile};
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::{reader::Reader, events::Event};

use lazy_static::lazy_static;
use crate::{get_num_from_ord, get_tuple_from_ord, CellValue, ColNum, Date32, MergedRange, RowNum, Timestamp, MAX_COL_NUM};

#[cfg(feature = "cached")]
use crate::is_merged_cell;

// ooxml： http://www.officeopenxml.com/

macro_rules! get_attr_val {
    ($e:expr, $tag:expr) => {
        match $e.try_get_attribute($tag)? {
            Some(v) => {v.unescape_value()?},
            None => return Err(anyhow!("attribute {} not exist", $tag))
        }
    };
    ($e:expr, $tag:expr, parse) => {
        match $e.try_get_attribute($tag)? {
            Some(v) => {v.unescape_value()?.parse()?},
            None => return Err(anyhow!("attribute {} not exist", $tag))
        }
    };
    ($e:expr, $tag:expr, to_string) => {
        match $e.try_get_attribute($tag)? {
            Some(v) => {v.unescape_value()?.to_string()},
            None => return Err(anyhow!("attribute {} not exist", $tag))
        }
    };
}

pub struct XlsxBook {
    ini_share: bool,
    str_share: Vec<String>,
    shts_hidden: Vec<String>,
    shts_visible: Vec<String>,
    map_style: HashMap<u32, u32>,
    map_sheet: HashMap<String, String>,
    zip_archive: ZipArchive<BufReader<File>>,
    datetime_fmts: HashMap<u32, u8>,
}

impl XlsxBook {
    /// load_share: if set to false, you should call load_share_strings before reading data. it should usually be true. If you only need to obtain the sheet names, you can set it false to open the file faster.
    pub fn new<T: AsRef<Path>>(path: T, load_share: bool) -> Result<XlsxBook> {
        // zip压缩文件
        let mut zip_archive = {
            let file = File::open(path)?;
            let zipreader = BufReader::new(file);
            ZipArchive::new(zipreader)?
        };

        let book_refs = {
            let file = zip_archive.by_name("xl/_rels/workbook.xml.rels")?;
            
            let mut buf = Vec::new();
            let mut refs = HashMap::new();
            let mut reader =  Reader::from_reader(BufReader::new(file));
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Empty(ref e)) => {
                        if e.name().as_ref() == b"Relationship"{
                            refs.insert(get_attr_val!(e, "Id", to_string), get_attr_val!(e, "Target", to_string));
                        };
                    },
                    Ok(Event::Start(ref e)) => {   // 解析 <sheet ..></sheet> 模式
                        if e.name().as_ref() == b"Relationship"{
                            refs.insert(get_attr_val!(e, "Id", to_string), get_attr_val!(e, "Target", to_string));
                        };
                    },
                    Ok(Event::Eof) => break, // exits the loop when reaching end of file
                    Err(e) => return Err(anyhow!("workbook.xml.refs broken: {:?}", e)),
                    _ => ()                  // There are several other `Event`s we do not consider here
                }
                buf.clear();
            };
            refs
        };


        // 初始化sheet列表
        let mut shts_hidden = Vec::<String>::new();
        let mut shts_visible = Vec::<String>::new();
        let map_sheet = {
            let file = zip_archive.by_name("xl/workbook.xml")?;
            let mut reader =  Reader::from_reader(BufReader::new(file));
            reader.trim_text(true);

            let mut buf = Vec::new();
            let mut map_share: HashMap<String, String> = HashMap::new();
            loop {
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Empty(ref e)) => {
                        if e.name().as_ref() == b"sheet"{
                            let name = get_attr_val!(e, "name", to_string);
                            let rid = get_attr_val!(e, "r:id", to_string);
                            let sheet = if book_refs.contains_key(&rid) {
                                format!("xl/{}", book_refs[&rid])
                            } else {
                                return Err(anyhow!("Relationship of sheet-{rid} not found"))
                            };
                            match e.try_get_attribute("state").unwrap_or(None) {
                                Some(attr) => {
                                    if attr.unescape_value()?.as_bytes() == b"hidden" {
                                        shts_hidden.push(name.clone());
                                    } else {
                                        shts_visible.push(name.clone());
                                    };
                                },
                                _ => {shts_visible.push(name.clone());}
                            };
                            map_share.insert(name, sheet);  // sheet名，对应的真是xml文件
                        };
                    },
                    Ok(Event::Start(ref e)) => {   // 解析 <sheet ..></sheet> 模式
                        if e.name().as_ref() == b"sheet"{
                            let name = get_attr_val!(e, "name", to_string);
                            let rid = get_attr_val!(e, "r:id", to_string);
                            let sheet = if book_refs.contains_key(&rid) {
                                format!("xl/{}", book_refs[&rid])
                            } else {
                                return Err(anyhow!("Relationship of sheet-rid not found!"))
                            };
                            match e.try_get_attribute("state").unwrap_or(None) {
                                Some(attr) => {
                                    if attr.unescape_value()?.as_bytes() != b"hidden" {
                                        shts_visible.push(name.clone());
                                    };
                                },
                                _ => {shts_visible.push(name.clone());}
                            };
                            map_share.insert(name, sheet);  // sheet名，对应的真是xml文件
                        };
                    },
                    Ok(Event::Eof) => break, // exits the loop when reaching end of file
                    Err(e) => return Err(anyhow!("workbook.xml is broken: {:?}", e)),
                    _ => ()                  // There are several other `Event`s we do not consider here
                }
                buf.clear();
            };
            map_share
        };

        // 初始化单元格格式
        let mut datetime_fmts = DATETIME_FMTS.clone();
        let map_style = {
            match zip_archive.by_name("xl/styles.xml") {
                Ok(file) => {
                    let mut reader =  Reader::from_reader(BufReader::new(file));
                    reader.trim_text(true);

                    let mut inx: u32 = 0;
                    let mut act = false;
                    let mut buf = Vec::new();
                    let mut map_style: HashMap<u32, u32> = HashMap::new();
                    loop {
                        match reader.read_event_into(&mut buf) {
                            Ok(Event::Start(ref e)) => {
                                if e.name().as_ref() == b"cellXfs" || e.name().as_ref() == b"numFmts" {
                                    act = true;
                                } else if act && (e.name().as_ref() == b"numFmt"){
                                    let code = get_attr_val!(e, "formatCode", to_string);
                                    if code.contains("yy") {
                                        if code.contains("h") || code.contains("ss") {
                                            datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_DATETIME);
                                        } else {
                                            datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_DATE);
                                        }
                                    } else if code.contains("ss") {
                                        datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_TIME);
                                    };
                                } else if act && (e.name().as_ref() == b"xf"){
                                    map_style.insert(inx, get_attr_val!(e, "numFmtId", parse));
                                    inx += 1;
                                };
                            },
                            Ok(Event::Empty(ref e)) => {
                                if act && (e.name().as_ref() == b"numFmt"){
                                    let code = get_attr_val!(e, "formatCode", to_string);
                                    if code.contains("yy") {
                                        if code.contains("h") || code.contains("ss") {
                                            datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_DATETIME);
                                        } else {
                                            datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_DATE);
                                        }
                                    } else if code.contains("ss") {
                                        datetime_fmts.insert(get_attr_val!(e, "numFmtId", parse), FMT_TIME);
                                    };
                                } else if act && (e.name().as_ref() == b"xf"){
                                    map_style.insert(inx, get_attr_val!(e, "numFmtId", parse));
                                    inx += 1;
                                };
                            },
                            Ok(Event::End(ref e)) => {
                                if e.name().as_ref() == b"numFmts" {
                                    act = false;
                                } else if e.name().as_ref() == b"cellXfs" {
                                    break;
                                };
                            },
                            Ok(Event::Eof) => break, // exits the loop when reaching end of file
                            Err(e) => return Err(anyhow!("styles.xml is broken: {:?}", e)),
                            _ => ()                  // There are several other `Event`s we do not consider here
                        }
                        buf.clear();
                    };
                    map_style
                },
                Err(_) => {
                    HashMap::new()
                }
            }
        };
        
        let mut book = XlsxBook{
                ini_share: false,
                str_share: Vec::new(),
                map_style,
                map_sheet,
                shts_hidden,
                shts_visible,
                zip_archive,
                datetime_fmts,
            };
        if load_share {
            book.load_share_strings()?;
        };
        Ok(book)
    }
    /// get hidden sheets
    pub fn get_hidden_sheets(&self) -> &Vec<String> {
        &self.shts_hidden
    } 
    /// get visible sheets
    pub fn get_visible_sheets(&self) -> &Vec<String> {
        &self.shts_visible
    }
    /// if set load_share to false, you should call load_share_strings before reading data
    pub fn load_share_strings(&mut self) -> Result<()>{
        if self.ini_share {
            return Ok(());
        };
        let str_share = {
            match self.zip_archive.by_name("xl/sharedStrings.xml") {
                Ok(file) => {
                    let mut reader =  Reader::from_reader(BufReader::new(file));
                    reader.trim_text(true);

                    let mut buf = Vec::with_capacity(3069);
                    let cap = loop {    // 获取ShareString容量
                        match reader.read_event_into(&mut buf) {
                            Ok(Event::Start(ref e)) => {
                                if e.name().as_ref() == b"sst"{
                                    let cnt: usize =  {
                                        match e.try_get_attribute("uniqueCount")? {
                                            Some(a) => {a.unescape_value()?.parse()?},
                                            None => {get_attr_val!(e, "count", parse)}
                                        }
                                    };
                                    break cnt
                                }
                            }
                            Ok(Event::Eof) => {return Ok(())}, // exits the loop when reaching end of file
                            Err(e) => return Err(anyhow!("sharedStrings.xml is broken: {:?}", e)),
                            _ => (),                     // There are several other `Event`s we do not consider here
                        }
                    };

                    let mut insert = false;
                    let mut shstring = String::new(); 
                    let mut vec_share: Vec<String> = Vec::with_capacity(cap);
                    loop {
                        match reader.read_event_into(&mut buf) {
                            Ok(Event::Start(ref e)) => {
                                match e.name().as_ref() {
                                    b"si" => {shstring.clear()},
                                    b"t" => {insert = true},
                                    _ => {insert = false},
                                }
                            },
                            Ok(Event::Text(ref t)) => {
                                if insert {
                                    shstring += &t.unescape()?.as_ref();
                                }
                            },
                            Ok(Event::End(ref e)) => {
                                if e.name().as_ref() == b"si" {
                                    vec_share.push(shstring.clone());
                                }
                            },
                            Ok(Event::Eof) => break, // exits the loop when reaching end of file
                            Err(e) => return Err(anyhow!("sharedStrings.xml is broken: {:?}", e)),
                            _ => ()                  // There are several other `Event`s we do not consider here
                        }
                        buf.clear();
                    };
                    if cap != vec_share.len() {  
                        return Err(anyhow!("shareString-lenth check error!！"));
                    };
                    vec_share
                },
                Err(_) => {
                    Vec::<String>::new()
                }
            }
        };
        self.ini_share = true;
        self.str_share = str_share;
        Ok(())
    }
    /// sht_name: sheet name  
    /// iter_batch: The number of rows per batch  
    /// skip_rows: number of skipped rows  
    /// left_ncol: Starting column (included), with 1 as the starting value  
    /// right_ncol: Terminate columns (including), MAX-COL_NUM to get non fixed termination columns  
    pub fn get_sheet_by_name(&mut self, sht_name: &String, iter_batch: usize, skip_rows: u32, left_ncol: ColNum, right_ncol: ColNum, first_row_is_header: bool) -> Result<XlsxSheet> {
        for (k, v) in self.map_sheet.clone() {
            if k.eq(sht_name) {
                if !self.ini_share {
                    self.load_share_strings()?;
                };

                match self.zip_archive.by_name(v.as_str()) {
                    Ok(file) => {
                        let mut reader = Reader::from_reader(BufReader::new(file));
                        reader.trim_text(true);
            
                        return Ok(XlsxSheet {
                            reader,
                            skip_rows,
                            left_ncol: left_ncol-1,
                            right_ncol,
                            iter_batch,
                            first_row_is_header,
                            first_row: None,
                            key: k,
                            buf: Vec::with_capacity(8*1024),
                            status: 1,
                            str_share: &self.str_share,
                            map_style: &self.map_style,
                            datetime_fmts: &self.datetime_fmts,
                            max_size: None,
                            merged_rects: None
                        });
                    },
                    Err(_) => {
                        return Err(anyhow!("sheet {} - {} lost！", k.as_str(), v.as_str()));
                    }
                };
            };
        };
        Err(anyhow!(format!("{} sheet not found!", sht_name)))
    }
    /// get cached sheet by name, all data will be cached in memory when sheet created
    #[cfg(feature = "cached")]
    pub fn get_cached_sheet_by_name(&mut self, sht_name: &String, iter_batch: usize, skip_rows: u32, left_ncol: ColNum, right_ncol: ColNum, first_row_is_header: bool) -> Result<CachedSheet> {
        Ok(self.get_sheet_by_name(sht_name, iter_batch, skip_rows, left_ncol, right_ncol, first_row_is_header)?.into_cached_sheet()?)
    }
}

pub struct XlsxSheet<'a> {
    key: String,
    str_share: &'a Vec<String>,
    map_style: &'a HashMap<u32, u32>,
    buf: Vec<u8>,
    status: u8,   // 0-closed; 1-new; 2-active; 3-get_cell; 4-skip_cell; 初始为1
    reader: Reader<BufReader<ZipFile<'a>>>,
    iter_batch: usize,
    skip_rows: u32,
    max_size: Option<(RowNum, ColNum)>,
    left_ncol: ColNum,
    right_ncol: ColNum,
    first_row_is_header: bool,
    first_row: Option<(u32, Vec<CellValue<'a>>)>,
    datetime_fmts: &'a HashMap<u32, u8>,
    merged_rects: Option<Vec<((RowNum, ColNum), (RowNum, ColNum))>>
}

impl<'a> XlsxSheet<'a> {
    /// into cached sheet
    #[cfg(feature = "cached")]
    pub fn into_cached_sheet(mut self) -> Result<CachedSheet<'a>> {
        let (data, bottom_nrow) =  match self.get_next_row() {
            Ok(Some((r, d))) => {
                let mut data = if let Some((rn, _)) = self.max_size {
                    HashMap::with_capacity(rn as usize)
                } else {
                    HashMap::new()
                };
                data.insert(r, d);
                let mut last_nrow = r;
                loop {
                    match self.get_next_row() {
                        Ok(Some((r, d))) => {
                            last_nrow = r;
                            data.insert(r, d);
                        },
                        Ok(None) => {
                            break;
                        },
                        Err(e) => {
                            return Err(e);
                        }
                    };
                };
                (data, last_nrow)
            },
            Ok(None) => {(HashMap::new(), 0)},
            Err(e) => {return Err(e);}
        };
        let merged_rects = self.get_merged_ranges()?.to_owned();
        let right_ncol = if self.right_ncol == MAX_COL_NUM {
            if let Some((_mr, mc)) = self.max_size {
                mc
            } else {
                self.right_ncol
            }
        } else {
            self.right_ncol
        };
        Ok(CachedSheet {
            data,
            merged_rects,
            key: self.key,
            _iter_batch: self.iter_batch,
            top_nrow: if self.first_row_is_header {self.skip_rows+2} else {self.skip_rows+1},
            bottom_nrow,
            left_ncol: self.left_ncol + 1,
            right_ncol,
            header_row: self.first_row,
        })
    }
    /// get sheet name
    pub fn sheet_name(&self) -> &String {
        &self.key
    }
    /// get column range
    pub fn column_range(&self) -> (ColNum, ColNum) {
        (self.left_ncol, self.right_ncol)
    }
    fn get_next_row(&mut self) -> Result<Option<(u32, Vec<CellValue<'a>>)>> {
        let mut col: ColNum = 0;
        let mut cell_type: Vec<u8> = Vec::new();
        let mut prev_head: Vec<u8> = Vec::new();
        let mut row_num: u32 = 0;
        let mut col_index: ColNum = 1;    // 当前需增加cell的col_index
        let mut row_value: Vec<CellValue<'_>> = Vec::new();
        let mut num_fmt_id: u32 = 0;
        if self.status == 0 {
            return Ok(None)
        }  //  已关闭的sheet直接返回None
        loop {
            match self.reader.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref e)) => {
                    prev_head = e.name().as_ref().to_vec();
                    if self.status == 0 {
                        break Ok(None)
                    } else if self.status == 1 {
                        if prev_head == b"dimension" {
                            let attr = get_attr_val!(e, "ref", to_string);
                            let dim: Vec<&str> = attr.split(':').collect();
                            if let Some(x) = dim.get(1) {
                                self.max_size = Some(get_tuple_from_ord(x.as_bytes())?);
                            };
                        } else if prev_head == b"sheetData" {
                            self.status = 2;
                        } else if prev_head == b"mergeCells" {
                            let cnt: usize = get_attr_val!(e, "count", parse);
                            self.process_merged_cells(cnt)?;
                        }; 
                    } else {
                        if prev_head == b"c" {
                            match e.try_get_attribute("t")? {
                                Some(attr) => {
                                    cell_type = attr.unescape_value()?.as_bytes().to_owned();
                                },
                                _ => {
                                    cell_type = b"n".to_vec();
                                }
                            };
                            match e.try_get_attribute("s")? {
                                Some(attr) => {
                                    num_fmt_id = self.map_style[&attr.unescape_value()?.parse::<u32>()?];
                                },
                                _ => {
                                    num_fmt_id = 0;
                                }
                            };
                            col = get_num_from_ord(get_attr_val!(e, "r").as_bytes()).unwrap_or(0);
                            
                            if row_num > self.skip_rows && col > self.left_ncol && col <= self.right_ncol {
                                self.status = 3;   // 3-get_cell; 4-skip_cell;
                            } else {
                                self.status = 4;   // 3-get_cell; 4-skip_cell;
                            }
                        } else if prev_head == b"row" {
                            row_num = get_attr_val!(e, "r", parse);
                            let cap = {
                                if self.right_ncol == MAX_COL_NUM {
                                    if let Some(x) = get_attr_val!(e, "spans").as_ref().split(":").last() {
                                        x.parse()?
                                    } else {
                                        1
                                    }
                                } else {
                                    self.right_ncol
                                }
                            } - self.left_ncol;
                            row_value = Vec::with_capacity(cap.into());
                            // row_value.push(CellValue::Number(row_num as f64));  // 行号单独返回
                        }; 
                    };
                },
                Ok(Event::Empty(ref e)) => {
                    prev_head = e.name().as_ref().to_vec();
                    if self.status == 1 && prev_head == b"dimension" {
                        let attr = get_attr_val!(e, "ref", to_string);
                        let dim: Vec<&str> = attr.split(':').collect();
                        if let Some(x) = dim.get(1) {
                            self.max_size = Some(get_tuple_from_ord(x.as_bytes())?);
                        };
                    }
                },
                Ok(Event::Text(ref t)) => {
                    // b for boolean
                    // d for date
                    // e for error
                    // inlineStr for an inline string (i.e., not stored in the shared strings part, but directly in the cell)
                    // n for number
                    // s for shared string (so stored in the shared strings part and not in the cell)
                    // str for a formula (a string representing the formula)
                    if self.status == 3 && prev_head == b"v"{
                        while col_index + self.left_ncol < col {
                            row_value.push(CellValue::Blank);
                            col_index += 1;
                        }
                        if cell_type == b"inlineStr" && prev_head == b"t" {
                            row_value.push(CellValue::String(t.unescape()?.to_string()));
                            col_index += 1;
                        } else if prev_head == b"v" {
                            if cell_type == b"s" {
                                row_value.push(CellValue::Shared(&self.str_share[String::from_utf8(t.to_vec())?.parse::<usize>()?]));
                            } else if cell_type == b"n" {
                                let fmt = self.datetime_fmts.get(&num_fmt_id).unwrap_or(&FMT_DEFAULT);
                                if *fmt == FMT_DATE {
                                    row_value.push(CellValue::Date(String::from_utf8(t.to_vec())?.parse::<f64>()?));
                                } else if *fmt == FMT_DATETIME {
                                    row_value.push(CellValue::Datetime(String::from_utf8(t.to_vec())?.parse::<f64>()?));
                                } else if *fmt == FMT_TIME {
                                    row_value.push(CellValue::Time(String::from_utf8(t.to_vec())?.parse::<f64>()?));
                                } else {
                                    row_value.push(CellValue::Number(String::from_utf8(t.to_vec())?.parse::<f64>()?));
                                };
                            } else if cell_type == b"b" {
                                if String::from_utf8(t.to_vec())?.parse::<usize>() == Ok(1) {
                                    row_value.push(CellValue::Bool(true));
                                } else {
                                    row_value.push(CellValue::Bool(false));
                                };
                            } else if cell_type == b"d" {
                                row_value.push(CellValue::String(String::from_utf8(t.to_vec())?));
                            } else if cell_type == b"e" {
                                row_value.push(CellValue::Error(String::from_utf8(t.to_vec())?));
                            } else if cell_type == b"str" {
                                row_value.push(CellValue::String(t.unescape()?.to_string()));
                            } else{
                                row_value.push(CellValue::Blank);
                            }
                            col_index += 1;
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    // 0-closed; 1-new; 2-active;
                    if (e.name().as_ref() == b"row") && self.status > 1 && row_value.len() > 0 {
                        if self.right_ncol != MAX_COL_NUM {
                            while row_value.len() < row_value.capacity() {
                                row_value.push(CellValue::Blank);
                            };
                        }
                        break Ok(Some((row_num, row_value)))
                    }else if e.name().as_ref() == b"sheetData" {
                        self.status = 0; 
                        break Ok(None)
                    }
                },
                Ok(Event::Eof) => {break Ok(None)}, // exits the loop when reaching end of file
                Err(e) => return Err(anyhow!("sharedStrings.xml is broken: {:?}", e)),
                _ => ()                  // There are several other `Event`s we do not consider here
            }
            self.buf.clear();
        }
    }
    /// get header if first_row_is_header is true
    pub fn get_header_row(&mut self) -> Result<(u32, Vec<CellValue<'a>>)> {
        if self.first_row_is_header {
            match self.get_next_row() {
                Ok(Some(v)) => {
                    self.first_row = Some(v);
                    self.first_row_is_header = false;
                },
                Ok(None) => {},
                Err(e) => {return Err(e)}
            }
        }
        match &self.first_row {
            Some(v) => Ok(v.clone()),
            None => Err(anyhow!("no header row！"))
        }
    }
    fn process_merged_cells(&mut self, count: usize) -> Result<()> {
        if self.status == 1 || self.status == 0 {
            if self.merged_rects.is_none() {
                self.merged_rects = Some(vec![]);
            }
            loop {
                match self.reader.read_event_into(&mut self.buf) {
                    Ok(Event::Start(ref e)) => {
                        if e.name().as_ref() == b"mergeCell" {
                            let attr = get_attr_val!(e, "ref", to_string);
                            let dim: Vec<&str> = attr.split(':').collect();
                            if let Some(x) = dim.get(0) {
                                let left_top = get_tuple_from_ord(x.as_bytes())?;
                                let right_end =  if let Some(x) = dim.get(1) {
                                    get_tuple_from_ord(x.as_bytes())?
                                } else {
                                    return Err(anyhow!("mergeCell error：{}", attr));
                                };
                                if let Some(ref mut mgs) = self.merged_rects {
                                    mgs.push((left_top, right_end))
                                };
                            } else {
                                return Err(anyhow!("mergeCell error：{}", attr));
                            }; 
                        }
                    },
                    Ok(Event::Empty(ref e)) => {
                        if e.name().as_ref() == b"mergeCell" {
                            let attr = get_attr_val!(e, "ref", to_string);
                            let dim: Vec<&str> = attr.split(':').collect();
                            if let Some(x) = dim.get(0) {
                                let left_top = get_tuple_from_ord(x.as_bytes())?;
                                let right_end =  if let Some(x) = dim.get(1) {
                                    get_tuple_from_ord(x.as_bytes())?
                                } else {
                                    return Err(anyhow!("mergeCell error：{}", attr));
                                };
                                if let Some(ref mut mgs) = self.merged_rects {
                                    mgs.push((left_top, right_end))
                                };
                            } else {
                                return Err(anyhow!("mergeCell error：{}", attr));
                            }; 
                        }
                    },
                    Ok(Event::End(ref e)) => {
                        if e.name().as_ref() != b"mergeCells" {
                            break;
                        }
                        else if e.name().as_ref() != b"mergeCell" {
                            break;
                        }
                    },
                    Ok(Event::Eof) => {
                        break;
                    }, // exits the loop when reaching end of file
                    _ => {}
                }
            };
            if let Some(ref rects) = self.merged_rects {
                if rects.len() != count {
                    self.merged_rects = None;
                    return Err(anyhow!("the number of merged ranges is not equal to the number of rows"));
                };
            }
        }
        Ok(())
    }
    /// get merged ranges, call after all data getched
    pub fn get_merged_ranges(&mut self) -> Result<&Vec<MergedRange>> {
        if self.merged_rects.is_none() {
            if self.status == 0 {  // 已关闭的情况下读取合并单元格
                loop {
                    match self.reader.read_event_into(&mut self.buf) {
                        Ok(Event::Start(ref e)) => {
                            if e.name().as_ref() == b"mergeCells" {
                                let cnt: usize = get_attr_val!(e, "count", parse);
                                self.process_merged_cells(cnt)?;
                                break;
                            };
                        },
                        _ => {}
                    }
                };
            } else {
                return Err(anyhow!("finish fetching data first"));
            }
        };
        if let Some(ref rects) = self.merged_rects {
            Ok(rects)
        } else {
            return Err(anyhow!("merged_rects error"));
        }
    }
    /// Get all the remaining data
    pub fn get_remaining_cells(&mut self) -> Result<Option<(Vec<u32>, Vec<Vec<CellValue<'_>>>)>> {
        match self.get_next_row() {
            Ok(Some((r, d))) => {
                let (mut rows, mut data) = if let Some((rn, _)) = self.max_size {
                    (Vec::with_capacity(max(1, rn-r+1) as usize), Vec::with_capacity(rn as usize))
                } else {
                    (Vec::new(), Vec::new())
                };
                rows.push(r);
                data.push(d);
                loop {
                    match self.get_next_row() {
                        Ok(Some((r, d))) => {
                            rows.push(r);
                            data.push(d);
                        },
                        Ok(None) => {
                            break Ok(Some((rows, data)));
                        },
                        Err(e) => {
                            break Err(e);
                        }
                    };
                }
            },
            Ok(None) => {
                Ok(None)
            },
            Err(e) => {Err(e)}
        }
    }
}

impl<'a> Iterator for XlsxSheet<'a> {
    type Item = Result<(Vec<u32>, Vec<Vec<CellValue<'a>>>)>;
    fn next(&mut self) -> Option<Self::Item> {
        let mut nums = Vec::with_capacity(self.iter_batch);
        let mut data = Vec::with_capacity(self.iter_batch);
        loop {
            match self.get_next_row() {
                Ok(Some(v)) => {
                    if self.first_row_is_header {
                        self.first_row = Some(v);
                        self.first_row_is_header = false;
                    } else {
                        nums.push(v.0);
                        data.push(v.1);
                        if nums.len() >= self.iter_batch { 
                            break Some(Ok((nums, data)))
                        }
                    }
                },
                Ok(None) => {
                    if nums.len() > 0 {
                        break Some(Ok((nums, data)))
                    } else {
                        break None
                    }
                },
                Err(e) => {
                    break Some(Err(e));
                }
            }
        }
        
    }
}

#[cfg(feature = "cached")]
pub struct CachedSheet<'a> {
    data: HashMap<RowNum, Vec<CellValue<'a>>>,
    key: String,
    _iter_batch: usize,   ///  TODO iter cached sheet
    top_nrow: RowNum,
    bottom_nrow: RowNum,
    left_ncol: ColNum,
    right_ncol: ColNum,
    header_row: Option<(u32, Vec<CellValue<'a>>)>,
    merged_rects: Vec<((RowNum, ColNum), (RowNum, ColNum))>
}

#[cfg(feature = "cached")]
impl <'a> CachedSheet<'a> {
    /// get sheet name
    pub fn sheet_name(&self) -> &String {
        &self.key
    }
    /// get row range
    pub fn row_range(&self) -> (RowNum, RowNum) {
        (self.top_nrow, self.bottom_nrow)
    }
    /// get column range
    pub fn column_range(&self) -> (ColNum, ColNum) {
        (self.left_ncol, self.right_ncol)
    }
    /// get header if first_row_is_header is true
    pub fn get_header_row(&self) -> Result<(u32, Vec<CellValue<'a>>)> {
        match &self.header_row {
            Some(v) => Ok(v.clone()),
            None => Err(anyhow!("no header row！"))
        }
    }
    /// get merged ranges, call as any time
    pub fn get_merged_ranges(&self) -> &Vec<MergedRange> {
        &self.merged_rects
    }
    /// Get all data
    pub fn get_all_cells(&self) -> &HashMap<RowNum, Vec<CellValue<'_>>> {
        &self.data
    }
    /// get cell value by address, if the cell is not exist, return &CellValue::Blank
    pub fn get_cell_value<A: AsRef<str>>(&self, addr: A) -> Result<&CellValue<'a>> {
        let (row, col) = get_tuple_from_ord(addr.as_ref().as_bytes())?;
        if row >= self.top_nrow && row <= self.bottom_nrow
            && col >= self.left_ncol && col <= self.right_ncol {
            if self.data.contains_key(&row) {
                Ok(self.data[&row].get((col-1) as usize).unwrap_or(&CellValue::Blank))
            } else {
                Ok(&CellValue::Blank)
            }
        } else {
            Err(anyhow!("Invalid address - out of range"))
        }
    }
    /// get cell value by address, if the cell is not exist, return &CellValue::Blank
    pub fn get_cell_value_with_merge_info<A: AsRef<str>>(&self, addr: A) -> Result<(&CellValue<'a>, (bool, Option<(RowNum, ColNum)>))> {
        let (row, col) = get_tuple_from_ord(addr.as_ref().as_bytes())?;
        if row >= self.top_nrow && row <= self.bottom_nrow
            && col >= self.left_ncol && col <= self.right_ncol {
            let (merge, spans) = is_merged_cell(&self.merged_rects, row, col);
            if self.data.contains_key(&row) {
                Ok((self.data[&row].get((col-1) as usize).unwrap_or(&CellValue::Blank), (merge, spans)))
            } else {
                Ok((&CellValue::Blank, (merge, spans)))
            }
        } else {
            Err(anyhow!("Invalid address - out of range"))
        }
    }
}

/// get another type of data from cell value
pub trait FromCellValue {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> 
        where Self: Sized;
}

impl FromCellValue for String {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => Ok(Some(n.to_string())),
            CellValue::Date(n) => {
                Ok(Some((BASE_DATE.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)).to_string()))
            },
            CellValue::Time(n) => {
                Ok(Some(NaiveTime::from_num_seconds_from_midnight_opt(((*n-n.trunc()) * 86400.0) as u32, 0).unwrap().format("%H:%M:%S").to_string()))
            }
            CellValue::Datetime(n) => {
                Ok(Some((BASE_DATETIME.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)+(Duration::try_seconds(((*n-n.trunc()) * 86400.0) as i64).ok_or(anyhow!("invalid date"))?)).to_string()))
            },
            CellValue::Shared(s) => Ok(Some((**s).to_owned())),
            CellValue::String(s) => Ok(Some((*s).to_owned())),
            CellValue::Error(s) => Ok(Some((*s).to_string())),
            CellValue::Bool(b) => Ok(Some(if *b {"true".to_string()}else{"false".to_string()})),
            CellValue::Blank => Ok(Some("".to_string())),
        }
    }
}

impl FromCellValue for f64 {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => Ok(Some(*n)),
            CellValue::Date(n) => Ok(Some(*n)),
            CellValue::Time(n) => Ok(Some(*n)),
            CellValue::Datetime(n) => Ok(Some(*n)),
            CellValue::Shared(s) => {
                match s.parse::<f64>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::String(s) => {
                match s.parse::<f64>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid value-{:?}", val))),
            CellValue::Bool(b) => Ok(Some(if *b {1.0}else{0.0})),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for i64 {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => Ok(Some(*n as i64)),
            CellValue::Date(n) => Ok(Some(*n as i64)),
            CellValue::Time(n) => Ok(Some(*n as i64)),
            CellValue::Datetime(n) => Ok(Some(*n as i64)),
            CellValue::Shared(s) => {
                match s.parse::<i64>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::String(s) => {
                match s.parse::<i64>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid value-{:?}", val))),
            CellValue::Bool(b) => Ok(Some(if *b {1}else{0})),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for bool {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => {if (*n).abs() > 0.009 {Ok(Some(true))} else {Ok(Some(false))}},
            CellValue::Date(n) => {if (*n).abs() > 0.009 {Ok(Some(true))} else {Ok(Some(false))}},
            CellValue::Time(n) => {if (*n).abs() > 0.009 {Ok(Some(true))} else {Ok(Some(false))}},
            CellValue::Datetime(n) => {if (*n).abs() > 0.009 {Ok(Some(true))} else {Ok(Some(false))}},
            CellValue::Shared(s) => {
                match s.parse::<bool>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::String(s) => {
                match s.parse::<bool>() {
                    Ok(n) => Ok(Some(n)),
                    Err(_) => Err(anyhow!(format!("invalid value-{:?}", val)))
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid value-{:?}", val))),
            CellValue::Bool(b) => Ok(Some(*b)),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for NaiveDate {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => Ok(Some(BASE_DATE.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid datetime"))?))),
            CellValue::Date(n) => Ok(Some(BASE_DATE.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid datetime"))?))),
            CellValue::Time(n) => Ok(Some(BASE_DATE.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid datetime"))?))),
            CellValue::Datetime(n) => Ok(Some(BASE_DATE.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid datetime"))?))),
            CellValue::Shared(s) => {
                match NaiveDate::parse_from_str(*s, "%Y-%m-%d") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveDate::parse_from_str(*s, "%Y/%m/%d") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid datetime-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::String(s) => {
                match NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveDate::parse_from_str(s, "%Y/%m/%d") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid datetime{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid datetime{:?}", val))),
            CellValue::Bool(_) => Err(anyhow!(format!("invalid datetime{:?}", val))),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for NaiveDateTime {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => {
                Ok(Some(BASE_DATETIME.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)+(Duration::try_seconds(((*n-n.trunc()) * 86400.0) as i64).ok_or(anyhow!("invalid date"))?)))
            },
            CellValue::Date(n) => {
                Ok(Some(BASE_DATETIME.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)+(Duration::try_seconds(((*n-n.trunc()) * 86400.0) as i64).ok_or(anyhow!("invalid date"))?)))
            },
            CellValue::Time(n) => {
                Ok(Some(BASE_DATETIME.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)+(Duration::try_seconds(((*n-n.trunc()) * 86400.0) as i64).ok_or(anyhow!("invalid date"))?)))
            },
            CellValue::Datetime(n) => {
                Ok(Some(BASE_DATETIME.clone()+(Duration::try_days(*n as i64).ok_or(anyhow!("invalid date"))?)+(Duration::try_seconds(((*n-n.trunc()) * 86400.0) as i64).ok_or(anyhow!("invalid date"))?)))
            },
            CellValue::Shared(s) => {
                match NaiveDateTime::parse_from_str(*s, "%Y-%m-%d %H:%M:%S") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveDateTime::parse_from_str(*s, "%Y/%m/%d %H:%M:%S") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid datetime-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::String(s) => {
                match NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M:%S") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid datetime{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid datetime{:?}", val))),
            CellValue::Bool(_) => Err(anyhow!(format!("invalid datetime{:?}", val))),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for NaiveTime {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            CellValue::Number(n) => {
                Ok(Some(NaiveTime::from_num_seconds_from_midnight_opt(((*n-n.trunc()) * 86400.0) as u32, 0).ok_or(anyhow!("invalid time"))?))
            },
            CellValue::Date(n) => {
                Ok(Some(NaiveTime::from_num_seconds_from_midnight_opt(((*n-n.trunc()) * 86400.0) as u32, 0).ok_or(anyhow!("invalid time"))?))
            },
            CellValue::Time(n) => {
                Ok(Some(NaiveTime::from_num_seconds_from_midnight_opt(((*n-n.trunc()) * 86400.0) as u32, 0).ok_or(anyhow!("invalid time"))?))
            },
            CellValue::Datetime(n) => {
                Ok(Some(NaiveTime::from_num_seconds_from_midnight_opt(((*n-n.trunc()) * 86400.0) as u32, 0).ok_or(anyhow!("invalid time"))?))
            },
            CellValue::Shared(s) => {
                match NaiveTime::parse_from_str(*s, "%H:%M:%S") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveTime::parse_from_str(*s, "%H:%M:%S") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid time-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::String(s) => {
                match NaiveTime::parse_from_str(s, "%H:%M:%S") {
                    Ok(v) => Ok(Some(v)),
                    Err(_) => {
                        match NaiveTime::parse_from_str(s, "%H:%M:%S") {
                            Ok(v) => Ok(Some(v)),
                            Err(_) => {Err(anyhow!(format!("invalid time{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid time{:?}", val))),
            CellValue::Bool(_) => Err(anyhow!(format!("invalid time{:?}", val))),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for Date32 {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            // 1970-01-01的Excel值为25569
            CellValue::Number(n) => Ok(Some((*n as i32)-25569)),
            CellValue::Date(n) => Ok(Some((*n as i32)-25569)),
            CellValue::Time(n) => Ok(Some((*n as i32)-25569)),
            CellValue::Datetime(n) => Ok(Some((*n as i32)-25569)),
            CellValue::Shared(s) => {
                match NaiveDate::parse_from_str(*s, "%Y-%m-%d") {
                    Ok(v) => Ok(Some((v - UNIX_DATE.clone()).num_days() as i32)),
                    Err(_) => {
                        match NaiveDate::parse_from_str(*s, "%Y/%m/%d") {
                            Ok(v) => Ok(Some((v - UNIX_DATE.clone()).num_days() as i32)),
                            Err(_) => {Err(anyhow!(format!("invalid date32-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::String(s) => {
                match NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                    Ok(v) => Ok(Some((v - UNIX_DATE.clone()).num_days() as i32)),
                    Err(_) => {
                        match NaiveDate::parse_from_str(s, "%Y/%m/%d") {
                            Ok(v) => Ok(Some((v - UNIX_DATE.clone()).num_days() as i32)),
                            Err(_) => {Err(anyhow!(format!("invalid date32-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid date32-{:?}", val))),
            CellValue::Bool(_) => Err(anyhow!(format!("invalid date32-{:?}", val))),
            CellValue::Blank => Ok(None),
        }
    }
}

impl FromCellValue for Timestamp {
    fn try_from_cval(val: &CellValue<'_>) -> Result<Option<Self>> {
        match val {
            // 1970-01-01的Excel值为25569
            CellValue::Number(n) => Ok(Some(((*n - 25569.0) * 86400.0).into())),
            CellValue::Date(n) => Ok(Some(((*n - 25569.0) * 86400.0).into())),
            CellValue::Time(n) => Ok(Some(((*n - 25569.0) * 86400.0).into())),
            CellValue::Datetime(n) => Ok(Some(((*n - 25569.0) * 86400.0).into())),
            CellValue::Shared(s) => {
                match NaiveDateTime::parse_from_str(*s, "%Y-%m-%d %H:%M:%S") {
                    Ok(v) => Ok(Some(v.and_utc().timestamp().into())),
                    Err(_) => {
                        match NaiveDateTime::parse_from_str(*s, "%Y-%m-%d %H:%M:%S") {
                            Ok(v) => Ok(Some(v.and_utc().timestamp().into())),
                            Err(_) => {Err(anyhow!(format!("invalid timestamp-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::String(s) => {
                match NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
                    Ok(v) => Ok(Some(v.and_utc().timestamp().into())),
                    Err(_) => {
                        match NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M:%S") {
                            Ok(v) => Ok(Some(v.and_utc().timestamp().into())),
                            Err(_) => {Err(anyhow!(format!("invalid timestamp-{:?}", val)))}
                        }
                    }
                }
            },
            CellValue::Error(_) => Err(anyhow!(format!("invalid timestamp-{:?}", val))),
            CellValue::Bool(_) => Err(anyhow!(format!("invalid timestamp-{:?}", val))),
            CellValue::Blank => Ok(None),
        }
    }
}

// datetime sign
static FMT_DATE: u8 = 0;
static FMT_TIME: u8 = 1;
static FMT_DATETIME: u8 = 2;
static FMT_DEFAULT: u8 = 255;

lazy_static! {
    static ref BASE_DATE: NaiveDate = NaiveDate::from_ymd_opt(1899, 12,30).unwrap();
    static ref BASE_DATETIME: NaiveDateTime = BASE_DATE.and_hms_opt(0, 0, 0).unwrap();
    static ref UNIX_DATE: NaiveDate = NaiveDate::from_ymd_opt(1970,  1, 1).unwrap();
    static ref DATETIME_FMTS: HashMap<u32, u8> = {
        let mut v = HashMap::new();
        v.extend((14..18).map(|n| (n, FMT_DATE)));
        v.extend((27..32).map(|n| (n, FMT_DATE)));
        v.extend((34..37).map(|n| (n, FMT_DATE)));
        v.extend((50..59).map(|n| (n, FMT_DATE)));  // FMT_DATE - 0
        v.extend((18..22).map(|n| (n, FMT_TIME)));
        v.extend((32..34).map(|n| (n, FMT_TIME)));
        v.extend((45..48).map(|n| (n, FMT_TIME)));  // FMT_TIME - 1
        v.insert(22, FMT_DATETIME);                 // FMT_DATETIME - 2
        v
    };
    static ref NUM_FMTS: HashMap<u32, String> = {
        let mut map: HashMap<u32, String> = HashMap::new();
        // General
        map.insert(0, "General".to_string());
        map.insert(1, "0".to_string());
        map.insert(2, "0.00".to_string());
        map.insert(3, "#,##0".to_string());
        map.insert(4, "#,##0.00".to_string());

        map.insert(9, "0%".to_string());
        map.insert(10, "0.00%".to_string());
        map.insert(11, "0.00E+00".to_string());
        map.insert(12, "# ?/?".to_string());
        map.insert(13, "# ??/??".to_string());
        map.insert(14, "m/d/yyyy".to_string()); // Despite ECMA 'mm-dd-yy");
        map.insert(15, "d-mmm-yy".to_string());
        map.insert(16, "d-mmm".to_string());
        map.insert(17, "mmm-yy".to_string());
        map.insert(18, "h:mm AM/PM".to_string());
        map.insert(19, "h:mm:ss AM/PM".to_string());
        map.insert(20, "h:mm".to_string());
        map.insert(21, "h:mm:ss".to_string());
        map.insert(22, "m/d/yyyy h:mm".to_string()); // Despite ECMA 'm/d/yy h:mm");

        map.insert(37, "#,##0_);(#,##0)".to_string()); //  Despite ECMA '#,##0 ;(#,##0)");
        map.insert(38, "#,##0_);[Red](#,##0)".to_string()); //  Despite ECMA '#,##0 ;[Red](#,##0)");
        map.insert(39, "#,##0.00_);(#,##0.00)".to_string()); //  Despite ECMA '#,##0.00;(#,##0.00)");
        map.insert(40, "#,##0.00_);[Red](#,##0.00)".to_string()); //  Despite ECMA '#,##0.00;[Red](#,##0.00)");

        map.insert(44, r###"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"###.to_string());
        map.insert(45, "mm:ss".to_string());
        map.insert(46, "[h]:mm:ss".to_string());
        map.insert(47, "mm:ss.0".to_string()); //  Despite ECMA 'mmss.0");
        map.insert(48, "##0.0E+0".to_string());
        map.insert(49, "@".to_string());

        // CHT
        map.insert(27, "[$-404]e/m/d".to_string());
        map.insert(30, "m/d/yy".to_string());
        map.insert(36, "[$-404]e/m/d".to_string());
        map.insert(50, "[$-404]e/m/d".to_string());
        map.insert(57, "[$-404]e/m/d".to_string());

        // THA
        map.insert(59, "t0".to_string());
        map.insert(60, "t0.00".to_string());
        map.insert(61, "t#,##0".to_string());
        map.insert(62, "t#,##0.00".to_string());
        map.insert(67, "t0%".to_string());
        map.insert(68, "t0.00%".to_string());
        map.insert(69, "t# ?/?".to_string());
        map.insert(70, "t# ??/??".to_string());

        // JPN
        map.insert(28, r###"[$-411]ggge"年"m"月"d"日""###.to_string());
        map.insert(29, r###"[$-411]ggge"年"m"月"d"日""###.to_string());
        map.insert(31, r###"yyyy"年"m"月"d"日""###.to_string());
        map.insert(32, r###"h"時"mm"分""###.to_string());
        map.insert(33, r###"h"時"mm"分"ss"秒""###.to_string());
        map.insert(34, r###"yyyy"年"m"月""###.to_string());
        map.insert(35, r###"m"月"d"日""###.to_string());
        map.insert(51, r###"[$-411]ggge"年"m"月"d"日""###.to_string());
        map.insert(52, r###"yyyy"年"m"月""###.to_string());
        map.insert(53, r###"m"月"d"日""###.to_string());
        map.insert(54, r###"[$-411]ggge"年"m"月"d"日""###.to_string());
        map.insert(55, r###"yyyy"年"m"月""###.to_string());
        map.insert(56, r###"m"月"d"日""###.to_string());
        map.insert(58, r###"[$-411]ggge"年"m"月"d"日""###.to_string());

        map
    };
}
