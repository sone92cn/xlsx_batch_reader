# Changelogs
### [0.4.2] - 2024.02.15
#### Changed
* public use chrono

### [0.4.1] - 2024.02.06
#### Changed
* the xlsxwriter feature keeps memory usage to a minimum when writing writing large files


### [0.3.5] - 2024.01.18
#### Changed
* rewrite xlsxwriter feature and allow to get raw Worksheet


### [0.3.3] - 2024.01.12
#### Changed
* update dependency zip to the latest version


### [0.3.2] - 2024.11.25
#### Added
* support to check whether the sheet is empty


### [0.3.1] - 2024.11.24
#### Fixed
* when retrieving numbers from CellValue(stored as string), comma separators will be recognized correctly.


### [0.3.0] - 2024.11.15
#### Changed
* when retrieving non String data from CellValue, return None for '', '-', '-', '#N/A'


### [0.2.3] - 2024.11.10
#### Added
* add with_capture_vals and get_captured_vals to XlsxSheet


### [0.2.2] - 2024.11.9
#### Added
* add data type Timesecond(seconds since midnight)


### [0.2.1] - 2024.11.9
#### Fixed
* unable to read data if skip_until is an empty hashmap


### [0.2.0] - 2024.11.5
#### Added
* support to read partial rows based on conditions


### [0.1.14] - 2024.6.1
#### Added
* add trait IntoCellValue and implement NaiveDate, NaiveDateTime, NaiveTime, Date32, Timestamp for it


### [0.1.13] - 2024.6.1
#### Changed
* update dependency rust_xlsxwriter to the latest version


### [0.1.12] - 2024.6.1
#### Added
* write row(s) by column name instead of position


### [0.1.11] - 2024.5.11
#### Fixed
* not full fetaures documnet


### [0.1.9] - 2024.5.5
#### Added
* add feature full and documnet full fetaures


### [0.1.8] - 2024.5.3
#### Added
* support xlsxwriter to append one row


### [0.1.7] - 2024.4.27
#### Added
* support to iter cached sheet by batches

#### Fixed
* column_range return the first and last column number


### [0.1.5] - 2024.4.26
#### Added
* support read all data into memory when sheet created(fearure `cached` should be enabled)

#### Fixed
* unable to read the size of sheet 


### [0.1.4] - 2024.4.15
#### Added
* get cell value as timestamp

#### Changed
* Optimaze date&time recognition algorithm for better performance


### [0.1.3] - 2024.4.14
#### Fixed
* unable to use feature xlsxwriter


### [0.1.2] - 2024.4.13
#### Added
* get cell value as datetime and time

#### Changed
* output error message in English


### [0.1.1] - 2024.4.13
#### Added
* simple writer example


### [0.1.0] - 2023.4.13
#### Added
* first release