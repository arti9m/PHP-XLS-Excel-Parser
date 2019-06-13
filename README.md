# PHP-XLS-Excel-Parser
Probably, the fastest possible and the most efficient parser for XLS excel files for PHP!

_Note:_ this parser is suitable __only for older XLS files__:MS Excel 1995 (BIFF5), 1997-2003 (BIFF8). It will not work with the newer ones, XLSX!

## 1. Requirements

At least __PHP 5.6__ 32-bit is required. Untested with PHP versions prior to 5.6. Works best with PHP 7.x 64-bit (faster, more memory efficient than PHP 5.6).

Also, this parser uses my [PHP MSCFB Parser](https://github.com/arti9m/PHP-MSCFB-Parser). Grab a copy of __MSCFB.php__ if you don't have one here: https://github.com/arti9m/PHP-MSCFB-Parser and put it to your PHP include directory or to the same directory where __MSXLS.php__ is. MSCFB is "required-once" inside MSXLS, so there's no need to include/require it manually.

## 2. Basic usage

1. Download __MSXLS.php__ from this repository and put it in your include directory or in the same directory where your script is.
2. Make sure that __MSCFB.php__ is in your include directory or in the same directory as MSXLS.php.
3. Add the following line to the beginning of your PHP script (specify full path to MSXLS.php, if needed):
```PHP
require_once 'MSXLS.php'; // MSCFB.php is 'required once' inside MSXLS.php
```
4. Create an instance of MSXLS (open XLS file):
```PHP
$excel = new MSXLS('path_to_file.xls');
```

5. If no errors occured up to this point, you are ready to read the cells from your file. There are two ways you can do it: either read all cells at once into a 2-dimensional array (the fastest method), or read cells row by row (slower, but is suitable for database upload and may use much less memory, depending on usage).

In any case, it's a good idea to check for errors first before reading anything:

```PHP
if($excel->error) die($excel->err_msg); // Terminate script execution, show error message.
```

6. Read all cells into a 2-dimensional array:
```PHP
$excel->read_everything(); // Read cells into $excel->cells
```
At this point all your cells data is contained inside `$excel->cells` array:
```PHP
var_dump($excel->cells); // Output all parsed cells from XLS file
```

7. Read all cells row by row:
```PHP
$excel->switch_to_row(); //switch to row-by-row mode

while($row = $excel->read_next_row()){
  // You can process $row here however you like.
  // For example, you can upload a row into a database.
  $rows[] = $row; // For now, just store a parsed row inside $rows array.
}
```

_Note:_ `$excel->cells` will be erased when `$excel->switch_to_row()` is executed, so make sure you save all the data you need before switching to another parsing mode. If you need to switch back to 'read all at once' mode, use `$excel->switch_to_array()` method.

8. If you need to perform some other memory-intensive tasks in the same script, it is a good idea to free some memory:
```PHP
$excel->free(); // This is also called in the destructor
unset($excel);
```

## 3. Advanced usage

### Sheet selection

If there is more than one worksheet in your file, and you want to parse the worksheet that is not the first valid worksheet, you will have to select your sheet manually. To do this, use `$excel->get_valid_sheets()` method to get an array with all available selectable worksheets. When the desired worksheet has been found, use its array index or `'number'` entry as `$sheet` parameter to `$excel->select_sheet($sheet)` method. For example:
```PHP
var_dump($excel->get_valid_sheets()); //outputs selectable sheets info
$excel->select_sheet(1); //select sheet with index 1
```
Alternatively, if you know sheet name, you can use it in the same method to select sheet:
```PHP
$excel->select_sheet('your_sheet_name'); //also works
```
Leave out sheet index/name to select the first available sheet:
```PHP
$excel->select_sheet(); //selects the first valid sheet in file
```

You can use `$excel->get_active_sheet()` method to return selected sheet info.

See __*Public properties and methods*__ below to get more information about methods mentioned above.

_Note:_ The first valid worksheet is selected automatically when file is opened or when Parsing mode is changed.


### Parsing modes

There are two modes which the parser can work in: __Array mode__ and __Row-by-row mode__. By default, Array mode is used.

#### Array mode

This mode lets you read all cells at once into `$excel->cells` array property. It is designed to read all available data as fast as possible when no additional cells processing needed. This mode is used by default. This mode can be selected with `$excel->switch_to_array()` method. Data is read with `$excel->read_everything()` method into `$excel->cells` array property. Example:
```PHP
$excel = new MSXLS('path_to_file.xls'); // Open file
$excel->read_everything(); // Read cells into $excel->cells
var_dump($excel->cells); // Output all parsed cells from XLS file
```

When `$excel->read_everything()` is invoked for the first time for your file, a private structure called __SST__ is built which contains all strings for all worksheets. It sits in memory until Parsing mode is changed or re-selected, or `$excel->free()` called, or your MSXLS instance (`$excel` variable in the examples here) had been destroyed. Therefore, it is rather memory-hungry mode if your file has a lot of unique strings. Non-unique strings are stored only once. Also, PHP is smart enough not to duplicate those strings in memory when a string is read into `$excel->cells` array from __SST__ storage.


In this mode, __empty rows and cells are ignored__. Boolean excel cells are parsed as `true` or `false`. If excel internally represents a whole number as _float_ (which is often the case), it will be parsed as _float_ without changes.


`$excel->cells` array first key is _zero based_ excel row number. For example, `$excel->cells[0]` will return an array of cells of excel row 1 (provided it is not empty). If your first populated row number in excel is 5, the first entry of `$excel->cells` will have index of 4.


`$excel->cells` array second key is _zero based_ excel column number. For example, `$excel->cells[0][3]` will return 'D' cell of the first row (cell `D1`), provided it is not empty.


Note that all empty rows and cells will create 'holes' in `$excel->cells` array, because empty cells are simply skipped. It is advisable to use `isset()` function to determine whether the cell is empty or not.


Array mode has only one additional setting for parsing: `$excel->set_fill_xl_errors($fill = false, $value = '#DEFAULT!')`, which defines whether or not to process excel cells with error values (such as #DIV/0!). Please refer to __*Public properties and methods*__ section for more information. In short, if `$fill` is `false`, error cells are skipped, otherwise they are filled with `$value`.

#### Row-by-row mode

This mode lets you read cells row by row. It is designed to let you process each row individually with using only necessary amount of memory. This mode is selected with `$excel->switch_to_row()` method. Data is read with `$excel->read_next_row()` method, which returns a single row as an array of cells.


As the method name implies, row number is advanced automatically, so next time you call `$excel->read_next_row()` method, it will read the next row. This method returns `null` if there are no more rows to read. You can manually set row number to read with `$excel->set_active_row($row_number)`, where `$row_number` is a valid zero-based excel row number. You can get the first and the last valid row number with `$excel->get_active_sheet()` method:
```PHP
$excel = new MSXLS('path_to_file.xls'); // Open file
$excel->switch_to_row(); // Switch mode to Row-by-row

$info = $excel->get_active_sheet(); //get selected sheet info
var_dump($info['first_row']); //displays first valid zero-based row index
var_dump($info['last_row']); //displays last valid zero-based row index
$excel->set_active_row($info['last_row']); //set active row to the last row of the sheet
$row = $excel->read_next_row(); //will read the last row of the sheet
```

The cell numbers in the returned row are _zero based_. Example:

```PHP
$row = $excel->read_next_row(); //$row now contains parsed cells from a single row
var_dump($row[0], $row[2]); //will output cell 'A' and cell 'C'
```

When `$excel->read_next_row()` is invoked for the first time for your file, __SST map__ will be built which is a structure that contains file stream offsets for every unique string in your excel file. It is similar to __SST__ structure in Array mode, but __SST__ contains the strings themselves, while __SST map__ only contains addresses of those strings.


When `$excel->read_next_row()` is invoked for the first time for selected sheet, __Rows map__ will be built. This structure contains file stream offsets for every excel row for currently selected worksheet.

Both of the structures mentioned above will be destroyed if Parsing mode is changed or re-selected, or if `$excel->free()` is called, or when your MSXLS instance is destroyed. Additionally, __Rows map__ will be destroyed when `$excel->select_sheet()` is called, because __Rows map__ is only valid for a selected sheet, unlike __SST map__, which is relevant for the whole file.

The main difference of Row-by-row mode is that it allowes many settings to be changed that affect which cells are proccessed and how. Please refer to __*Public properties and methods*__ section for more information. Methods that are relevant to Row-by-row mode settings are marked with __\[Row-by-row\]__ string.

### Debug mode

Debug mode enables output (echo) of all error and warning messages. To enable Debug mode, set the 2nd parameter to `true` in constructor:
```PHP
$file = new MSCFB("path_to_cfb_file.bin", true); // Show errors and warnings
```
It is also possible to show errors from MSCFB helper class. To do this, set the 4th parameter to `true` in constructor:
```PHP
$file = new MSCFB("path_to_cfb_file.bin", true, null, true);
```

**Warning!** PHP function name in which error occured is displayed alongside the actual message. Do not enable Debug mode in your production code since it may pose a security risk! This warning applies both to MSXLS class and MSCFB class.

### Temporary files and memory

If XLS file was saved as a Compound File (which is almost always the case), then MSXLS must use a temporary PHP stream resource to store Workbook stream that is extracted from the Compound File. It is stored either in memory or as a temporary file, depending on data size. By default, data that exceeds 2MiB (PHP's default value) is stored as a temporary file. XLS file may sometimes be stored as a Workbook stream itself, in which case a temporary file or stream is not needed and not created.

You can control when a temporary file is used instead of memory by specifying the threshold in bytes as the 3rd parameter to constructor. If Workbook stream size (in bytes) is less than this value, it will be stored in memory.
```PHP
$excel = new MSXLS("path_to_file.xls", false, 1024); //data with size > 1KiB is stored in a temp file
```
You can instruct PHP not to use a temporary file (thus always storing Workbook stream in memory) by setting this parameter to zero:
```PHP
$excel = new MSXLS("path_to_file.xls", false, 0); //temporary data is always stored in memory
```
Set this parameter to `null` to use default value:
```PHP
$excel = new MSXLS("path_to_file.xls", false, null); //default temp file settings
```
_Note:_ MSCFB helper class may also need to use a temporary stream resource. It will behave the same way as described above, and will also use that 3rd parameter as its memory limiter.

_Note:_ temporary files are automatically managed (created and deleted) by PHP.


## 4. How it works

## 5. Public properties and methods

### Properties

`(bool) $debug` — whether or not to display error and warning messages. Can be set as the 2nd parameter to constructor.

`(string) $err_msg` — a string that contains all error messages concatenated into one.

`(string) $warn_msg` — same as above, but for warnings.

`(array) $error` — array of error codes, empty if no errors occured.

`(array) $warn` — array of warning codes, empty if no warnings occured.

`(array) $cells` — two-dimensional array which is the storage for cells parsed in __Array__ mode. Filled when _read_everything()_ is invoked. This propertry is made public (instead of using a getter) mainly for performance reasons.

### Methods (functions)

---
#### General

---
`get_biff_ver()` — returns version of excel file. _5_ means 1995 XLS file, _8_ means 1997-2003 XLS file.

`get_codepage()` — returns CODEPAGE string. Relevant only for 1995 BIFF5 files, in which strings are encoded using a specific codepage. In BIFF8 (1997-2003) all strings are unicode (UTF-16 little endian).

---
`get_sheets()` — returns array of structures that represent all sheet info. See the code below.
```PHP
$excel = new MSXLS('file.xls');
$sheets = $excel->get_sheets(); //$sheets is array of sheet info structures
$sheet = reset($sheets); //$sheet now contains the first element of $sheets array

// Here is complete description of the sheet info structure:
$sheet['error'];         //[Boolean] whether an error occured while collecting sheet information
$sheet['err_msg'];       //[String] Error messages, if any
$sheet['name'];          //[String] Sheet name
$sheet['hidden'];        //[Integer] 0: normal, 1: hidden, 2: very hidden (set via excel macro)
$sheet['type'];          //[String] Sheet type: Worksheet, Macro, Chart, VB module or Dialog
$sheet['BOF_offset'];    //[Integer] Sheet byte offset in Workbook stream of XLS file
$sheet['empty'];         //*[Boolean] Whether the worksheet is empty
$sheet['first_row'];     //*[Integer] First non-empty row number of the worksheet
$sheet['last_row'];      //*[Integer] Last non-empty row number of the worksheet
$sheet['first_col'];     //*[Integer] First non-empty column number of the worksheet
$sheet['last_col'];      //*[Integer] Last non-empty column number of the worksheet
$sheet['cells_offset'];  //*[Integer] Byte offset of the 1st cell structure in Workbook stream

// Entries marked with * exist only for sheets of "Worksheet" type.
```

---
`get_valid_sheets()` — same as above, but returns only valid non-empty selectable worksheets. Additional _$sheet\['number'\]_ entry is present, which is the same number as the index of this sheet in the array returned by  _get_sheets()_.

`get_active_sheet()` — returns currently selected sheet info in the same structure that _get_valid_sheets()_ array consists of.

`get_filename()` — returns a file name string originally supplied to the constructor.

`get_filesize()` — returns size of the file supplied to the constructor (in bytes).

---
`get_margins($which = 'all')` — returns currently set margins for the selected worksheet. They are set automatically when the sheet is selected. Margins can be set manually with _set_margins()_ method. They define what rows and columns are read by _read_next_row()_ method.

_**$which**_ can be set to _'first_row'_, _'last_row'_, _'first_col'_, or _'last_col'_ string, in which cases a corresponding value will be returned. _**$which**_ also can be set to _'all'_ or left out, in which case an array of all four margins will be returned. If _**$which**_ is set to something not mentioned above, _false_ will be returned.

---
`set_encodings($enable = true, $from = null, $to = null, $use_iconv = false)` — manually set transcoding parameters for BIFF5 (1995 XLS file). This is usually not needed since the script detects these settings when file is opened.

_**$enable**_ enables encoding conversion of BIFF5 strings.

_**$from**_ is source encoding string, for example _'CP1252'_. Leaving it out or setting it to _null_ resets this parameter to detected internal BIFF5 codepage.

_**$to**_ is target encoding string, for example _'UTF-8'_. Leaving it out or setting it to _null_ resets this parameter to the value returned by _mb_internal_encoding()_.

_**$use_iconv**_ — If _true_, _iconv()_ will be used for convertion. Otherwise, _mb_convert_encoding()_ will be used.

---
`set_output_encoding($enc = null)` — sets output encoding which excel strings should be decoded to. _**$enc**_ is target encoding string. If parameter set to _null_ or left out, a value returned by `mb_internal_encoding()` will be used.

_Note:_ Setting _$to_ parameter in _set_encodings()_ and using _set_output_encoding()_ do the same thing. _set_output_encoding()_ is provided for simplicity if BIFF8 files are used.

---
`select_sheet($sheet = -1)` — Select a worksheet to read data from.

_**$sheet**_ must be either a sheet number or a sheet name. Use _get_valid_sheets()_ to get those, if needed. _-1_ or leaving out the parameter will select first valid worksheet.

---
`switch_to_row()` — switch to __Row-by-row__ parsing mode. Will also execute _free(false)_ and _select_sheet()_.

`switch_to_array()` — switch __Array__ parsing mode. Will also execute _free(false)_ and _select_sheet()_.

`read_everything()` — read all cells from file into _cells_ property. Works only in __Array__ mode.

`read_next_row()` — parses next row and returns array of parsed cells. Works only in __Row-by-row__ mode.

---
#### Memory free-ers

---
`free_stream()` — Close Workbook stream, free memory associated with it and delete temporary files.

`free_cells()` — re-initialize _cells_ array storage (parsed cell data from __Array__ mode).

`free_sst()` — re-initialize SST structure (Shared Strings Table from __Array__ mode).

`free_rows_map()` — re-initialize rows map storage used for __Row-by-row__ mode.

`free_sst_maps()` — re-initialize SST offsets map and SST lengths storage used for __Row-by-row__ mode.

`free_maps()` — execute both _free_row_map()_ and _free_sst_maps()_.

`free($stream = true)` — free memory by executing all "free"-related methods mentioned above. _free_stream()_ is called only if __*$stream*__ evaluates to _true_.

---
#### Reading settings (mostly for Row-by-row mode)

---
`set_fill_xl_errors($fill = false, $value = '#DEFAULT!')` — setup how cells with excel errors are processed. If __*$fill*__ evaluates to _true_, cells will be parsed as __*$value*__. _'#DEFAULT!'_ value is special as it will expand to actual excel error value. For example, if a cell has a number divided by zero, it will be parsed as _#DIV/0!_ string. If __*$value*__ is set to some other value, error cells will be parsed directly as __*$value*__. If __*$fill*__ evaluates to _false_, cells with errors will be treated as empty cells.  
Note: this is the only setting that also works in __Array__ mode.


`set_margins($first_row = null, $last_row = null, $first_col = null, $last_col = null)` — sets first row, last row, first column and last column that are parsed. If a parameter is _null_ or left out, the corresponding margin is not changed. If a parameter is _-1_, the corresponding margin is set to the default value. The default values correspond to the first/last non-empty row/column in a worksheet.

`set_active_row($row_number)` — set which row to read next. __*$row_number*__ is zero-based excel row number and it must not be out of bounds set by _set_margins()_ method.

`last_read_row_number()` — returns most recently parsed row number. Valid only if called immediately after _read_next_row()_.

`next_row_number()` -- returns row number that is to be parsed upon next call of _read_next_row()_.  Returns _-1_ if there is no more rows left to parse.

`set_empty_value($value = null)` — set __*$value*__ as _empty value_, a value that is used to parse empty cells as.

`use_empty_cols($set = false)` — whether or not to parse empty columns to _empty value_.

`use_empty_rows($set = false)` — whether or not to parse empty rows.
Note: if empty columns parsing is disabled (it is disabled by default), _read_next_row()_ will return _-1_ when an empty row is encountered. If empty columns parsing is enabled with _use_empty_cols(true)_, it will return array of cells filled with _empty value_.

`set_boolean_values($true = true, $false = false)` — set values which excel boolean cells are parsed as. By default, TRUE cells are parsed as PHP _true_ value, FALSE cells are parsed as PHP _false_ value.

`set_float_to_int($tf = false)` — whether or not to parse excel cells with whole float numbers to integers. Often whole numbers are stored as float internally in XLS file, and by default they are parsed as floats. This setting allows to parse such numbers as integer type. Note: cells with numbers internally stored as integers are always parsed as integers.

---
#### Constructor and destructor

---
`__construct($filename, $debug = false, $mem = null, $debug_MSCFB = false)` — open file, extract Workbook stream (or use the file as Workbook stream), execute _set_output_encoding()_ and _get_data()_ methods.

__*$filename*__ — path to XLS file.

__*$debug*__ — if evaluates to _true_, enables [Debug mode](#debug-mode). 

__*$mem*__ — sets memory limit for [temporary memory streams vs temporary files](#temporary-files-and_memory "Temporary files and memory").

__*$debug_MSCFB*__ — if evaluates to _true_, enables Debug mode in MSCFB helper class.

---
`__destruct()` — execute _free()_ method, thus closing all opened streams, deleting temporary files and erasing big structures.

## 6. Error handling

Each time an _error_ occures, the script places an error code into `$excel->error` array and appends an error message to `this->err_msg`. If an error occures, it prevents execution of parts of the script that depend on successful execution of the part where the error occured. _Warnings_ work similarly to errors except they do not prevent execution of other parts of the script, because they always occur in non-critical places. Warnings use `$excel->warn` to store warning codes and `$excel->warn_msg` for warning texts.

If an error occurs in constructor and Debug mode is disabled, the user should check if `$excel->error` non-strictly evaluates to `true` (for example, `if($excel->error){ /*error processing here*/ }`, in which case error text can be read from `$excel->err_msg` and the most recent error code can be obtained as the last element of `$excel->error` array. Same applies to Warnings, which use `$excel->warn_msg` and `$excel->warn`, respectively.

If Debug mode is enabled, errors and warnings are printed (echoed) to standart output.

## 7. Security considerations

There are extensive error checks in every function that should prevent any potential problems no matter what file is supplied to the constructor. The only potential security risk can come from the Debug mode, which prints a function name in which an error or a warning has occured, but even then I do not see how such information can lead to problems with this particular class. It's pretty safe to say that this code can be safely run in (automated) production of any kind. Same applies to MSCFB class.

## 8. Performance and memory

The MSXLS class has been optimized for fast parsing and data extraction, while still performing error checks for safety. It is possible to marginally increase constructor performance by leaving those error checks out, but I would strongly advise against it, because if a specially crafted mallicious file is supplied, it becomes possible to cause a memory hog or an infinite loop.

The following numbers were obtained on a Windows machine (AMD Phenom II x4 940), with a 97.0MiB test XLS file (96.2MiB Workbook stream) using WAMP server. XLS file entirely consists of unique strings.

```
  Time   Memory     Time   Memory
  
 - Open XLS file and parse its structure
 - Extract cells (Array mode)
 - Extract cells (Row-by-row, save extracted data to array)
 - Extract cells (Row-by-row, don't save extracted data)
 
  5.6.25 32-bit |  7.0.10 64-bit  - PHP Version
```

## 9. More documentation

All the code in __MSXLS.php__ file is heavily commented, feel free to take a look it. To understand how XLS file is structured, please refer to [MS documentation](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/ "Open official Microsoft XLS file documentation on Microsoft website"), or to [OpenOffice.org's Documentation of MS Compound File](https://www.openoffice.org/sc/excelfileformat.pdf "Open OpenOffice.org's Documentation of the Microsoft Excel File Format (PDF)") (also provided as a PDF file in this repository).

