# PHP-XLS-Excel-Parser
Probably, the fastest possible and the most efficient parser for XLS excel files for PHP!

_Note:_ this parser works __only with older XLS files__ that were used in Microsoft Excel 95 (BIFF5) and 97-2003 (BIFF8).  
It will not work with the newer ones, XLSX!

## 1. Requirements

At least __PHP 5.6__ 32-bit is required. Untested with PHP versions prior to 5.6.  
Works best with PHP 7.x 64-bit (faster, more memory efficient than PHP 5.6).


Also, this parser uses my [PHP MSCFB Parser](https://github.com/arti9m/PHP-MSCFB-Parser). Grab a copy of __MSCFB.php__ if you don't have one here: https://github.com/arti9m/PHP-MSCFB-Parser and put it in your PHP include directory or in the same directory where __MSXLS.php__ is. MSCFB is "required-once" inside MSXLS, so there's no need to include/require it manually.

## 2. Basic usage

1. Download __MSXLS.php__ from this repository and put it in your include directory or in the same directory where your script is.
2. Make sure that __MSCFB.php__ is in your include directory or in your script directory.
3. Add the following line to the beginning of your PHP script (specify full path to MSXLS.php, if needed):
```PHP
require_once 'MSXLS.php'; //MSCFB.php is 'required once' inside MSXLS.php
```
4. Create an instance of MSXLS (open XLS file):
```PHP
$excel = new MSXLS('path_to_file.xls');
```

5. If no errors occured up to this point, you are ready to read the cells from your file. There are two ways you can do it: either read all cells at once into a two-dimensional array using [Array mode](#1-array-mode "Array mode description") (faster), or read the cells in [Row-by-row mode](#2-row-by-row-mode "Row-by-row mode description"), which is slower, but is more configurable, suitable for database upload and may use much less memory depending on usage scenario.  
In any case, it's a good idea to check for errors before trying to read anything:

```PHP
if($excel->error) die($excel->err_msg); //Terminate script execution, show error message.
```

6. You can read all cells at once into a two-dimensional array:
```PHP
$excel->read_everything(); //Read cells into $excel->cells
```
At this point all your cells data is contained inside `$excel->cells` array:
```PHP
var_dump($excel->cells); //Output all parsed cells from XLS file
```

7. Or you can read the cells row by row:
```PHP
$excel->switch_to_row(); //switch to Row-by-row mode

while($row = $excel->read_next_row()){
  //You can process $row however you want here.
  //For example, you can upload a row into a database.
  $rows[] = $row; //For now, just store a parsed row inside $rows array.
}
```

_Note:_ `$excel->cells` will be erased when `$excel->switch_to_row()` is executed, so make sure you save the contents of `$excel->cells` (if any) to some other variable before switching to [Row-by-row mode](#2-row-by-row-mode). If you need to switch back to [Array mode](#1-array-mode), use `$excel->switch_to_array()` method.

8. If you need to perform some other memory-intensive tasks in the same script, it is a good idea to free some memory:
```PHP
$excel->free(); //This is also called in the destructor
unset($excel);
```

## 3. Advanced usage
_Note:_ every example in this section assumes that `$excel` is your MSXLS instance: `$excel = new MSXLS('file.xls')`.

---
### Sheet selection
If there is more than one worksheet in your file, and you want to parse the worksheet that is not the first valid non-empty worksheet, you will have to select your sheet manually. To do this, use `$excel->get_valid_sheets()` method to get an array with all available selectable worksheets. When the desired worksheet has been found, use its array index or _'number'_ entry as a parameter to `$excel->select_sheet($sheet)` method. For example:
```PHP
var_dump($excel->get_valid_sheets()); //outputs selectable sheets info
$excel->select_sheet(1); //select sheet with index 1
```
Alternatively, if you know sheet name, you can use it with the same method to select sheet:
```PHP
$excel->select_sheet('your_sheet_name'); //also works
```
Leave out sheet index/name to select the first available valid sheet:
```PHP
$excel->select_sheet(); //selects the first valid non-empty sheet in XLS file
```
You can use `$excel->get_active_sheet()` method to get information about selected sheet.  
Refer to [Methods (functions)](#methods-functions) subsection to get more information about methods mentioned above.

_Note:_ The first valid worksheet is selected automatically when the file is opened or when Parsing mode is changed.

---
### Parsing modes

There are two modes which the parser can work in: __Array__ mode and __Row-by-row__ mode. By default, Array mode is used.

#### 1. Array mode

This mode lets you read all cells at once into `$excel->cells` array property. It is designed to read all available data as fast as possible when no additional cells processing is needed. This mode is used by default. This mode can be selected with `$excel->switch_to_array()` method. Data is read with `$excel->read_everything()` method into `$excel->cells` array property. Example:
```PHP
$excel->read_everything(); //Read cells into $excel->cells
var_dump($excel->cells); //Output all parsed cells from XLS file
```

When `$excel->read_everything()` is invoked for the first time for your file, a private structure called __SST__ is built which contains all strings for all worksheets. It sits in memory until Parsing mode is changed or re-selected, or `$excel->free()` is called, or your MSXLS instance is destroyed. Therefore, it is rather memory-hungry mode if your file has a lot of unique strings. Non-unique strings are stored only once. Also, PHP is usually smart enough not to duplicate those strings in memory when a string is read into `$excel->cells` array from __SST__ storage, or when you copy `$excel->cells` to some other variable.


In this mode, __empty rows and cells are ignored__. Boolean excel cells are parsed as _true_ or _false_. If excel internally represents a whole number as _float_ (which is often the case), it will be parsed as _float_ type.


`$excel->cells` is a two-dimentional array. Its first dimension represents rows and its second dimension represents columns, both have zero-based numeration. See [Rows and columns numeration](#rows-and-columns-numeration) for more information.


Note that all empty rows and cells will create 'holes' in `$excel->cells` array, because empty cells are simply skipped. It is advisable to use `isset()` function to determine whether the cell is empty or not.


Array mode has only one additional setting for parsing: `$excel->set_fill_xl_errors($fill, $value)`, which defines whether or not to process excel cells with error values (such as division by zero). Please refer to [Methods (functions)](#methods-functions) subsection for more information. In short, if `$fill` is `false`, error cells are skipped, otherwise they are filled with `$value`.

#### 2. Row-by-row mode

This mode lets you read the cells row by row. It is designed to let you process each row individually while using as little memory as possible. This mode is selected with `$excel->switch_to_row()` method. Data is read with `$excel->read_next_row()` method, which returns a single row as an array of cells.


As the method name implies, row number is advanced automatically, so next time you call `$excel->read_next_row()`, it will read the next row. This method returns _null_ if there are no more rows to read. You can manually set row number to read with `$excel->set_active_row($row_number)`, where `$row_number` is a valid zero-based excel row number. You can get the first and the last valid row number with `$excel->get_active_sheet()` method:
```PHP
$info = $excel->get_active_sheet(); //get selected sheet info
var_dump($info['first_row']); //displays first valid zero-based row index
var_dump($info['last_row']); //displays last valid zero-based row index
$excel->set_active_row($info['last_row']); //set active row to the last row of the sheet
$row = $excel->read_next_row(); //will read the last row of the sheet
```

Cell numeration in the returned row is zero-based. See [Rows and columns numeration](#rows-and-columns-numeration) for more information.

When `$excel->read_next_row()` is invoked for the first time for your file, __SST map__ will be built which is a structure that contains file stream offsets for every unique string in your excel file. It is similar to __SST__ structure in Array mode, but __SST__ contains the strings themselves, while __SST map__ only contains addresses of those strings.


When `$excel->read_next_row()` is invoked for the first time for selected sheet, __Rows map__ will be built. This structure contains file stream offsets for every excel row for currently selected worksheet.

Both of the structures mentioned above will be destroyed if Parsing mode is changed or re-selected, or if `$excel->free()` is called, or when your MSXLS instance is destroyed. Additionally, __Rows map__ will be destroyed when `$excel->select_sheet()` is called, because __Rows map__ is only valid for a selected sheet, unlike __SST map__, which is relevant for the whole file.

One advantage of Row-by-row mode is that it allowes many settings to be changed that affect which cells are proccessed and how. Please refer to [Reading settings](#reading-settings-mostly-for-row-by-row-mode) part of [Methods (functions)](#methods-functions) subsection for more information.

---
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

---
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

### Rows and columns numeration
Rows and columns numeration in this parser is zero-based. Excel row numeration is numeric and starts from _1_, and column numeration is alphabetical and starts with _A_. Excel references a single cell by its column letter and row number, for example: A1, B3, C4, F9. If __Array__ mode is used, cells are stored in _$cells_ property, which is a two-dimensional array. The 1st index corresponds to row number, and the 2nd index is the column number. In __Row-by-row__ mode, a single row is returned as an array of cells. If `$row` contains a row returned by _read_next_row()_, Column A is `$row[0]`, column D is `$row[3]`, etc. In this mode, the user can get zero-based row number with _last_read_row_number()_ method. The table below illustrates how the cells are numerated.

|     | A | B | C | D | E | F |
|:---:|:-:|:-:|:-:|:-:|:-:|:-:|
|__1__| `$cells[0][0]`| `$cells[0][1]`| `$cells[0][2]`| `$cells[0][3]`| `$cells[0][4]`| `$cells[0][5]`|
|__2__| `$cells[1][0]`| `$cells[1][1]`| `$cells[1][2]`| `$cells[1][3]`| `$cells[1][4]`| `$cells[1][5]`|
|__3__| `$cells[2][0]`| `$cells[2][1]`| `$cells[2][2]`| `$cells[2][3]`| `$cells[2][4]`| `$cells[2][5]`|
|__4__| `$cells[3][0]`| `$cells[3][1]`| `$cells[3][2]`| `$cells[3][3]`| `$cells[3][4]`| `$cells[3][5]`|
|__5__| `$cells[4][0]`| `$cells[4][1]`| `$cells[4][2]`| `$cells[4][3]`| `$cells[4][4]`| `$cells[4][5]`|
|...|
|__row__| `$row[0]`| `$row[1]`| `$row[2]`| `$row[3]`| `$row[4]`| `$row[5]`|

### Some terms

A _Compound File_, or Microsoft Binary Compound File, is a special file format which is essentially a FAT-like container for other files.

_Workbook stream_, or just _Workbook_ is a binary bytestream that essentially represents excel BIFF file.

Excel file format is known as _BIFF_, or _Binary Interchangeable File Format_. There are several versions exist which differ in how they store excel data from version to version. This parser supports BIFF version 5, or BIFF5, which is the file format used in Excel 95, and BIFF version 8 (BIFF8), which is used in Excel 97-2003 versions. The biggest difference between BIFF5 and BIFF8 is that they store strings differently. In BIFF5, strings are stored inside cells in locale-specific 8-bit codepage (for example, CP1252), while BIFF8 has a special structure called _SST_ (Shared Strings Table), which stores unique strings inside itself in UTF16 little-endian encoding, and reference to SST entry is stored in cell.

Workbook stream consists of Workbook Globals substream and one or more Sheet substreams. Workbook Globals contains information about the file such as BIFF5 encoding, encryption, sheets information and much more (we do not actually need much more). Sheet substreams, or Sheets represent actual sheets that are created in Excel. They can be Worksheets, Charts, Visual Basic modules and some more, but only regular Worksheets can be parsed.

Excel keeps track of cells starting with first non-empty row and non-empty column, ending with last non-empty row and non-empty column. All other cells are completely ignored by this parser like they don't exist at all.

### What happens when I open XLS file

Note: during every stage extensive error checking is performed. See [Error handling](#error-handling) for more info.

When a user opens XLS file, for example by executing `$excel = new MSXLS('file.xls');`, first thing happens is the script checks whether XLS file is stored as a Compound File (most of the time it is) or as a Workbook stream. If it is a Compound File, the script attempts to extract Workbook stream to a temporary file and use that file in the future for all operations. Otherwise, it will directly use the supplied XLS file. The script never opens the supplied XLS file for writing.

After Workbook stream is accessed, the output encoding is set to _mb_internal_encoding()_ return value. Then _get_data()_ is executed: the script extracts information such as sheets count, codepage, sheets byte offsets, etc.

After that, either the first non-empty worksheet will be selected and ready for parsing and all other sheets information will be available to the user, or some error will be created (for example, when no non-empty worksheet was found).

By default, __Array__ parsing mode is active.

Attempts to invoke a __Row-by-row__ related method that is suitable for __Array__ mode only (and vice versa) will create an error, disabling any further actions most of the time.

If no errors occured, it is now possible to select and setup parsing mode.

When a worksheet is parsed, you can select another worksheet for parsing (if any) with _select_sheet()_ method. When you are finished parsing a file, it is a good idea to clean up, especially if something else is going on in your script later on. `$excel->free()` method and `unset($excel)` function called one after another is the best way to do it.

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

`get_data()` — Checks XLS file for errors and encryption, gathers information such as CODEPAGE for BIFF5, SST location for BIFF8. Gathers information about all sheets in the file. Also executes _select_sheet()_ to select first valid worksheet for parsing. This method is called automatically when XLS file is opened. Invoking it manually makes sence only if BIFF5 codepage was detected incorrectly and you cannot see sheet names (and you really need them). In this case, encoding settings must be configured with _set_encodings()_ after file opening and _get_data()_ should be called manually after it. 

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
`set_encodings($enable = true, $from = null, $to = null, $use_iconv = false)` — manually set transcoding parameters for BIFF5 (1995 XLS file). This is usually not needed since the script detects these settings when the file is opened.

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

`read_everything()` — read all cells from XLS file into _cells_ property. Works only in __Array__ mode.

`read_next_row()` — parses next row and returns array of parsed cells. Works only in __Row-by-row__ mode.

---
#### Memory free-ers
`free_stream()` — Close Workbook stream, free memory associated with it and delete temporary files.

`free_cells()` — re-initialize _cells_ array storage (parsed cell data from __Array__ mode).

`free_sst()` — re-initialize SST structure (Shared Strings Table from __Array__ mode).

`free_rows_map()` — re-initialize rows map storage used for __Row-by-row__ mode.

`free_sst_maps()` — re-initialize SST offsets map and SST lengths storage used for __Row-by-row__ mode.

`free_maps()` — execute both _free_row_map()_ and _free_sst_maps()_.

`free($stream = true)` — free memory by executing all "free"-related methods mentioned above. _free_stream()_ is called only if __*$stream*__ evaluates to _true_.

---
#### Reading settings (mostly for Row-by-row mode)
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
`__construct($filename, $debug = false, $mem = null, $debug_MSCFB = false)` — open file, extract Workbook stream (or use the file as Workbook stream), execute _set_output_encoding()_ and _get_data()_ methods.

__*$filename*__ — path to XLS file.

__*$debug*__ — if evaluates to _true_, enables [Debug mode](#debug-mode). 

__*$mem*__ — sets memory limit for [temporary memory streams vs temporary files](#temporary-files-and-memory "Temporary files and memory").

__*$debug_MSCFB*__ — if evaluates to _true_, enables Debug mode in MSCFB helper class.

---
`__destruct()` — execute _free()_ method, thus closing all opened streams, deleting temporary files and erasing big structures.

## 6. Error handling

Each time an __error__ occures, the script places an error code into __*$error*__ array property and appends an error message to __*$err_msg*__ string property. If an error occures, it prevents execution of parts of the script that depend on successful execution of the part where the error occured. __Warnings__ work similarly to errors except they do not prevent execution of other parts of the script, because they always occur in non-critical places. Warnings use __*$warn*__ property to store warning codes and __*$warn_msg*__ for warning texts.

If Debug mode is disabled, you should check if __*$error__ property evaluates to _true_, which would mean that __*$error*__ array is not empty, i.e. has one or multiple error codes as its elements. Error handling example:
```PHP
$excel = new MSXLS('nofile.xls'); //Try to open non-existing file

if($excel->error){
  var_dump(end($excel->error)); //Will output last error code
  var_dump($excel->err_msg); //Will output all errors texts
  die(); //Terminate script execution
}

if($excel->warn){
  var_dump(end($excel->warn)); //Will output last warning code
  var_dump($excel->warn_msg); //Will output all warnings texts
}
```

If Debug mode is enabled, errors and warnings are printed (echoed) to standart output automatically.

## 7. Security considerations

There are extensive error checks in every function that should prevent any potential problems no matter what file is supplied to the constructor. The only potential security risk can come from the Debug mode, which prints a function name in which an error or a warning has occured, but even then I do not see how such information can lead to problems with this particular class. It's pretty safe to say that this code can be safely run in (automated) production of any kind. Same applies to MSCFB class.

## 8. Performance and memory

The MSXLS class has been optimized for fast parsing and data extraction, while still performing error checks for safety. It is possible to marginally increase performance by leaving those error checks out, but I would strongly advise against it, because if a specially crafted mallicious file is supplied, it becomes possible to cause a memory hog or an infinite loop.

The following numbers were obtained on a Windows machine (AMD Phenom II x4 940), with a 97.0 MiB test XLS file (96.2 MiB Workbook stream) using WAMP server. XLS file consists entirely of unique strings.

| Time | Memory | Time | Memory | Action | 
|:-:|:-:|:-:|:-:|---|
|  7.52s |   1.0 MiB |  3.48s |   0.6 MiB | Open XLS File (create MSXLS instance)
| 77.77s | 213.2 MiB | 16.41s | 128.8 MiB | Open XLS File and parse in __Array__ mode
| 91.08s | 192.2 MiB | 27.20s | 204.3 MiB | Open file, parse in __Row-by-row__ mode to variable
| 54.71s |  82.9 MiB | 21.49s |  82.1 MiB | Open file, parse in __Row-by-row__ mode (don't save)
|__PHP 5.6.25__ |__PHP 5.6.25__ |__PHP 7.0.10__ |__PHP 7.0.10__ |

_Note:_ It took 1.65 seconds and 12.0 MiB of memory to parse a real-life XLS pricelist of 13051 entries in __Array__ mode in PHP 7.0.10. That XLS file was 3.45 MiB in size.

## 9. More documentation

All code in __MSXLS.php__ file is heavily commented, feel free to take a look at it. To understand how XLS file is structured, please refer to [MS documentation](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/ "Open official Microsoft XLS file documentation on Microsoft website"), or to [OpenOffice.org's Documentation of MS Compound File](https://www.openoffice.org/sc/excelfileformat.pdf "Open OpenOffice.org's Documentation of the Microsoft Excel File Format (PDF)") (also provided as a PDF file in this repository).
