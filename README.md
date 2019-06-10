# PHP-XLS-Excel-Parser
Probably, the fastest possible and the most efficient parser for XLS excel files for PHP!

Note: documentation writing is in progress.
Note 2: this parser is suitable __only for older XLS files__, not the newer ones, XLSX!

## Usage

1. This parser requires my MSCFB parser. Download __MSCFB.php__ from here: https://github.com/arti9m/PHP-MSCFB-Parser and put it in your PHP include directory or to your script directory.
2. Download __MSXLS.php__ from this repository and put it in the same directory as MSCFB.php.
3. Use the example below to extract data from your XLS file:
```PHP
require_once 'MSXLS.php'; // MSCFB.php is 'required once' in MSXLS.php
$excel = new MSXLS('path_to_file.xls'); // Open file and read basic information
$excel->read_everything(); // Read all non-empty cells from excel file
var_dump($excel->cells); // All read cells are saved in $excel->cells as 2-dimmensional array
```

That's it! When you are done, you can free memory and resources with these two lines:
```PHP
$excel->free();
unset($excel);
```
Note that `free()` is called in the destructor, so it is not strictly necessary to call it manually.

### Get data row by row
You can also extract cells row-by-row. It is slower, but much more memory efficient.
Use the code below as an example:
```PHP
require_once 'MSXLS.php'; // MSCFB.php is 'required once' in MSXLS.php
$excel = new MSXLS('path_to_file.xls'); // Open file and read basic information
$excel->switch_to_row(); // Switch to row-by-row mode
while($row = $excel->read_next_row()){
  $rows[] = $row; // $row contains cells from last read excel row
}
var_dump($rows); // $rows now contains each and every read row
```
If needed, free resources manually:
```PHP
$excel->free();
unset($excel);
```

### How it works
1. Main excel workbook stream is extracted from XLS file into a temporary file or into memory.
2. Basic data is collected and error checking is performed.
3. Upon user request, data is read and parsed. Some additional structures are created and stored in memory, if needed by selected reading mode.

A complete documentation writing is in progress, which will include information about public methods and properties, various reading parameters (row/column limiting, empty cells processing, etc), and also some benchmarks are to be included.
