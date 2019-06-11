# PHP-XLS-Excel-Parser
Probably, the fastest possible and the most efficient parser for XLS excel files for PHP!

_Note:_ this parser is suitable __only for older XLS files__, not the newer ones, XLSX!

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

### Parsing modes
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
_Note:_ MSCFB helper class may also need to use a temporary stream resource. It will behave the same way as described above, and will also use that 3rd parameter as its memory limiter.

_Note:_ temporary files are automatically managed (created and deleted) by PHP.


## 4. How it works

## 5. Public properties and methods

### Properties

### Methods (functions)


## 6. Error handling

Each time an _error_ occures, the script places an error code into `$this->error` array and appends an error message to `this->err_msg`. If an error occures, it prevents execution of parts of the script that depend on successful execution of the part where the error occured. _Warnings_ work similarly to errors except they do not prevent execution of other parts of the script, because they always occur in non-critical places. Warnings use `$this->warn` to store warning codes and `$this->warn_msg` for warning texts.

If an error occurs in constructor and Debug mode is disabled, the user should check if `$this->error` non-strictly evaluates to `true` (for example, `if($this->error){ /*error processing here*/ }`, in which case error text can be read from `$this->err_msg` and the most recent error code can be obtained as the last element of `$this->error` array. Same applies to Warnings, which use `$this->warn_msg` and `$this->warn`, respectively.

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

