<?php
require_once 'MSCFB.php';

class MSXLS{

  /* --------------------------------------------- */
  /*                   CONSTANTS                   */
  /* --------------------------------------------- */

  // RECORD NAMES
  const BLANK = 0x201;
  const BOOLERR = 0x205;
  const LABEL = 0x204;
  const LABELSST = 0xFD;
  const MULBLANK = 0xBE;
  const MULRK = 0xBD;
  const NUMBER = 0x203;
  const RK = 0x27e;
  const RSTRING = 0xD6;
  const ROW = 0x208;
  const NOTE = 0x1C;
  const WINDOW2 = 0x23E;
  const EOF = 0x0A;
  const BOF = 0x809;
  const DIMENSION = 0x200;
  const SHEET = 0x85;
  const WINDOW1 = 0x3D;
  const FILEPASS = 0x2F;
  const CODEPAGE = 0x42;
  const SST = 0xFC;
  const CONT = 0x3C;
  const DBCELL = 0xD7;
  const FORMULA = 0x06;
  const STR = 0x207;
  const SHEETPR = 0x81;

  // ERROR CODES
  const E_EOF = 0;
  const E_SEEK = 1;
  const E_TELL = 2;
  const E_OPEN = 3;
  const E_CLOSE = 4;
  const E_PREV = 5;
  const E_SST5 = 6;
  const E_SST_NOPOS = 7;
  const E_SST_EXPECT = 8;
  const E_ARRAYMODE = 9;
  const E_MARG_NOSHEET = 10;
  const E_MARG_OUTOFBOUNDS = 11;
  const E_SEL_NOSHEETS = 12;
  const E_SEL_WRONGINDEX = 13;
  const E_SEL_WRONGNAME = 14;
  const E_READ_NOSHEET = 15;
  const E_READ_UNEXP_REC = 16;
  const E_HDR_BIFF = 17;
  const E_HDR_SIZE5 = 18;
  const E_HDR_SIZE8 = 19;
  const E_HDR_SIZEMAX = 20;
  const E_HDR_CRYPT = 21;
  const E_NOFILE = 22;
  const E_HDR_NOVALIDSHEET = 23;
  const E_HDR_NOSHEETS = 24;
  const E_HDR_NOWSHEETS = 25;
  const E_HDR_2GLOBALS = 26;
  const E_ROWBYROW = 27;
  const E_NOACTSHEET = 28;
  const E_INT = 29;
  const E_CFB = 30;
  const E_NOSTREAM = 31;
  const E_TEMP = 32;
  const E_BADHANDLE = 33;
  const E_EXTRACT = 34;
  const E_UNREACH = 35;
  const SH_OFFSET = 36;
  const SH_TYPE_GLOB = 37;
  const SH_WORKSPACE = 38;
  const SH_UNKNOWN = 39;
  const SH_TYPE_SHEETPR = 40;
  const SH_NODIMM = 41;
  const SH_OUB_SR = 42;
  const SH_OUB_LR = 43;
  const SH_OUB_FC = 44;
  const SH_OUB_LC = 45;
  const E_NOBOF = 46;
  const E_NOWBSTREAM = 47;
  
  

  /* --------------------------------------------- */
  /*                   PROPERTIES                  */
  /* --------------------------------------------- */

  public $debug = false; // whether or not errors and warnings are echoed

  // error texts
  private $E = array(
  self::E_EOF => 'Unexpected EOF or stream read error!',
  self::E_SEEK => 'Unable to set stream position!',
  self::E_TELL => 'Unable to get current stream position!',
  self::E_OPEN => 'Unable to open file (stream)!',
  self::E_CLOSE => 'Unable to close file (stream)!',
  self::E_PREV => 'Function exited because of previous error.',
  self::E_SST5 => 'SST-related function called for BIFF5!',
  self::E_SST_NOPOS => 'Attempted to read SST data when SST position is unknown!',
  self::E_SST_EXPECT => 'Expected SST record, got something else!',
  self::E_ARRAYMODE => 'Attempted action suitable only for "array" mode, but "row-by-row" mode is active.',
  self::E_MARG_NOSHEET => 'Attempted to set margins when no sheet is selected!',
  self::E_MARG_OUTOFBOUNDS => 'Attempted to set margins out of sheet bounds',
  self::E_SEL_NOSHEETS => 'There are no valid non-empty worksheets!',
  self::E_SEL_WRONGINDEX => 'Incorrect sheet index! Sheet does not exist or is not a valid worksheet!',
  self::E_SEL_WRONGNAME => 'Incorrect sheet name! Sheet does not exist or is not a valid worksheet!',
  self::E_READ_NOSHEET => 'Tried to read data when no sheet is selected!',
  self::E_READ_UNEXP_REC => 'Got unexpected record!',
  self::E_HDR_BIFF => 'Unsupported BIFF version!',
  self::E_HDR_SIZE5 => 'Minimum filesize for BIFF5 is 166 bytes!',
  self::E_HDR_SIZE8 => 'Minimum filesize for BIFF8 is 194 bytes!',
  self::E_HDR_SIZEMAX => 'Maximum filesize is 2GB!',
  self::E_HDR_CRYPT => 'File is encrypted!',
  self::E_NOFILE => 'File does not exist!',
  self::E_HDR_NOVALIDSHEET => 'No valid non-empty sheet found in file!',
  self::E_HDR_NOSHEETS => 'No sheets found in file!',
  self::E_HDR_NOWSHEETS => 'No worksheets found in file!',
  self::E_HDR_2GLOBALS => 'Second Globals substream found, this is impossible!',
  self::E_ROWBYROW => 'Attempted action suitable only for "row-by-row" mode, but "array" mode is active.',
  self::E_NOACTSHEET => 'No valid sheet is selected!',
  self::E_INT => 'Expected integer value!',
  self::E_CFB => 'Error occured while opening Compound File!',
  self::E_NOSTREAM => 'No "Workbook" or "Book" stream found in Compound File!',
  self::E_TEMP => 'Failed to create temporary file or stream!',
  self::E_BADHANDLE => 'Stream handle does not point to valid stream (stream closed or does not exist)!',
  self::E_EXTRACT => 'Unable to extract Workbook stream from Compound File!',
  self::E_UNREACH => 'Unreachable place reached!',
  self::SH_OFFSET => 'Invalid sheet offset!',
  self::SH_TYPE_GLOB => 'Unknown sheet type (in Globals substream)!',
  self::SH_WORKSPACE => 'Workspace substream, unsupported!',
  self::SH_UNKNOWN => 'Unknown substream!',
  self::SH_TYPE_SHEETPR => 'Type unknown, failed to find SHEETPR record!',
  self::SH_NODIMM => 'Failed to find DIMENSION record!',
  self::SH_OUB_SR => 'Starting row out of bounds!',
  self::SH_OUB_LR => 'Last row out of bounds!',
  self::SH_OUB_FC => 'First column out of bounds!',
  self::SH_OUB_LC => 'Last column out of bounds!',
  self::E_NOBOF => 'BOF record must be the first record in a Workbook stream but it is not!',
  self::E_NOWBSTREAM => 'No "Workbook" or "Book" stream found in XLS file! File is invalid!',
  );

  private $stream = null; // stream handle: either file itself or a temporary stream

  private $filename = null; // Getter: get_filename()
  private $filesize = null; // Getter:  get_filesize()

  private $sheets = array(); // array of XL sheets info, Getter: get_sheets()
  private $valid_sheets = array(); // non-empty worksheets, Getter: get_valid_sheets()

  private $BIFF_VER = 0; // 5 or 8, Getter: get_biff_ver();
  private $CODEPAGE = 'ASCII'; // Default BIFF5 codepage, Getter: get_codepage()

  public $err_msg = ''; // error messages storage
  public $warn_msg = ''; // warning message storage
  public $error = array(); // active error codes container
  public $warn = array(); // active warning codes container

  // default excel cell error value, '#DEFAULT!' is special: for example,
  // it expands to '#DIV/0!' if cell has div by zero condition.
  // See also: $err_vals.
  // Setter: set_fill_xl_errors()
  private $err_val = '#DEFAULT!';

  // whether to fill cells with errors using $err_val or skip them
  // Setter: set_fill_xl_errors()
  private $fill_err = false;

  // values for 'true' and 'false' cells
  // Setter: set_boolean_values()
  private $true_val = true;
  private $false_val = false;

  // boundaries storage
  // Setter: set_margins(); Getter: get_margins().
  private $first_row = 0;
  private $last_row = 0;
  private $first_col = 0;
  private $last_col = 0;

  private $empty_cols = false; // process empty cols or not. Setter: use_empty_cols()
  private $empty_val = null; // empty value to fill empty cell. Setter: set_empty_value()
  private $empty_rows = false; // process empty rows or not. Setter: use_empty_rows()

  // set with set_encodings(convert_enc,source_enc,target_enc,use_iconv)
  // set $target_enc with set_output_encoding(target_enc)
  private $convert_enc = true; // whether to convert encoding or not
  private $source_enc = ''; // encoding to convert from
  private $target_enc = ''; // encoding to convert to
  private $use_iconv = false; // use mb_convert_encoding or iconv
  private $CP_set = false; // needed only once for get_data()

  // Mode: get cells as an 'array' (false), or get them 'row-by-row' (true).
  // switch_to_array() sets this to false, switch_to_row() sets this to true
  private $row_by_row = false;

  // whether to convert whole floats to integer. Setter: set_float_to_int()
  private $float_to_int = false;

  // result storage for 'array' mode
  public $cells = array();

  private $last_parsed_row = -1; // needed for row-by-row only

  private $SST = array(); // shared strings table, filled with read_SST()
  private $SST_pos = 0; // file offset for SST record, filled with get_data();

  // SST map: [sst_index][0] = file_offset, [sst_index][1] = continue1_offset etc
  // filled with build_SST_map(); needed only for row-by-row mode
  private $SST_map = array();

  // array[sst_index][x] = bytes to read (x>0 means that CONTINUE record is involved)
  // filled with build_SST_map();
  private $SST_lengths = array();

  // holds file offsets to first record of [N]th row
  // filled with build_rows_map();
  private $rows_map = array();

  //how many bytes of string left to read, helper for some string related functions
  private $strstate = 0;

  private $err_vals = array( //excel cell error values
    0x00 => '#NULL!',
    0x07 => '#DIV/0!',
    0x0F => '#VALUE!',
    0x17 => '#REF!',
    0x1D => '#NAME?',
    0x24 => '#NUM!',
    0x2A => '#N/A!'
  );

  private $active_sheet = null; // active (selected) sheet info storage

  /* --------------------------------------------- */
  /*                    METHODS                    */
  /* --------------------------------------------- */


  /* ----------- 1. GENERAL FUNCTIONS ------------ */
  /* --------------------------------------------- */

  // Generate error or warning message and set error or warning flag.
  // If Debug mode enabled, function name from which error originates is appended to error message
  private function gen_err($code, $func_name = 'general', $warn = false){
    $h = $warn ? 'WARNING: ' : 'ERROR: '; //heading for the message

    // if $code is integer, get error text using $code
    if(gettype($code)==='integer') $txt = $this->E[$code];
    else { // otherwise assume that $code is error text itself
      $txt = $code;
      $code = -1;
    }

    //if Debug mode enabled, create html formatted message
    if($this->debug){
      $html = '<br>'.$h.'<b>['.$func_name.']</b> '.$txt.'<br>';
      $txt = '['.$func_name.'] '.$txt;
    }

    $txt = preg_replace('/\s+/',' ', $txt); //replace all whitespaces with ' '
    if($this->debug) echo $html; //if Debug mode enabled, echo message

    if($warn){
      $this->warn_msg = $txt.' '.$this->err_msg; //append warning to warnings string
      $this->warn[] = $code; //add code to active codes list
    } else {
      $this->err_msg = $txt.' '.$this->err_msg; //append error to errors string
      $this->error[] = $code; //add code to active codes list
    }
    return $txt;
  }

  //Generates codepage name from excel codepage number
  private function gen_cp($cp_number){
    switch($cp_number){
      case 1200: return 'UTF-16LE';
      case 367: return 'ASCII';
      case 10000: return 'MAC';
      case 32768: return 'MAC2';
      default: return 'CP'.$cp_number; //ex: CP1251
    }
  }


  /* ----------- 2. GETTERS & SETTERS ------------ */
  /* --------------------------------------------- */

  public function get_biff_ver(){ //return BIFF version
    return $this->BIFF_VER;
  }

  public function get_codepage(){ //return codepage string (relevant for BIFF 5 only)
    return $this->CODEPAGE;
  }

  public function get_sheets(){ //return array with info about all sheets of all types
    return $this->sheets;
  }

  public function get_valid_sheets(){ //return array with info about valid selectable sheets
    return $this->valid_sheets;
  }

  public function get_active_sheet(){ //return currently selected sheet info
    return $this->active_sheet;
  }

  public function get_filename(){ //return name of the main stream (usually XLS file)
    return $this->filename;
  }

  public function get_filesize(){ //return size of the main stream (usually XLS file)
    return $this->filesize;
  }

  //whether or not fill cells with excel errors and default error value
  public function set_fill_xl_errors($fill = false, $value = '#DEFAULT!'){
    $this->fill_err = (bool) $fill;
    $this->err_val = $value;
  }

  //get 'first_row', 'last_row', 'first_col', 'last_col' or 'all' of them in array
  //returns false if $which is not recognized
  public function get_margins($which = 'all'){
    switch($which){
      case 'all':
        $ret = array();
        $ret['first_row'] = $this->first_row;
        $ret['last_row'] = $this->last_row;
        $ret['first_col'] = $this->first_col;
        $ret['last_col'] = $this->last_col;
        return $ret;
      case 'first_row':
        return $this->first_row;
      case 'last_row':
        return $this->last_row;
      case 'first_col':
        return $this->first_col;
      case 'last_col':
        return $this->last_col;
      default:
        return false;
    }
  }

  //set transcoding parameters (relevant only for BIFF5)
  //$enable: whether or not enable strings transcoding
  //$from: source encoding name (example: 'CP1251'), null resets to $this->CODEPAGE
  //$to: target encoding name (example: 'UTF-8'), null resets to mb_internal_encoding()
  //$use_iconv: whether to use iconv() (true) or mb_convert_encoding() (false)

  public function set_encodings($enable = true, $from = null, $to = null, $use_iconv = false){
    $this->convert_enc = (bool) $enable;
    $this->use_iconv = (bool) $use_iconv;

    if($from===null) $this->source_enc = $this->CODEPAGE;
    elseif(gettype($from)==='string') $this->source_enc = $from;

    if($to===null) $this->target_enc = mb_internal_encoding();
    elseif(gettype($to)==='string') $this->target_enc = $to;
  }

  //set output encoding for strings decoding (shares same variable with set_encodings!!!)
  //$enc: target encoding name, defaults to mb_internal_encoding() if $enc is null
  public function set_output_encoding($enc = null){
    if($enc===null){
      $this->target_enc = mb_internal_encoding();
    }
    if(gettype($enc)==='string') $this->target_enc = $enc;
  }

  //[row-by-row] Whether include empty rows or not
  public function use_empty_rows($set = false){
    $this->empty_rows = (bool) $set;
  }

  //[row-by-row] whether include empty columns or not
  public function use_empty_cols($set = false){
    $this->empty_cols = (bool) $set;
  }

  //[row-by-row] Set a value for an empty cell
  public function set_empty_value($value = null){
    $this->empty_val = $value;
  }

  //[row-by-row] Set custom cell boolean values
  public function set_boolean_values($true = true, $false = false){
    $this->true_val = $true;
    $this->false_val = $false;
  }

  //[row-by-row] Set row number to read next. See valid margins using get_active_sheet().
  //returns false on failure, true on success.
  public function set_active_row($row_number){
    if($this->error){
      $this->gen_err(self::E_PREV, __FUNCTION__);
      return false;
    }
    if(!$this->active_sheet){
      $this->gen_err(self::E_NOACTSHEET, __FUNCTION__);
      return false;
    }
    if(!$this->row_by_row){
      $this->gen_err(self::E_ROWBYROW, __FUNCTION__);
      return false;
    }
    if(gettype($row_number)!=='integer'){
      $this->gen_err(self::E_INT, __FUNCTION__);
      return false;
    }
    if($row_number < $this->first_row || $row_number > $this->last_row){
      $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
      return false;
    }

    $this->last_parsed_row = ($row_number - 1);
    return true;
  }

  // [row-by-row] Returns index of last parsed excel row
  // VALID ONLY IMMEDIATELY AFTER read_next_row()!
  // returns false on error

  public function last_read_row_number(){
    if(!$this->active_sheet){
      $this->gen_err(self::E_NOACTSHEET, __FUNCTION__);
      return false;
    }
    if(!$this->row_by_row){
      $this->gen_err(self::E_ROWBYROW, __FUNCTION__);
      return false;
    }
    return $this->last_parsed_row;
  }

  //[row-by-row] Returns excel row index of next row to parse
  //returns -1 if all rows parsed and there's no next row
  //returns false if error
  public function next_row_number(){
    if(!$this->active_sheet){
      $this->gen_err(self::E_NOACTSHEET, __FUNCTION__);
      return false;
    }
    if(!$this->row_by_row){
      $this->gen_err(self::E_ROWBYROW, __FUNCTION__);
      return false;
    }
    if($this->last_parsed_row===$this->last_row) return -1;
    return $this->last_parsed_row-1;
  }

  //[row-by-row] If float is a whole number, convert it to integer
  public function set_float_to_int($tf = false){
    $this->float_to_int = (bool) $tf;
  }


  /* --- 3. BINARY READERS AND FILE TRAVERSERS --- */
  /* --------------------------------------------- */

  //Search records from record_list until a record from stop_records occur
  //returns false on error and when EOF reached
  //returns Record Header if specified record has been found
  //note: don't forget about file position which is essentially a shared state!
  //$file: stream handle
  //$record_list: array of record name constants representing records we are looking for
  //$stop_records: array of record name constants representing records that stop searching
  //$once: if set to true, once record has been found, it is removed from $record_list
  private function find_recs_before_rec($file, &$record_list, &$stop_records, $once = true){
    while($record_list){ //while there are record IDs in $record_list
      $header = fread($file,4); //read record headers
      if(false===$header || strlen($header)!==4){ //EOF reached
        return false;
      }

      $header = unpack('vID/vL', $header);

      foreach($stop_records as $key => $value){ //loop through stop-records
        if($header['ID']===$value){ //if this record is a stop-record
          //seek to the beginning of stop-record
          if(-1 === fseek($file,-4,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            //no file close here!
            return false;
          }
          unset($stop_records[$key]); //remove stop-record from stop-records list
          return false; //exit from function
        }
      }

      foreach($record_list as $key => $value){ //loop through list of records we are looking for
        if($header['ID']===$value){ //if we found one
          if($once) unset($record_list[$key]); //if $once is set, remove record from list
          return $header; //return record header. File position is the beginning of record data
        }
      }

      //seek to the beginning of next record (skip current record data)
      if(-1 === fseek($file,$header['L'],SEEK_CUR)){
        $this->gen_err(self::E_SEEK, __FUNCTION__);
        //no file close here!
        return false;
      }
    }
    return false;
  }


  /* -------- 4. STRING RELATED FUNCTIONS -------- */
  /* --------------------------------------------- */

  //Parse unformatted excel string (uses $strstate)
  //$input: raw input data
  //$len_bits: string length size: 8bit or 16bit
  //returns processed string on success, false on error
  private function parse_unf_str($input, $len_bits){ //len_bits is 8 or 16
    $cut = $len_bits/8; //how many bytes to cut later
    $par = ($len_bits === 8) ? 'c' : 'v'; //parameter for unpack()

    $strlen = unpack("{$par}sl", $input)['sl']; //unpack string length (characters, not bytes)

    if($this->BIFF_VER===5){ //version 5 simply uses 1-byte characters
      $input = substr($input, $cut, $strlen); //cut input to $strlen bytes
      $this->strstate = $strlen - strlen($input); //update strstate (needed for CONTINUE records)
      if($this->convert_enc){ //convert encoding if needed
        if($this->use_iconv){ //either using iconv
          return iconv($this->source_enc, $this->target_enc, $input);
        } else { //or using mb_convert_encoding
          return mb_convert_encoding($input, $this->target_enc, $this->source_enc);
        }
      }
      return $input;
    }

    if($this->BIFF_VER===8){
      $opt = unpack("x$cut/copt",$input)['opt']; //skip $cut bytes and unpack opt flag

      $comp = ($opt & 0b1) === 0; //if bit 0 is set, then there's no compression
      $cmp = $comp ? 1 : 2; //multiplier for character length to get byte length
      $bytelen = $cmp*$strlen; //length of the string in bytes

      $input = substr($input, $cut+1, $bytelen); //skup $cut bytes and opt flag
      $this->strstate = $bytelen - strlen($input); //update strstate
      if($comp) $input = implode("\0",str_split($input))."\0"; //if compression used, uncompress
      $input = mb_convert_encoding($input, $this->target_enc, 'UTF-16LE'); //convert from utf16
      return $input;
    }
    $this->gen_err(self::E_UNREACH, __FUNCTION__, true); //generate warning
    return false; //should never reach here
  }

  //Continue parsing unformatted excel string (uses $strstate)
  private function parse_unf_str_cont($input){
    //contents of this function is pretty much the same as parse_unf_str()
    if(!$this->strstate) return ''; //if no bytes remain - return empty string
    if($this->BIFF_VER===5){
      $input = substr($input, $this->strstate);
      $this->strstate -= strlen($input);
      if($this->convert_enc){
        if($this->use_iconv){
          return iconv($this->source_enc, $this->target_enc, $input);
        } else {
          return mb_convert_encoding($input, $this->target_enc, $this->source_enc);
        }
      }
      return $input;
    }
    if($this->BIFF_VER===8){
      $comp = !ord($input[0]); //compression flag is present in CONTINUE record, 'unpack' it
      $input = substr($input, 1, $this->strstate);
      $this->strstate -= strlen($input);
      if($comp) $input = implode("\0",str_split($input))."\0";
      $input = mb_convert_encoding($input,  $this->target_enc, 'UTF-16LE');
      return $input;
    }
    $this->gen_err(self::E_UNREACH, __FUNCTION__, true); //generate warning
    return false;
  }

  //Handle all sequential CONTINUE records. $coord holds (r)ow and (c)olumn
  //This function is helper for read_everything()
  //Returns true on success, false on error
  //This function directly modifies $this->cells
  private function parse_str_cont($file,$coord){
    while(true){
      $hdr = fread($file,4); //read record headers
      if($hdr===false || strlen($hdr)!==4){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        return false;
      }
      $hdr = unpack('vID/vL',$hdr); //unpack record headers

      if($hdr['ID']!==self::CONT){ //if record is not CONTINUE
        //seek back 4 bytes (before header)
        if(-1===fseek($file,-4,SEEK_CUR)){
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        return true; //exit function
      }

      $part = fread($file, $hdr['L']); //read record data
      if($part===false || strlen($part)!==$hdr['L']){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      
      //append parsed part of string to cell value
      $this->cells[$coord['r']][$coord['c']] .= $this->parse_unf_str_cont($part);
    }
  }

  // Handle all sequential CONTINUE records.
  // Returns string part, or false on error
  // This function is helper for read_next_row()
  private function parse_str_cont_row($file){
    $ret = ''; //initial return string
    while(true){
      $hdr = fread($file,4); //read record header
      if($hdr===false || strlen($hdr)!==4){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $hdr = unpack('vID/vL',$hdr); //unpack record header

      if($hdr['ID']!==self::CONT){ //if record is not CONTINUE
        //seek back 4 bytes (before header)
        if(-1===fseek($file,-4,SEEK_CUR)){
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        return $ret; //exit function
      }

      $str_data = fread($file, $hdr['L']); //read string binary data
      if($str_data===false || strlen($str_data)!==$hdr['L']){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      //parse and append string to ret value
      $ret .= $this->parse_unf_str_cont($str_data);
    }
  }

  //Return string value using SST map, false on error
  private function decode_sst_string($file, $sst_index){
    
    if($this->error){ //exit if previous error exists
      $this->gen_err(self::E_PREV, __FUNCTION__);
      return false;
    }

    if(false === ($pos = ftell($file))){ //get file position
      $this->gen_err(self::E_TELL, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    //make copies of relevant data and use them to get our strings
    $readlen = $this->SST_lengths[$sst_index]; //lengths of string parts
    $map = $this->SST_map[$sst_index]; //offsets of string parts

    $i_max = count($map); //how many string parts are there

    $ret_str = ''; //string placeholder

    for($i=0; $i<$i_max; $i++){

      //seek to SST location
      if(-1===fseek($file,$map[$i])){
        $this->gen_err(self::E_SEEK,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      if($i===0){ // this is not CONTINUE record

        $strhdr = fread($file,3); //read string headers
        if($strhdr===false || strlen($strhdr)!==3){
          $this->gen_err(self::E_EOF,__FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }

        $readlen[$i] -= 3; //keep track of how much did we read

        $strhdr = unpack('vL/copt', $strhdr); //unpack string headers

        $comp = ($strhdr['opt'] & 1) === 0; //whether or not utf-16 string is compressed

        //if bit3 is set, there's rich formatting info after headers, skip it
        if(($strhdr['opt'] & 0b1000) > 0){
          if(-1===fseek($file,2,SEEK_CUR)){ //skip rich info
            $this->gen_err(self::E_SEEK,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $readlen[$i] -= 2;
        }

        //if bit2 is set, there's asian transcription info after headers, skip it
        if(($strhdr['opt'] & 0b100) > 0){
          if(-1===fseek($file,4,SEEK_CUR)){ //skip asian stuff
            $this->gen_err(self::E_SEEK,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $readlen[$i] -= 4;
        }
      } else { //this is continue record
        $comp = fread($file,1); //read string comp flag
        if($comp===false || strlen($comp)!==1){
          $this->gen_err(self::E_EOF,__FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $comp = !ord($comp); //'convert' comp to integer (1 or 0) and invert
        --$readlen[$i]; //substract compression byte
      }

      $str = fread($file, $readlen[$i]); //read string binary data
      if($str===false || strlen($str)!==$readlen[$i]){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      if($comp) $str = implode("\0",str_split($str))."\0"; //'uncompress' string
      $str = mb_convert_encoding($str, $this->target_enc, 'UTF-16LE'); //decode from utf16
      $ret_str .= $str; //append to return value
    }
    if(-1===fseek($file,$pos)){ //restore file position to the value it was in the beginning
      $this->gen_err(self::E_SEEK,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    return $ret_str;
  }


  /* ------------ 5. SPECIFIC PARSERS ------------ */
  /* --------------------------------------------- */

  //Unpacks number from excel RK record
  //$rk is raw RK value
  private function readRK($rk){
    $x100 = ($rk & 0b1) > 0; //if bit 0 is set, number is multiplied by 100
    $int = ($rk & 0b10) > 0; //if bit 1 is set, number is integer (otherwise float)

    if($int){ //parse as integer
      $rk = $rk >> 2; //shifting bits 2-32 to the right gives us correct int value
    } else { //parse as float
      $rk = ($rk & ~1); //unset bit 1 (bit 0 is not set anyway because we are here)
      $enc = pack('V',$rk); //convert to binary
      $enc = "\0\0\0\0".$enc; //append nul-bytes to the beginning
      $rk = unpack('da',$enc)['a']; //unpack as 64-bit float
    }
    if($x100){
      $rk /= 100; //if x100, value is stored multiplied, so we must divide it by 100
    }
    return $rk;
  }

  //same as above, but returns whole floats as ints
  private function readRK_fl($rk){
    $x100 = ($rk & 0x1) > 0;
    $int = ($rk & 0x2) > 0;

    if($int){
      $rk = $rk >> 2;
    } else { //parse as float
      $rk = ($rk & ~1); //unset bit 1 (bit 0 is not set since we are here)
      $enc = pack('V',$rk); //convert to binary
      $enc = "\0\0\0\0".$enc; //append nul-bytes to the beginning
      $rk = unpack('da',$enc)['a']; //unpack as 64-bit float
      if($this->float_to_int){ //if float_to_int is enabled by user
        $int_val = (int) $rk; //create closest integer
        if($rk==$int_val) $rk = $int_val; //if int value is equal to float value, use int value
      }
    }
    if($x100){
      $rk /= 100; //if x100, value is stored multiplied, so we must divide it by 100
    }
    return $rk;
  }

  //Parse excel error code.
  //$err_code is a value read from excel cell
  private function parse_xl_err($err_code){
    if($this->err_val==='#DEFAULT!'){ //fill with default excel values
      if(isset($this->err_vals[$err_code])){ //find error code in err_vals array
        return $this->err_vals[$err_code]; //fill cell
      } else {
        return '#ERROR!'; //fill with string #ERROR!
      }
    } else {
      return $this->err_val; //use user-set error value
    }
  }

  //Parse boolean
  private function parse_bool($bool){
    return ($bool ? $this->true_val : $this->false_val);
  }


  /* -- 6. MAP BUILDERS AND BIG DATA GENERATORS -- */
  /* --------------------------------------------- */

  //builds array: array[sst_string_number] = file_offset
  // This is relevant for all sheets.
  //Returns true on success, false on error.
  private function build_SST_map(){
    if($this->error){ //check for previous errors
      $this->gen_err(self::E_PREV, __FUNCTION__);
      return false;
    }
    if($this->BIFF_VER===5){ //this function only works for BIFF8
      $this->gen_err(self::E_SST5,__FUNCTION__);
      return false;
    }
    if(!$this->row_by_row){ //mode must be set to row-by-row
      $this->gen_err(self::E_ROWBYROW, __FUNCTION__);
      return false;
    }
    if(!$this->SST_pos){ //SST File offset must be known at this point
      $this->gen_err(self::E_SST_NOPOS,__FUNCTION__);
      return false;
    }

    $this->SST_map = array(); //clear storage for SST map
    $this->SST_lengths = array(); //clear SST lengths helper array

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream; //$file variable is here for convenience

    //set cursor to where we set SST_pos in get_data()
    if(-1 === fseek($file, $this->SST_pos)){
      $this->gen_err(self::E_SEEK, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $header = fread($file,4); //read record headers
    if(false === $header || strlen($header)!==4){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    $header = unpack('vID/vL',$header); //unpack record headers
    if($header['ID']!==self::SST){ //record must be SST
      $this->gen_err(self::E_SST_EXPECT,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    $bytes_left = $header['L']; //how many bytes in this record

    $sst_init = fread($file,8); //read SST info
    if(false === $sst_init|| strlen($sst_init)!==8){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    $bytes_left -= 8; //every time we read smth from record, we do this to keep track of remaining bytes

    $CNT = unpack('x4/Vu',$sst_init)['u']; //unique strings count

    for($i=0; $i<$CNT; $i++){
      //add position of current string to SST_map
      if(false === ($this->SST_map[$i][0] = ftell($file))){
        $this->gen_err(self::E_TELL, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $this->SST_lengths[$i][0] = 0; //initialize length of first part of string
      $strhdr = fread($file,3); //read string headers
      if(false === $strhdr || strlen($strhdr)!==3){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $bytes_left -= 3;
      $this->SST_lengths[$i][0] += 3; //update string part length
      $strhdr = unpack('vL/copt', $strhdr); //unpack string headers

      $comp = ($strhdr['opt'] & 0b1) === 0; //utf-16 string compression
      $asia = ($strhdr['opt'] & 0b100) > 0; //asian transcription
      $rich = ($strhdr['opt'] & 0b1000) > 0; //rich formatting

      $richlen = 0; //byte length of rich section
      $asialen = 0; //bytes length of asian section

      if($rich){
        $richlen = fread($file,2);
        if(false === $richlen || strlen($richlen)!==2){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= 2;
        $this->SST_lengths[$i][0] += 2;
        $richlen = unpack('vL',$richlen)['L']*4; //each reach section occupies 4 bytes
      }
      if($asia){
        $asialen = fread($file,4);
        if(false === $asialen || strlen($asialen)!==4){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= 4;
        $this->SST_lengths[$i][0] += 4;
        $asialen = unpack('VL',$asialen)['L'];
      }

      $remain = ($comp ? 1 : 2)*$strhdr['L']; //remaining string bytes to parse

      //if we need to read more bytes than there are bytes in this record,
      //then CONTINUE record will occur
      $j = 0; //counter for continues
      while($remain>$bytes_left){
         //skip string
        if(-1 === fseek($file,$bytes_left,SEEK_CUR)){
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }

        $this->SST_lengths[$i][$j] += $bytes_left; //string[0] length
        ++$j;
        $remain -= $bytes_left;
        //at this point current value of $bytes_left is irrelevant and will be updated below

        //update bytes_left: read CONTINUE headers (4 bytes) and string compression flag (1 byte),
        //then unpack record byte size, then substract compression byte from record byte size
        $cont_size = fread($file,5);
        if(false === $cont_size || strlen($cont_size)!==5){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $cont_size = unpack('x2/vL', $cont_size)['L'];

         //add offset of string in CONTINUE record
        if(false === ($this->SST_map[$i][$j] = ftell($file))){
          $this->gen_err(self::E_TELL, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $this->SST_map[$i][$j] -= 1; //pointer was set after comp byte, correct it here
        $this->SST_lengths[$i][$j] = 1; //string[1] length is now 1 (comp flag)
        $bytes_left =  $cont_size - 1; //use CONTINUE headers (we need Length-1 (comp byte))
      }

      //if we have bytes to read
      if($remain>0){
        $this->SST_lengths[$i][$j] += $remain;

         //skip string part
        if(-1 === fseek($file,$remain,SEEK_CUR)){
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= $remain;
      }

      //if there's rich section, then skip it (incl CONTINUE recs)
      //option flag is not repeated if we are CONTINUEing into rich section
      if($rich){
        $remain = $richlen;
        while($remain>$bytes_left){
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $remain -= $bytes_left;
          $bytes_left = fread($file,4);
          if(false === $bytes_left || strlen($bytes_left)!==4){
            $this->gen_err(self::E_EOF, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = unpack('vID/vL', $bytes_left)['L'];
        }

        if($remain>0){
          if(-1 === fseek($file,$remain,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left -= $remain;
        }
      }

      //if there's asian section, then skip it (incl CONTINUE recs)
      //option flag is not repeated if we are CONTINUEing into asian section
      if($asia){
        $remain = $asialen;
        while($remain>$bytes_left){
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $remain -= $bytes_left;

          $bytes_left = fread($file,4);
          if(false === $bytes_left || strlen($bytes_left)!==4){
            $this->gen_err(self::E_EOF, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = unpack('vID/vL', $bytes_left)['L'];
        }
        if($remain>0){
          if(-1 === fseek($file,$remain,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left -= $remain;
        }
      }

      //if there are more strings to read, but not enough bytes for str header + 1 char,
      //then there's CONTINUE record after SST record to which we should skip
      if($CNT>0 && $bytes_left<4){ //4 = 2(str length) + 1(opt flag) + 1(first char)
        if($bytes_left){
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){ //skip till the end of record
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
        }
        $bytes_left = fread($file,4);
        if(false === $bytes_left || strlen($bytes_left)!==4){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left = unpack('vID/vL', $bytes_left)['L']; //read CONTINUE headers and update bytes_left
      }
    }
    return true; //exit function
  }

  //builds array: array[row_number] = file offset of first cell
  // This is relevant only for active sheet.
  //returns true on success, false on error
  private function build_rows_map(){
    if($this->error){ //check for previous errors
      $this->gen_err(self::E_PREV, __FUNCTION__);
      return false;
    }
    if(!$this->row_by_row){ //mode must be set to row-by-row
      $this->gen_err(self::E_ROWBYROW, __FUNCTION__);
      return false;
    }
    $this->rows_map = array(); //initialize container

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream; //$file is here for convenience

    //set cursor to where sheet points to in its properties (see $this->sheets)
    if(-1 === fseek($file,$this->active_sheet['cells_offset'])){
      $this->gen_err(self::E_SEEK, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $reclist = array(self::LABELSST, self::NUMBER, self::RK, self::MULRK, self::RSTRING, self::LABEL, self::BOOLERR, self::FORMULA);
    $stoprec = array(self::NOTE, self::WINDOW1, self::WINDOW2);

    $last_row = -1;

    while($rec = $this->find_recs_before_rec($file, $reclist, $stoprec, false)){
      $row_num_bin = fread($file, 2); //read cell row number (2 bytes)
      if(false === $row_num_bin || strlen($row_num_bin)!==2){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $row_num = unpack('vr',$row_num_bin)['r']; //unpack cell row number
      
      //if row number of this cell is the same as in previous iteration,
      //then skip this record and continue to the next iteration of while()
      if($row_num===$last_row){
        if(-1 === fseek($file,$rec['L']-2,SEEK_CUR)){ //seek to the end of record
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        continue;
      }

      //save position of the beginning of record
      if(false === ($this->rows_map[$row_num] = ftell($file))){
        $this->gen_err(self::E_TELL, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      
      $this->rows_map[$row_num] -= 6; //correct position: -2 (row num) -4 (rec headers)
      
      //seek to the end of record
      if(-1 === fseek($file,$rec['L']-2,SEEK_CUR)){
        $this->gen_err(self::E_SEEK, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $last_row = $row_num; //set last read row number to current row number
    }
    return true; //exit function
  }

  //builds array: array[sst_string_number] = string value (decoded)
  // This is relevant for all sheets.
  //True on success, false on error
  private function read_SST(){
    if($this->error){ //check for previous errors
      $this->gen_err(self::E_PREV, __FUNCTION__);
      return false;
    }
    if($this->row_by_row){ //mode must be set to array mode
      $this->gen_err(self::E_ARRAYMODE, __FUNCTION__);
      return false;
    }
    if(!$this->SST_pos){ //SST file offset must be known
      $this->gen_err(self::E_SST_NOPOS,__FUNCTION__);
      return false;
    }
    if($this->BIFF_VER===5){ //Function should never be called for BIFF5
      $this->gen_err(self::E_SST5,__FUNCTION__);
      return false;
    }
    $this->SST = array(); //initialize SST storage

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream; //'$file' is shorter

    //set cursor to where we set SST position in get_data()
    if(-1 === fseek($file,$this->SST_pos)){
      $this->gen_err(self::E_SEEK, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $header = fread($file,4); //read SST record header
    if(false === $header || strlen($header)!==4){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    $header = unpack('vID/vL',$header); //unpack SST record header
    $bytes_left = $header['L']; //how much bytes to read. Keep track of it when reading!

    $sst_init = fread($file,8); //SST strings count info
    if(false === $sst_init || strlen($sst_init)!==8){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    $bytes_left -= 8; //keeping track...


    $CNT = unpack('x4/Vu',$sst_init)['u']; //unpack SST strings count

    while($CNT>0){ //--$CNT below will be loop breaker
      $strhdr = fread($file,3); //read string length and option byte
      if(false === $strhdr || strlen($strhdr)!==3){
        $this->gen_err(self::E_EOF, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $bytes_left -= 3; //keeping track...
      $strhdr = unpack('vL/copt', $strhdr); //unpack string length and opts

      $comp = ($strhdr['opt'] & 0b1) === 0;
      $asia = ($strhdr['opt'] & 0b100) > 0;
      $rich = ($strhdr['opt'] & 0b1000) > 0;

      $richlen = 0;
      $asialen = 0;

      if($rich){
        $richlen = fread($file,2); //read rich headers
        if(false === $richlen || strlen($richlen)!==2){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= 2;
        $richlen = unpack('vL',$richlen)['L']*4; //each rich section consumes 4 bytes
      }
      if($asia){
        $asialen = fread($file,4); //read asian headers
        if(false === $asialen || strlen($asialen)!==4){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= 4;
        $asialen = unpack('VL',$asialen)['L'];
      }

      $strfull = ''; //placeholder for string
      $remain = ($comp ? 1 : 2)*$strhdr['L']; //remaining string bytes to parse, keep track!

      while($remain>$bytes_left){ //handle this record and CONTINUE records
        $str = fread($file,$bytes_left); //read string data
        if(false === $str || strlen($str)!==$bytes_left){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        if($comp) $str = implode("\0",str_split($str))."\0"; //'uncompress' string
        $strfull .= $str; //append uncompressed part to string
        $remain -= $bytes_left; //keeping track again...
        //$bytes_left -= $bytes_left; // no bytes left in this record

        $cont_hdr = fread($file,5); //read CONTINUE headers + 1 byte (compression)
        if(false === $cont_hdr || strlen($cont_hdr)!==5){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $cont_hdr = unpack('vID/vL/ccmp', $cont_hdr); //unpack CONTINUE headers
        if($cont_hdr['ID']!==self::CONT){ //only CONTINUE record allowed
          $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left = $cont_hdr['L']-1; //substract compression byte

        $comp = ($cont_hdr['cmp'] === 0); //set compression for next section
      }

      if($remain>0){ //if we still have bytes to read
        $str = fread($file,$remain); //read remaining bytes
        if(false === $str || strlen($str)!==$remain){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left -= $remain;
        if($comp) $strfull .= implode("\0",str_split($str))."\0";
        else $strfull .= $str;
      }

      //append decoded string to SST storage
      $this->SST[] = mb_convert_encoding($strfull, $this->target_enc, 'UTF-16LE');

      if($rich){ //skip rich data
        $remain = $richlen;
        while($remain>$bytes_left){ //handle this rec and CONTINUE records
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){ //skip to the end of record data
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $remain -= $bytes_left; //keep track...
          $bytes_left = fread($file,4); //read next record (CONTINUE) headers (reuse $bytes_left)
          if(false === $bytes_left || strlen($bytes_left)!==4){
            $this->gen_err(self::E_EOF, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = unpack('vID/vL', $bytes_left); //unpack CONTINUE record headers
          if($bytes_left['ID']!==self::CONT){ //only CONTINUE records allowed
            $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = $bytes_left['L']; //$bytes_left just reused all over again
        }

        if($remain>0){ //if there's still rich data to skip
          if(-1 === fseek($file,$remain,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left -= $remain;
        }
      }

      if($asia){ //everything is the same as with rich data skipping
        $remain = $asialen;
        while($remain>$bytes_left){
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $remain -= $bytes_left;

          $bytes_left = fread($file,4);
          if(false === $bytes_left || strlen($bytes_left)!==4){
            $this->gen_err(self::E_EOF, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = unpack('vID/vL', $bytes_left);
          if($bytes_left['ID']!==self::CONT){ //only CONTINUE records allowed
            $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left = $bytes_left['L']; //$bytes_left just reused all over again
        }
        if($remain>0){
          if(-1 === fseek($file,$remain,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $bytes_left -= $remain;
        }
      }

      --$CNT; //at this point a string is completely parsed, so decrement counter

      // If there are still strings to read, but less than 4bytes left in this record, then
      // skip data, because headers and 1st character must always occur in the same record.
      // Theoretically, this should never happen in a proper excel file.
      if($CNT>0 && $bytes_left<4){
        if($bytes_left){
          if(-1 === fseek($file,$bytes_left,SEEK_CUR)){
            $this->gen_err(self::E_SEEK, __FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
        }

        $bytes_left = fread($file,4); //read next record headers
        if(false === $bytes_left || strlen($bytes_left)!==4){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left = unpack('vID/vL', $bytes_left); //unpack headers
        if($bytes_left['ID']!==self::CONT){ //only CONTINUE records allowed
          $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bytes_left = $bytes_left['L']; //$bytes_left just reused all over again
      }
    }
    return true; //exit function
  }


  /* ---------- 7. SET-UP AFTER LOADING ---------- */
  /* --------------------------------------------- */

  public function switch_to_row(){ //switch to 'get excel lines one by one' mode
    $this->free(false); //free everything except file pointer
    $this->row_by_row = true;
    $this->select_sheet();
  }

  public function switch_to_array(){ //'get all data at once' mode
    $this->free(false); //free everything except file pointer
    $this->row_by_row = false;
    $this->select_sheet();
  }

  // Sets first row, last row, first col and last col of active sheet.
  // Also will reset active row (next row to read)
  // True on OK, false on error. null: don't set new value. -1: reset to default.
  public function set_margins($first_row = null, $last_row = null, $first_col = null, $last_col = null){
    if(!$this->active_sheet){ //if somehow active sheet is not set
      $this->gen_err(self::E_MARG_NOSHEET,__FUNCTION__);
      return false;
    }
    if($first_row!==null&&$first_row!==-1){
      $first_row = (int) $first_row;
      if($first_row < $this->active_sheet['first_row']){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      if($first_row > $this->last_row){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      $this->first_row = $first_row;
    }

    if($last_row!==null&&$last_row!==-1){
      $last_row = (int) $last_row;
      if($last_row > $this->active_sheet['last_row']){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      if($last_row < $this->first_row){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      $this->last_row = $last_row;
    }

    if($first_col!==null&&$first_col!==-1){
      $first_col = (int) $first_col;
      if($first_col < $this->active_sheet['first_col']){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      if($first_col > $this->last_col){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      $this->first_col = $first_col;
    }

    if($last_col!==null&&$last_col!==-1){
      $last_col = (int) $last_col;
      if($last_col > $this->active_sheet['last_col']){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      if($last_col < $this->first_col){
        $this->gen_err(self::E_MARG_OUTOFBOUNDS, __FUNCTION__);
        return false;
      }
      $this->last_col = $last_col;
    }

    if($first_row===-1){
      $this->first_row = $this->active_sheet['first_row'];
    }
    if($last_row===-1){
      $this->last_row = $this->active_sheet['last_row'];
    }
    if($first_col===-1){
      $this->first_col = $this->active_sheet['first_col'];
    }
    if($last_col===-1){
      $this->last_col = $this->active_sheet['last_col'];
    }

    $this->set_active_row($this->first_row); //reset active row to new first row
    return true;
  }

  //Selects active sheet to read data from
  //$sheet: sheet name or number. -1 means first valid sheet. True=OK, False=Error
  public function select_sheet($sheet = -1){
    //$this->free(false); //free everything except file pointer

    if(!$this->valid_sheets){
      $this->gen_err(self::E_SEL_NOSHEETS,__FUNCTION__);
      return false;
    }

    if($sheet === -1){ //set active sheet to first valid sheet
      $this->active_sheet = reset($this->valid_sheets);
    }

    //handle $sheet as sheet index in $this->valid_sheets
    if(gettype($sheet)==='integer' && $sheet>-1){
      if(array_key_exists($sheet, $this->valid_sheets)){
        $this->active_sheet = $this->valid_sheets[$sheet];
      } else {
        $this->gen_err(self::E_SEL_WRONGINDEX,__FUNCTION__);
        return false;
      }
    }

    //handle $sheet as sheet name
    if(gettype($sheet)==='string'){
      $err = true; //helper
      foreach($this->valid_sheets as $index => $valid_sheet){
        if($sheet===$valid_sheet['name']){
          $this->active_sheet = $this->valid_sheets[$index];
          $err = false;
          break;
        }
      }
      if($err){
        $this->gen_err(self::E_SEL_WRONGNAME,__FUNCTION__);
        return false;
      } unset($err);
    }

    if($this->row_by_row){
      $this->set_margins(-1,-1,-1,-1); //reset margins
      $this->free_rows_map(); //since rows map only relevant per sheet, erase it
    }
    return true;
  }


  /* ---------- 8. DATA READ FUNCTIONS ----------- */
  /* --------------------------------------------- */

  //read row ($this->last_parsed_row + 1) and increment $this->last_parsed_row
  //returns array of cells, false if error, null if out of bounds reached
  public function read_next_row(){
    if($this->error){ //check for previous error
      $this->gen_err(self::E_PREV,__FUNCTION__);
      return false;
    }

    if(!$this->active_sheet){ //active sheet must exist!
      $this->gen_err(self::E_READ_NOSHEET,__FUNCTION__);
      return false;
    }

    if($this->BIFF_VER===8 && !$this->SST_map){ //SST map must exist for BIFF8
      if(!$this->build_SST_map() || !$this->SST_map){ //try to build SST map
        $this->gen_err(self::E_PREV,__FUNCTION__);
        return false;
      }
    }

    if(!$this->rows_map){ //rows map must exist
      if(!$this->build_rows_map() || !$this->rows_map){ //try to build rows map
        $this->gen_err(self::E_PREV,__FUNCTION__);
        return false;
      }
    }

    $r = ++$this->last_parsed_row; //$r = row number that is in use now

    //check if current row is in bounds
    if($r < $this->first_row || $r > $this->last_row){
      --$this->last_parsed_row; //set last parsed row to correct value and exit
      return null;
    }

    //check if current row is empty
    if(!array_key_exists($r,$this->rows_map)){
      if($this->empty_rows){ //if empty_rows==true, fill empty rows
        $out = array(); //row storage
        if($this->empty_cols){ //if empty_cols==true, fill empty cols
          for($i = $this->first_col; $i < $this->last_col+1; $i++){ //set bounds
            $out[$i] = $this->empty_val; //fill with empty value
          }
          return $out; //return filled empty row
        } else {
          return -1; //-1 will be return as an indicator of empty row
        }
      } else { //if empty_rows===false, skip empty rows
        while(true){
          $r = ++$this->last_parsed_row; //get next row number
          if(array_key_exists($r,$this->rows_map)){ //check if it's not empty
            break; //if not empty - break infinite loop and continue down the function
          }
          if($r > $this->last_row){ //check if out of bounds
            --$this->last_parsed_row; //set last parsed row to correct value and exit
            return null;
          }
        }
      }
    }

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream; //'$file' is shorter

    //set cursor to where current row cells records start
    if(-1 === fseek($file,$this->rows_map[$r])){
      $this->gen_err(self::E_SEEK,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $row = array(); //row storage

    while(true){

      $rechdr = fread($file,6); //read record header + next two bytes (should be row number)
      if(false===$rechdr || strlen($rechdr)!==6){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      $rechdr = unpack('vID/vL/vr',$rechdr); //unpack record ID, record length and row number

      //if row number in our unpacked record is not $r, then there are no more cells
      //that correspond to out row. Rest of row is empty. Handle empty cells:
      if($rechdr['r']!==$r){
        if($this->empty_cols){ //if empty_cols==true, fill empty cells
          for($i=$this->first_col; $i<$this->last_col+1; $i++){ //also set bounds
            //if cell $i is not filled previously, fill it now with empty value
            if(!array_key_exists($i,$row)) $row[$i] = $this->empty_val;
          }
        }
        ksort($row); //sort row cells by key, so they appear in series
        return $row; //return row
      }

      $rechdr['L'] -= 2; //we already got first two bytes with header, which were row number

      switch($rechdr['ID']){
        case self::LABELSST: //LABELSST is used in BIFF8 as a reference to string in SST
          $data = fread($file,$rechdr['L']); //read record data
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc/x2/Vs',$data); //[c]olumn of current cell, [s]st index
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2; //check bounds
          $row[$u['c']] = $this->decode_sst_string($file, $u['s']);
          continue 2; //continue while

        case self::NUMBER: //NUMBER is used to store 64-bit float
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc/x2/dflt',$data);
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2;

          if($this->float_to_int){ //if float_to_int==true, convert whole floats to ints
            $int = (int) $u['flt']; //cast float to int
            if($int==$u['flt']) $row[$u['c']] = $int; //check if equal and assign
            else $row[$u['c']] = $u['flt'];
          } else $row[$u['c']] = $u['flt'];

          continue 2;

        case self::RK: //RK is special excel format for numbers, see readRK() descr.
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc/x2/Vrk', $data);
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2;

          if($this->float_to_int) $row[$u['c']] = $this->readRK_fl($u['rk']);
          else $row[$u['c']] = $this->readRK($u['rk']);
          continue 2;

        case self::MULRK: //same as RK but specifies cell range instead of single cell
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $rklen = $rechdr['L'] - 4; //rk numbers length is data length -2(firstcol) - 2(lastcol)
          $u = unpack("vfc/x$rklen/vlc", $data); //[fc]:first column, [lc]:last column
          $rks = substr($data,2,$rklen); //get rk numbers binary data

          $cols = $u['lc'] - $u['fc'] + 1; //number of affected columns

          //read RK entries
          for($i=0; $i<$cols; $i++){
            //skip less-than-first cells
            if(($u['fc']+$i) < $this->first_col) continue;
            //break on more-than-last cell
            if(($u['fc']+$i) > $this->last_col) break;

            // Parse RK data with offset of $i*6 + 2 bytes:
            // number * size of RK (6 bytes) + XF reference (2 bytes)
            if($this->float_to_int){
              $row[$u['fc']+$i] = $this->readRK_fl(unpack("x".($i*6+2)."/Vrk", $rks)['rk']);
            } else {
              $row[$u['fc']+$i] = $this->readRK(unpack("x".($i*6+2)."/Vrk", $rks)['rk']);
            }
          }
          continue 2;

        case self::RSTRING: //this is used for storing rich-formatted string in BIFF5
        case self::LABEL: //this is used for storing simple string in BIFF5
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc',$data);
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2;
          $data = substr($data,4); //cut column number (2bytes) and XF reference (2bytes)
          $row[$u['c']] = $this->parse_unf_str($data,16); //no error handling is needed

          //this record theoretically can be extended with CONTINUE record(s), so handle it
          if(false === ($data = $this->parse_str_cont_row($file))) return false; //error handled in func.
          $row[$u['c']] .= $data;
          continue 2;

        case self::BOOLERR: //is used to store error codes or boolean values
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc/x2/cval/ctype',$data); //unpack [c]olumn, skip 2 bytes, [val]ue, [type]
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2;
          if(!$u['type']){ //this is boolean
            $row[$u['c']] = $this->parse_bool($u['val']); //see parse_bool() for info
          } else { //this is error
            //if fill_err set to true, fill cells with errors, otherwise do nothing
            if($this->fill_err) $row[$u['c']] = $this->parse_xl_err($u['val']);
          }
          continue 2;

        case self::FORMULA: //only last calculated (and saved) result is parsed
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vc',$data); //[c]olumn
          if($u['c']<$this->first_col || $u['c']>$this->last_col) continue 2;
          $data = substr($data,4,8); //read binary result data (skip col number and XF reference)

          //if last two bytes are not 0xFF 0xFF, it's float
          if(unpack('x6/vflt',$data)['flt'] !== 0xFFFF){
            $float = unpack('dflt',$data)['flt'];
            if($this->float_to_int){
              $int = (int) $float;
              if($int==$float) $row[$u['c']] = $int;
              else $row[$u['c']] = $float;
            } else $row[$u['c']] = $float;
            continue 2; //we are done, continue to the next record
          }

          //if last two bytes are 0xFFFF...
          switch(unpack('ctype',$data)['type']){ //read first byte
            case 0: //non-empty STRING for biff8, any STRING for biff5

              $str_hdr = fread($file,4); //read next record, MUST be STRING!
              if(false===$str_hdr || strlen($str_hdr)!==4){ //check if EOF
                $this->gen_err(self::E_EOF,__FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }

              $str_hdr = unpack('vID/vL',$str_hdr); //read record headers
              if($str_hdr['ID']!==self::STR){ //if the record is not STRING, file is invalid
                $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }

              //set cell value to parsed STRING
              $str = fread($file,$str_hdr['L']);
              if(false===$str || strlen($str)!==$str_hdr['L']){
                $this->gen_err(self::E_EOF,__FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }
              $row[$u['c']] = $this->parse_unf_str($str,16);

              //STRING can be extended with CONTINUE record. See RSTRING/LABEL above for comments
              if(false === ($str = $this->parse_str_cont_row($file))) return false;
              $row[$u['c']] .= $str;

              continue 3; //go on to the next record

            case 1: //boolean
              $row[$u['c']] = $this->parse_bool(unpack('x2/cbool',$data)['bool']);
              continue 3;

            case 2: //error
              $errcode = unpack('x2/cerr',$data)['err'];
              if($this->fill_err) $row[$u['c']] = $this->parse_xl_err($errcode);
              continue 3;

            case 3: //blank string for biff8 (only biff8)
              if($this->empty_cols) $row[$u['c']] = $this->empty_val;
              continue 3;
          }
          continue 2; //skip this record and go to next one

        // DBCELL may occur in either BIFF5 or BIFF8 after some rows
        // This record will indicate end of row.
        // Alternatively, there's code above that handles end of rows in a different way.
        case self::DBCELL:
          //check for empty cells and fill them, if needed
          if($this->empty_cols){
            for($i=$this->first_col; $i<$this->last_col+1; $i++){
              if(!array_key_exists($i,$row)) $row[$i] = $this->empty_val;
            }
          }
          ksort($row);
          return $row;

        // if this record is unknown or bad data read, just skip the data and go to next record
        default:
          //skip record data
          if(-1 === fseek($file,$rechdr['L'],SEEK_CUR)){
            $this->gen_err(self::E_SEEK,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          continue 2;       
      }
    }
    //this place should never be reached
    $this->gen_err(self::E_UNREACH, __FUNCTION__, true); //generate warning
    return false;
  }

  //parse all data into $this->cells
  //no float to int, no row/col limit, no empty values
  public function read_everything(){
    if($this->error){ //check for previous errors
      $this->gen_err(self::E_PREV,__FUNCTION__);
      return false;
    }

    if(!$this->active_sheet){ //active sheet must be selected
      $this->gen_err(self::E_READ_NOSHEET,__FUNCTION__);
      return false;
    }

    if($this->BIFF_VER===8 && !$this->SST){ //SST must be present for BIFF8
      if(!$this->read_SST() || !$this->SST){
        $this->gen_err(self::E_PREV,__FUNCTION__);
        return false;
      }
    }

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream;

    //set cursor to where select_sheet want us to
    if(-1 === fseek($file,$this->active_sheet['cells_offset'])){
      $this->gen_err(self::E_SEEK,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $this->cells = array(); //this clears 'cells' in case they where already filled

    while(true){

      $rechdr = fread($file,4); //read record header
      if(false===$rechdr || strlen($rechdr)!==4){ //record header size must be 4
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $rechdr = unpack('vID/vL',$rechdr); //get record ID and byte length

      switch($rechdr['ID']){ //check if current record ID corresponds to those that describe CELLs
        case self::LABELSST: //LABELSST is used in BIFF8 as a reference to string in SST
          $data = fread($file,$rechdr['L']); //read record data
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc/x2/Vs',$data); //unpack [r]ow, [c]olumn of current cell, [s]st index
          $this->cells[$u['r']][$u['c']] = $this->SST[$u['s']]; //cells[r][c] = SST[s]
          continue 2; //continue to the next record (continues 'while')

        case self::NUMBER: //NUMBER is used to store 64-bit float
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc/x2/dflt',$data);
          $this->cells[$u['r']][$u['c']] = $u['flt'];
          continue 2;

        case self::RK: //RK is special excel format for numbers, see readRK() descr.
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc/x2/Vrk', $data);
          $this->cells[$u['r']][$u['c']] = $this->readRK($u['rk']);
          continue 2;

        case self::MULRK: //same as RK but specifies cell range instead of single cell
          $data = fread($file,$rechdr['L']); //read record data
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          //rk numbers data length is data length -2 (row) -2(firstcol) - 2(lastcol)
          $rklen = $rechdr['L'] - 6;

          $u = unpack("vr/vfc/x$rklen/vlc", $data); //[r]ow, [fc]:first column, [lc]:last column
          $rks = substr($data,4,$rklen); //get binary rk numbers data

          $cols = $u['lc'] - $u['fc'] + 1; //count of affected columns

          // Unpack RK, decode RK to int/float, assign int/float to the cell.
          // Offset is $i (col number) * 6 (rk data) + 2 (XF)
          for($i=0; $i<$cols; $i++){
            $this->cells[$u['r']][$u['fc']+$i] = $this->readRK(unpack("x".($i*6+2)."/Vrk", $rks)['rk']);
          }
          continue 2;

        case self::RSTRING: //this is used for storing rich-formatted string in BIFF5
        case self::LABEL: //this is used for storing simple string in BIFF5
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc',$data);
          $data = substr($data,6); //skip first 6 bytes from record data (row number, col number, XF ref)
          $this->cells[$u['r']][$u['c']] = $this->parse_unf_str($data,16);

          //this record theoretically can be extended with CONTINUE record(s), so handle it
          if(!$this->parse_str_cont($file, $u)){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          continue 2;

        case self::BOOLERR: //is used to store error codes or boolean values
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc/x2/cval/ctype',$data); //[r]ow, [c]olumn, skip 2 bytes, [val]ue, [type]
          if($u['type']){ //this is error value
            //if fill_err set to true, fill cells with errors, otherwise skip
            if($this->fill_err)
              $this->cells[$u['r']][$u['c']] = $this->parse_xl_err($u['val']);
          } else { //this is boolean
            $this->cells[$u['r']][$u['c']] = (bool) $u['val']; //set cell to true or false
          }
          continue 2;

        case self::FORMULA: //read last calculated (and therefore saved in file) result
          $data = fread($file,$rechdr['L']);
          if(false===$data || strlen($data)!==$rechdr['L']){
            $this->gen_err(self::E_EOF,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          $u = unpack('vr/vc',$data); //row and column
          $data = substr($data,6,8); //read binary result data (skip row, col, XF ref)
          if(unpack('x6/vflt',$data)['flt'] !== 0xFFFF){ //if last two bytes are not 0xFF 0xFF
            $this->cells[$u['r']][$u['c']] = unpack('dflt',$data)['flt']; //the whole data is 64-bit float
            continue 2; //we are done, continue to the next record
          }
          //if data is not float, then it's either of the following:
          switch(unpack('ctype',$data)['type']){ //read first byte
            case 0: //non-empty STRING for biff8, any STRING for biff5

              $str_hdr = fread($file,4); //read next record, MUST be STRING!

              if(false===$str_hdr || strlen($str_hdr)!==4){ //as usual, check for EOF
                $this->gen_err(self::E_EOF, __FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }

              $str_hdr = unpack('vID/vL',$str_hdr); //read record headers
              if($str_hdr['ID']!==self::STR){ //if the record is not STRING, file is invalid
                $this->gen_err(self::E_READ_UNEXP_REC, __FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }

              //set cell value to parsed STRING
              $data = fread($file,$str_hdr['L']);
              if(false===$data || strlen($data)!==$str_hdr['L']){
                $this->gen_err(self::E_EOF, __FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }
              $this->cells[$u['r']][$u['c']] = $this->parse_unf_str($data,16);

              //STRING can be extended with CONTINUE record. See RSTRING/LABEL above for comments
              if(!$this->parse_str_cont($file, $u)){
                $this->gen_err(self::E_PREV,__FUNCTION__);
                if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
                return false;
              }

              continue 3; //go on to the next record

            case 1: //boolean
              $this->cells[$u['r']][$u['c']] = (bool) unpack('x2/cbool',$data)['bool'];
              continue 3;

            case 2: //error,
              $errcode = unpack('x2/cerr',$data)['err'];
              if($this->fill_err)
                $this->cells[$u['r']][$u['c']] = $this->parse_xl_err($errcode);
              continue 3;
            
            // case 3: //blank string for biff8 (only biff8)
            // $this->cells[$u['r']][$u['c']] = '';
            // continue 3;
           
          }
          continue 2; //skip this record and go to next one

        // BLANK and MULBLANK records are in specs, but usually not used. Ignore them anyway.
        // case self::BLANK:
          // fseek($file,$rechdr['L'],SEEK_CUR);
          // continue 2;
        // case self::MULBLANK:
          // fseek($file,$rechdr['L'],SEEK_CUR);
          // continue 2;

        // DBCELL record means we finished reading a ROW BLOCK. Ignore..
        // case self::DBCELL:
          // fseek($file,$rechdr['L'],SEEK_CUR); //skip DBCELL record
          // continue 2;

        // either of (NOTE, WINDOW1, WINDOW2) must occur in either BIFF5 or BIFF8 after all cells
        case self::NOTE:
        case self::WINDOW1:
        case self::WINDOW2:
          return true;

        // if this record is unknown or bad data read, just skip the data and go to next record
        default:
          //skip record data
          if(-1 === fseek($file,$rechdr['L'],SEEK_CUR)){
            $this->gen_err(self::E_SEEK,__FUNCTION__);
            if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
            return false;
          }
          continue 2;
      }
    }
    //this place should never be reached
    $this->gen_err(self::E_UNREACH, __FUNCTION__, true); //generate warning
    return false;
  }

  //get all data after file loading: codepage, encryption, sheets info, boundaries, offsets, etc
  public function get_data(){
    if($this->error){ // check for previous error
      $this->gen_err(self::E_PREV,__FUNCTION__);
      return false;
    }

    //check if main stream is ok
    if(!$this->stream || !gettype($this->stream)==='resource' ||
    !get_resource_type($this->stream)==='stream'){
      $this->gen_err(self::E_BADHANDLE, __FUNCTION__);
      return false;
    }

    $file = $this->stream;

    //rewind stream
    if(!rewind($file)){
      $this->gen_err(self::E_SEEK, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $header = fread($file,4); //read first 4 bytes, they must be first record header
    if(false===$header || strlen($header)!==4){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    $header = unpack('vID/vL',$header); //unpack record ID and record data byte length
    
    //First record MUST be BOF!
    if($header['ID']!==self::BOF){
      $this->gen_err(self::E_NOBOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    
    $data = fread($file,$header['L']); //read record data
    if(false===$data || strlen($data)!==$header['L']){
      $this->gen_err(self::E_EOF, __FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    //unpack BOF data and determine substream type
    $data = unpack('vver/vtype', $data);
    if($data['type']===5){ //5 means Workbook Globals substream
    
      // Decode BIFF version
      if($data['ver']===0x600) $this->BIFF_VER = 8;
      else if($data['ver']===0x500) $this->BIFF_VER = 5;
      else {
        $this->gen_err(self::E_HDR_BIFF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
    } else {
      $this->gen_err(self::E_READ_UNEXP_REC,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }
    
    // Re-check file size for BIFF8, which has a different minimum size limit
    if($this->BIFF_VER === 8 && $this->filesize < 194){
      $this->gen_err(self::E_HDR_SIZE8,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }


    // CHECK IF ENCRYPTED AND GET CODEPAGE

    $reclist = array(self::FILEPASS, self::CODEPAGE);
    $stoprec = array(self::WINDOW1);

    while($rec = $this->find_recs_before_rec($file, $reclist, $stoprec, true)){
      if($rec['ID']===self::FILEPASS){ // Encryption record found, file not supported!
        $this->gen_err(self::E_HDR_CRYPT,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      if($rec['ID']===self::CODEPAGE){ //codepage record found
        $cp_bin = fread($file,$rec['L']); //read record data
        if(false===$cp_bin || strlen($cp_bin)!==$rec['L']){
          $this->gen_err(self::E_EOF,__FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $this->CODEPAGE = $this->gen_cp(unpack('vcp',$cp_bin)['cp']); //save codepage
      }
    }

    // SET ENCODING STUFF

    //this only matters for BIFF5
    if($this->BIFF_VER===5 && !$this->CP_set){
      //set transcoding parameters
      $this->set_encodings(true, $this->CODEPAGE, $this->target_enc, false);
      $this->CP_set = true; //don't reset transcoding parameters if get_data() called again
    }

    // GET SHEETs

    $reclist = array(self::SHEET);
    $stoprec = array(self::EOF);

    if($this->BIFF_VER===8){ //for BIFF8 stop before SST
      $stoprec[] = self::SST;
    }

    $this->sheets = array(); //container for sheets info
    $i = 0; //will be used as sheet index
    $ws_found = false; //stands for "worksheet is found"
    while($rec = $this->find_recs_before_rec($file, $reclist, $stoprec, false)){
      $data = fread($file,$rec['L']); //read record
      if(false===$data || strlen($data)!==$rec['L']){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      $unp = unpack('VBOF_offset/chidden/ctype', $data); //unpack BOF offset, hidden attribute and type

      $this->sheets[$i]['error'] = false; //initialize error flag for current sheet
      $this->sheets[$i]['err_msg'] = ''; //initialize error message for current sheet

      // check if BOF offset is within possible limits
      if($unp['BOF_offset']<0 || $unp['BOF_offset']>0x7fffffff){
        $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OFFSET];
        $this->sheets[$i]['error'] = true;
      }

      //GET TYPE, WILL BE OVERWRITTEN LATER
      switch($unp['type']){
        case 0:
          if(!$ws_found) $ws_found = true;
          $type = 'Worksheet'; //'Worksheet' can be also 'Dialog', this is determined later
          break;
        case 1:
          $type = 'Macro';
          break;
        case 2:
          $type = 'Chart';
          break;
        case 6:
          $type = 'VB module'; //Visual Basic module
          break;
        default:
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_TYPE_GLOB];
          $this->sheets[$i]['error'] = true;
      }

      $name = substr($data,6,$rec['L']-6); //6 = 4 (BOF offset) + 1 (hidden) + 1(type)
      $this->sheets[$i]['name'] = $this->parse_unf_str($name, 8); //sheet name
      $this->sheets[$i]['hidden'] = $unp['hidden']; //sheet 'hidden' attribute
      $this->sheets[$i]['type'] = $type; //this will be overwritten below
      $this->sheets[$i]['BOF_offset'] = $unp['BOF_offset']; //pointer to BOF in main stream
      ++$i;
    }

    if(!$this->sheets){ //if no sheets found at all, file is invalid
      $this->gen_err(self::E_HDR_NOSHEETS,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    if(!$ws_found){ //if no worksheets found, there's nothing to parse in file
      $this->gen_err(self::E_HDR_NOWSHEETS,__FUNCTION__);
      if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
      return false;
    }

    //For BIFF8, we should have SST and we should have stopped right before it
    if($this->BIFF_VER===8&&!in_array(self::SST, $stoprec)){
      if(false === ($this->SST_pos = ftell($file))){ //save SST main stream offset
        $this->gen_err(self::E_TELL, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
    }

    //if BIFF_VER is 5 or there's no SST, we reached the end of Globals Substream

    // SEARCH THROUGH ALL WORKSHEETS AND GET SHEET PROPERTIES

    foreach($this->sheets as $i => $sheet){
      if($sheet['type']!=='Worksheet') continue; //skip non worksheets
      if($sheet['error']) continue; //skip sheets with errors
       //seek to BOF of the sheet
      if(-1 === fseek($file,$sheet['BOF_offset'])){
        $this->gen_err(self::E_SEEK, __FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $rec = fread($file, 4); //read BOF headers
      if(false===$rec || strlen($rec)!==4){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      $rec = unpack('vID/vL', $rec); //unpack BOF headers
      if($rec['ID']!==self::BOF){
        $this->gen_err(self::E_READ_UNEXP_REC,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      $data = fread($file, $rec['L']); //read BOF data
      if(false===$data || strlen($data)!==$rec['L']){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }
      $type = unpack('x2/vtype', $data)['type']; //get sheet type

      $sheet = false; //whether a substream is a worksheet or dialog

      //get type and overwrite previously written type
      switch($type){
        case 0x5: //Globals (we shouldn't encounter it!)
          $this->gen_err(self::E_HDR_2GLOBALS, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;

        case 0x6: //VB
          $this->sheets[$i]['type'] = 'VB module';
          break;

        case 0x10: //Sheet or Dialog
          $sheet = true; //additional code below
          break;

        case 0x20: //Chart
          $this->sheets[$i]['type'] = 'Chart';
          break;

        case 0x40: //Macro
          $this->sheets[$i]['type'] = 'Macro';
          break;

        case 0x100:
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_WORKSPACE];
          $this->sheets[$i]['error'] = true;
          continue;

        default:
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_UNKNOWN];
          $this->sheets[$i]['error'] = true;
          continue;
      }

      if(!$sheet) continue; //nothing interesting is this sheet

      $reclist = array(self::SHEETPR); //find SHEETPR record
      $stoprec = array(self::DIMENSION); //before DIMENSION record

      $rec = $this->find_recs_before_rec($file, $reclist, $stoprec, true);
      if($rec){
        $data = fread($file, $rec['L']);
        if(false===$data || strlen($data)!==$rec['L']){
          $this->gen_err(self::E_EOF, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        $bits = unpack('vbits', $data)['bits'];

        if(($bits & 0x10) > 0){
          //this is dialog
          $this->sheets[$i]['type'] = 'Dialog';
          continue; //continue to the next sheet
        }
      } else {
        //no SHEETPR found!! Unable to determine sheet type!
        $this->sheets[$i]['err_msg'] .= $this->E[self::SH_TYPE_SHEETPR];
        $this->sheets[$i]['error'] = true;
        continue;
      }

      //if we reached here, this is regular sheet
      $this->sheets[$i]['type'] = 'Worksheet';

      //find and read DIMENSION
      $reclist = array(self::DIMENSION); //find SHEETPR record
      $stoprec = array(self::ROW, self::EOF); //before DIMENSION record

      $rec = $this->find_recs_before_rec($file, $reclist, $stoprec, true);
      if($rec){
        if(!in_array((self::EOF), $stoprec)){
          //if we stopped at EOF instead of ROW, this sheet is empty
          $this->sheets[$i]['empty'] = true;
          continue;
        } else {
          $this->sheets[$i]['empty'] = false;
        }
      } else {
        $this->sheets[$i]['err_msg'] .= $this->E[self::SH_NODIMM];
        $this->sheets[$i]['error'] = true;
        continue;
      }

      $margins = fread($file, $rec['L']); //read DIMENSION record
      if(false===$margins || strlen($margins)!==$rec['L']){
        $this->gen_err(self::E_EOF,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
      }

      //unpack and process DIMENSION record, error-check margins
      if($this->BIFF_VER===5){
        $margins = unpack('vfirst_row/vlast_row/vfirst_col/vlast_col', $margins);
        if($margins['first_row']<0 || $margins['first_row']>0x3FFF){
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_SR];
          $this->sheets[$i]['error'] = true;
          continue;
        }
        if($margins['last_row']<0 || $margins['last_row']>0x4000){
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_LR];
          $this->sheets[$i]['error'] = true;
          continue;
        }
      } else {
        $margins = unpack('Vfirst_row/Vlast_row/vfirst_col/vlast_col', $margins);
        if($margins['first_row']<0 || $margins['first_row']>0x0000FFFF){
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_SR];
          $this->sheets[$i]['error'] = true;
          continue;
        }
        if($margins['last_row']<0 || $margins['last_row']>0x00010000){
          $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_LR];
          $this->sheets[$i]['error'] = true;
          continue;
        }
      }
      if($margins['first_col']<0 || $margins['first_col']>0x00FF){
        $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_FC];
        $this->sheets[$i]['error'] = true;
        continue;
      }
      if($margins['last_col']<0 || $margins['last_col']>0x100){
        $this->sheets[$i]['err_msg'] .= $this->E[self::SH_OUB_LC];
        $this->sheets[$i]['error'] = true;
        continue;
      }
      if($margins['last_row']===0 || $margins['last_col']===0){
        $this->sheets[$i]['empty'] = true;
        continue;
      }

      //excel actually stores increased value for last row and col, we don't want it
      --$margins['last_row'];
      --$margins['last_col'];

      //append margins to sheets info
      foreach($margins as $key => $value){
        $this->sheets[$i][$key] = $value;
      }

      //find first cell record and remember its offset
      $reclist = array(self::LABELSST, self::NUMBER, self::RK, self::MULRK, self::RSTRING, self::LABEL, self::BOOLERR, self::FORMULA);
      $stoprec = array(self::DBCELL, self::NOTE, self::WINDOW1, self::WINDOW2, self::EOF);

      $rec = $this->find_recs_before_rec($file, $reclist, $stoprec, true);
      if($rec){
        if(-1 === fseek($file,-4,SEEK_CUR)){
          $this->gen_err(self::E_SEEK, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
        if(false === ($this->sheets[$i]['cells_offset'] = ftell($file))){
          $this->gen_err(self::E_TELL, __FUNCTION__);
          if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
          return false;
        }
      } else {
        $this->sheets[$i]['empty'] = true;
      }
    }

    $this->valid_sheets = array(); //placeholder for valid selectable sheets
    foreach($this->sheets as $key => $sheet){
      if($sheet['type']==='Worksheet' && !$sheet['empty'] && !$sheet['error']){
        $sheet['number'] = $key; //for convenience
        $this->valid_sheets[$key] = $sheet;
        continue;
      }
    } unset($key,$sheet);
    if(!$this->valid_sheets){
        $this->gen_err(self::E_HDR_NOVALIDSHEET,__FUNCTION__);
        if(!fclose($file)) $this->gen_err(self::E_CLOSE,__FUNCTION__);
        return false;
    }
    $this->select_sheet(); //set active sheet to first valid sheet and set margins
    return true;
  }


  /* ---------- 9. FREE-ERS, UNSETTERS ----------- */
  /* --------------------------------------------- */

  public function free_stream(){ //free stream (delete temporary files)
    if($this->stream){
      if(gettype($this->stream)==='resource' && get_resource_type($this->stream)==='stream'){
        if(!fclose($this->stream)){
          $this->gen_err(self::E_CLOSE, __FUNCTION__, true);
        }
      }
      $this->stream = null;
    }
  }

  public function free_cells(){ //free result got by get_everything (for array mode)
    $this->cells = array();
  }

  public function free_sst(){ //free SST created for 'array' mode
    $this->SST = array();
  }
  
  public function free_rows_map(){ //free rows map created for 'row-by-row' mode
    $this->rows_map = array();
  }

  public function free_sst_maps(){ //free SST map, SST lengths created for 'row-by-row' mode
    $this->SST_map = array();
    $this->SST_lengths = array();
  }
  
  public function free_maps(){
    $this->free_rows_map();
    $this->free_sst_maps();
  }

  public function free($stream = true){ //free everything
    if($stream) $this->free_stream();
    $this->free_cells();
    $this->free_sst();
    $this->free_maps();
  }


  /* ------ 10. CONSTRUCTOR AND DESTRUCTOR ------- */
  /* --------------------------------------------- */

  function __construct($filename, $debug = false, $mem = null, $debug_MSCFB = false){
    $this->debug = (bool) $debug;

    if(!file_exists($filename)){
      $this->gen_err(self::E_NOFILE,__FUNCTION__);
      return;
    }

    $raw_stream = false; //whether the file is MS Compound File or a raw Workbook stream

    do{
      // Attempt to open file as Compound File
      $cfb = new MSCFB($filename, $debug_MSCFB, $mem);
      if($cfb->error){
        unset($cfb);
        $raw_stream = true;
        break;
      }

      // at this point Compound File is opened

      // Try to get stream index of Workbook stream
      $i = $cfb->get_by_name('Workbook');
      if($i===-1){
        $i = $cfb->get_by_name('Book'); //for BIFF5
        if($i===-1){
          $cfb->free();
          unset($cfb);
          $this->gen_err(self::E_NOWBSTREAM,__FUNCTION__);
          return;
        }
      }

      //at this point we know that Workbook stream exists

      $temp_str = 'php://temp'; //temp file address for fopen()
      if($mem !== null){
        $size = (int) $mem;
        if($size>0) $temp_str .= '/maxmemory:'.$size; //tempfile size adjustment
        if($size===0) $temp_str = 'php://memory'; //always store data in memory
      }

      $temp = null;

      //create temporary stream
      if(false === ($temp = fopen($temp_str, 'w+b'))){
        $this->gen_err(self::E_TEMP, __FUNCTION__);
        return;
      }

      $bytes = $cfb->DE[$i]['sizeL'];
      if(false === $cfb->extract_stream($i, $temp)){
        $this->gen_err(self::E_EXTRACT, __FUNCTION__);
        return;
      }
      $cfb->free();
      unset($cfb);
      $this->stream = $temp;
    } while(false);

    // at this point we either extracted Workbook stream to temporary storage,
    // or an error occured so we try to open file directly as Workbook stream

    if($raw_stream){
      if(!($this->stream = fopen($filename, 'rb'))){
        $this->gen_err(self::E_OPEN, __FUNCTION__);
        return;
      }
    }

    //at this point $this->stream always contains Workbook stream

    $this->filename = $filename;

    $this->filesize = filesize($filename);
    if($this->filesize > 0x7fffffff || $this->filesize < 0){
      $this->gen_err(self::E_HDR_SIZEMAX,__FUNCTION__);
      return;
    }
    if($this->filesize < 166){
      $this->gen_err(self::E_HDR_SIZE5,__FUNCTION__);
      return;
    }

    $this->set_output_encoding(); //set enc to mb_internal_encoding
    $this->get_data(); //get data, set encoding parameters, select first valid sheet
  }

  function __destruct(){ //just free()
    $this->free();
  }
}
?>