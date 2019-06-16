<?php
require_once 'MSXLS.php'; //MSCFB.php is 'required once' inside MSXLS.php
$excel = new MSXLS('example.xls', true, null, true);

if($excel->error) die($excel->err_msg); //Terminate script execution, show error message.
$excel->set_fill_xl_errors(true);
$excel->read_everything(); //Read cells into $excel->cells

// var_dump($excel->cells); //Output all parsed cells from XLS file
// var_dump($excel->get_sheets()); //Show all sheets
// var_dump($excel->get_valid_sheets()); //Show valid sheets
// var_dump($excel->get_active_sheet()); //Show selected sheet

$sheet = $excel->get_active_sheet(); //Get it for margins

$table_style = 'border: 1px solid gray;'; //CSS Style of <table>
$td_style = 'border: 1px solid gray; padding: 5px 10px; text-align: center;'; //CSS of <td>

echo 'This was read via read_everything() method:';
echo '<table style="'.$table_style.'">';
for($row = $sheet['first_row']; $row <= $sheet['last_row']; $row++){
  echo '<tr>';
  
  for($col = $sheet['first_col']; $col <= $sheet['last_col']; $col++){
    echo '<td style="'.$td_style.'">';
    if(isset($excel->cells[$row][$col])) echo($excel->cells[$row][$col]);
    else echo '-'; //echo '-' as empty element
    echo '</td>';
  }
  
  echo '</tr>';
}
echo '</table>';

echo '<br><br>';

// ---------------------------------------------------
// END OF ARRAY MODE EXAMPLE, ROW-BY-ROW EXAMPLE BELOW
// ---------------------------------------------------

$excel->switch_to_row(); //switch to Row-by-row mode

$excel->set_fill_xl_errors(true, 'ERROR'); //Fill excel error cells with 'ERROR' string
$excel->set_margins(-1, $sheet['last_row']-1); //Set last row as next-to-last
$excel->set_active_row($sheet['first_row']+1); //Set first row to read as the 2nd row

$excel->set_empty_value('(empty)'); //set '(empty)' string as empty value
$excel->use_empty_cols(true); //fill empty columns with  '(empty)'
$excel->use_empty_rows(true); //pretend that there are empty cells in empty row

$excel->set_boolean_values('[TRUE]', '[FALSE]'); //instead of PHP true and false values
$excel->set_float_to_int(true); //convert whole floats to ints

echo 'This was read in row-by-row mode.<br>';
echo 'Last row is set to be next to last via set_margins().<br>';
echo 'Reading starts from the second row via set_active_row().<br>';
echo 'First row is manually read after all other rows.';

echo '<table style="'.$table_style.'">';
while($row = $excel->read_next_row()){
  echo '<tr>';
  foreach($row as $cell){
    echo '<td style="'.$td_style.'">';
    echo $cell;
    echo '</td>';
  }
}

$excel->set_active_row($sheet['first_row']); //now we will read first row
$row = $excel->read_next_row();

echo '<tr>';
foreach($row as $cell){
  echo '<td style="'.$td_style.'">';
  echo $cell;
  echo '</td>';
}
echo '</tr>';

echo '</table>';
$excel->free();
unset($excel);
?>