<!--The purpose of this HTML/PHP code is to enable contact data to recognized & found by file:StaffPhonebook.html -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en-gb">
<head>
<title>Staff Phonebook</title>
<meta name="Author" content="emailAddress@gmail.com" />
<meta name="Description" content="Shows all records in staff phonebook" />
<style type="text/css">
body { padding:15px; font-family:sans-serif; }
h1 { background:#9999FF; color:white; padding:1ex; }
</style>
</head>
<body>
<h1>Staff Phonebook</h1>
<p>
Return to the <a href="Phonebook.html">Staff Phonebook Homepage</a>
</p>
<table>
<tr>
<th>First Name</th>
<th>Surname</th>
<th>Telephone</th>
</tr>
<?php

//Seting up the connection
if (!$conn = new COM('ADODB.Connection')) exit('Unable to create an ADODB connection');

//Define database string, specify database drive 
$strConn = 'DRIVER={Microsoft Access Driver (*.mdb)} DBQ=E \WWW.\GRE\33085962\DB1\Phonebook.mdb';

//Open connection to the database
$conn->open($strConn);

//SQL statment to retieve data 
$strSQL = 'SELECT id, firstname, surname, phone FROM staff ORDER BY id';

//Execute SQL statment and return records
$rs = $conn->execute($strSQL);


//While loop used to carry on looping while there are records, exits at end of the file
while (!$rs->EOF) {
  echo '<tr><td>' . $rs->Fields(0)->value . '</td>';
  echo '<td>' . $rs->Fields(1)->value . '</td>';
  echo '<td>' . $rs->Fields(2)->value . '</td>';
  echo '<td>' . $rs->Fields(3)->value . '</td></tr>';
  
  //Moving on to next record 
  $rs->MoveNext(); 
}
?>

</table>
</body>
</html>
