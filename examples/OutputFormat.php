<?php

// All file is writen in UTF-8

use odsPhpGenerator\ods;
use odsPhpGenerator\odsTable;
use odsPhpGenerator\odsTableRow;
use odsPhpGenerator\odsTableCellString;

// Load library
require_once '../vendor/autoload.php';

// Create Ods object
$ods  = new ods();
$table = new odsTable('table 1');
$row   = new odsTableRow();
$row->addCell( new odsTableCellString("Hello") );
$row->addCell( new odsTableCellString("World") );
$table->addRow($row);
$ods->addTable($table);

// Download the file
// $ods->downloadOdsFile("HelloWorld.ods");
// or $ods->downloadOdsFile("HelloWorld.ods", ods::OUTPUT_ODS);

// If open/libre Office installed on the computer
// You can export on other format

// PDF
//$ods->downloadOdsFile("HelloWorld.pdf", ods::OUTPUT_PDF);

// XLS ( excel )
//$ods->downloadOdsFile("HelloWorld.xsl", ods::OUTPUT_XLS);

// XLSX ( excel )
$ods->downloadOdsFile("HelloWorld.xlsx", ods::OUTPUT_XLSX);
