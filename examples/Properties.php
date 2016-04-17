<?php

// Load library
require_once '../vendor/autoload.php';

use odsPhpGenerator\ods;
use odsPhpGenerator\odsTable;
use odsPhpGenerator\odsTableRow;
use odsPhpGenerator\odsTableCellString;

// Create Ods object
$ods  = new ods();

// Set properties
$ods->setTitle("My title is cool");
$ods->setSubject("My subject is cool too");
$ods->setKeyword("ods, internet");
$ods->setDescription("I love odsPhpGenerator !!\n\nOk, i'm a programmer, sorry");


// Set page Format
$ods->setPageDimention('210mm','297mm');
// or use
$ods->setPageWidth('21cm'); $ods->setPageHeight('297mm');
// or use
$ods->setPageFormat(ods::PAGE_A3, ods::PAGE_LANDSCAPE);

// Set margins
$ods->setPagesMargins('20mm', '20mm', '20mm', '20mm');
// or use ( null: default )
$ods->setPageMarginTop(null);
$ods->setPageMarginBottom(null);
$ods->setPageMarginLeft(null);
$ods->setPageMarginRight(null);

// Center table to print
$ods->setPageTableCentering(ods::PAGE_TABLE_CENTERING_BOTH);

// Remove header et footer to print
$ods->setPageHeaderDisplay(false);
$ods->setPageFooterDisplay(false);


// Create table named 'table 1'
$table = new odsTable('table 1');
$row   = new odsTableRow();
$row->addCell( new odsTableCellString("Hello") );
$row->addCell( new odsTableCellString("World") );
$table->addRow($row);
$ods->addTable($table);



// Download the file
$ods->downloadOdsFile("Properties.ods"); 

?>
