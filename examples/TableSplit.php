<?php

// Load library
require_once '../vendor/autoload.php';

use odsPhpGenerator\ods;
use odsPhpGenerator\odsTable;


// Create Ods object
$ods  = new ods();

//Set vertical split
$table = new odsTable('V Split');
$table->setVerticalSplit(1);
$ods->addTable($table);

//Set Horizontal split
$table = new odsTable('H Split');
$table->setHorizontalSplit(1);
$ods->addTable($table);

// Both split
$table = new odsTable('B Split');
$table->setVerticalSplit(1);
$table->setHorizontalSplit(1);
$ods->addTable($table);

// Download the file
$ods->downloadOdsFile("TableSplit.ods");  

?>
