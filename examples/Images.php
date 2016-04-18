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

// Create table named 'table 1'
$table = new odsTable('table 1');

$image = new \odsPhpGenerator\odsDrawImage('imgs/img.png','1cm','1cm','2cm','2cm');
$table->addDraw($image);

$image = new \odsPhpGenerator\odsDrawImage('imgs/img.jpg','4cm','1cm','2cm','2cm');
$table->addDraw($image);

$image = new \odsPhpGenerator\odsDrawImage('imgs/img.gif','7cm','1cm','2cm','2cm');
$table->addDraw($image);

$image = new \odsPhpGenerator\odsDrawImage('imgs/img.svg','10cm','1cm','2cm','2cm');
$table->addDraw($image);

$image = new \odsPhpGenerator\odsDrawImage('imgs/img.tiff','13cm','1cm','2cm','2cm');
$table->addDraw($image);

$line = new \odsPhpGenerator\odsDrawLine('1cm','4cm','3cm','6cm');
$table->addDraw($line);

$rect = new \odsPhpGenerator\odsDrawRect('4cm','4cm','2cm','2cm');
$table->addDraw($rect);

// Attach talble to ods
$ods->addTable($table);

// Download the file
$ods->downloadOdsFile("Images.ods");
