<?php
// Load library
require_once '../vendor/autoload.php';

use odsPhpGenerator\ods;
use odsPhpGenerator\odsTable;
use odsPhpGenerator\odsStyleTableCell;
use odsPhpGenerator\odsTableRow;
use odsPhpGenerator\odsTableCellString;
use odsPhpGenerator\odsStyleTableColumn;
use odsPhpGenerator\odsTableColumn;

$data = [
	['Apple', 12],
	['Bananas', 5],
	['Orange', 7]
];


// Create Ods object
$ods  = new ods();

// First Table
$ods->addTable('Table 1', $data);


// Style
$style = new odsStyleTableCell();
$style->setColor('#0000ff');
$style->setBackgroundColor('#00ff00');

// Second Table
$table = new odsTable('Table2');
$table->setVerticalSplit();

//Header
// create column width 7cm
$styleColumn = new odsStyleTableColumn();
$styleColumn->setColumnWidth("7cm");
$column = new odsTableColumn($styleColumn);
$table->addTableColumn($column);


$row = new odsTableRow();
$row->addCell(new odsTableCellString('Fruit', $style));
$row->addCell(new odsTableCellString('Number', $style));
$table->addRow($row);

// Data
$table->addRows($data);

$ods->addTable($table);


// Download the file
$ods->downloadOdsFile("EasyTable.ods");
