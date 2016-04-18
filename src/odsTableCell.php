<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

namespace  odsPhpGenerator;

abstract class odsTableCell {
	protected $styleName;
	protected $numberColumnsSpanned;
	protected $numberRowsSpanned;
	protected $formula;
	protected $numberColumnsRepeated;
	
	//abstract protected function __construct();
	
	protected function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = $dom->createElement('table:table-cell');
		if( $this->styleName ) {
			$ods->addTmpStyles($this->styleName);
			$table_table_cell->setAttribute("table:style-name", $this->styleName->getName());
		}
		$this->cellOpts($table_table_cell);
		return $table_table_cell;
	}

	protected function cellOpts( $table_table_cell ) {
		if($this->numberColumnsSpanned)
			$table_table_cell->setAttribute("table:number-columns-spanned", $this->numberColumnsSpanned);
		if($this->numberRowsSpanned)
			$table_table_cell->setAttribute("table:number-rows-spanned", $this->numberRowsSpanned);
		if($this->formula)
			$table_table_cell->setAttribute("table:formula", "of:=".$this->formula);
		if($this->numberColumnsRepeated)
			$table_table_cell->setAttribute("table:number-columns-repeated", $this->numberColumnsRepeated);
	}

	public function setNumberColumnsSpanned($numberColumnsSpanned) {
		$this->numberColumnsSpanned = $numberColumnsSpanned;
	}

	public function getNumberColumnsSpanned() {
		if(!$this->numberColumnsSpanned) return 1;
		return $this->numberColumnsSpanned;
	}

	public function setNumberRowsSpanned($numberRowsSpanned) {
		$this->numberRowsSpanned = $numberRowsSpanned;
	}

	public function setFormula($formula) {
		$this->formula = $formula;
	}

	public function setNumberColumnsRepeated($numberColumnsRepeated) {
		$this->numberColumnsRepeated = $numberColumnsRepeated;
	}

}

class odsTableCellEmpty extends odsTableCell {

	public function __construct(odsStyleTableCell $odsStyleTableCell = null) {
		$this->styleName = $odsStyleTableCell;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		return odsTableCell::getContent($ods,$dom);
	}
}

class odsCoveredTableCell extends odsTableCell {
	public function __construct() {}
	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = $dom->createElement('table:covered-table-cell');
		$this->cellOpts($table_table_cell);
		return $table_table_cell;
	}
}

class odsTableCellStringHttp extends odsTableCell {
	public $value;
	public $styleName;

	public function __construct($value,odsStyleTableCell $odsStyleTableCell = null) {
		$this->value = $value;
		$this->styleName = $odsStyleTableCell;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = odsTableCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");

			// text:p
			$text_p = $dom->createElement('text:p',$this->value);
				$table_table_cell->appendChild($text_p);
		return $table_table_cell;
	}
}

class odsTableCellString extends odsTableCellStringHttp {
	public function __construct($value,odsStyleTableCell $odsStyleTableCell = null) {
		$this->value = str_replace('&', '&amp;', $value);
		$this->styleName = $odsStyleTableCell;
	}
}

class odsTableCellStringEmail extends odsTableCellString {

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = odsTableCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");

			// text:p
			$text_p = $dom->createElement('text:p');
				$table_table_cell->appendChild($text_p);

				// text:a
				$text_a = $dom->createElement('text:a',$this->value);
					$text_a->setAttribute("xlink:href", "mailto:".$this->value);
					$text_a->setAttribute("xlink:type", "simple");
					$text_p->appendChild($text_a);
		return $table_table_cell;
	}
}

class odsTableCellStringUrl extends odsTableCellString {

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = odsTableCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");

			// text:p
			$text_p = $dom->createElement('text:p');
				$table_table_cell->appendChild($text_p);

				// text:a
				$text_a = $dom->createElement('text:a',$this->value);
					$text_a->setAttribute("xlink:href", (substr($this->value,0,7)=="http://"?'':"http://").$this->value);
					$text_a->setAttribute("xlink:type", "simple");
					$text_p->appendChild($text_a);
		return $table_table_cell;
	}
}



class odsTableCellFloat extends odsTableCell {
	public $value;
	public $styleName;

	public function __construct($value,odsStyleTableCell $odsStyleTableCell = null) {
		$this->value = $value;
		$this->styleName = $odsStyleTableCell;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = odsTableCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "float");
			$table_table_cell->setAttribute("office:value", $this->value);

			// text:p
			$text_p = $dom->createElement('text:p',$this->value);
				$table_table_cell->appendChild($text_p);
		return $table_table_cell;
	}
}

class odsTableCellCurrency extends odsTableCell {
	public $value;
	public $styleName;
	public $currency;

	public function __construct($value, $currency, $odsStyleTableCell = null) {
		$this->value    = $value;
		$this->currency = $currency;
		$this->styleName = $odsStyleTableCell;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		switch($this->currency) {
		case 'EUR':
			$ods->addTmpStyles(new odsStyleMoneyEUR());
			$ods->addTmpStyles(new odsStyleMoneyEURNeg());
			break;
		case 'USD':
			$ods->addTmpStyles(new odsStyleMoneyUSD());
			$ods->addTmpStyles(new odsStyleMoneyUSDNeg());
			break;
		case 'GBP':
			$ods->addTmpStyles(new odsStyleMoneyGBP());
			$ods->addTmpStyles(new odsStyleMoneyGBPNeg());
			break;
		default:
			//FIXME: send error;
		}

		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);
			if( $this->styleName ) {
				$style = $ods->getStyleByName($this->styleName->getName()."-".$this->currency);
				if(!$style) {
					$style = clone $this->styleName;
					$style->setName($this->styleName->getName()."-".$this->currency);
					$style->setStyleDataName('NCur-'.$this->currency);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			} else {
				$style = $ods->getStyleByName("ce1-".$this->currency);
				if(!$style) {
					$style = clone $ods->getStyleByName("ce1");
					$style->setName("ce1-".$this->currency);
					$style->setStyleDataName('NCur-'.$this->currency);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			}

			$table_table_cell->setAttribute("office:value-type", "currency");
			$table_table_cell->setAttribute("office:currency", $this->currency);
			$table_table_cell->setAttribute("office:value", $this->value);

			// text:p
			$text_p = $dom->createElement('text:p');
				$table_table_cell->appendChild($text_p);

		return $table_table_cell;
	}
}

class odsTableCellDate extends odsTableCell {
	private $date;
	private $format;
	private $language;

	public function __construct($date,  $format="MMDDYYYY", $language=null, odsStyleGraphic $odsStyleCellDate = null) {
		$this->date = $date;
		$this->format = $format;
		$this->language = $language;
		$this->styleName = $odsStyleCellDate;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		switch($this->format) {
		case 'DDMMYYYY':
			$ods->addTmpStyles(new odsStyleDateDDMMYYYY($this->language));
			break;
		case 'DDMMYY':
			$ods->addTmpStyles($style = new odsStyleDateDDMMYY($this->language));
			break;
		case 'MMDDYYYY':
			$ods->addTmpStyles($style = new odsStyleDateMMDDYYYY($this->language));
			break;
		case 'MMDDYY':
			$ods->addTmpStyles($style = new odsStyleDateMMDDYY($this->language));
			break;
		case 'DMMMYYYY':
			$ods->addTmpStyles($style = new odsStyleDateDMMMYYYY($this->language));
			break;
		case 'DMMMYY':
			$ods->addTmpStyles($style = new odsStyleDateDMMMYY($this->language));
			break;
		case 'DMMMMYYYY':
			$ods->addTmpStyles($style = new odsStyleDateDMMMMYYYY($this->language));
			break;
		case 'DMMMMYY':
			$ods->addTmpStyles($style = new odsStyleDateDMMMMYY($this->language));
			break;
		case 'MMMDYYYY':
			$ods->addTmpStyles($style = new odsStyleDateMMMDYYYY($this->language));
			break;
		case 'MMMDYY':
			$ods->addTmpStyles($style = new odsStyleDateMMMDYY($this->language));
			break;
		default:
			//FIXME: send error;
		}

		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);

		if( $this->styleName ) {
				$style = $ods->getStyleByName($this->styleName->getName()."-".$this->format);
				if(!$style) {
					$style = clone $this->styleName;
					$style->setName($this->styleName->getName()."-".$this->format);
					$style->setStyleDataName('Date-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			} else {
				$style = $ods->getStyleByName("ce1-".$this->format);
				if(!$style) {
					$style = clone $ods->getStyleByName("ce1");
					$style->setName("ce1-".$this->format);
					$style->setStyleDataName('Date-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			}

		$table_table_cell->setAttribute("office:value-type","date");
		$table_table_cell->setAttribute("office:date-value",$this->date);
		return $table_table_cell;
	}

}

class odsTableCellTime extends odsTableCell {
	private $time;
	private $format;

	public function __construct($time, $format="HHMM", odsStyleGraphic $odsStyleCellDate = null) {
		$this->time = $time;
		$this->format = $format;
		$this->styleName = $odsStyleCellDate;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		switch($this->format) {
		case 'HHMMSS':
			$ods->addTmpStyles(new odsStyleTimeHHMMSS());
			break;
		case 'HHMM':
			$ods->addTmpStyles(new odsStyleTimeHHMM());
			break;
		case 'HHMMSSAMPM':
			$ods->addTmpStyles(new odsStyleTimeHHMMSSAMPM());
			break;
		case 'HHMMAMPM':
			$ods->addTmpStyles(new odsStyleTimeHHMMAMPM());
			break;
		default:
			//FIXME: send error;
		}

		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);

		if( $this->styleName ) {
				$style = $ods->getStyleByName($this->styleName->getName()."-".$this->format);
				if(!$style) {
					$style = clone $this->styleName;
					$style->setName($this->styleName->getName()."-".$this->format);
					$style->setStyleDataName('Time-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			} else {
				$style = $ods->getStyleByName("ce1-".$this->format);
				if(!$style) {
					$style = clone $ods->getStyleByName("ce1");
					$style->setName("ce1-".$this->format);
					$style->setStyleDataName('Time-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			}

		$table_table_cell->setAttribute("office:value-type","time");
		$table_table_cell->setAttribute("office:time-value",$this->time);
		return $table_table_cell;
	}
}

class odsTableCellDateTime extends odsTableCell {
	private $dateTime;
	private $format;

	public function __construct($dateTime, $format="MMDDYYHHMMAMPM", $language=null, odsStyleGraphic $odsStyleCellDate = null) {
		$this->dateTime = $dateTime;
		$this->format = $format;
		$this->language = $language;
		$this->styleName = $odsStyleCellDate;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		switch($this->format) {
		case 'MMDDYYHHMMSSAMPM':
			$ods->addTmpStyles(new odsStyleDateTimeMMDDYYHHMMSSAMPM($this->language));
			break;
		case 'MMDDYYHHMMAMPM':
			$ods->addTmpStyles(new odsStyleDateTimeMMDDYYHHMMAMPM($this->language));
			break;
		case 'DDMMYYHHMMSS':
			$ods->addTmpStyles(new odsStyleDateTimeDDMMYYHHMMSS($this->language));
			break;
		case 'DDMMYYHHMM':
			$ods->addTmpStyles(new odsStyleDateTimeDDMMYYHHMM($this->language));
			break;
		default:
			//FIXME: send error;
		}

		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);

		if( $this->styleName ) {
				$style = $ods->getStyleByName($this->styleName->getName()."-".$this->format);
				if(!$style) {
					$style = clone $this->styleName;
					$style->setName($this->styleName->getName()."-".$this->format);
					$style->setStyleDataName('DateTime-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			} else {
				$style = $ods->getStyleByName("ce1-".$this->format);
				if(!$style) {
					$style = clone $ods->getStyleByName("ce1");
					$style->setName("ce1-".$this->format);
					$style->setStyleDataName('DateTime-'.$this->format);
					$ods->addTmpStyles($style);
				}
				$table_table_cell->setAttribute("table:style-name", $style->getName());
			}

		$table_table_cell->setAttribute("office:value-type","date");
		$table_table_cell->setAttribute("office:date-value",$this->dateTime);
		return $table_table_cell;
	}
}

class odsTableCellDraw extends odsTableCell {
	private $odsDraw;

	public function __construct(odsDraw $odsDraw) {
		$this->odsDraw = $odsDraw;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);
		$table_table_cell->appendChild($this->odsDraw->getContent($ods, $dom));
		return $table_table_cell;
	}

}

class odsTableCellImage extends odsTableCell {

	protected $image; // odsDrawImage

	public function __construct($file, odsStyleGraphic $odsStyleGraphic = null) {
		$this->image = new odsDrawImage($file, '0mm', '0mm', null, null, $odsStyleGraphic);
	}

	public function setWidth($width) {
		$this->image->setWidth($width);
	}

	public function setHeight($heigth) {
		$this->image->setHeight($heigth);
	}

	public function setZIndex($zIndex) {
		$this->image->setZIndex($zIndex);
	}

	public function setX($x) {
		$this->image->setX($x);
	}

	public function setY($y) {
		$this->image->setY($y);
	}

	public function getContent(ods $ods, \DOMDocument $dom) {
		$table_table_cell = $dom->createElement('table:table-cell');
		$this->cellOpts($table_table_cell);
		$table_table_cell->appendChild($this->image->getContent($ods, $dom));

		return $table_table_cell;
	}
}

?>
