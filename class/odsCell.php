<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

abstract class odsCell {
	protected $styleName;
	
	abstract protected function __construct();
	
	protected function getContent(ods $ods, DOMDocument $dom) {
		$table_table_cell = $dom->createElement('table:table-cell');
		if( $this->styleName ) {
			$ods->addTmpStyles($this->styleName);
			$table_table_cell->setAttribute("table:style-name", $this->styleName->getName());
		}
		return $table_table_cell;
	}
}

class odsCellEmpty extends odsCell {
	
	public function __construct(odsStyleTableCell $odsStyleTableCell = null) {
		$this->styleName = $odsStyleTableCell;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		return parent::getContent($ods,$dom);
		//$table_table_row->setAttribute("table:number-columns-repeated", "1022");
		//return $table_table_cell;
	}
}

class odsCellString extends odsCell {
	public $value;
	public $styleName;
	
	public function __construct($value,odsStyleTableCell $odsStyleTableCell = null) {
		$this->value = $value;
		$this->styleName = $odsStyleTableCell;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$table_table_cell = parent::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");
			
			// text:p
			$text_p = $dom->createElement('text:p',$this->value);
				$table_table_cell->appendChild($text_p);
		return $table_table_cell;
	}
}

class odsCellStringEmail extends odsCellString {
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$table_table_cell = odsCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");
			
			// text:p
			$text_p = $dom->createElement('text:p');
				$table_table_cell->appendChild($text_p);
				
				// text:a
				$text_a = $dom->createElement('text:a',$this->value);
					$text_a->setAttribute("xlink:href", "mailto:".$this->value);
					$text_p->appendChild($text_a);
		return $table_table_cell;
	}
}

class odsCellStringUrl extends odsCellString {
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$table_table_cell = odsCell::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "string");
			
			// text:p
			$text_p = $dom->createElement('text:p');
				$table_table_cell->appendChild($text_p);
				
				// text:a
				$text_a = $dom->createElement('text:a',$this->value);
					$text_a->setAttribute("xlink:href", (substr($this->value,0,7)=="http://"?'':"http://").$this->value);
					$text_p->appendChild($text_a);
		return $table_table_cell;
	}
}



class odsCellFloat extends odsCell {
	public $value;
	public $styleName;
	
	public function __construct($value,odsStyleTableCell $odsStyleTableCell = null) {
		$this->value = $value;
		$this->styleName = $odsStyleTableCell;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$table_table_cell = parent::getContent($ods,$dom);
			$table_table_cell->setAttribute("office:value-type", "float");
			$table_table_cell->setAttribute("office:value", $this->value);
			
			// text:p
			$text_p = $dom->createElement('text:p',$this->value);
				$table_table_cell->appendChild($text_p);
		return $table_table_cell;
	}
}

class odsCellCurrency extends odsCell {
	public $value;
	public $styleName;
	public $currency;
	
	public function __construct($value, $currency, $odsStyleTableCell = null) {
		$this->value    = $value;
		$this->currency = $currency;
		$this->styleName = $odsStyleTableCell;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
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

?>
