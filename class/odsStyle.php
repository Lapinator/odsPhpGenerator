<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

abstract class odsStyle {
	protected $name;
	protected $family; // table-column, table-row, table, table-cell
	
	protected function __construct($name, $family) {
		if($name == null) $name = $this->getType().'-'.$this->randString();
		
		$this->name   = $name;
		$this->family = $family;
	}
	
	protected function getContent(ods $ods, DOMDocument $dom) {
		$style_style = $dom->createElement('style:style');
			$style_style->setAttribute("style:name", $this->name);
			$style_style->setAttribute("style:family", $this->family);
		return $style_style;
	} 

	public function getName() {
		return $this->name;
	}
	
	public function __clone() {
		$this->name = 'clone-'.$this->name.'-'.rand(0,99999999999);
	}
	
	public function setName($name) {
		$this->name = $name;
	}
	
	abstract protected function getType();
	
	protected function randString() {
		return md5(time().rand().$this->getType());
	}
	
}

class odsStyleTableColumn extends odsStyle {
	private $breakBefore;         // auto
	private $columnWidth;         // 2.267cm
	
	public function __construct($name = null) {
		parent::__construct($name, "table-column");
		$this->breakBefore = "auto";
		$this->columnWidth = "2.267cm";
	}
	
	public function getContent(ods $ods,DOMDocument $dom) {
		$style_style = parent::getContent($ods,$dom);
			
			// style:table-column-properties
			$style_table_column_properties = $dom->createElement('style:table-column-properties');
				$style_table_column_properties->setAttribute("fo:break-before", $this->breakBefore);
				$style_table_column_properties->setAttribute("style:column-width", $this->columnWidth);
				$style_style->appendChild($style_table_column_properties);

		return $style_style;
	}
	
	public function getType() {
		return 'odsStyleTableColumn';
	}
}

class odsStyleTable extends odsStyle {
	private $masterPageName;      // Default
	private $display;             // true
	private $writingMode;         // lr-tb

	public function __construct($name = null) {
		parent::__construct($name, "table");
		$this->masterPageName = "Default";
		$this->display        = "true";
		$this->writingMode    = "lr-tb";
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$style_style = parent::getContent($ods,$dom);
			$style_style->setAttribute("style:master-page-name", $this->masterPageName);
			
			// style:table-properties
			$style_table_properties = $dom->createElement('style:table-properties');
				$style_table_properties->setAttribute("table:display", $this->display);
				$style_table_properties->setAttribute("style:writing-mode", $this->writingMode);
				$style_style->appendChild($style_table_properties);
		return $style_style;
	}
	
	public function getType() {
		return 'odsStyleTable';
	}
}

class odsStyleTableRow extends odsStyle {
	private $rowHeight;           // 0.52cm
	private $breakBefore;         // auto
	private $useOptimalRowHeight; // true, false
	
	public function __construct($name = null) {
		parent::__construct($name, "table-row");
		$this->name                = $name;
		$this->family              = "table-row";
		$this->rowHeight           = "0.52cm";
		$this->breakBefore         = "auto";
		$this->useOptimalRowHeight = "true";
	}

	public function setRowHeight($rowHeight) {
		$this->rowHeight = $rowHeight;
	}

	public function setBreakBefore($breakBefore) {
		$this->breakBefore = $breakBefore;
	}

	public function setUseOptimalRowHeight($useOptimalRowHeight) {
		$this->useOptimalRowHeight = $useOptimalRowHeight;
	}

	public function getContent(ods $ods, DOMDocument $dom) {
		$style_style = parent::getContent($ods,$dom);
			
			// style:table-row-properties
			$style_table_row_properties = $dom->createElement('style:table-row-properties');
				$style_table_row_properties->setAttribute("style:row-height", $this->rowHeight);
				$style_table_row_properties->setAttribute("fo:break-before", $this->breakBefore);
				$style_table_row_properties->setAttribute("style:use-optimal-row-height", $this->useOptimalRowHeight);
				$style_style->appendChild($style_table_row_properties);
				
		return $style_style; 
	}
	
	public function getType() {
		return 'odsStyleTableRow';
	}
}

class odsStyleTableCell extends odsStyle {
	private $parentStyleName;     // Default
	private $textAlignSource;     // fix
	private $repeatContent;       // true, false
	private $color;               // opt: #ffffff
	private $backgroundColor;     // opt: #ffffff
	private $border;              // opt: 0.002cm solid #000000
	private $textAlign;           // opt: center
	private $marginLeft;          // opt: 0cm
	private $fontWeight;          // opt: bold
	private $fontSize;            // opt: 18pt;
	private $fontStyle;           // opt: italic, normal
	private $underline;           // opt: font-color, #000000, null
	private $styleDataName;       // opt: interne
	
	public function __construct($name = null) {
		parent::__construct($name, "table-cell");
		$this->parentStyleName     = "Default";
		$this->textAlignSource     = "fix";
		$this->repeatContent       = "false";
		$this->color               = false;
		$this->backgroundColor     = false;
		$this->border              = false;
		$this->textAlign           = false;
		$this->marginLeft          = false;
		$this->fontWeight          = false;
		$this->fontSize            = false;
		$this->fontStyle           = false;
		$this->underline           = false;
		$this->styleDataName       = false;
		
	}
	
	public function setColor($color) {
		$this->color = $color;
	}
	
	public function setBackgroundColor($color) {
		$this->backgroundColor = $color;
	}

	public function setBorder($border) {
		$this->border = $border;
	}
	
	public function setTextAlign($textAlign) {
		$this->textAlign = $textAlign;
	}
	
	public function setFontWeight($weight) {
		$this->fontWeight = $weight;
	}
	
	public function setFontStyle($fontStyle) {
		$this->fontStyle = $fontStyle;
	}
	
	public function setUnderline($underline) {
		$this->underline = $underline;
	}
	
	public function setStyleDataName($styleDataName) {
		$this->styleDataName = $styleDataName;
	}
	
	public function setFontSize($fontSize) {
		$this->fontSize = $fontSize;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		// style:style
		$style_style = parent::getContent($ods,$dom);
			$style_style->setAttribute("style:parent-style-name", $this->parentStyleName);
			if($this->styleDataName)
				$style_style->setAttribute("style:data-style-name", $this->styleDataName);
			
			// style:table-cell-properties
			$style_table_cell_properties = $dom->createElement('style:table-cell-properties');
				$style_table_cell_properties->setAttribute("style:text-align-source", $this->textAlignSource);
				$style_table_cell_properties->setAttribute("style:repeat-content", $this->repeatContent);
				
				if($this->backgroundColor)
					$style_table_cell_properties->setAttribute("fo:background-color", $this->backgroundColor);
					
				if($this->border)
					$style_table_cell_properties->setAttribute("fo:border", $this->border);
				
				$style_style->appendChild($style_table_cell_properties);

				if($this->textAlign) {
					// style:paragraph-properties
					$style_paragraph_properties = $dom->createElement('style:paragraph-properties');
						$style_paragraph_properties->setAttribute("fo:text-align", $this->textAlign);
						$style_paragraph_properties->setAttribute("fo:margin-left", "0cm");
						$style_style->appendChild($style_paragraph_properties);
				}
				
				if($this->color OR $this->fontWeight OR $this->fontStyle OR $this->underline OR $this->fontSize) {
					// style:text-properties
					$style_text_properties = $dom->createElement('style:text-properties');
					
						if($this->color)
							$style_text_properties->setAttribute("fo:color", $this->color);
							
						if($this->fontWeight) {
							$style_text_properties->setAttribute("fo:font-weight", $this->fontWeight);
							$style_text_properties->setAttribute("style:font-weight-asian", $this->fontWeight);
							$style_text_properties->setAttribute("style:font-weight-complex", $this->fontWeight);
						}
						
						if($this->fontStyle) {
							$style_text_properties->setAttribute("fo:font-style", $this->fontStyle);
							$style_text_properties->setAttribute("fo:font-style-asian", $this->fontStyle);
							$style_text_properties->setAttribute("fo:font-style-complex", $this->fontStyle);
						}
						
						if($this->underline) {
							$style_text_properties->setAttribute("style:text-underline-style", 'solid');
							$style_text_properties->setAttribute("style:text-underline-width", 'auto');
							$style_text_properties->setAttribute("style:text-underline-color", $this->underline);
						}
						
						if($this->fontSize) {
							$style_text_properties->setAttribute("fo:font-size", $this->fontSize);
							$style_text_properties->setAttribute("style:font-size-asian", $this->fontSize);
							$style_text_properties->setAttribute("style:font-size-complex", $this->fontSize);
						}
						
						$style_style->appendChild($style_text_properties);
				}
				
				
		return $style_style;
	}
	 
	public function getType() {
		return 'odsStyleTableCell';
	}
}

abstract class odsStyleMoney extends odsStyle {
	//abstract protected function __construct();
	//abstract protected function getContent();
	//abstract protected function getType();
}


class odsStyleMoneyEUR extends odsStyleMoney {
	
	public function __construct() {
		$this->name='NCur-EUR-P0';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-EUR-P0");
			$number_currency_style->setAttribute("style:volatile", "true");
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
				
			$number_text = $dom->createElement('number:text', ' ');
				$number_currency_style->appendChild($number_text);
		
			$number_currency_symbol = $dom->createElement('number:currency-symbol', '€');
				$number_currency_symbol->setAttribute("number:language", "fr");
				$number_currency_symbol->setAttribute("number:country", "FR");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_currency_symbol);
		return $number_currency_style;
	}
	
	public function getType() {
		return 'odsStyleMoneyEUR';
	}
}

class odsStyleMoneyEURNeg extends odsStyleMoney {
	
	public function __construct() {
		$this->name='NCur-EUR';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
	
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-EUR");
	
			$style_text_properties = $dom->createElement('style:text-properties');
				$style_text_properties->setAttribute("fo:color", "#ff0000");
				$number_currency_style->appendChild($style_text_properties);
	
			$number_text = $dom->createElement('number:text', '-');
				$number_currency_style->appendChild($number_text);
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
				
			$number_text = $dom->createElement('number:text', ' ');
				$number_currency_style->appendChild($number_text);
		
			$number_currency_symbol = $dom->createElement('number:currency-symbol', '€');
				$number_currency_symbol->setAttribute("number:language", "fr");
				$number_currency_symbol->setAttribute("number:country", "FR");
				$number_currency_style->appendChild($number_currency_symbol);
		
			$style_map = $dom->createElement('style:map');
				$style_map->setAttribute("style:condition", "value()>=0");
				$style_map->setAttribute("style:apply-style-name", "NCur-EUR-P0");
				$number_currency_style->appendChild($style_map);

		return $number_currency_style;
	}

	public function getType() {
		return 'odsStyleMoneyEURNeg';
	}
}

class odsStyleMoneyUSD extends odsStyleMoney {
	
	public function __construct() {
		$this->name='NCur-USD-P0';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-USD-P0");
			$number_currency_style->setAttribute("style:volatile", "true");

			$number_currency_symbol = $dom->createElement('number:currency-symbol', '$');
				$number_currency_symbol->setAttribute("number:language", "en");
				$number_currency_symbol->setAttribute("number:country", "US");
				$number_currency_symbol->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_currency_symbol);
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
		return $number_currency_style;
		
	}
	
	public function getType() {
		return 'odsStyleMoneyUSD';
	}
}

class odsStyleMoneyUSDNeg extends odsStyleMoney {
	
	public function __construct(){
		$this->name='NCur-USD';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-USD");

			$style_text_properties = $dom->createElement('style:text-properties');
				$style_text_properties->setAttribute("fo:color", "#ff0000");
				$number_currency_style->appendChild($style_text_properties);

			$number_text = $dom->createElement('number:text', '-');
				$number_currency_style->appendChild($number_text);
		
			$number_currency_symbol = $dom->createElement('number:currency-symbol', '$');
				$number_currency_symbol->setAttribute("number:language", "en");
				$number_currency_symbol->setAttribute("number:country", "US");
				$number_currency_style->appendChild($number_currency_symbol);
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
								
			$style_map = $dom->createElement('style:map');
				$style_map->setAttribute("style:condition", "value()>=0");
				$style_map->setAttribute("style:apply-style-name", "NCur-USD-P0");
				$number_currency_style->appendChild($style_map);
		return $number_currency_style;
	}
	
	public function getType() {
		return 'odsStyleMoneyUSDNeg';
	}
}

class odsStyleMoneyGBP extends odsStyleMoney {
	
	public function __construct() {
		$this->name='NCur-GBP-P0';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-GBP-P0");
			$number_currency_style->setAttribute("style:volatile", "true");

			$number_currency_symbol = $dom->createElement('number:currency-symbol', '£');
				$number_currency_symbol->setAttribute("number:language", "en");
				$number_currency_symbol->setAttribute("number:country", "GB");
				$number_currency_style->appendChild($number_currency_symbol);
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
		return $number_currency_style;
	}
	
	public function getType() {
		return 'odsStyleMoneyGBP';
	}
}

class odsStyleMoneyGBPNeg extends odsStyleMoney {
	
	public function __construct() {
		$this->name='NCur-GBP';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_currency_style = $dom->createElement('number:currency-style');
			$number_currency_style->setAttribute("style:name", "NCur-GBP");

			$style_text_properties = $dom->createElement('style:text-properties');
				$style_text_properties->setAttribute("fo:color", "#ff0000");
				$number_currency_style->appendChild($style_text_properties);

			$number_text = $dom->createElement('number:text', '-');
				$number_currency_style->appendChild($number_text);
		
			$number_currency_symbol = $dom->createElement('number:currency-symbol', '£');
				$number_currency_symbol->setAttribute("number:language", "en");
				$number_currency_symbol->setAttribute("number:country", "GB");
				$number_currency_style->appendChild($number_currency_symbol);
		
			$number_number = $dom->createElement('number:number');
				$number_number->setAttribute("number:decimal-places", "2");
				$number_number->setAttribute("number:min-integer-digits", "1");
				$number_number->setAttribute("number:grouping", "true");
				$number_currency_style->appendChild($number_number);
								
			$style_map = $dom->createElement('style:map');
				$style_map->setAttribute("style:condition", "value()>=0");
				$style_map->setAttribute("style:apply-style-name", "NCur-GBP-P0");
				$number_currency_style->appendChild($style_map);
		return $number_currency_style;
	}
	
	public function getType() {
		return 'odsStyleMoneyGBPNeg';
	}
}

?>
