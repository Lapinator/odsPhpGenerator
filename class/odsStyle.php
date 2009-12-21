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
	
	public function setColumnWidth($columnWidth) {
		$this->columnWidth = $columnWidth;
	}
	
	public function setBreakBefore($breakBefore) {
		$this->breakBefore = $breakBefore;
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
	private $fontFace;            // opt: fontFace
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
		$this->fontFace            = false;
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
	
	public function setFontFace(odsFontFace $fontFace) {
		$this->fontFace = $fontFace;
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
				
				if($this->color OR $this->fontWeight OR $this->fontStyle OR $this->underline OR $this->fontSize OR $this->fontFace) {
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
						
						if($this->fontFace) {
							$style_text_properties->setAttribute("style:font-name", $this->fontFace->getFontName());
						}
						
						$style_style->appendChild($style_text_properties);
				}
				
				
		return $style_style;
	}
	 
	public function getType() {
		return 'odsStyleTableCell';
	}
}

class odsStyleGraphic extends odsStyle {
	private $stroke;    // none
	private $fill;      // none
 	private $luminance; // 0%
 	private $contrast;  // 0%
 	private $gamma;     // 100%
 	private $red;       // 0%
 	private $green;     // 0%
 	private $blue;      // 0%  
 	private $opacity;   // 100%

	public function __construct($name = null) {
		parent::__construct($name, "graphic");
		$this->stroke    = "none";
		$this->fill      = "none";
		$this->luminance = "0%";
		$this->contrast  = "0%";
		$this->gamma     = "100%";
		$this->red       = "0%";
		$this->green     = "0%";
		$this->blue      = "0%";
		$this->opacity   = "100%";
	}
	
	public function setStroke($stroke) {
		$this->stroke = $stroke;
	}
	
	public function setFill($fill) {
		$this->fill = $fill;
	}
	
	public function setLuminance($luminance) {
		$this->luminance = $luminance;
	}
	
	public function setContrast($contrast) {
		$this->contrast = $contrast;
	}
	
	public function setGamma($gamma) {
		$this->gamma = $gamma;
	}
	
	public function setRed($red) {
		$this->red = $red;
	}
	
	public function setGreen($green) {
		$this->green = $green;
	}
	
	public function setBlue($blue) {
		$this->blue = $blue;
	}
	
	public function setOpacity($opacity) {
		$this->opacity = $opacity;
	}
	
	public function getContent(ods $ods,DOMDocument $dom) {
		$style_style = parent::getContent($ods,$dom);
		
			// style:table-row-properties
			$style_graphic_properties = $dom->createElement('style:graphic-properties');
				$style_graphic_properties->setAttribute("draw:stroke",        $this->stroke);
				$style_graphic_properties->setAttribute("draw:fill",          $this->fill);
				$style_graphic_properties->setAttribute("draw:textarea-horizontal-align", "center");
				$style_graphic_properties->setAttribute("draw:textarea-vertical-align", "middle");
				$style_graphic_properties->setAttribute("draw:color-mode", "standard");
				$style_graphic_properties->setAttribute("draw:luminance",     $this->luminance);
				$style_graphic_properties->setAttribute("draw:contrast",      $this->contrast);
				$style_graphic_properties->setAttribute("draw:gamma",         $this->gamma);
				$style_graphic_properties->setAttribute("draw:red",           $this->red);
				$style_graphic_properties->setAttribute("draw:green",         $this->green);
				$style_graphic_properties->setAttribute("draw:blue",          $this->blue);
				$style_graphic_properties->setAttribute("fo:clip", "rect(0cm, 0cm, 0cm, 0cm)");
				$style_graphic_properties->setAttribute("draw:image-opacity", $this->opacity);
				$style_graphic_properties->setAttribute("style:mirror", "none");
				$style_style->appendChild($style_graphic_properties);
				
		return $style_style;
	}

	public function getType() {
		return 'odsStyleGraphic';
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

abstract class odsStyleDate extends odsStyle {
	protected $language;
	
	protected function setLanguage($number_date_style) {
		if(!isset($this->language)) return;
		$number_date_style->setAttribute("number:language", $this->language);
		$number_date_style->setAttribute("number:country", $this->language);
	}
}

class odsStyleDateMMDDYYYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-MMDDYYYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_day->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_year->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateMMDDYYYY';
	}
}

class odsStyleDateMMDDYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-MMDDYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_day->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateMMDDYY';
	}
}

class odsStyleDateDDMMYYYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DDMMYYYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_day->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_year->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDDMMYYYY';
	}
}

class odsStyleDateDDMMYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DDMMYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_day->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDDMMYY';
	}
}

class odsStyleDateDMMMYYYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DMMMYYYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_year->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMYYYY';
	}
}

class odsStyleDateDMMMYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DMMMYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMYY';
	}
}

class odsStyleDateDMMMMYYYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DMMMMYYYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_month->setAttribute(" number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_year->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMMYYYY';
	}
}

class odsStyleDateDMMMMYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-DMMMMYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);
	
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_month->setAttribute(" number:style", "long");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMMYY';
	}
}

class odsStyleDateMMMDYYYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-MMMDYYYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', ', ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_year->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMMYY';
	}
}

class odsStyleDateMMMDYY extends odsStyleDate {
	public function __construct($language) {
		$this->name='Date-MMMDYY';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_month->setAttribute("number:textual", "true");
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', ', ');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateDMMMM';
	}
}

abstract class odsStyleTime extends odsStyle {
}

class odsStyleTimeHHMMSS extends odsStyleTime {
	public function __construct() {
		$this->name='Time-HHMMSS';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_time_style = $dom->createElement('number:time-style');
			$number_time_style->setAttribute("style:name", $this->name);
			
			$number_hours = $dom->createElement('number:hours');
				$number_hours->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
				$number_minutes->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_minutes);	

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_seconds = $dom->createElement('number:seconds');
				$number_seconds->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_seconds);	
			
		return $number_time_style;
	}

	public function getType() {
		return 'odsStyleTimeHHMMSS';
	}
}

class odsStyleTimeHHMM extends odsStyleTime {
	public function __construct() {
		$this->name='Time-HHMM';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_time_style = $dom->createElement('number:time-style');
			$number_time_style->setAttribute("style:name", $this->name);
			
			$number_hours = $dom->createElement('number:hours');
				$number_hours->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
				$number_minutes->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_minutes);	
			
		return $number_time_style;
	}

	public function getType() {
		return 'odsStyleTimeHHMM';
	}
}

class odsStyleTimeHHMMSSAMPM extends odsStyleTime {
	public function __construct() {
		$this->name='Time-HHMMSSAMPM';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_time_style = $dom->createElement('number:time-style');
			$number_time_style->setAttribute("style:name", $this->name);
			
			$number_hours = $dom->createElement('number:hours');
				$number_hours->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
				$number_minutes->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_minutes);	

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_seconds = $dom->createElement('number:seconds');
				$number_seconds->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_seconds);	

			$number_text = $dom->createElement('number:text', ' ');
				$number_time_style->appendChild($number_text);

			$number_am_pm = $dom->createElement('number:am-pm');
				$number_am_pm->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_am_pm);	
			
		return $number_time_style;
	}

	public function getType() {
		return 'odsStyleTimeHHMMSSAMPM';
	}
}

class odsStyleTimeHHMMAMPM extends odsStyleTime {
	public function __construct() {
		$this->name='Time-HHMMAMPM';
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_time_style = $dom->createElement('number:time-style');
			$number_time_style->setAttribute("style:name", $this->name);
			
			$number_hours = $dom->createElement('number:hours');
				$number_hours->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_time_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
				$number_minutes->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_minutes);	

			$number_text = $dom->createElement('number:text', ' ');
				$number_time_style->appendChild($number_text);
			
			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_am_pm = $dom->createElement('number:am-pm');
				$number_am_pm->setAttribute("number:style", "long");
				$number_time_style->appendChild($number_am_pm);	
			
		return $number_time_style;
	}

	public function getType() {
		return 'odsStyleTimeHHMMAMPM';
	}
}

abstract class odsStyleDateTime extends odsStyleDate {
}

class odsStyleDateTimeMMDDYYHHMMSSAMPM extends odsStyleDateTime {
	public function __construct($language) {
		$this->name='DateTime-MMDDYYHHMMSSAMPM';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_hours = $dom->createElement('number:hours');
			$number_hours->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
			$number_minutes->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_minutes);
				
			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_seconds = $dom->createElement('number:seconds');
				$number_seconds->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_seconds);
				
			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);
			
			$number_am_pm = $dom->createElement('number:am-pm');
				$number_date_style->appendChild($number_am_pm);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateTimeMMDDYYHHMMSSAMPM';
	}
}

class odsStyleDateTimeMMDDYYHHMMAMPM extends odsStyleDateTime {
	public function __construct($language) {
		$this->name='DateTime-MMDDYYHHMMAMPM';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_month = $dom->createElement('number:month');
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_hours = $dom->createElement('number:hours');
			$number_hours->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
			$number_minutes->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_minutes);
			
			$number_am_pm = $dom->createElement('number:am-pm');
				$number_date_style->appendChild($number_am_pm);
			
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateTimeMMDDYYHHMMAMPM';
	}
}

class odsStyleDateTimeDDMMYYHHMMSS extends odsStyleDateTime {
	public function __construct($language) {
		$this->name='DateTime-DDMMYYHHMMSS';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);
			
			$number_month = $dom->createElement('number:month');
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_hours = $dom->createElement('number:hours');
			$number_hours->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
			$number_minutes->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_minutes);
				
			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_seconds = $dom->createElement('number:seconds');
				$number_seconds->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_seconds);
				
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateTimeDDMMYYHHMMSS';
	}
}

class odsStyleDateTimeDDMMYYHHMM extends odsStyleDateTime {
	public function __construct($language) {
		$this->name='DateTime-DDMMYYHHMM';
		$this->language = $language;
	}
	
	public function getContent(ods $ods, DOMDocument $dom) {
		$number_date_style = $dom->createElement('number:date-style');
			$number_date_style->setAttribute("style:name", $this->name);
			$number_date_style->setAttribute("number:automatic-order", "true");
			$this->setLanguage($number_date_style);
			
			$number_day = $dom->createElement('number:day');
				$number_date_style->appendChild($number_day);	

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);
			
			$number_month = $dom->createElement('number:month');
				$number_date_style->appendChild($number_month);

			$number_text = $dom->createElement('number:text', '/');
				$number_date_style->appendChild($number_text);

			$number_year = $dom->createElement('number:year');
				$number_date_style->appendChild($number_year);

			$number_text = $dom->createElement('number:text', ' ');
				$number_date_style->appendChild($number_text);

			$number_hours = $dom->createElement('number:hours');
			$number_hours->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_hours);

			$number_text = $dom->createElement('number:text', ':');
				$number_date_style->appendChild($number_text);

			$number_minutes = $dom->createElement('number:minutes');
			$number_minutes->setAttribute("number:style", "long");
				$number_date_style->appendChild($number_minutes);
				
		return $number_date_style;
	}

	public function getType() {
		return 'odsStyleDateTimeDDMMYYHHMM';
	}
}

?>