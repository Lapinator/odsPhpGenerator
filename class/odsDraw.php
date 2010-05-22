<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */
 
abstract class odsDraw {
	protected $styleGraphic;
	protected $zIndex;           // number 
	protected $tableBackground;  // "true", "false", null
	
	abstract function __construct();
	abstract function getContent(ods $ods, DOMDocument $dom);
	abstract protected function getType();
	
	public function setZIndex($zIndex){
		$this->zIndex = $zIndex;
	}
	
	public function setTableBackground($tableBackground) {
		$this->tableBackground = $tableBackground;
	}
	
}

class odsDrawLine extends odsDraw {
	private $styleParagraph;
	private $x1; // 1cm
	private $y1; // 1cm
	private $x2; // 2cm
	private $y2; // 2cm
	
	static private $defaultStyle = null;
	static private $defaultStyleParagraph = null;

	
	public function __construct($x1, $y1, $x2, $y2, $odsStyleGraphic = null, $odsStyleParagraph = null ) {
		$this->styleGraphic    = $odsStyleGraphic;
		$this->styleParagraph  = $odsStyleParagraph;
		$this->tableBackground = null; 
		$this->x1              = $x1;
		$this->y1              = $y1;
		$this->x2              = $x2;
		$this->y2              = $y2;
		$this->zIndex          = "0";		
	}
	
	function getContent(ods $ods, DOMDocument $dom) {

		if($this->styleGraphic)
			$style = $this->styleGraphic;
		else 
			$style = self::getOdstyleGraphic();

		if($this->styleParagraph)
			$styleParagraph = $this->styleParagraph;
		else 
			$styleParagraph = self::getOdstyleParagraph();
		
		$ods->addTmpStyles($style);
		$ods->addTmpStyles($styleParagraph);
				
		$draw_line = $dom->createElement('draw:line');
			$draw_line->setAttribute('draw:z-index', $this->zIndex);
			$draw_line->setAttribute('draw:style-name', $style->getName());
			$draw_line->setAttribute('draw:text-style-name', $styleParagraph->getName());
			$draw_line->setAttribute('svg:x1', $this->x1);
			$draw_line->setAttribute('svg:y1', $this->y1);
			$draw_line->setAttribute('svg:x2', $this->x2);
			$draw_line->setAttribute('svg:y2', $this->y2);
			
			if($this->tableBackground)
				$draw_line->setAttribute('table:table-background', $this->tableBackground);
			
			$text_p = $dom->createElement('text:p');
				$draw_line->appendChild($text_p);
			
		return $draw_line;
	}
	
	static public function getOdstyleGraphic() {
		if(!self::$defaultStyle)
			self::$defaultStyle = new odsStyleGraphic("gr-line"); 
		return self::$defaultStyle;
	}
	
	static public function getOdstyleParagraph() {
		if(!self::$defaultStyleParagraph)
			self::$defaultStyleParagraph = new odsStyleParagraph("p1");
		return self::$defaultStyleParagraph;
	}
	
	public function getType() {
		return "odsDrawLine";
	}
}

class odsDrawRect extends odsDraw {
	private $styleParagraph;
	private $x;
	private $y;
	private $width;
	private $height;
	
	static private $defaultStyle = null;
	static private $defaultStyleParagraph = null;
	
	public function __construct($x, $y, $width, $height, $odsStyleGraphic = null, $odsStyleParagraph = null ) {
		$this->styleGraphic    = $odsStyleGraphic;
		$this->styleParagraph  = $odsStyleParagraph;
		$this->tableBackground = null;
		$this->x               = $x;
		$this->y               = $y;
		$this->width           = $width;
		$this->height          = $height;
		$this->zIndex          = "0";	
	}
	
	function getContent(ods $ods, DOMDocument $dom) {
		
		if($this->styleGraphic)
			$style = $this->styleGraphic;
		else 
			$style = self::getOdstyleGraphic();
			
		if($this->styleParagraph)
			$styleParagraph = $this->styleParagraph;
		else 
			$styleParagraph = self::getOdstyleParagraph();
		
		$ods->addTmpStyles($style);
		$ods->addTmpStyles($styleParagraph);
		
		$draw_rect = $dom->createElement('draw:rect');
			$draw_rect->setAttribute('draw:z-index', $this->zIndex);
			$draw_rect->setAttribute('draw:style-name', $style->getName());
			$draw_rect->setAttribute('draw:text-style-name', $styleParagraph->getName());
			$draw_rect->setAttribute('svg:x',      $this->x);
			$draw_rect->setAttribute('svg:y',      $this->y);
			$draw_rect->setAttribute('svg:width',  $this->width);
			$draw_rect->setAttribute('svg:height', $this->height);
			
			if($this->tableBackground)
				$draw_rect->setAttribute('table:table-background', $this->tableBackground);
			
			$text_p = $dom->createElement('text:p');
				$draw_rect->appendChild($text_p);
			
		return $draw_rect;
	}
	
	static public function getOdstyleGraphic() {
		if(!self::$defaultStyle)
			self::$defaultStyle = new odsStyleGraphic("gr-rect"); 
		return self::$defaultStyle;
	}
	
	static public function getOdstyleParagraph() {
		if(!self::$defaultStyleParagraph)
			self::$defaultStyleParagraph = new odsStyleParagraph("p1");
		return self::$defaultStyleParagraph;
	}
	
	public function getType() {
		return "odsDrawRect";
	}
	
}

?>
