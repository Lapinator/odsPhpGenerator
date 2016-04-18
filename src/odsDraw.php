<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

namespace  odsPhpGenerator;
 
abstract class odsDraw {
	protected $styleGraphic;
	protected $zIndex;           // number 
	protected $tableBackground;  // "true", "false", null
	
	//abstract function __construct();
	abstract function getContent(ods $ods, \DOMDocument $dom);
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
	
	public function getContent(ods $ods, \DOMDocument $dom) {

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
	
	public function getContent(ods $ods, \DOMDocument $dom) {
		
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

class odsDrawImage extends odsDraw {

	private $file;
	private $styleParagraph;
	private $x;
	private $y;
	private $width;
	private $height;

	static private $defaultStyle = null;
	static private $defaultStyleParagraph = null;

	public function __construct($file, $x='0mm', $y='0mm', $width=null, $height=null, $odsStyleGraphic = null, $odsStyleParagraph = null ) {
		$this->file            = $file;
		$this->styleGraphic    = $odsStyleGraphic;
		$this->styleParagraph  = $odsStyleParagraph;
		$this->tableBackground = null;
		$this->x               = $x;
		$this->y               = $y;
		$this->width           = $width;
		$this->height          = $height;
		$this->zIndex          = "0";

		if(!$this->width OR !$this->height) {
			$im = imagecreatefromstring( file_get_contents( $file ));
			if(!$this->width) $this->width  = (imagesx($im)*0.035276875)."cm";
			if(!$this->height) $this->height = (imagesy($im)*0.035276875)."cm";
			imagedestroy($im);
		}
	}

	public function setWidth($width) {
		$this->width = $width;
	}

	public function setHeight($heigth) {
		$this->heigth = $heigth;
	}

	public function setZIndex($zIndex) {
		$this->zIndex = $zIndex;
	}

	public function setX($x) {
		$this->$x = $x;
	}

	public function setY($y) {
		$this->$y = $y;
	}

	public function getContent(ods $ods, \DOMDocument $dom) {

		if($this->styleGraphic) $style = $this->styleGraphic;
		else $style = self::getOdstyleGraphic();

		if($this->styleParagraph) $styleParagraph = $this->styleParagraph;
		else $styleParagraph = self::getOdstyleParagraph();

		$draw_frame = $dom->createElement('draw:frame');
			$draw_frame->setAttribute("draw:z-index", $this->zIndex);
			$draw_frame->setAttribute("draw:name", "Images ".md5(time().rand()));
			$draw_frame->setAttribute("draw:style-name", $style->getName());
			$draw_frame->setAttribute("draw:text-style-name", $styleParagraph->getName());
			$draw_frame->setAttribute("svg:width", $this->width);
			$draw_frame->setAttribute("svg:height", $this->height);
			$draw_frame->setAttribute("svg:x", $this->x);
			$draw_frame->setAttribute("svg:y", $this->y);

			$draw_image = $dom->createElement('draw:image');
				$draw_image->setAttribute("xlink:href", $ods->addTmpPictures($this->file));
				$draw_image->setAttribute("xlink:type", "simple");
				$draw_image->setAttribute("xlink:show", "embed");
				$draw_image->setAttribute("xlink:actuate", "onLoad");
				$draw_frame->appendChild($draw_image);

				$text_p = $dom->createElement('text:p');
					$draw_image->appendChild($text_p);

		return $draw_frame;
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
		return "odsDrawImages";
	}

}
