<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

require_once("class/odsFontFace.php");
require_once("class/odsStyle.php");
require_once("class/odsTable.php");
require_once("class/odsTableColumn.php");
require_once("class/odsTableRow.php");
require_once("class/odsTableCell.php");

class ods {
	private $defaultTable;
	
	private $scripts;        // FIXME: Looking
	private $fontFaces;
	private $styles;
	private $tmpStyles;
	private $tmpPictures;
	private $tables;
	
	private $title;
	private $subject;
	private $keyword;
	private $description;
	
	public function __construct() {
		$this->title       = null;
		$this->subject     = null;
		$this->keyword     = null;
		$this->description = null;
		
		$this->defaultTable = null;
		
		$this->scripts   = array();
		$this->fontFaces = array();
		$this->styles    = array();
		$this->tables    = array();
		
		$this->addFontFaces( new odsFontFace( "Nimbus Sans L", "'Nimbus Sans L'", "swiss", "variable" ) );
		$this->addFontFaces( new odsFontFace( "Bitstream Vera Sans", "'Bitstream Vera Sans'", "system", "variable" ) );
		
		$this->addStyles( new odsStyleTableColumn("co1") );
		$this->addStyles( new odsStyleTable("ta1") );
		$this->addStyles( new odsStyleTableRow("ro1") );
		$this->addStyles( new odsStyleTableCell("ce1") );
		
	}
	
	public function setTitle($title) {
		$this->title = $title;
	}
	
	public function setSubject($subject) {
		$this->subject = $subject;
	}
	
	public function setKeyword($keyword) {
		$this->keyword = $keyword;
	}
	
	public function setDescription($description) {
		$this->description = $description;
	}
	
	public function addFontFaces(odsFontFace $odsFontFace) {
		if(in_array($odsFontFace,$this->fontFaces)) return;
		$this->fontFaces[$odsFontFace->getName()] = $odsFontFace;
	}
	
	public function addStyles(odsStyle $odsStyle) {
		if(in_array($odsStyle,$this->styles)) return;
		$this->styles[$odsStyle->getName()] = $odsStyle;
	}
	
	public function addTmpStyles(odsStyle $odsStyle) {
		if(in_array($odsStyle,$this->styles)) return;
		if(in_array($odsStyle,$this->tmpStyles)) return;
		$this->tmpStyles[$odsStyle->getName()] = $odsStyle;
	}
	
	public function addTmpPictures($file) {
		if(in_array($file,$this->tmpPictures)) return;
		$this->tmpPictures[$file] = "Pictures/".md5(time().rand()).'.jpg';
		return $this->tmpPictures[$file];
	}
	
	public function getStyleByName($name) {
		if(isset($this->styles[$name])) return $this->styles[$name];
		if(isset($this->tmpStyles[$name])) return $this->tmpStyles[$name];
		return null; 
	}
	
	public function addTable(odsTable $odsTable) {
		if(in_array($odsTable,$this->tables)) return;
		$this->tables[$odsTable->getName()] = $odsTable;
	}
	
	public function setDefaultTable(odsTable $odsTable) {
		$this->defaultTable = $odsTable;
	}
	
	private function getDefaultTableName() {
		if($this->defaultTable) {
			return $this->defaultTable->getName();
		} elseif(count($this->tables)){
			$keys = array_keys($this->tables);
			//var_dump($keys);
			return $this->tables[$keys[0]]->getName();
		} else {
			return "feuille1";
		}
	}
	
	public function getContent() {
		$this->tmpStyles = array();
		$this->tmpPictures = array();
		
		$dom = new DOMDocument('1.0', 'UTF-8');
		$root = $dom->createElement('office:document-content');
			$root->setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
			$root->setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
			$root->setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
			$root->setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
			$root->setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
			$root->setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
			$root->setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
			$root->setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
			$root->setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
			$root->setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
			$root->setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0");
			$root->setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
			$root->setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0");
			$root->setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0");
			$root->setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML");
			$root->setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0");
			$root->setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0");
			$root->setAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
			$root->setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer");
			$root->setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc");
			$root->setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events");
			$root->setAttribute("xmlns:xforms", "http://www.w3.org/2002/xforms");
			$root->setAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
			$root->setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
			$root->setAttribute("xmlns:rpt", "http://openoffice.org/2005/report");
			$root->setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
			$root->setAttribute("xmlns:rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#");
			$root->setAttribute("xmlns:field", "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:field:1.0");
			$root->setAttribute("xmlns:formx", "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0");
			$root->setAttribute("office:version", "1.2");
			$dom->appendChild($root);

		// office:scripts
		$root->appendChild($dom->createElement('office:scripts'));
		
		//office:font-face-decls
		$office_font_face_decls =  $dom->createElement('office:font-face-decls');
		$root->appendChild($office_font_face_decls);
		
			foreach($this->fontFaces as $fontFace)
				$office_font_face_decls->appendChild($fontFace->getContent($this,$dom));

		// office:automatic-styles
		$office_automatic_styles =  $dom->createElement('office:automatic-styles');
			$root->appendChild($office_automatic_styles);
			
			foreach($this->styles as $style)
				$office_automatic_styles->appendChild($style->getContent($this,$dom));

		// office:body
		$office_body =  $dom->createElement('office:body');
			$root->appendChild($office_body);
		
			// office:spreadsheet
			$office_spreadsheet = $dom->createElement('office:spreadsheet');
				$office_body->appendChild($office_spreadsheet);

					foreach($this->tables as $table)
						$office_spreadsheet->appendChild($table->getContent($this,$dom));
		
			foreach($this->tmpStyles as $style)
				$office_automatic_styles->appendChild($style->getContent($this,$dom));
		
		return $dom->saveXML();
	}
	
	public function getMeta() {
		$dom = new DOMDocument('1.0', 'UTF-8');
		
		$root = $dom->createElement('office:document-meta');
			$root->setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
			$root->setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
			$root->setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
			$root->setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
			$root->setAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
			$root->setAttribute("office:version", "1.2");
			$dom->appendChild($root);
		
		$meta =  $dom->createElement('office:meta');
			$root->appendChild($meta);
		
		$meta->appendChild($dom->createElement('meta:creation-date',date("Y-m-d\TH:j:s")));
		$meta->appendChild($dom->createElement('meta:generator','ods générator'));
		$meta->appendChild($dom->createElement('dc:date',date("Y-m-d\TH:j:s")));
		$meta->appendChild($dom->createElement('meta:editing-duration','PT1S'));
		$meta->appendChild($dom->createElement('meta:editing-cycles','1'));
		if($this->title)
			$meta->appendChild($dom->createElement('dc:title',$this->title));
		if($this->subject)
			$meta->appendChild($dom->createElement('dc:subject',$this->subject));
		if($this->keyword)
			$meta->appendChild($dom->createElement('meta:keyword',$this->keyword));
		if($this->description)
			$meta->appendChild($dom->createElement('dc:description',$this->description));
		$elm = $dom->createElement('meta:document-statistic');
			$elm->setAttribute("meta:table-count", "1");
			$elm->setAttribute("meta:cell-count", "4");
			$elm->setAttribute("meta:object-count", "0");
			$meta->appendChild($elm);
		
		return $dom->saveXML();
	}
	
	public function getSettings() {
		$dom = new DOMDocument('1.0', 'UTF-8');
		
		$root = $dom->createElement('office:document-settings');
			$root->setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
			$root->setAttribute("xmlns:xlink",  "http://www.w3.org/1999/xlink");
			$root->setAttribute("xmlns:config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0");
			$root->setAttribute("xmlns:ooo",    "http://openoffice.org/2004/office");
			$root->setAttribute("office:version", "1.2");
			$dom->appendChild($root);
		
			$office_settings =  $dom->createElement('office:settings');
				$root->appendChild($office_settings);
			
				$config_config_item_set = $dom->createElement('config:config-item-set');
					$config_config_item_set->setAttribute("config:name", "ooo:view-settings");
					$office_settings->appendChild($config_config_item_set);
				
					$config_config_item = $dom->createElement('config:config-item',0);
						$config_config_item->setAttribute("config:name", "VisibleAreaTop");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);
			
					$config_config_item = $dom->createElement('config:config-item',0);
						$config_config_item->setAttribute("config:name", "VisibleAreaLeft");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item',10018);
						$config_config_item->setAttribute("config:name", "VisibleAreaWidth");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item',2592);
						$config_config_item->setAttribute("config:name", "VisibleAreaHeight");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);
					
					$config_config_item_map_indexed = $dom->createElement('config:config-item-map-indexed');
						$config_config_item_map_indexed->setAttribute("config:name", "Views");
						$config_config_item_set->appendChild($config_config_item_map_indexed);
			
						$config_config_item_map_entry1 = $dom->createElement('config:config-item-map-entry');
							$config_config_item_map_indexed->appendChild($config_config_item_map_entry1);
			
							//<config:config-item config:name="ViewId" config:type="string">View1</config:config-item>
							$config_config_item = $dom->createElement('config:config-item', 'View1');
								$config_config_item->setAttribute("config:name", "ViewId");
								$config_config_item->setAttribute("config:type", "string");
								$config_config_item_map_entry1->appendChild($config_config_item);
								
							//<config:config-item-map-named config:name="Tables">
							$config_config_item_map_named = $dom->createElement('config:config-item-map-named');
								$config_config_item_map_named->setAttribute("config:name", "Tables");
								$config_config_item_map_entry1->appendChild($config_config_item_map_named);
								
								foreach($this->tables as $table)
									$config_config_item_map_named->appendChild($table->getSettings($this,$dom));

							$config_config_item = $dom->createElement('config:config-item', $this->getDefaultTableName());
								$config_config_item->setAttribute("config:name", "ActiveTable");
								$config_config_item->setAttribute("config:type", "string");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '270');
								$config_config_item->setAttribute("config:name", "HorizontalScrollbarWidth");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);


							$config_config_item = $dom->createElement('config:config-item', '0');
								$config_config_item->setAttribute("config:name", "ZoomType");
								$config_config_item->setAttribute("config:type", "short");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '100');
								$config_config_item->setAttribute("config:name", "ZoomValue");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '60');
								$config_config_item->setAttribute("config:name", "PageViewZoomValue");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'false');
								$config_config_item->setAttribute("config:name", "ShowPageBreakPreview");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "ShowZeroValues");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "ShowNotes");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "ShowGrid");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '12632256');
								$config_config_item->setAttribute("config:name", "GridColor");
								$config_config_item->setAttribute("config:type", "long");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "ShowPageBreaks");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);
										
							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "HasColumnRowHeaders");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "HasSheetTabs");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "IsOutlineSymbolsSet");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'false');
								$config_config_item->setAttribute("config:name", "IsSnapToRaster");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'false');
								$config_config_item->setAttribute("config:name", "RasterIsVisible");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '1000');
								$config_config_item->setAttribute("config:name", "RasterResolutionX");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '1000');
								$config_config_item->setAttribute("config:name", "RasterResolutionY");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '1');
								$config_config_item->setAttribute("config:name", "RasterSubdivisionX");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', '1');
								$config_config_item->setAttribute("config:name", "RasterSubdivisionY");
								$config_config_item->setAttribute("config:type", "int");
								$config_config_item_map_entry1->appendChild($config_config_item);

							$config_config_item = $dom->createElement('config:config-item', 'true');
								$config_config_item->setAttribute("config:name", "IsRasterAxisSynchronized");
								$config_config_item->setAttribute("config:type", "boolean");
								$config_config_item_map_entry1->appendChild($config_config_item);

				$config_config_item_set = $dom->createElement('config:config-item-set');
					$config_config_item_set->setAttribute("config:name", "ooo:configuration-settings");
					$office_settings->appendChild($config_config_item_set);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "ShowZeroValues");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "ShowNotes");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "ShowGrid");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '12632256');// test 0
						$config_config_item->setAttribute("config:name", "GridColor");
						$config_config_item->setAttribute("config:type", "long");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "ShowPageBreaks");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '3');
						$config_config_item->setAttribute("config:name", "LinkUpdateMode");
						$config_config_item->setAttribute("config:type", "short");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "HasColumnRowHeaders");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "HasSheetTabs");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "IsOutlineSymbolsSet");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "IsSnapToRaster");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "RasterIsVisible");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '1000');
						$config_config_item->setAttribute("config:name", "RasterResolutionX");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '1000');
						$config_config_item->setAttribute("config:name", "RasterResolutionY");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '1');
						$config_config_item->setAttribute("config:name", "RasterSubdivisionX");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '1');
						$config_config_item->setAttribute("config:name", "RasterSubdivisionY");
						$config_config_item->setAttribute("config:type", "int");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "IsRasterAxisSynchronized");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "AutoCalculate");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item');
						$config_config_item->setAttribute("config:name", "PrinterName");
						$config_config_item->setAttribute("config:type", "string");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item');
						$config_config_item->setAttribute("config:name", "PrinterSetup");
						$config_config_item->setAttribute("config:type", "base64Binary");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "ApplyUserData");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', '0');
						$config_config_item->setAttribute("config:name", "CharacterCompressionType");
						$config_config_item->setAttribute("config:type", "short");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "IsKernAsianPunctuation");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "SaveVersionOnClose");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "UpdateFromTemplate");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'true');
						$config_config_item->setAttribute("config:name", "AllowPrintJobCancel");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "LoadReadonly");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

					$config_config_item = $dom->createElement('config:config-item', 'false');
						$config_config_item->setAttribute("config:name", "IsDocumentShared");
						$config_config_item->setAttribute("config:type", "boolean");
						$config_config_item_set->appendChild($config_config_item);

			return $dom->saveXML();
	}

	public function getStyles() {
		$dom = new DOMDocument('1.0', 'UTF-8');
		
		$root = $dom->createElement('office:document-styles');
			$root->setAttribute("xmlns:office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
			$root->setAttribute("xmlns:style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
			$root->setAttribute("xmlns:text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
			$root->setAttribute("xmlns:table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
			$root->setAttribute("xmlns:draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
			$root->setAttribute("xmlns:fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
			$root->setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
			$root->setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
			$root->setAttribute("xmlns:meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
			$root->setAttribute("xmlns:number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
			$root->setAttribute("xmlns:presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0");
			$root->setAttribute("xmlns:svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
			$root->setAttribute("xmlns:chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0");
			$root->setAttribute("xmlns:dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0");
			$root->setAttribute("xmlns:math", "http://www.w3.org/1998/Math/MathML");
			$root->setAttribute("xmlns:form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0");
			$root->setAttribute("xmlns:script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0");
			$root->setAttribute("xmlns:ooo", "http://openoffice.org/2004/office");
			$root->setAttribute("xmlns:ooow", "http://openoffice.org/2004/writer");
			$root->setAttribute("xmlns:oooc", "http://openoffice.org/2004/calc");
			$root->setAttribute("xmlns:dom", "http://www.w3.org/2001/xml-events");
			$root->setAttribute("xmlns:rpt", "http://openoffice.org/2005/report");
			$root->setAttribute("xmlns:of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");
			$root->setAttribute("xmlns:rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#");
			$root->setAttribute("office:version", "1.2");
			$dom->appendChild($root);
		
		
			$office_font_face_decls = $dom->createElement('office:font-face-decls');
				$root->appendChild($office_font_face_decls);
				
				foreach($this->fontFaces as $fontFace)
					$office_font_face_decls->appendChild($fontFace->getStyles($this,$dom));
			
			$office_styles = $dom->createElement('office:styles');
				$root->appendChild($office_styles);
				
				$style_default_style = $dom->createElement('style:default-style');
					$style_default_style->setAttribute("style:family", "table-cell");
					$office_styles->appendChild($style_default_style);
				
					$style_table_cell_properties = $dom->createElement('style:table-cell-properties');
						$style_table_cell_properties->setAttribute("style:decimal-places", "2");
						$style_default_style->appendChild($style_table_cell_properties);
						
					$style_paragraph_properties = $dom->createElement('style:paragraph-properties');
						$style_paragraph_properties->setAttribute("style:tab-stop-distance", "1.25cm");
						$style_default_style->appendChild($style_paragraph_properties);
				
					$style_text_properties = $dom->createElement('style:text-properties');
						$style_text_properties->setAttribute("style:font-name", "Nimbus Sans L");
						$style_text_properties->setAttribute("fo:language", "fr");
						$style_text_properties->setAttribute("fo:country", "FR");
						$style_text_properties->setAttribute("style:font-name-asian", "Bitstream Vera Sans");
						$style_text_properties->setAttribute("style:language-asian", "zxx");
						$style_text_properties->setAttribute("style:country-asian", "none");
						$style_text_properties->setAttribute("style:font-name-complex", "Bitstream Vera Sans");
						$style_text_properties->setAttribute("style:language-complex", "zxx");
						$style_text_properties->setAttribute("style:country-complex", "none");
						$style_default_style->appendChild($style_text_properties);

				$style_default_style = $dom->createElement('style:default-style');
					$style_default_style->setAttribute("style:family", "graphic");
					$office_styles->appendChild($style_default_style);
				
					$style_graphic_properties = $dom->createElement('style:graphic-properties');
						$style_graphic_properties->setAttribute("draw:shadow-offset-x", "0.3cm");
						$style_graphic_properties->setAttribute("draw:shadow-offset-y", "0.3cm");
						$style_default_style->appendChild($style_graphic_properties);
					
					$style_paragraph_properties = $dom->createElement('style:paragraph-properties');
						$style_paragraph_properties->setAttribute("style:text-autospace", "ideograph-alpha");
						$style_paragraph_properties->setAttribute("style:punctuation-wrap", "simple");
						$style_paragraph_properties->setAttribute("style:line-break", "strict");
						$style_paragraph_properties->setAttribute("style:writing-mode", "page");
						$style_paragraph_properties->setAttribute("style:font-independent-line-spacing", "false");
						$style_default_style->appendChild($style_paragraph_properties);
						
						$style_tab_stops = $dom->createElement('style:tab-stops');
							$style_paragraph_properties->appendChild($style_tab_stops);
							
					$style_text_properties =  $dom->createElement('style:text-properties');
						$style_text_properties->setAttribute("style:use-window-font-color", "true");
						$style_text_properties->setAttribute("fo:font-family", "'Nimbus Roman No9 L'");
						$style_text_properties->setAttribute("style:font-family-generic", "roman");
						$style_text_properties->setAttribute("style:font-pitch", "variable");
						$style_text_properties->setAttribute("fo:font-size", "12pt");
						$style_text_properties->setAttribute("fo:language", "fr");
						$style_text_properties->setAttribute("fo:country", "FR");
						$style_text_properties->setAttribute("style:letter-kerning", "true");
						$style_text_properties->setAttribute("style:font-size-asian", "24pt");
						$style_text_properties->setAttribute("style:language-asian", "zxx");
						$style_text_properties->setAttribute("style:country-asian", "none");
						$style_text_properties->setAttribute("style:font-size-complex", "24pt");
						$style_text_properties->setAttribute("style:language-complex", "zxx");
						$style_text_properties->setAttribute("style:country-complex", "none");
						$style_default_style->appendChild($style_text_properties);

				//<number:number-style style:name="N0">
				$number_number_style = $dom->createElement('number:number-style');
					$number_number_style->setAttribute("style:name", "N0");
					$office_styles->appendChild($number_number_style);

					$number_number = $dom->createElement('number:number');
						$number_number->setAttribute("number:min-integer-digits", "1");
						$number_number_style->appendChild($number_number);				

				$style_style = $dom->createElement('style:style');
					$style_style->setAttribute("style:name", "Default");
					$style_style->setAttribute("style:family", "table-cell");
					$office_styles->appendChild($style_style);
					
				$style_style = $dom->createElement('style:style');
					$style_style->setAttribute("style:name", "Result");
					$style_style->setAttribute("style:family", "table-cell");
					$style_style->setAttribute("style:parent-style-name", "Default");
					$office_styles->appendChild($style_style);
					
					$style_text_properties = $dom->createElement('style:text-properties');
						$style_text_properties->setAttribute("fo:font-style", "italic");
						$style_text_properties->setAttribute("style:text-underline-style", "solid");
						$style_text_properties->setAttribute("style:text-underline-width", "auto");
						$style_text_properties->setAttribute("style:text-underline-color", "font-color");
						$style_text_properties->setAttribute("fo:font-weight", "bold");
						$style_style->appendChild($style_text_properties);

				$style_style = $dom->createElement('style:style');
					$style_style->setAttribute("style:name", "Result2");
					$style_style->setAttribute("style:family", "table-cell");
					$style_style->setAttribute("style:parent-style-name", "Result");
					$style_style->setAttribute("style:data-style-name", "N106");
					$office_styles->appendChild($style_style);
						
				$style_style = $dom->createElement('style:style');
					$style_style->setAttribute("style:name", "Heading");
					$style_style->setAttribute("style:family", "table-cell");
					$style_style->setAttribute("style:parent-style-name", "Default");
					$office_styles->appendChild($style_style);

					$style_table_cell_properties = $dom->createElement('style:table-cell-properties');
						$style_table_cell_properties->setAttribute("style:text-align-source", "fix");
						$style_table_cell_properties->setAttribute("style:repeat-content", "false");
						$style_style->appendChild($style_table_cell_properties);
					
					$style_paragraph_properties = $dom->createElement('style:paragraph-properties');
						$style_paragraph_properties->setAttribute("fo:text-align", "center");
						$style_style->appendChild($style_paragraph_properties);
						
					$style_text_properties = $dom->createElement('style:text-properties');
						$style_text_properties->setAttribute("fo:font-size", "16pt");
						$style_text_properties->setAttribute("fo:font-style", "italic");
						$style_text_properties->setAttribute("fo:font-weight", "bold");
						$style_style->appendChild($style_text_properties);

				$style_style = $dom->createElement('style:style');
					$style_style->setAttribute("style:name", "Heading1");
					$style_style->setAttribute("style:family", "table-cell");
					$style_style->setAttribute("style:parent-style-name", "Heading");
					$office_styles->appendChild($style_style);

					$style_table_cell_properties = $dom->createElement('style:table-cell-properties');
						$style_table_cell_properties->setAttribute("style:rotation-angle", "90");
						$style_style->appendChild($style_table_cell_properties);

			$office_automatic_styles = $dom->createElement('office:automatic-styles');
				$root->appendChild($office_automatic_styles);
				
				$style_page_layout = $dom->createElement('style:page-layout');
					$style_page_layout->setAttribute("style:name", "Mpm1");
					$office_automatic_styles->appendChild($style_page_layout);
				
					$style_page_layout_properties = $dom->createElement('style:page-layout-properties');
						$style_page_layout_properties->setAttribute("style:writing-mode", "lr-tb");
						$style_page_layout->appendChild($style_page_layout_properties);
					
					$style_header_style = $dom->createElement('style:header-style');
						$style_page_layout->appendChild($style_header_style);
			
						$style_header_footer_properties = $dom->createElement('style:header-footer-properties');
							$style_header_footer_properties->setAttribute("fo:min-height", "0.751cm");
							$style_header_footer_properties->setAttribute("fo:margin-left", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-right", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-bottom", "0.25cm");
							$style_header_style->appendChild($style_header_footer_properties);
					
					$style_footer_style = $dom->createElement('style:footer-style');
						$style_page_layout->appendChild($style_footer_style);
			
						$style_header_footer_properties = $dom->createElement('style:header-footer-properties');
							$style_header_footer_properties->setAttribute("fo:min-height", "0.751cm");
							$style_header_footer_properties->setAttribute("fo:margin-left", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-right", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-top", "0.25cm");
							$style_footer_style->appendChild($style_header_footer_properties);
					
				$style_page_layout = $dom->createElement('style:page-layout');
					$style_page_layout->setAttribute("style:name", "Mpm2");
					$office_automatic_styles->appendChild($style_page_layout);

					$style_page_layout_properties = $dom->createElement('style:page-layout-properties');
						$style_page_layout_properties->setAttribute("style:writing-mode", "lr-tb");
						$style_page_layout->appendChild($style_page_layout_properties);
				
					$style_header_style = $dom->createElement('style:header-style');
						$style_page_layout->appendChild($style_header_style);
				
						$style_header_footer_properties = $dom->createElement('style:header-footer-properties');
							$style_header_footer_properties->setAttribute("fo:min-height", "0.751cm");
							$style_header_footer_properties->setAttribute("fo:margin-left", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-right", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-bottom", "0.25cm");
							$style_header_footer_properties->setAttribute("fo:border", "0.088cm solid #000000");
							$style_header_footer_properties->setAttribute("fo:padding", "0.018cm");
							$style_header_footer_properties->setAttribute("fo:background-color", "#c0c0c0");
							$style_header_style->appendChild($style_header_footer_properties);
						
							$style_background_image = $dom->createElement('style:background-image');
								$style_header_footer_properties->appendChild($style_background_image);
				
					$style_footer_style = $dom->createElement('style:footer-style');
						$style_page_layout->appendChild($style_footer_style);
				
						$style_header_footer_properties = $dom->createElement('style:header-footer-properties');
							$style_header_footer_properties->setAttribute("fo:min-height", "0.751cm");
							$style_header_footer_properties->setAttribute("fo:margin-left", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-right", "0cm");
							$style_header_footer_properties->setAttribute("fo:margin-top", "0.25cm");
							$style_header_footer_properties->setAttribute("fo:border", "0.088cm solid #000000");
							$style_header_footer_properties->setAttribute("fo:padding", "0.018cm");
							$style_header_footer_properties->setAttribute("fo:background-color", "#c0c0c0");
							$style_footer_style->appendChild($style_header_footer_properties);
						
							$style_background_image = $dom->createElement('style:background-image');
								$style_header_footer_properties->appendChild($style_background_image);

			$office_master_styles = $dom->createElement('office:master-styles');
				$root->appendChild($office_master_styles);
		
				$style_master_page = $dom->createElement('style:master-page');
					$style_master_page->setAttribute("style:name", "Default");
					$style_master_page->setAttribute("style:page-layout-name", "Mpm1");
					$office_master_styles->appendChild($style_master_page);

					$style_header = $dom->createElement('style:header');
						$style_master_page->appendChild($style_header);

						$text_p = $dom->createElement('text:p');
							$style_header->appendChild($text_p);
						
							$text_sheet_name = $dom->createElement('text:sheet-name', '???');
								$text_p->appendChild($text_sheet_name);

					$style_header_left = $dom->createElement('style:header-left');
						$style_header_left->setAttribute("style:display", "false");
						$style_master_page->appendChild($style_header_left);

					$style_footer = $dom->createElement('style:footer');
						$style_master_page->appendChild($style_footer);
						
						$text_p = $dom->createElement('text:p', "Page");
							$style_footer->appendChild($text_p);
						
							$text_page_number = $dom->createElement('text:page-number', '1');
								$text_p->appendChild($text_page_number);

					$style_footer_left = $dom->createElement('style:footer-left');
						$style_footer_left->setAttribute("style:display", "false");
						$style_master_page->appendChild($style_footer_left);

				$style_master_page = $dom->createElement('style:master-page');
					$style_master_page->setAttribute("style:name", "Report");
					$style_master_page->setAttribute("style:page-layout-name", "Mpm2");
					$office_master_styles->appendChild($style_master_page);

					$style_header = $dom->createElement('style:header');
						$style_master_page->appendChild($style_header);

						$style_region_left = $dom->createElement('style:region-left');
							$style_header->appendChild($style_region_left);
						
							$text_p = $dom->createElement('text:p');
								$style_region_left->appendChild($text_p);
								
								$text_sheet_name = $dom->createElement('text:sheet-name', '???');
									$text_p->appendChild($text_sheet_name);

								$note_text = $dom->createTextNode('(');
									$text_p->appendChild($note_text);

								$text_title = $dom->createElement('text:title', '???');
									$text_p->appendChild($text_title);
									
								$note_text = $dom->createTextNode(')');
									$text_p->appendChild($note_text);

						$style_region_right = $dom->createElement('style:region-right');
							$style_header->appendChild($style_region_right);	
							
							$text_p = $dom->createElement('text:p');
								$style_region_right->appendChild($text_p);
								
								$text_date = $dom->createElement('text:date','31/10/2009');
									$text_date->setAttribute("style:data-style-name", "N2");
									$text_date->setAttribute("text:date-value", "2009-10-31");
									$text_p->appendChild($text_date);
						
								$note_text = $dom->createTextNode(',');
									$text_p->appendChild($note_text);
									
								$text_date = $dom->createElement('text:time','18:09:40');
									$text_p->appendChild($text_date);
					
					$style_header_left = $dom->createElement('style:header-left');
						$style_header_left->setAttribute("style:display", "false");
						$style_master_page->appendChild($style_header_left);

					$style_footer = $dom->createElement('style:footer');
						$style_master_page->appendChild($style_footer);
						
						$text_p = $dom->createElement('text:p', 'Page');
							$style_footer->appendChild($text_p);
							
							$text_page_number = $dom->createElement('text:page-number','1');
								$text_p->appendChild($text_page_number);
								
							$note_text = $dom->createTextNode('/');
								$text_p->appendChild($note_text);
						
							$text_page_count = $dom->createElement('text:page-count','99');
								$text_p->appendChild($text_page_count);
						
					$style_footer_left = $dom->createElement('style:footer-left');
						$style_footer_left->setAttribute("style:display", "false");
						$style_master_page->appendChild($style_footer_left);
				
		
		return $dom->saveXML();
	}
		

	public function getOds() {
		$zip = new ZipArchive();
		$filename = tempnam("tmp", "genodp");
		//var_dump($filename);
		
		if ($zip->open($filename, ZIPARCHIVE::OVERWRITE)!==TRUE) {
		   exit("cannot open $filename\n");
		}
		
		$zip->addFromString("meta.xml", $this->getMeta());
		$zip->addFromString("content.xml", $this->getContent());
		$zip->addFile("files/mimetype","mimetype");
		//$zip->addFile("files/settings.xml","settings.xml");
		$zip->addFromString("settings.xml", $this->getSettings());
		//$zip->addFile("files/styles.xml","styles.xml");
		$zip->addFromString("styles.xml", $this->getStyles());
		$zip->addFile("files/Configurations2/accelerator/current.xml","Configurations2/accelerator/current.xml");
		$zip->addFile("files/META-INF/manifest.xml","META-INF/manifest.xml");
		$zip->addFile("files/Thumbnails/thumbnail.png","Thumbnails/thumbnail.png");
		
		foreach($this->tmpPictures AS $file => $name)
			$zip->addFile($file,$name);
		
		$zip->close();
		
		header('Content-type: application/vnd.oasis.opendocument.spreadsheet');
		header('Content-Disposition: attachment; filename="downloaded.ods"');
		
		readfile($filename);
		unlink($filename);
	}


}
 
?>
