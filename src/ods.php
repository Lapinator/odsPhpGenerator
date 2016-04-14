<?php
/*-
 * Copyright (c) 2009 Laurent VUIBERT
 * License : GNU Lesser General Public License v3
 */

namespace  odsPhpGenerator;

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
		$this->title         = null;
		$this->subject       = null;
		$this->keyword       = null;
		$this->description   = null;
		$this->path2OdsFiles = ".";
		
		$this->defaultTable = null;
		
		$this->scripts   = array();
		$this->fontFaces = array();
		$this->styles    = array();
		$this->tables    = array();
		
		$this->addFontFaces( new odsFontFace( "Nimbus Sans L", "swiss" ) );
		$this->addFontFaces( new odsFontFace( "Bitstream Vera Sans", "system" ) );
		
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
	
	// Deprecated
	public function setPath2OdsFiles($path) {}
	
	public function addFontFaces(odsFontFace $odsFontFace) {
		if(in_array($odsFontFace,$this->fontFaces)) return;
		$this->fontFaces[$odsFontFace->getFontName()] = $odsFontFace;
	}
	
	public function addStyles(odsStyle $odsStyle) {
		if(in_array($odsStyle,$this->styles)) return;
		$this->styles[$odsStyle->getName()] = $odsStyle;
	}
	
	public function addTmpStyles(odsStyle $odsStyle) {
		if(in_array($odsStyle,$this->styles)) return;
		if(in_array($odsStyle,$this->tmpStyles)) return;
		$this->tmpStyles[$odsStyle->getName()] = $odsStyle;
		//echo "addTmpStyles:".$odsStyle->getName()."\n";
	}
	
	public function addTmpPictures($file) {
		if(in_array($file,$this->tmpPictures)) return  $this->tmpPictures[$file];
		$this->tmpPictures[$file] = "Pictures/".md5(time().rand()).'.png';
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
		
		$dom = new \DOMDocument('1.0', 'UTF-8');
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
		
			// the $this->tmpStyle can change in for ( add new elemements only )
			for($i=0; $i<count($this->tmpStyles); $i++) {
				$keys = array_keys($this->tmpStyles);
				$style = $this->tmpStyles[$keys[$i]];
				//echo "createTmpStyle:".$style->getName()."\n";
				$office_automatic_styles->appendChild($style->getContent($this,$dom));
			}
		
		return $dom->saveXML();
	}
	
	public function getMeta() {

		$dom = new \DOMDocument('1.0', 'UTF-8');
		
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
		$dom = new \DOMDocument('1.0', 'UTF-8');
		
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
		$dom = new \DOMDocument('1.0', 'UTF-8');
		
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
	
	private function getMimeType() {
		return "application/vnd.oasis.opendocument.spreadsheet";
	}
	
	private function getAcceleratorCurrent() {
		return "";
	}
	
	private function getManifest() {
		$dom = new \DOMDocument('1.0', 'UTF-8');
		$root = $dom->createElement('manifest:manifest');
			$root->setAttribute("xmlns:manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
			$root->setAttribute("manifest:version","1.2");
			$dom->appendChild($root);
		
			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "application/vnd.oasis.opendocument.spreadsheet");
				$manifest_file_entry->setAttribute("manifest:version", "1.2");
				$manifest_file_entry->setAttribute("manifest:full-path", "/");
				$root->appendChild($manifest_file_entry);

			foreach ($this->tmpPictures AS $pictures) {
				$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:full-path", $pictures);
				$manifest_file_entry->setAttribute("manifest:media-type", "image/png");
				$root->appendChild($manifest_file_entry);
			}
			
			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "text/xml");
				$manifest_file_entry->setAttribute("manifest:full-path", "content.xml");
				$root->appendChild($manifest_file_entry);

			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "text/xml");
				$manifest_file_entry->setAttribute("manifest:full-path", "styles.xml");
				$root->appendChild($manifest_file_entry);
			
			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "text/xml");
				$manifest_file_entry->setAttribute("manifest:full-path", "meta.xml");
				$root->appendChild($manifest_file_entry);

			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "");
				$manifest_file_entry->setAttribute("manifest:full-path", "Thumbnails/thumbnail.png");
				$root->appendChild($manifest_file_entry);

			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "");
				$manifest_file_entry->setAttribute("manifest:full-path", "Configurations2/accelerator/current.xml");
				$root->appendChild($manifest_file_entry);

			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "application/vnd.sun.xml.ui.configuration");
				$manifest_file_entry->setAttribute("manifest:full-path", "Configurations2/");
				$root->appendChild($manifest_file_entry);

			$manifest_file_entry = $dom->createElement("manifest:file-entry");
				$manifest_file_entry->setAttribute("manifest:media-type", "text/xml");
				$manifest_file_entry->setAttribute("manifest:full-path", "settings.xml");
				$root->appendChild($manifest_file_entry);

		return $dom->saveXML();
	}
	
	private function getThumbnail() {
		return base64_decode("
			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");
	}
	
	public function genOdsFile($file) {
		$zip = new \ZipArchive();
		
		if ($zip->open($file, \ZipArchive::OVERWRITE)!==TRUE) {
		   exit("cannot open $file\n");
		}

		$zip->addFromString("mimetype", $this->getMimeType());
		$zip->addFromString("meta.xml", $this->getMeta());
		$zip->addFromString("content.xml", $this->getContent());
		$zip->addFromString("settings.xml", $this->getSettings());
		$zip->addFromString("styles.xml", $this->getStyles());
		$zip->addFromString("Configurations2/accelerator/current.xml", $this->getAcceleratorCurrent());
		$zip->addFromString("META-INF/manifest.xml", $this->getManifest());
		$zip->addFromString("Thumbnails/thumbnail.png", $this->getThumbnail());

		//$zip->setCompressionIndex(0, \ZipArchive::CM_STORE);

		foreach($this->tmpPictures AS $imgfile => $name)
			$zip->addFile($imgfile,$name);
		
		$zip->close();
	}
	
	public function downloadOdsFile($fileName) {
		header('Content-type: application/vnd.oasis.opendocument.spreadsheet');
		header('Content-Disposition: attachment; filename="'.$fileName.'"');
		$tmpfile = tempnam("tmp", "genods");
		$this->genOdsFile($tmpfile);
		readfile($tmpfile);
		unlink($tmpfile);
	}

}
 
?>
