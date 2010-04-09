<?php
define('SPREADSHEET_EXCEL_READER_BIFF8',             0x600);
define('SPREADSHEET_EXCEL_READER_BIFF7',             0x500);
define('SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS',   0x5);
define('SPREADSHEET_EXCEL_READER_WORKSHEET',         0x10);

define('SPREADSHEET_EXCEL_READER_TYPE_BOF',          0x809);
define('SPREADSHEET_EXCEL_READER_TYPE_EOF',          0x0a);
define('SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET',   0x85);
define('SPREADSHEET_EXCEL_READER_TYPE_DIMENSION',    0x200);
define('SPREADSHEET_EXCEL_READER_TYPE_ROW',          0x208);
define('SPREADSHEET_EXCEL_READER_TYPE_DBCELL',       0xd7);
define('SPREADSHEET_EXCEL_READER_TYPE_FILEPASS',     0x2f);
define('SPREADSHEET_EXCEL_READER_TYPE_NOTE',         0x1c);
define('SPREADSHEET_EXCEL_READER_TYPE_TXO',          0x1b6);
define('SPREADSHEET_EXCEL_READER_TYPE_RK',           0x7e);
define('SPREADSHEET_EXCEL_READER_TYPE_RK2',          0x27e);
define('SPREADSHEET_EXCEL_READER_TYPE_MULRK',        0xbd);
define('SPREADSHEET_EXCEL_READER_TYPE_MULBLANK',     0xbe);
define('SPREADSHEET_EXCEL_READER_TYPE_INDEX',        0x20b);
define('SPREADSHEET_EXCEL_READER_TYPE_SST',          0xfc);
define('SPREADSHEET_EXCEL_READER_TYPE_EXTSST',       0xff);
define('SPREADSHEET_EXCEL_READER_TYPE_CONTINUE',     0x3c);
define('SPREADSHEET_EXCEL_READER_TYPE_LABEL',        0x204);
define('SPREADSHEET_EXCEL_READER_TYPE_LABELSST',     0xfd);
define('SPREADSHEET_EXCEL_READER_TYPE_NUMBER',       0x203);
define('SPREADSHEET_EXCEL_READER_TYPE_NAME',         0x18);
define('SPREADSHEET_EXCEL_READER_TYPE_ARRAY',        0x221);
define('SPREADSHEET_EXCEL_READER_TYPE_STRING',       0x207);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA',      0x406);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMULA2',     0x6);
define('SPREADSHEET_EXCEL_READER_TYPE_FORMAT',       0x41e);
define('SPREADSHEET_EXCEL_READER_TYPE_XF',           0xe0);
define('SPREADSHEET_EXCEL_READER_TYPE_BOOLERR',      0x205);
define('SPREADSHEET_EXCEL_READER_TYPE_UNKNOWN',      0xffff);
define('SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR', 0x22);
define('SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS',  0xE5);

define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS' ,    25569);
define('SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904', 24107);
define('SPREADSHEET_EXCEL_READER_MSINADAY',          86400);
define('SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT',    "%s");

require_once('oleread.inc.php');

/**
 * A class for reading Microsoft Excel (old-school) Spreadsheets.
 *
 * Originally developed by Vadim Tkachenko under the name PHPExcelReader.
 * (http://sourceforge.net/projects/phpexcelreader)
 * Based on the Java version by Andy Khan (http://www.andykhan.com).  Now
 * maintained by David Sanders.  Reads only Biff 7 and Biff 8 formats.
 * Edited Object Oriented and PHP5 updated by Jesus Lara <jesuslara@gmail.com>
 *
 * @category   Spreadsheet
 * @package    Spreadsheet_Excel_Reader
 * @author     Jesus Lara <jesuslara@gmail.com>
 * @copyright  2007-2010 Soft Project
 * @license    http://www.php.net/license/3_0.txt  PHP License 3.0
 * @version    Release: @package_version@
 */


class excel_reader Implements Iterator, Countable {

	#iterators
	protected $_current = 0;
	protected $_count = 0;

	/**
	 * Filename
	 * @var string
	 * @access protected
	 */
	protected $_filename = '';

	/**
	 * Array of worksheets found
	 *
	 * @var array
	 * @access public
	 */
	public $boundsheets = array();

	/**
	 * Array of format records found
	 *
	 * @var array
	 * @access public
	 */
	public $formatRecords = array();

	/**
	 * todo
	 *
	 * @var array
	 * @access public
	 */
	public $sst = array();

	/**
	 * Array of worksheets
	 *
	 * The data is stored in 'cells' and the meta-data is stored in an array
	 * called 'cellsInfo'
	 *
	 * Example:
	 *
	 * $sheets  -->  'cells'  -->  row --> column --> Interpreted value
	 *          -->  'cellsInfo' --> row --> column --> 'type' - Can be 'date', 'number', or 'unknown'
	 *                                            --> 'raw' - The raw data that Excel stores for that data cell
	 *
	 * @var array
	 * @access public
	 */
	public $sheets = array();

	/**
	 * The data returned by OLE
	 *
	 * @var string
	 * @access private
	 */
	protected $data;

	/**
	 * The associated column with column names
	 *
	 * @var string
	 * @access private
	 */
	protected $_row = 0;

	/**
	 * The columns names of the sheet
	 *
	 * @var string
	 * @access private
	 */
	protected $_columns;


	/**
	 * OLE object for reading the file
	 *
	 * @var OLE object
	 * @access private
	 */
	protected $_ole;

	/**
	 * Default encoding
	 *
	 * @var string
	 * @access private
	 */
	protected $_defaultEncoding = 'UTF-8';

	/**
	 * Default number format
	 *
	 * @var integer
	 * @access private
	 */
	protected $_defaultFormat = SPREADSHEET_EXCEL_READER_DEF_NUM_FORMAT;

	/**
	 * todo
	 * List of formats to use for each column
	 *
	 * @var array
	 * @access private
	 */
	protected $_columnsFormat = array();

	/**
	 * todo
	 *
	 * @var integer
	 * @access private
	 */
	protected $_rowoffset = 1;

	/**
	 * todo
	 *
	 * @var integer
	 * @access private
	 */
	protected $_coloffset = 1;

	/**
	 * List of default date formats used by Excel
	 *
	 * @var array
	 * @access public
	 */
	public $dateFormats = array (
	0xe => "d/m/Y",
	0xf => "d-M-Y",
	0x10 => "d-M",
	0x11 => "M-Y",
	0x12 => "h:i a",
	0x13 => "h:i:s a",
	0x14 => "H:i",
	0x15 => "H:i:s",
	0x16 => "d/m/Y H:i",
	0x2d => "i:s",
	0x2e => "H:i:s",
	0x2f => "i:s.S");

	/**
	 * Default number formats used by Excel
	 *
	 * @var array
	 * @access public
	 */
	public $numberFormats = array(
	0x1 => "%1.0f",     // "0"
	0x2 => "%1.2f",     // "0.00",
	0x3 => "%1.0f",     //"#,##0",
	0x4 => "%1.2f",     //"#,##0.00",
	0x5 => "%1.0f",     /*"$#,##0;($#,##0)",*/
	0x6 => '$%1.0f',    /*"$#,##0;($#,##0)",*/
	0x7 => '$%1.2f',    //"$#,##0.00;($#,##0.00)",
	0x8 => '$%1.2f',    //"$#,##0.00;($#,##0.00)",
	0x9 => '%1.0f%%',   // "0%"
	0xa => '%1.2f%%',   // "0.00%"
	0xb => '%1.2f',     // 0.00E00",
	0x25 => '%1.0f',    // "#,##0;(#,##0)",
	0x26 => '%1.0f',    //"#,##0;(#,##0)",
	0x27 => '%1.2f',    //"#,##0.00;(#,##0.00)",
	0x28 => '%1.2f',    //"#,##0.00;(#,##0.00)",
	0x29 => '%1.0f',    //"#,##0;(#,##0)",
	0x2a => '$%1.0f',   //"$#,##0;($#,##0)",
	0x2b => '%1.2f',    //"#,##0.00;(#,##0.00)",
	0x2c => '$%1.2f',   //"$#,##0.00;($#,##0.00)",
	0x30 => '%1.0f');   //"##0.0E0";

	public function __construct($filename = '') {
		$this->_ole = new OLERead();
		$this->setUTFEncoder('iconv');
		if ($filename!='') {
			$this->filename($filename);
		}
	}

	/**
	 * Define filename
	 *
	 * @param string $filename
	 * @access public
	 */
	public function filename($filename) {
		if (is_file($filename)) {
			$this->_filename = $filename;
		} else {
			throw new exception("No existe el archivo {$filename}");
		}
	}

	/**
	 * Set the encoding method
	 *
	 * @param string Encoding to use
	 * @access public
	 */
	public function setOutputEncoding($encoding) {
		$this->_defaultEncoding = $encoding;
		return $this;
	}

	/**
	 *  $encoder = 'iconv' or 'mb'
	 *  set iconv if you would like use 'iconv' for encode UTF-16LE to your encoding
	 *  set mb if you would like use 'mb_convert_encoding' for encode UTF-16LE to your encoding
	 *
	 * @access public
	 * @param string Encoding type to use.  Either 'iconv' or 'mb'
	 */
	public function setUTFEncoder($encoder = 'iconv') {
		$this->_encoderFunction = '';
		if ($encoder == 'iconv') {
			$this->_encoderFunction = function_exists('iconv') ? 'iconv' : '';
		} elseif ($encoder == 'mb') {
			$this->_encoderFunction = function_exists('mb_convert_encoding') ? 'mb_convert_encoding' :'';
		}
		return $this;
	}

	/**
	 * set the offset of rows
	 *
	 * @access public
	 * @param offset
	 */
	public function setRowColOffset($iOffset) {
		$this->_rowoffset = $iOffset;
		$this->_coloffset = $iOffset;
		return $this;
	}

	/**
	 * Set the default number format
	 *
	 * @access public
	 * @param Default format
	 */
	public function setDefaultFormat($sFormat) {
		$this->_defaultFormat = $sFormat;
		return $this;
	}

	/**
	 * Force a column to use a certain format
	 *
	 * @access public
	 * @param integer Column number
	 * @param string Format
	 */
	public function setColumnFormat($column, $sFormat) {
		$this->_columnsFormat[$column] = $sFormat;
	}

	/**
	 * Read the spreadsheet file using OLE, then parse
	 *
	 * @access public
	 * @todo return a valid value
	 */
	public function read() {
		if ($this->_filename) {
			$res = $this->_ole->read($this->_filename);
			if ($res) {
				$this->data = $this->_ole->getWorkBook();
				$this->_parse();
				$this->_count = count($this->sheets);
			} else {
				return false;
			}
		} else {
			echo "No ha definido un nombre de archivo";
		}
	}

	/**
	 * retorna el tipo de datos requerido
	 * @access public
	 * @return un array con los datos en bruto de la hoja de excel
	 */
	public function getData() {
		return $this->data;
	}

	/**
	 * define el numero de la columna donde se extraeran los nombres
	 * @access public
	 * @param int $row fila de donde se extraeran los nombres de columna
	 */
	public function setColumnName($row) {
		if(is_integer($row)) {
			$this->_row = $row;
		}
	}

	/**
	 * devuelve las columnas de la hoja solicitada
	 * @access public
	 * @param int $sheet numero de la hoja
	 */
	public function getColumns($sheet) {
		if (isset($this->_columns[$sheet])) {
			return $this->_columns[$sheet];
		}
	}

	/**
	 * Parse a workbook
	 *
	 * @access private
	 * @return bool
	 */
	private function _parse() {
		$pos = 0;
		$code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
		$length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;

		$version = ord($this->data[$pos + 4]) | ord($this->data[$pos + 5])<<8;
		$substreamType = ord($this->data[$pos + 6]) | ord($this->data[$pos + 7])<<8;
		//echo "Start parse code=".base_convert($code,10,16)." version=".base_convert($version,10,16)." substreamType=".base_convert($substreamType,10,16).""."\n";

		if (($version != SPREADSHEET_EXCEL_READER_BIFF8) && ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
			echo 'bad Excel version';
			return false;
		}

		if ($substreamType != SPREADSHEET_EXCEL_READER_WORKBOOKGLOBALS) {
			return false;
		}

		$pos += $length + 4;

		$code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
		$length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;

		while ($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {
			switch ($code) {
				case SPREADSHEET_EXCEL_READER_TYPE_SST:
					//echo "Type_SST\n";
					$spos = $pos + 4;
					$limitpos = $spos + $length;
					$uniqueStrings = $this->_GetInt4d($this->data, $spos+4);
					$spos += 8;
					for ($i = 0; $i < $uniqueStrings; $i++) {
						if ($spos == $limitpos) {
							$opcode = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
							$conlength = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
							if ($opcode != 0x3c) {
								return -1;
							}
							$spos += 4;
							$limitpos = $spos + $conlength;
						}
						$numChars = ord($this->data[$spos]) | (ord($this->data[$spos+1]) << 8);
						//echo "i = $i pos = $pos numChars = $numChars ";
						$spos += 2;
						$optionFlags = ord($this->data[$spos]);
						$spos++;
						$asciiEncoding = (($optionFlags & 0x01) == 0) ;
						$extendedString = ( ($optionFlags & 0x04) != 0);

						// See if string contains formatting information
						$richString = ( ($optionFlags & 0x08) != 0);

						if ($richString) {
							// Read in the crun
							$formattingRuns = ord($this->data[$spos]) | (ord($this->data[$spos+1]) << 8);
							$spos += 2;
						}

						if ($extendedString) {
							// Read in cchExtRst
							$extendedRunLength = $this->_GetInt4d($this->data, $spos);
							$spos += 4;
						}

						$len = ($asciiEncoding)? $numChars : $numChars*2;
						if ($spos + $len < $limitpos) {
							$retstr = substr($this->data, $spos, $len);
							$spos += $len;
						} else {
							// found countinue
							$retstr = substr($this->data, $spos, $limitpos - $spos);
							$bytesRead = $limitpos - $spos;
							$charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
							$spos = $limitpos;
							while ($charsLeft > 0) {
								$opcode = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
								$conlength = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
								if ($opcode != 0x3c) {
									return -1;
								}
								$spos += 4;
								$limitpos = $spos + $conlength;
								$option = ord($this->data[$spos]);
								$spos += 1;
								if ($asciiEncoding && ($option == 0)) {
									$len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($this->data, $spos, $len);
									$charsLeft -= $len;
									$asciiEncoding = true;
								} elseif (!$asciiEncoding && ($option != 0)){
									$len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($this->data, $spos, $len);
									$charsLeft -= $len/2;
									$asciiEncoding = false;
								} elseif (!$asciiEncoding && ($option == 0)) {
									// Bummer - the string starts off as Unicode, but after the
									// continuation it is in straightforward ASCII encoding
									$len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
									for ($j = 0; $j < $len; $j++) {
										$retstr .= $this->data[$spos + $j].chr(0);
									}
									$charsLeft -= $len;
									$asciiEncoding = false;
								} else {
									$newstr = '';
									for ($j = 0; $j < strlen($retstr); $j++) {
										$newstr = $retstr[$j].chr(0);
									}
									$retstr = $newstr;
									$len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($this->data, $spos, $len);
									$charsLeft -= $len/2;
									$asciiEncoding = false;
									//echo "Izavrat\n";
								}
								$spos += $len;
							}
						}
						$retstr = ($asciiEncoding) ? $retstr : $this->_encodeUTF16($retstr);
						if ($richString){
							$spos += 4 * $formattingRuns;
						}

						// For extended strings, skip over the extended string data
						if ($extendedString) {
							$spos += $extendedRunLength;
						}
						$this->sst[]=$retstr;
					}
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
						
					return false;
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_NAME:
					//echo "Type_NAME\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
					$indexCode = ord($this->data[$pos+4]) | ord($this->data[$pos+5]) << 8;
					if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
						$numchars = ord($this->data[$pos+6]) | ord($this->data[$pos+7]) << 8;
						if (ord($this->data[$pos+8]) == 0){
							$formatString = substr($this->data, $pos+9, $numchars);
						} else {
							$formatString = substr($this->data, $pos+9, $numchars*2);
						}
					} else {
						$numchars = ord($this->data[$pos+6]);
						$formatString = substr($this->data, $pos+7, $numchars*2);
					}
					$this->formatRecords[$indexCode] = $formatString;
					//echo $formatString;
					//echo "Type.FORMAT\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_XF:
					$indexCode = ord($this->data[$pos+6]) | ord($this->data[$pos+7]) << 8;
					if (array_key_exists($indexCode, $this->dateFormats)) {
						echo "isdate ".$dateFormats[$indexCode];
						$this->formatRecords['xfrecords'][] = array(
                                    'type' => 'date',
                                    'format' => $this->dateFormats[$indexCode]
						);
					} elseif (array_key_exists($indexCode, $this->numberFormats)) {
						//echo "isnumber ".$this->numberFormats[$indexCode];
						$this->formatRecords['xfrecords'][] = array(
                                    'type' => 'number',
                                    'format' => $this->numberFormats[$indexCode]
						);
					} else {
						$isdate = FALSE;
						if ($indexCode > 0) {
							if (isset($this->formatRecords[$indexCode])) {
								$formatstr = $this->formatRecords[$indexCode];
							}
							if ($formatstr) {
								if (preg_match("/[^hmsday\/\-:\s]/i", $formatstr) == 0) { // found day and time format
									$isdate = TRUE;
									$formatstr = str_replace('mm', 'i', $formatstr);
									$formatstr = str_replace('h', 'H', $formatstr);
									//echo "\ndate-time $formatstr \n";
								}
							}
						}
						if ($isdate) {
							$this->formatRecords['xfrecords'][] = array(
                                        'type' => 'date',
                                        'format' => $formatstr,
							);
						} else {
							$this->formatRecords['xfrecords'][] = array(
                                        'type' => 'other',
                                        'format' => '',
                                        'code' => $indexCode
							);
						}
					}
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR:
					//echo "Type.NINETEENFOUR\n";
					$this->nineteenFour = (ord($this->data[$pos+4]) == 1);
					//var_dump($this->nineteenFour);
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:
					//echo "Type.BOUNDSHEET\n";
					$rec_offset = $this->_GetInt4d($this->data, $pos+4);
					$rec_typeFlag = ord($this->data[$pos+8]);
					$rec_visibilityFlag = ord($this->data[$pos+9]);
					$rec_length = ord($this->data[$pos+10]);
					if ($version == SPREADSHEET_EXCEL_READER_BIFF8){
						$chartype =  ord($this->data[$pos+11]);
						if ($chartype == 0){
							$rec_name    = substr($this->data, $pos+12, $rec_length);
						} else {
							$rec_name    = $this->_encodeUTF16(substr($this->data, $pos+12, $rec_length*2));
						}
					} elseif ($version == SPREADSHEET_EXCEL_READER_BIFF7){
						$rec_name = substr($this->data, $pos+11, $rec_length);
					}
					$this->boundsheets[] = array('name'=>$rec_name, 'offset'=>$rec_offset);
					break;
			}
			$pos += $length + 4;
			$code = ord($this->data[$pos]) | ord($this->data[$pos+1])<<8;
			$length = ord($this->data[$pos+2]) | ord($this->data[$pos+3])<<8;
		}
		foreach ($this->boundsheets as $key=>$val){
			$this->sn = $key;
			$this->_parsesheet($val['offset']);
		}
		return true;
	}

	/**
	 * Parse a worksheet
	 *
	 * @access private
	 * @param todo
	 * @todo fix return codes
	 */
	function _parsesheet($spos) {
		$cont = true;
		// read BOF
		$code = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
		$length = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;

		$version = ord($this->data[$spos + 4]) | ord($this->data[$spos + 5])<<8;
		$substreamType = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8;

		if (($version != SPREADSHEET_EXCEL_READER_BIFF8) && ($version != SPREADSHEET_EXCEL_READER_BIFF7)) {
			return -1;
		}

		if ($substreamType != SPREADSHEET_EXCEL_READER_WORKSHEET){
			return -2;
		}

		//echo "Start parse code=".base_convert($code,10,16)." version=".base_convert($version,10,16)." substreamType=".base_convert($substreamType,10,16).""."\n";
		$spos += $length + 4;
		//var_dump($this->formatRecords);
		//echo "code $code $length";
		while($cont) {
			$lowcode = ord($this->data[$spos]);
			if ($lowcode == SPREADSHEET_EXCEL_READER_TYPE_EOF) {
				break;
			}
			$code = $lowcode | ord($this->data[$spos+1])<<8;
			$length = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
			$spos += 4;
			$this->sheets[$this->sn]['maxrow'] = $this->_rowoffset - 1;
			$this->sheets[$this->sn]['maxcol'] = $this->_coloffset - 1;
			//echo "Code=".base_convert($code,10,16)." $code\n";
			unset($this->rectype);
			$this->multiplier = 1; // need for format with %
			switch ($code) {
				case SPREADSHEET_EXCEL_READER_TYPE_DIMENSION:
					//echo 'Type_DIMENSION ';
					if (!isset($this->numRows)) {
						if (($length == 10) ||  ($version == SPREADSHEET_EXCEL_READER_BIFF7)){
							$this->sheets[$this->sn]['numRows'] = ord($this->data[$spos+2]) | ord($this->data[$spos+3]) << 8;
							$this->sheets[$this->sn]['numCols'] = ord($this->data[$spos+6]) | ord($this->data[$spos+7]) << 8;
						} else {
							$this->sheets[$this->sn]['numRows'] = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
							$this->sheets[$this->sn]['numCols'] = ord($this->data[$spos+10]) | ord($this->data[$spos+11]) << 8;
						}
					}
					//echo 'numRows '.$this->numRows.' '.$this->numCols."\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_MERGEDCELLS:
					$cellRanges = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					for ($i = 0; $i < $cellRanges; $i++) {
						$fr =  ord($this->data[$spos + 8*$i + 2]) | ord($this->data[$spos + 8*$i + 3])<<8;
						$lr =  ord($this->data[$spos + 8*$i + 4]) | ord($this->data[$spos + 8*$i + 5])<<8;
						$fc =  ord($this->data[$spos + 8*$i + 6]) | ord($this->data[$spos + 8*$i + 7])<<8;
						$lc =  ord($this->data[$spos + 8*$i + 8]) | ord($this->data[$spos + 8*$i + 9])<<8;
						//$this->sheets[$this->sn]['mergedCells'][] = array($fr + 1, $fc + 1, $lr + 1, $lc + 1);
						if ($lr - $fr > 0) {
							$this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['rowspan'] = $lr - $fr + 1;
						}
						if ($lc - $fc > 0) {
							$this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['colspan'] = $lc - $fc + 1;
						}
					}
					//echo "Merged Cells $cellRanges $lr $fr $lc $fc\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_RK:
				case SPREADSHEET_EXCEL_READER_TYPE_RK2:
					//echo 'SPREADSHEET_EXCEL_READER_TYPE_RK'."\n";
					$row = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$rknum = $this->_GetInt4d($this->data, $spos + 6);
					$numValue = $this->_GetIEEE754($rknum);
					//echo $numValue." ";
					if ($this->isDate($spos)) {
						list($string, $raw) = $this->createDate($numValue);
					}else{
						$raw = $numValue;
						if (isset($this->_columnsFormat[$column + 1])){
							$this->curformat = $this->_columnsFormat[$column + 1];
						}
						$string = sprintf($this->curformat, $numValue * $this->multiplier);
						//$this->addcell(RKRecord($r));
					}
					$this->addcell($row, $column, $string, $raw);
					//echo "Type_RK $row $column $string $raw {$this->curformat}\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_LABELSST:
					$row        = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column     = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$xfindex    = ord($this->data[$spos+4]) | ord($this->data[$spos+5])<<8;
					$index  = $this->_GetInt4d($this->data, $spos + 6);
					//var_dump($this->sst);
					$this->addcell($row, $column, $this->sst[$index]);
					//echo "LabelSST $row $column $string\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_MULRK:
					$row        = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$colFirst   = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$colLast    = ord($this->data[$spos + $length - 2]) | ord($this->data[$spos + $length - 1])<<8;
					$columns    = $colLast - $colFirst + 1;
					$tmppos = $spos+4;
					for ($i = 0; $i < $columns; $i++) {
						$numValue = $this->_GetIEEE754($this->_GetInt4d($this->data, $tmppos + 2));
						//echo $numValue;
						if ($this->isDate($tmppos-4)) {
							list($string, $raw) = $this->createDate($numValue);
						}else{
							$raw = $numValue;
							if (isset($this->_columnsFormat[$colFirst + $i + 1])){
								$this->curformat = $this->_columnsFormat[$colFirst + $i + 1];
							}
							$string = sprintf($this->curformat, $numValue * $this->multiplier);
						}
						//$rec['rknumbers'][$i]['xfindex'] = ord($rec['data'][$pos]) | ord($rec['data'][$pos+1]) << 8;
						$tmppos += 6;
						$this->addcell($row, $colFirst + $i, $string, $raw);
						//echo "MULRK $row ".($colFirst + $i)." $string\n";
					}
					//MulRKRecord($r);
					// Get the individual cell records from the multiple record
					//$num = ;

					break;
				case SPREADSHEET_EXCEL_READER_TYPE_NUMBER:
					$row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$tmp = unpack("ddouble", substr($this->data, $spos + 6, 8)); // It machine machine dependent
					//var_dump($row);
					if ($this->isDate($spos)) {
						list($string, $raw) = $this->createDate($tmp['double']);
						//   $this->addcell(DateRecord($r, 1));
					}else{
						//$raw = $tmp[''];
						if (isset($this->_columnsFormat[$column + 1])) {
							$this->curformat = $this->_columnsFormat[$column + 1];
						}
						$raw = $this->createNumber($spos);
						$string = sprintf($this->curformat, $raw * $this->multiplier);
						//   $this->addcell(NumberRecord($r));
					}
					$this->addcell($row, $column, $string, $raw);
					//echo "Number $row $column $string\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_FORMULA:
				case SPREADSHEET_EXCEL_READER_TYPE_FORMULA2:
					$row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					if ((ord($this->data[$spos+6])==0) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
						//String formula. Result follows in a STRING record
						//echo "FORMULA $row $column Formula with a string<br>\n";
						$this->addcell($row, $column, '', $raw);
					} elseif ((ord($this->data[$spos+6])==1) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
						//Boolean formula. Result is in +2; 0=false,1=true
					} elseif ((ord($this->data[$spos+6])==2) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
						//Error formula. Error code is in +2;
					} elseif ((ord($this->data[$spos+6])==3) && (ord($this->data[$spos+12])==255) && (ord($this->data[$spos+13])==255)) {
						//Formula result is a null string.
					} else {
						// result is a number, so first 14 bytes are just like a _NUMBER record
						$tmp = unpack("ddouble", substr($this->data, $spos + 6, 8)); // It machine machine dependent
						if ($this->isDate($spos)) {
							list($string, $raw) = $this->createDate($tmp['double']);
							//   $this->addcell(DateRecord($r, 1));
						}else {
							//$raw = $tmp[''];
							if (isset($this->_columnsFormat[$column + 1])){
								$this->curformat = $this->_columnsFormat[$column + 1];
							}
							$raw = $this->createNumber($spos);
							$string = sprintf($this->curformat, $raw * $this->multiplier);
							//$this->addcell(NumberRecord($r));
						}
						$this->addcell($row, $column, $string, $raw);
						//echo "Number $row $column $string\n";
					}
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_BOOLERR:
					$row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$string = ord($this->data[$spos+6]);
					$this->addcell($row, $column, $string);
					//echo 'Type_BOOLERR '."\n";
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_ROW:
				case SPREADSHEET_EXCEL_READER_TYPE_DBCELL:
				case SPREADSHEET_EXCEL_READER_TYPE_MULBLANK:
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_LABEL:
					$row    = ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$column = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					//var_dump((ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8)+1);
					//$n = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8;
					//var_dump(utf8_encode(mb_convert_encoding(substr($this->data, $spos + 8, $n+1), 'UTF-8')));
					$this->addcell($row, $column, $this->_extract_data($spos));
					break;
				case SPREADSHEET_EXCEL_READER_TYPE_EOF:
					$cont = false;
					break;
				default:
					//echo ' unknown :'.base_convert($r['code'],10,16)."\n";
					break;

			}
			$spos += $length;
		}
		if (!isset($this->sheets[$this->sn]['numRows'])) {
			$this->sheets[$this->sn]['numRows'] = $this->sheets[$this->sn]['maxrow'];
		}
		if (!isset($this->sheets[$this->sn]['numCols'])) {
			$this->sheets[$this->sn]['numCols'] = $this->sheets[$this->sn]['maxcol'];
		}
		#obtengo el nombre de las columnas:
		$this->_get_columns($this->sn);
	}


	/**
	 * obtengo el nombre de las columnas de la celda actual
	 * @access private
	 * @param $sheet numero de la hoja actual
	 *
	 */
	private function _get_columns($sheet) {
		if ($this->_row) {
			$this->_columns[$sheet] = array();
			$cols = $this->sheets[$this->sn]['numCols'];
			for($col = 0; $col < $cols; $col++) {
				$n = $col + $this->_coloffset;
				$column = $this->sheets[$sheet]['cells'][$this->_row][$n];
				$this->_columns[$sheet][$column] = $n;
			}
		}
	}

	/**
	 * extract data from row
	 *
	 * @param todo
	 * @return data from row & column
	 */
	private function _extract_data($spos) {
		$n = ord($this->data[$spos + 6]) | ord($this->data[$spos + 7])<<8;
		return substr($this->data, $spos + 8, $n+1);
	}

	/**
	 * Check whether the current record read is a date
	 *
	 * @param todo
	 * @return boolean True if date, false otherwise
	 */
	function isDate($spos) {
		$xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
		//echo 'check is date '.$xfindex.' '.$this->formatRecords['xfrecords'][$xfindex]['type']."\n";
		//var_dump($this->formatRecords['xfrecords'][$xfindex]);
		if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'date') {
			$this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
			$this->rectype = 'date';
			return true;
		} else {
			if ($this->formatRecords['xfrecords'][$xfindex]['type'] == 'number') {
				$this->curformat = $this->formatRecords['xfrecords'][$xfindex]['format'];
				$this->rectype = 'number';
				if (($xfindex == 0x9) || ($xfindex == 0xa)){
					$this->multiplier = 100;
				}
			}else{
				$this->curformat = $this->_defaultFormat;
				$this->rectype = 'unknown';
			}
			return false;
		}
	}

	/**
	 * Convert the raw Excel date into a human readable format
	 *
	 * Dates in Excel are stored as number of seconds from an epoch.  On
	 * Windows, the epoch is 30/12/1899 and on Mac it's 01/01/1904
	 *
	 * @access private
	 * @param integer The raw Excel value to convert
	 * @return array First element is the converted date, the second element is number a unix timestamp
	 */
	function createDate($numValue) {
		if ($numValue > 1) {
			$utcDays = $numValue - ($this->nineteenFour ? SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS1904 : SPREADSHEET_EXCEL_READER_UTCOFFSETDAYS);
			$utcValue = round(($utcDays+1) * SPREADSHEET_EXCEL_READER_MSINADAY);
			$string = date ($this->curformat, $utcValue);
			$raw = $utcValue;
		} else {
			$raw = $numValue;
			$hours = floor($numValue * 24);
			$mins = floor($numValue * 24 * 60) - $hours * 60;
			$secs = floor($numValue * SPREADSHEET_EXCEL_READER_MSINADAY) - $hours * 60 * 60 - $mins * 60;
			$string = date ($this->curformat, mktime($hours, $mins, $secs));
		}

		return array($string, $raw);
	}

	function createNumber($spos) {
		$rknumhigh = $this->_GetInt4d($this->data, $spos + 10);
		$rknumlow = $this->_GetInt4d($this->data, $spos + 6);
		//for ($i=0; $i<8; $i++) { echo ord($this->data[$i+$spos+6]) . " "; } echo "<br>";
		$sign = ($rknumhigh & 0x80000000) >> 31;
		$exp =  ($rknumhigh & 0x7ff00000) >> 20;
		$mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
		$mantissalow1 = ($rknumlow & 0x80000000) >> 31;
		$mantissalow2 = ($rknumlow & 0x7fffffff);
		$value = $mantissa / pow( 2 , (20- ($exp - 1023)));
		if ($mantissalow1 != 0) $value += 1 / pow (2 , (21 - ($exp - 1023)));
		$value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));
		//echo "Sign = $sign, Exp = $exp, mantissahighx = $mantissa, mantissalow1 = $mantissalow1, mantissalow2 = $mantissalow2<br>\n";
		if ($sign) {
			$value = -1 * $value;
		}
		return  $value;
	}

	function addcell($row, $col, $string, $raw = '') {
		//echo "ADD cel {$row}-{$col} {$string}\n";
		$string =  mb_convert_encoding(trim($string), 'UTF-8');
		//$raw = $string;
		$this->sheets[$this->sn]['maxrow'] = max($this->sheets[$this->sn]['maxrow'], $row + $this->_rowoffset);
		$this->sheets[$this->sn]['maxcol'] = max($this->sheets[$this->sn]['maxcol'], $col + $this->_coloffset);
		$this->sheets[$this->sn]['cells'][$row + $this->_rowoffset][$col + $this->_coloffset] = $string;
		if ($raw) {
			$this->sheets[$this->sn]['cellsInfo'][$row + $this->_rowoffset][$col + $this->_coloffset]['raw'] = $raw;
		}
		if (isset($this->rectype)) {
			$this->sheets[$this->sn]['cellsInfo'][$row + $this->_rowoffset][$col + $this->_coloffset]['type'] = $this->rectype;
		}

	}


	private function _GetIEEE754($rknum) {
		if (($rknum & 0x02) != 0) {
			$value = $rknum >> 2;
		} else {
			//mmp
			// first comment out the previously existing 7 lines of code here
			//                $tmp = unpack("d", pack("VV", 0, ($rknum & 0xfffffffc)));
			//                //$value = $tmp[''];
			//                if (array_key_exists(1, $tmp)) {
			//                    $value = $tmp[1];
			//                } else {
			//                    $value = $tmp[''];
			//                }
			// I got my info on IEEE754 encoding from
			// http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
			// The RK format calls for using only the most significant 30 bits of the
			// 64 bit floating point value. The other 34 bits are assumed to be 0
			// So, we use the upper 30 bits of $rknum as follows...
			$sign = ($rknum & 0x80000000) >> 31;
			$exp = ($rknum & 0x7ff00000) >> 20;
			$mantissa = (0x100000 | ($rknum & 0x000ffffc));
			$value = $mantissa / pow(2 ,(20-($exp - 1023)));
			if ($sign) {
				$value = -1 * $value;
			}
			//end of changes by mmp
		}
		if (($rknum & 0x01) != 0) {
			$value /= 100;
		}
		return $value;
	}

	private function _encodeUTF16($string) {
		$result = $string;
		if ($this->_defaultEncoding) {
			switch ($this->_encoderFunction) {
				case 'iconv' :
					$result = iconv('UTF-16LE', $this->_defaultEncoding, $string);
					break;
				case 'mb_convert_encoding' :
					$result = mb_convert_encoding($string, $this->_defaultEncoding, 'UTF-16LE' );
					break;
			}
		}
		return $result;
	}

	private function _GetInt4d($data, $pos) {
		$_or_24 = ord($data[$pos+3]);
		$_or_24 >=128 ? $_ord_24 = -abs((256-$_or_24) << 24) : $_ord_24 = ($_or_24&127) << 24;
		return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $_ord_24;
	}

	// --- magical methods
	public function __call($name, array $args) {
		switch($name) {
			case 'columns':
				return $this->getColumns($this->_current);
				break;
			case 'numRows':
				return $this->sheets[$this->_current]['numRows'];
				break;
			case 'numCols':
				return $this->sheets[$this->_current]['numCols'];
				break;
			case 'rows':
				if ($args) {
					return $this->sheets[$this->_current]['cells'][$args[0]];
				} else {
					return $this->sheets[$this->_current]['cells'];
				}
		}
	}

	public function first() {
		$this->_current = 0;
		return $this->sheets[0];
	}

	public function last() {
		$this->_current = $this->_count - 1;
		return $this->sheets[$this->_current];
	}

	public function rewind() {
		return $this->first();
	}

	public function next() {
		$s = $this->_sheets[$this->_current];
		$this->_current++;
	}

	public function valid() {
		if ($this->_current <= count($this->_count)) {
			return true;
		} else {
			return false;
		}
	}

	public function current() {
		return $this->sheets[$this->_current];
	}

	public function count() {
		return $this->_count;
	}

	public function key() {
		return $this->_current;
	}

	public function value() {
		return $this->current();
	}
}
?>