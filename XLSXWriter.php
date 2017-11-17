<?php

/*
 * @license MIT License
 * */

class XLSXWriter {

    //http://www.ecma-international.org/publications/standards/Ecma-376.htm
    //http://officeopenxml.com/SSstyles.php
    //------------------------------------------------------------------
    //http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
    const EXCEL_2007_MAX_ROW = 1048576;
    const EXCEL_2007_MAX_COL = 16384;

    //------------------------------------------------------------------
    protected $author = 'Doc Author';
    protected $sheets = array();
    protected $temp_files = array();
    protected $cell_styles = array();
    protected $number_formats = array();
    protected $current_sheet = '';

    public function __construct() {
        if (!ini_get('date.timezone')) {
            //using date functions can kick out warning if this isn't set
            date_default_timezone_set('UTC');
        }
    }

    public function setAuthor($author = '') {
        $this->author = $author;
    }

    public function setTempDir($tempdir = '') {
        $this->tempdir = $tempdir;
    }

    public function __destruct() {
        if (!empty($this->temp_files)) {
            foreach ($this->temp_files as $temp_file) {
                @unlink($temp_file);
            }
        }
    }

    protected function tempFilename() {
        $tempdir = !empty($this->tempdir) ? $this->tempdir : sys_get_temp_dir();
        $filename = tempnam($tempdir, "xlsx_writer_");
        $this->temp_files[] = $filename;
        return $filename;
    }

    public function writeToStdOut() {
        $temp_file = $this->tempFilename();
        $this->writeToFile($temp_file);
        $file = readfile($temp_file);
        $this->removeFile($temp_file);
        return $file;
    }

    public function writeToString() {
        $temp_file = $this->tempFilename();
        $this->writeToFile($temp_file);
        $string = file_get_contents($temp_file);
        $this->removeFile($temp_file);
        return $string;
    }

    public function writeToFile($filename) {
        foreach ($this->sheets as $sheet_name => $sheet) {
            self::finalizeSheet($sheet_name); //making sure all footers have been written
        }

        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename); //if the zip already exists, remove it
            } else {
                self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
                return;
            }
        }
        $zip = new ZipArchive();
        if (empty($this->sheets)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", no worksheets defined.");
            return;
        }
        if (!$zip->open($filename, ZipArchive::CREATE)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
            return;
        }

        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", self::buildAppXML());
        $zip->addFromString("docProps/core.xml", self::buildCoreXML());

        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            $zip->addFile($sheet->filename, "xl/worksheets/" . $sheet->xmlname);
        }
        $zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString("[Content_Types].xml", self::buildContentTypesXML());

        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML());
        $zip->close();
    }

    public function removeFile($filename) {
        $removed = true;
        if (file_exists($filename)) {
            $removed = @unlink($filename);
            if (!$removed) {
                self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not removed.");
            }
        }
        return $removed;
    }

    private $dafault_column_width = 12;

    protected function initializeSheet($sheet_name, $col_widths = array()) {
        //if already initialized
        if ($this->current_sheet == $sheet_name || isset($this->sheets[$sheet_name]))
            return;

        $sheet_filename = $this->tempFilename();
        $sheet_xmlname = 'sheet' . (count($this->sheets) + 1) . ".xml";
        $this->sheets[$sheet_name] = (object) array(
                    'filename' => $sheet_filename,
                    'sheetname' => $sheet_name,
                    'xmlname' => $sheet_xmlname,
                    'row_count' => 0,
                    'file_writer' => new XLSXWriter_BuffererWriter($sheet_filename),
                    'columns' => array(),
                    'merge_cells' => array(),
                    'max_cell_tag_start' => 0,
                    'max_cell_tag_end' => 0,
                    'finalized' => false,
        );
        $sheet = &$this->sheets[$sheet_name];
        $tabselected = count($this->sheets) == 1 ? 'true' : 'false'; //only first sheet is selected
        $max_cell = XLSXWriter::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL); //XFE1048577
        $sheet->file_writer->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $sheet->file_writer->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
//        $sheet->file_writer->write('<sheetPr filterMode="false">');
//        $sheet->file_writer->write('<pageSetUpPr fitToPage="false"/>');
//        $sheet->file_writer->write('</sheetPr>');
        $sheet->max_cell_tag_start = $sheet->file_writer->ftell();
        $sheet->file_writer->write('<dimension ref="A1:' . $max_cell . '"/>');
        $sheet->file_writer->write("\n");
        $sheet->max_cell_tag_end = $sheet->file_writer->ftell();
        $sheet->file_writer->write('<sheetViews>');
        $sheet->file_writer->write('<sheetView tabSelected="' . $tabselected . '" workbookViewId="0">');
        //$sheet->file_writer->write('<sheetView defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
        $sheet->file_writer->write('<selection activeCell="A1" sqref="A1"/>');
        $sheet->file_writer->write('</sheetView>');
        $sheet->file_writer->write('</sheetViews>');
        $sheet->file_writer->write("\n");

        $sheet->file_writer->write('<cols>');
        $i = 0;
        if (!empty($col_widths)) {
            foreach ($col_widths as $column_width) {
                $sheet->file_writer->write('<col collapsed="false" hidden="false" max="' . ($i + 1) . '" min="' . ($i + 1) . '" style="0" width="' . floatval($column_width) . '"/>');
                $i++;
            }
        }
        $sheet->file_writer->write('<col collapsed="false" hidden="false" max="1024" min="' . ($i + 1) . '" style="0" width="' . floatval($this->dafault_column_width) . '"/>');
        $sheet->file_writer->write('</cols>');
        $sheet->file_writer->write("\n");
        $sheet->file_writer->write('<sheetData>');
    }

    private function addCellStyle($number_format, $cell_style_string) {
        $number_format_idx = self::add_to_list_get_index($this->number_formats, $number_format);
        $lookup_string = $number_format_idx . ";" . $cell_style_string;
        $cell_style_idx = self::add_to_list_get_index($this->cell_styles, $lookup_string);
        return $cell_style_idx;
    }

    private function initializeColumnTypes($header_types) {
        $column_types = array();
        foreach ($header_types as $v) {
            $number_format = self::numberFormatStandardized($v);
            $number_format_type = self::determineNumberFormatType($number_format);
            $cell_style_idx = $this->addCellStyle($number_format, $style_string = null);
            $column_types[] = array(
                'number_format' => $number_format, //contains excel format like 'YYYY-MM-DD HH:MM:SS'
                'number_format_type' => $number_format_type, //contains friendly format like 'datetime'
                'default_cell_style' => $cell_style_idx,
            );
        }
        return $column_types;
    }

    public function writeSheetHeader($sheet_name, array $header_types, $col_options = null) {
        if (empty($sheet_name) || empty($header_types) || !empty($this->sheets[$sheet_name])) {
            return;
        }
        $suppress_row = isset($col_options['suppress_row']) ? intval($col_options['suppress_row']) : false;
        $style = &$col_options;

        $column_widths = array();
        for ($i = 0; $i < count($header_types); $i++) {
            $column_width = isset($col_options[0]) ? floatval(@$col_options[$i]['width']) : floatval(@$col_options['width']);
            if (!($column_width > 0)) {
                $column_width = floatval($this->dafault_column_width);
            }
            $column_widths[] = $column_width;
        }
        self::initializeSheet($sheet_name, $column_widths);
        $sheet = &$this->sheets[$sheet_name];
        $sheet->columns = $this->initializeColumnTypes($header_types);
        $base_style_index = 1; //for index //1 placeholders for static xml later
        if (!$suppress_row) {
            $header_row = array_keys($header_types);
            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
            foreach ($header_row as $c => $v) {
                $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle('GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style));
                $cell_style_idx = $cell_style_idx + $base_style_index; //fix index
                $this->writeCell($sheet->file_writer, 0, $c, $v, $number_format_type = 'n_string', $cell_style_idx);
            }
            $sheet->file_writer->write('</row>');
            $sheet->file_writer->write("\n");
            $sheet->row_count++;
        }
        $this->current_sheet = $sheet_name;
    }

    public function writeSheetRow($sheet_name, array $row, $row_options = null) {
        if (empty($sheet_name)) {
            return;
        }

        self::initializeSheet($sheet_name);
        $sheet = &$this->sheets[$sheet_name];
        if (count($sheet->columns) < count($row)) {
            $default_column_types = $this->initializeColumnTypes(array_fill($from = 0, $until = count($row), 'GENERAL')); //will map to n_auto
            $sheet->columns = array_merge((array) $sheet->columns, $default_column_types);
        }

        if (!empty($row_options)) {
            $ht = isset($row_options['height']) ? floatval($row_options['height']) : 12.1;
            $customHt = isset($row_options['height']) ? true : false;
            $hidden = isset($row_options['hidden']) ? boolval($row_options['hidden']) : false;
            $collapsed = isset($row_options['collapsed']) ? boolval($row_options['collapsed']) : false;
            $sheet->file_writer->write('<row collapsed="' . ($collapsed) . '" customFormat="false" customHeight="' . ($customHt) . '" hidden="' . ($hidden) . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        } else {
            $sheet->file_writer->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->row_count + 1) . '">');
        }

        $style = &$row_options;
        $base_style_index = 1; //for index //1 placeholders for static xml later
        $c = 0;
        foreach ($row as $v) {
            $number_format = $sheet->columns[$c]['number_format'];
            $number_format_type = $sheet->columns[$c]['number_format_type'];
            $cell_style_idx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle($number_format, json_encode(isset($style[0]) ? $style[$c] : $style));
            $cell_style_idx = $cell_style_idx + $base_style_index; //fix index
            $this->writeCell($sheet->file_writer, $sheet->row_count, $c, $v, $number_format_type, $cell_style_idx);
            $c++;
        }
        $sheet->file_writer->write('</row>');
        $sheet->file_writer->write("\n");
        $sheet->row_count++;
        $this->current_sheet = $sheet_name;
    }

    public function countSheetRows($sheet_name = '') {
        $sheet_name = $sheet_name ?: $this->current_sheet;
        return array_key_exists($sheet_name, $this->sheets) ? $this->sheets[$sheet_name]->row_count : 0;
    }

    protected function finalizeSheet($sheet_name) {
        if (empty($sheet_name) || $this->sheets[$sheet_name]->finalized)
            return;

        $sheet = &$this->sheets[$sheet_name];

        $sheet->file_writer->write('</sheetData>');

        if (!empty($sheet->merge_cells)) {
            $sheet->file_writer->write('<mergeCells>');
            foreach ($sheet->merge_cells as $range) {
                $sheet->file_writer->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->file_writer->write('</mergeCells>');
        }

        $sheet->file_writer->write('</worksheet>');

        $max_cell = self::xlsCell($sheet->row_count - 1, count($sheet->columns) - 1);
        $max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
        $padding_length = $sheet->max_cell_tag_end - $sheet->max_cell_tag_start - strlen($max_cell_tag);
        $sheet->file_writer->fseek($sheet->max_cell_tag_start);
        $sheet->file_writer->write($max_cell_tag . str_repeat(" ", $padding_length));
        $sheet->file_writer->close();
        $sheet->finalized = true;
    }

    public function markMergedCell($sheet_name, $start_cell_row, $start_cell_column, $end_cell_row, $end_cell_column) {
        if (empty($sheet_name) || $this->sheets[$sheet_name]->finalized)
            return;

        self::initializeSheet($sheet_name);
        $sheet = &$this->sheets[$sheet_name];

        $startCell = self::xlsCell($start_cell_row, $start_cell_column);
        $endCell = self::xlsCell($end_cell_row, $end_cell_column);
        $sheet->merge_cells[] = $startCell . ":" . $endCell;
    }

    public function writeSheet(array $data, $sheet_name = '', array $header_types = array()) {
        $sheet_name = empty($sheet_name) ? 'Sheet1' : $sheet_name;
        $data = empty($data) ? array(array('')) : $data;
        if (!empty($header_types)) {
            $this->writeSheetHeader($sheet_name, $header_types);
        }
        foreach ($data as $i => $row) {
            $this->writeSheetRow($sheet_name, $row);
        }
        $this->finalizeSheet($sheet_name);
    }

    protected function writeCell(XLSXWriter_BuffererWriter &$file, $row_number, $column_number, $value, $num_format_type, $cell_style_idx) {
        $cell_name = self::xlsCell($row_number, $column_number);
        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" />');
        } elseif (is_string($value) && $value{0} == '=') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="s"><f>' . self::xmlspecialchars($value) . '</f></c>');
        } elseif ($num_format_type == 'n_date') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . intval(self::convert_date_time($value)) . '</v></c>');
        } elseif ($num_format_type == 'n_datetime') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::convert_date_time($value) . '</v></c>');
        } elseif ($num_format_type == 'n_numeric') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::xmlspecialchars($value) . '</v></c>'); //int,float,currency
        } elseif ($num_format_type == 'n_string') {
            $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . self::xmlspecialchars($value) . '</t></is></c>');
        } elseif ($num_format_type == 'n_auto' || 1) { //auto-detect unknown column types
            if (!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^\-?[1-9][0-9]*(\.[0-9]+)?$/", $value)) {
                $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="n"><v>' . self::xmlspecialchars($value) . '</v></c>'); //int,float,currency
            } else { //implied: ($cell_format=='string')
                $file->write('<c r="' . $cell_name . '" s="' . $cell_style_idx . '" t="inlineStr"><is><t>' . self::xmlspecialchars($value) . '</t></is></c>');
            }
        }
    }

    protected function styleFontIndexes() {
        static $border_allowed = array('left', 'right', 'top', 'bottom');
        static $border_style_allowed = array('thin'/* default */, 'none', 'double', 'thin', 'medium', 'dashed', 'hair', 'thick');
        // CellStyle.BORDER_DOUBLE      双边线   
        // CellStyle.BORDER_THIN        细边线   
        // CellStyle.BORDER_MEDIUM      中等边线   
        // CellStyle.BORDER_DASHED      虚线边线   
        // CellStyle.BORDER_HAIR        小圆点虚线边线   
        // CellStyle.BORDER_THICK       粗边线   
        static $horizontal_allowed = array('general', 'left', 'right', 'justify', 'center');
        static $vertical_allowed = array('bottom', 'center', 'distributed');
        //format
        $base_fmt_index = 1; //for index //2 placeholders for static xml later
        //fill
        $fills = array();
        $base_fill_index = 2; //for index //2 placeholders for static xml later
        $fills_tmp_list = array();
        //font
        $fonts = array();
        $base_font_index = 2; //for index //2 placeholders for static xml later
        $default_font = array('size' => '10', 'name' => 'Arial', 'family' => '2');
        //border
        $borders = array();
        $base_border_index = 1; //for index //1 placeholders for static xml later
        $borders_tmp_list = array();
        //
        $style_indexes = array();
        foreach ($this->cell_styles as $i => $cell_style_string) {
            $semi_colon_pos = strpos($cell_style_string, ";");
            $number_format_idx = substr($cell_style_string, 0, $semi_colon_pos);
            $style_json_string = substr($cell_style_string, $semi_colon_pos + 1);
            $style = @json_decode($style_json_string, $as_assoc = true);
            $style_indexes[$i] = array('num_fmt_idx' => $number_format_idx); //initialize entry
            if (isset($style['border'])) {
                if (is_array($style['border'])) {
                    $border_input = $style['border'];
                    //All
                    if (isset($border_input['style']) || isset($border_input['color'])) {
                        foreach ($border_allowed as $key => $value) {
                            if (!isset($border_input[$value])) {
                                $border_input[$value] = $border_input;
                            }
                        }
                    }
                    //each
                    foreach ($border_input as $key => $border_input_value) {
                        if (!in_array($key, $border_allowed)) {
                            continue;
                        }
                        $border_style_input = [
                            'style' => $border_style_allowed[0],
                            'color' => 'FFBFBFBF',
                        ];
                        $border_input_style_value = '';
                        if (is_array($border_input_value)) {
                            $border_input_style_value = trim($border_input_value['style']);
                        }
                        if (is_string($border_input_value)) {
                            $border_input_style_value = $border_input_value;
                        }
                        if (in_array($border_input_style_value, $border_style_allowed)) {
                            $border_style_input['style'] = $border_input_style_value;
                        }
                        $border_input_color_value = '';
                        if (is_array($border_input_value)) {
                            $color = substr(strtoupper(trim(strval(@$border_input_value['color']))), 0, 8);
                            $color = $this->getColorStandardized($color);
                            if (strlen($color) == 6) {
                                $color = 'FF' . $color;
                            }
                            if ($color) {
                                $border_style_input['color'] = $color;
                            }
                        }
                        $border_input[$key] = $border_style_input;
                    }
                    //var_dump($border_input);
                    $border_idx = self::get_list_index($borders_tmp_list, json_encode($border_input));
                    if ($border_idx < 0) {
                        $borders_tmp_list[] = json_encode($border_input);
                        $borders[] = $border_input;
                    }
                    $border_idx = self::get_list_index($borders_tmp_list, json_encode($border_input));
                    $style_indexes[$i]['border_idx'] = $border_idx + $base_border_index; //fix index
                }
            }
            $fill_input = array(
                'frontgroud' => '',
                'backgroud' => '',
            );
            if (isset($style['frontgroud'])) {
                $color = substr(strtoupper(trim(strval($style['frontgroud']))), 0, 8);
                $color = $this->getColorStandardized($color);
                if (strlen($color) == 6) {
                    $color = 'FF' . $color;
                }
                $fill_input['frontgroud'] = $color;
            }
            if (isset($style['backgroud'])) {
                $color = substr(strtoupper(trim(strval($style['backgroud']))), 0, 8);
                $color = $this->getColorStandardized($color);
                if (strlen($color) == 6) {
                    $color = 'FF' . $color;
                }
                $fill_input['backgroud'] = $color;
                $fill_input['frontgroud'] = $color; // use fgColor
            }
            $fill_idx = self::get_list_index($fills_tmp_list, json_encode($fill_input));
            if ($fill_idx < 0) {
                $fills_tmp_list[] = json_encode($fill_input);
                $fills[] = $fill_input;
            }
            $fill_idx = self::get_list_index($fills_tmp_list, json_encode($fill_input));
            $style_indexes[$i]['fill_idx'] = $fill_idx + $base_fill_index; //fix index

            if (isset($style['halign'])) {
                $halign = strtolower(trim($style['halign']));
                if (in_array($halign, $horizontal_allowed)) {
                    $style_indexes[$i]['alignment'] = true;
                    $style_indexes[$i]['halign'] = $halign;
                }
            }
            if (isset($style['valign'])) {
                $valign = strtolower(trim($style['valign']));
                if (in_array($valign, $vertical_allowed)) {
                    $style_indexes[$i]['alignment'] = true;
                    $style_indexes[$i]['valign'] = $valign;
                }
            }
            if (isset($style['wrap_text'])) {
                $style_indexes[$i]['alignment'] = true;
                $style_indexes[$i]['wrap_text'] = $style['wrap_text'];
            }

            $font = $default_font;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']); //floatval to allow "10.5" etc
            }
            if (isset($style['font-family']) && is_string($style['font-family'])) {
                if ($style['font-family'] == 'Comic Sans MS') {
                    $font['family'] = 4;
                }
                if ($style['font-family'] == 'Times New Roman') {
                    $font['family'] = 1;
                }
                if ($style['font-family'] == 'Courier New') {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font-family']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                $font_style = strtolower(trim($style['font-style']));
                if (strpos($font_style, 'bold') !== false) {
                    $font['bold'] = true;
                }
                if (strpos($font_style, 'italic') !== false) {
                    $font['italic'] = true;
                }
                if (strpos($font_style, 'strike') !== false) {
                    $font['strike'] = true;
                }
                if (strpos($font_style, 'underline') !== false) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['font-color']) && is_string($style['font-color'])) {
                $color = substr(strtoupper(trim(strval($style['font-color']))), 0, 8);
                $color = $this->getColorStandardized($color);
                if (strlen($color) == 6) {
                    $color = 'FF' . $color;
                }
                $font['color'] = $color;
            }
            if ($font != $default_font) {
                $style_indexes[$i]['font_idx'] = self::add_to_list_get_index($fonts, json_encode($font)) + $base_font_index; //fix index
            }
        }
        return array('fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $style_indexes);
    }

    protected function writeStylesXML() {
        $r = self::styleFontIndexes();
        $fills = $r['fills'];
        $fonts = $r['fonts'];
        $borders = $r['borders'];
        $style_indexes = $r['styles'];

        $temporary_filename = $this->tempFilename();
        $file = new XLSXWriter_BuffererWriter($temporary_filename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' . "\n");
        $file->write('<numFmts count="' . (count($this->number_formats) + 1) . '">' . "\n");
        $file->write('<numFmt formatCode="GENERAL" numFmtId="76"/>' . "\n");
        foreach ($this->number_formats as $i => $v) {
            $v = self::xmlspecialchars($v);
            $v = self::numberFormatStandardized($v);
            $file->write('<numFmt numFmtId="' . (80 + $i) . '" formatCode="' . ($v) . '" />' . "\n");
        }
        $file->write('</numFmts>');
        $file->write("\n");

        $file->write('<fonts count="' . (count($fonts) + 2) . '">' . "\n");
        $file->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>' . "\n");
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>' . "\n");

        foreach ($fonts as $font) {
            if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
                $f = json_decode($font, true);
                $file->write('<font>');
                $file->write('<name val="' . htmlspecialchars($f['name']) . '"/><charset val="1"/><family val="' . intval($f['family']) . '"/>');
                $file->write('<sz val="' . intval($f['size']) . '"/>');
                if (!empty($f['color'])) {
                    $file->write('<color rgb="' . strval($f['color']) . '"/>');
                }
                if (!empty($f['bold'])) {
                    $file->write('<b val="true"/>');
                }
                if (!empty($f['italic'])) {
                    $file->write('<i val="true"/>');
                }
                if (!empty($f['underline'])) {
                    $file->write('<u val="single"/>');
                }
                if (!empty($f['strike'])) {
                    $file->write('<strike val="true"/>');
                }
                $file->write('</font>' . "\n");
            }
        }
        $file->write('</fonts>');
        $file->write("\n");

        $file->write('<fills count="' . (count($fills) + 2) . '">' . "\n");
        $file->write('<fill><patternFill patternType="none"/></fill>' . "\n");
        $file->write('<fill><patternFill patternType="gray125"/></fill>' . "\n");
        foreach ($fills as $fill) {
            if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
                $fg_color = '';
                if (is_array($fill) && isset($fill['frontgroud'])) {
                    $fg_color = strtoupper(strval($fill['frontgroud']));
                }
                $bg_color = '';
                if (is_array($fill) && isset($fill['backgroud'])) {
                    $bg_color = strtoupper(strval($fill['backgroud']));
                }
                $file->write(
                        '<fill>' .
                        '<patternFill patternType="solid">' .
                        '<fgColor ' . ($fg_color ? 'rgb="' . $fg_color . '" ' : '') . '/>' .
                        '<bgColor ' . ($bg_color ? 'rgb="' . $bg_color . '" ' : '') . '/>' .
                        '</patternFill>' .
                        '</fill>' . "\n"
                );
            }
        }
        $file->write('</fills>');
        $file->write("\n");

        $file->write('<borders count="' . (count($borders) + 1) . '">' . "\n");
        $file->write('<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>' . "\n");
        foreach ($borders as $border) {
            if (!empty($border)) { //fonts have an empty placeholder in the array to offset the static xml entry above
                $file->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (['left', 'right', 'top', 'bottom'] as $border_direction) {
                    if (in_array($border_direction, array_keys($border))) {
                        $file->write(
                                '<' . $border_direction . ' ' . (@$border[$border_direction]['style'] ? 'style="' . $border[$border_direction]['style'] . '" ' : '') . '>' .
                                (@$border[$border_direction]['color'] ? '<color rgb="' . $border[$border_direction]['color'] . '"/>' : '') .
                                '</' . $border_direction . '>'
                        );
                    }
                }
                $file->write('<diagonal/>');
                $file->write('</border>' . "\n");
            }
        }
        $file->write('</borders>');
        $file->write("\n");
        $file->write('<cellStyleXfs count="4">' . "\n");
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="79">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>' . "\n");
        $file->write('</cellStyleXfs>');
        $file->write("\n");

        $file->write('<cellXfs count="' . (count($style_indexes) + 1) . '">' . "\n");
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="79" xfId="0"/>' . "\n");

        foreach ($style_indexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText = isset($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? strval($v['halign']) : 'general';
            $vertAlignment = isset($v['valign']) ? strval($v['valign']) : 'bottom';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont = 'true';
            $borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
            $numFmtIdx = isset($v['num_fmt_idx']) ? intval($v['num_fmt_idx']) : 0;
            //$file->write('<xf applyAlignment="'.$applyAlignment.'" applyBorder="'.$applyBorder.'" applyFont="'.$applyFont.'" applyProtection="false" borderId="'.($borderIdx).'" fillId="'.($fillIdx).'" fontId="'.($fontIdx).'" numFmtId="'.(80+$v['num_fmt_idx']).'" xfId="0"/>');
            $file->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (80 + $numFmtIdx) . '" xfId="0">');
            $file->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $file->write('	<protection locked="true" hidden="false"/>');
            $file->write('</xf>' . "\n");
        }
        $file->write('</cellXfs>');
        $file->write("\n");

        $file->write("\n");
        $file->write('</styleSheet>');
        $file->write("\n");
        $file->close();
        return $temporary_filename;
    }

    protected function buildAppXML() {
        $app_xml = "";
        $app_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $app_xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' . "\n";
        $app_xml .= '<TotalTime>0</TotalTime>' . "\n";
        $app_xml .= '</Properties>';
        return $app_xml;
    }

    protected function buildCoreXML() {
        $core_xml = "";
        $core_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $core_xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $core_xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>' . "\n"; //$date_time = '2014-10-25T15:54:37.00Z';
        $core_xml .= '<dc:creator>' . self::xmlspecialchars($this->author) . '</dc:creator>' . "\n";
        $core_xml .= '<cp:revision>0</cp:revision>' . "\n";
        $core_xml .= '</cp:coreProperties>';
        return $core_xml;
    }

    protected function buildRelationshipsXML() {
        $rels_xml = "";
        $rels_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $rels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' . "\n";
        $rels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' . "\n";
        $rels_xml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' . "\n";
        $rels_xml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' . "\n";
        $rels_xml .= '</Relationships>';
        return $rels_xml;
    }

    protected function buildWorkbookXML() {
        $i = 0;
        $workbook_xml = "";
        $workbook_xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbook_xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' . "\n";
        $workbook_xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>' . "\n";
        $workbook_xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>' . "\n";
        $workbook_xml .= '<sheets>' . "\n";
        foreach ($this->sheets as $sheet_name => $sheet) {
            $sheetname = self::sanitize_sheetname($sheet->sheetname);
            $workbook_xml .= '<sheet name="' . self::xmlspecialchars($sheetname) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>' . "\n";
            $i++;
        }
        $workbook_xml .= '</sheets>' . "\n";
        $workbook_xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>' . "\n";
        $workbook_xml .= '</workbook>' . "\n";
        return $workbook_xml;
    }

    protected function buildWorkbookRelsXML() {
        $i = 0;
        $wkbkrels_xml = "";
        $wkbkrels_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $wkbkrels_xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' . "\n";
        $wkbkrels_xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' . "\n";
        foreach ($this->sheets as $sheet_name => $sheet) {
            $wkbkrels_xml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet->xmlname) . '"/>' . "\n";
            $i++;
        }
        $wkbkrels_xml .= '</Relationships>';
        return $wkbkrels_xml;
    }

    protected function buildContentTypesXML() {
        $content_types_xml = "";
        $content_types_xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $content_types_xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' . "\n";
        $content_types_xml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' . "\n";
        $content_types_xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' . "\n";
        $content_types_xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' . "\n";
        $content_types_xml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' . "\n";
        $content_types_xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' . "\n";
        $content_types_xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' . "\n";
        foreach ($this->sheets as $sheet_name => $sheet) {
            $content_types_xml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlname) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' . "\n";
        }
        $content_types_xml .= '</Types>' . "\n";
        return $content_types_xml;
    }

    //------------------------------------------------------------------
    /*
     * @param $row_number int, zero based
     * @param $column_number int, zero based
     * @return Cell label/coordinates, ex: A1, C3, AA42
     * */
    public static function xlsCell($row_number, $column_number) {
        $n = $column_number;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        return $r . ($row_number + 1);
    }

    //------------------------------------------------------------------
    public static function log($string) {
        file_put_contents("php://stderr", date("Y-m-d H:i:s:") . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }

    //------------------------------------------------------------------
    public static function sanitize_filename($filename) { //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
        $nonprinting = array_map('chr', range(0, 31));
        $invalid_chars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
        $all_invalids = array_merge($nonprinting, $invalid_chars);
        return str_replace($all_invalids, "", $filename);
    }

    //------------------------------------------------------------------
    public static function sanitize_sheetname($sheetname) {
        static $badchars = '\\/?*:[]';
        static $goodchars = '        ';
        $sheetname = strtr($sheetname, $badchars, $goodchars);
        $sheetname = substr($sheetname, 0, 31);
        $sheetname = trim(trim(trim($sheetname), "'")); //trim before and after trimming single quotes
        return !empty($sheetname) ? $sheetname : 'Sheet' . ((rand() % 900) + 100);
    }

    //------------------------------------------------------------------
    public static function xmlspecialchars($val) {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodchars = "                              ";
        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badchars, $goodchars); //strtr appears to be faster than str_replace
    }

    //------------------------------------------------------------------
    public static function array_first_key(array $arr) {
        reset($arr);
        $first_key = key($arr);
        return $first_key;
    }

    //------------------------------------------------------------------
    private static function determineNumberFormatType($num_format) {
        $num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $num_format);
        if ($num_format == 'GENERAL')
            return 'n_auto';
        if ($num_format == 'STRING')
            return 'n_string';
        if ($num_format == 'NUMBER')
            return 'n_numeric';
        if (preg_match("/[H]{1,2}:[M]{1,2}/", $num_format))
            return 'n_datetime';
        if (preg_match("/[M]{1,2}:[S]{1,2}/", $num_format))
            return 'n_datetime';
        if (preg_match("/[YY]{2,4}/", $num_format))
            return 'n_date';
        if (preg_match("/[D]{1,2}/", $num_format))
            return 'n_date';
        if (preg_match("/[M]{1,2}/", $num_format))
            return 'n_date';
        if (preg_match("/$/", $num_format))
            return 'n_numeric';
        if (preg_match("/%/", $num_format))
            return 'n_numeric';
        if (preg_match("/0/", $num_format))
            return 'n_numeric';
        return 'n_auto';
    }

    //------------------------------------------------------------------
    private function getColorStandardized($color) {
        switch (strtolower($color)) {
            case 'white':
                return 'FFFFFF';
            case 'black':
                return '000000';
            case 'gray':
                return 'CCCCCC';
            case 'red':
                return 'FF0000';
        }
        return $color;
    }

    //------------------------------------------------------------------
    private static function numberFormatStandardized($num_format) {
        if ($num_format == 'money') {
            $num_format = 'dollar';
        }
        if ($num_format == 'number') {
            $num_format = 'integer';
        }
        if ($num_format == 'string')
            $num_format = 'STRING';
        else if ($num_format == 'integer')
            $num_format = 'NUMBER';
        else if ($num_format == 'date')
            $num_format = 'YYYY-MM-DD';
        else if ($num_format == 'datetime')
            $num_format = 'YYYY-MM-DD HH:MM:SS';
        else if ($num_format == 'price')
            $num_format = '#,##0.00';
        else if ($num_format == 'dollar')
            $num_format = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
        else if ($num_format == 'euro')
            $num_format = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
        $ignore_until = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($num_format); $i < $ix; $i++) {
            $c = $num_format[$i];
            if ($ignore_until == '' && $c == '[')
                $ignore_until = ']';
            else if ($ignore_until == '' && $c == '"')
                $ignore_until = '"';
            else if ($ignore_until == $c)
                $ignore_until = '';
            if ($ignore_until == '' && ($c == ' ' || $c == '-' || $c == '(' || $c == ')') && ($i == 0 || $num_format[$i - 1] != '_'))
                $escaped .= "\\" . $c;
            else
                $escaped .= $c;
        }
        return $escaped;
    }

    //------------------------------------------------------------------
    public static function get_list_index(&$haystack, $needle) {
        $existing_idx = array_search($needle, $haystack, $strict = true);
        if ($existing_idx === false) {
            $existing_idx = -1;
        }
        return $existing_idx;
    }

    //------------------------------------------------------------------
    public static function add_to_list_get_index(&$haystack, $needle) {
        $existing_idx = array_search($needle, $haystack, $strict = true);
        if ($existing_idx === false) {
            $existing_idx = count($haystack);
            $haystack[] = $needle;
        }
        return $existing_idx;
    }

    //------------------------------------------------------------------
    public static function convert_date_time($date_input) { //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
        $days = 0;    # Number of days since epoch
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;
        $hour = $min = $sec = 0;

        $date_time = $date_input;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case
        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31')
            return $seconds;# Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00')
            return $seconds;# Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29')
            return 60 + $seconds;# Excel false leapday
        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
        $mdays = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

        # Some boundary checks
        if ($year < $epoch || $year > 9999)
            return 0;
        if ($month < 1 || $month > 12)
            return 0;
        if ($day < 1 || $day > $mdays[$month - 1])
            return 0;

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += intval(( $range ) / 4);             # Add leapdays
        $days -= intval(( $range + $offset ) / 100); # Subtract 100 year leapdays
        $days += intval(( $range + $offset + $norm ) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above
        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }

    //------------------------------------------------------------------
}

class XLSXWriter_BuffererWriter {

    protected $fd = null;
    protected $buffer = '';
    // if xlsx file content is very large , and the processing time is long ,
    // you can change the value of $buffer_size to 32KB or 128KB etc.
    protected $buffer_size = 8192; // Byte
    protected $check_utf8 = false;

    public function __construct($filename, $fd_fopen_flags = 'w', $check_utf8 = false) {
        $this->check_utf8 = $check_utf8;
        $this->fd = fopen($filename, $fd_fopen_flags);
        if ($this->fd === false) {
            XLSXWriter::log("Unable to open $filename for writing.");
        }
    }

    public function write($string) {
        $this->buffer .= $string;
        if (isset($this->buffer[$this->buffer_size - 1])) {
            $this->purge();
        }
    }

    protected function purge() {
        if ($this->fd) {
            if ($this->check_utf8 && !self::isValidUTF8($this->buffer)) {
                XLSXWriter::log("Error, invalid UTF8 encoding detected.");
                $this->check_utf8 = false;
            }
            fwrite($this->fd, $this->buffer);
            $this->buffer = '';
        }
    }

    public function close() {
        $this->purge();
        if ($this->fd) {
            fclose($this->fd);
            $this->fd = null;
        }
    }

    public function __destruct() {
        $this->close();
    }

    public function ftell() {
        if ($this->fd) {
            $this->purge();
            return ftell($this->fd);
        }
        return -1;
    }

    public function fseek($pos) {
        if ($this->fd) {
            $this->purge();
            return fseek($this->fd, $pos);
        }
        return -1;
    }

    protected static function isValidUTF8($string) {
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($string, 'UTF-8') ? true : false;
        }
        return preg_match("//u", $string) ? true : false;
    }

}

// end code
