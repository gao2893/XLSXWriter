# XLSXWriter
php excel export lib


## Fork from mk-j/PHP_XLSXWriter

Remove the `default` `hair` border line style of all cells.

Add ecxcel cell style 

## Examples
``` php
        
        $cell_style1 = array(
            //'width' => '40', //default 11.5
            'halign' => 'center',
            'valign' => 'center',
            'wrap_text' => 'center',
            'border' => array(
                //'style' => 'hair', // all border style
                //'color' => 'FFF3F3F3', // all border style
                'left' => [
                    'style' => 'hair',
                    'color' => 'E26868',
                ], // over write left border style
                //'right' => ['style' => 'medium'],// no right border style
                'top' => [], // over write top border style by default style
            ),
            'backgroud' => 'gray',
            'font-family' => '宋体',
            'font-size' => '16',
            'font-color' => 'red',
            'font-style' => 'bold,italic,strike,underline',
        );
        $cell_style2 = array_merge($cell_style1, array(
            'width' => '30',
        ));
        $column_styles = [$cell_style1, $cell_style2, /* ... columns ... */];
        $writer->writeSheetHeader($SheetName, $header, $header_styles = $column_styles);
        foreach ($rows as $row)
            $writer->writeSheetRow($SheetName, $row, $row_styles = $column_styles);
        $str = $writer->writeToString();
        
```   
## Options
``` php
        
        static $border_allowed = array('left', 'right', 'top', 'bottom');
        static $border_style_allowed = array('thin'/* default */,'none', 'double', 'thin', 'medium', 'dashed', 'hair', 'thick');
        // CellStyle.BORDER_DOUBLE      双边线   
        // CellStyle.BORDER_THIN        细边线   
        // CellStyle.BORDER_MEDIUM      中等边线   
        // CellStyle.BORDER_DASHED      虚线边线   
        // CellStyle.BORDER_HAIR        小圆点虚线边线   
        // CellStyle.BORDER_THICK       粗边线   
        
```   
        
