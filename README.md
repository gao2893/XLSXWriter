# XLSXWriter
php excel export lib


## Fork from mk-j/PHP_XLSXWriter

Remove the `hair border` line style of all cells.

Add ecxcel cell style 

## Examples
``` php
        $cell_style1 = array(
            'halign' => 'center',
            'border' => array(
                    'left' => ['style'=>'thin', 'color'=>'FFCCCCCC'],
                    'right' => ['style'=>'dashed' ],
                    'top' => [],
                    // 'bottom' => [],
                    ),
            'backgroud' => 'gray',
            'frontgroud' => 'black',
        );         
        $style1 = $cell_style1; // extend cell style to row
        $writer->writeSheetHeader($SheetName, $header, $style1);
        foreach ($rows as $row) {
            $writer->writeSheetRow($SheetName, $row, $style1);
        }
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
        
