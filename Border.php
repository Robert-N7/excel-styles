<?php

namespace App\Reports\Excel;

class Border implements ExcelStyle
{
    public function __construct($border_style = 'thick', $rgb = '000000')
    {
        $this->border = ['borders' => [
            'outline' => [
                'borderStyle' => $border_style,
                'color' => ['argb' => 'FF' . $rgb],
            ]
        ]];
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->applyFromArray($this->border);
    }
}