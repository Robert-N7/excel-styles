<?php

namespace App\Reports\Excel;

use PhpOffice\PhpSpreadsheet\Style\Fill;

class BackgroundColor implements ExcelStyle
{
    public function __construct($rgb = 'd0d0d0')
    {
        $this->color = 'FF' . $rgb;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFill()->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB($this->color);
    }
}