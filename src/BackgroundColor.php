<?php

namespace RobertN7\ExcelStyles;

use PhpOffice\PhpSpreadsheet\Style\Fill;

class BackgroundColor implements ExcelStyle
{
    public function __construct($rgb = 'd0d0d0')
    {
        $this->color = strlen($rgb) > 6 ? $rgb : 'FF' . $rgb;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFill()->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB($this->color);
    }

    public function __toString(): string
    {
        return 'background-color: ' . $this->color;
    }
}