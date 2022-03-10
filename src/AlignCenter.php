<?php

namespace RobertN7\ExcelStyles;

class AlignCenter implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getAlignment()->setHorizontal('center');
    }

    public function __toString(): string
    {
        return 'align-center';
    }
}