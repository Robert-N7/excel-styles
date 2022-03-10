<?php

namespace RobertN7\ExcelStyles;

class Underline implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setUnderline(true);
    }

    public function toString(): string
    {
        return 'underline';
    }
}