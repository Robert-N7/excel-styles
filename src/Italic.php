<?php

namespace RobertN7\ExcelStyles;

class Italic implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setItalic(true);
    }

    public function toString(): string
    {
        return 'italic';
    }
}