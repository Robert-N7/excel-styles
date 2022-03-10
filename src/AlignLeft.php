<?php

namespace RobertN7\ExcelStyles;

class AlignLeft implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getAlignment()->setHorizontal('left');
    }

    public function toString(): string
    {
        return 'align-left';
    }
}