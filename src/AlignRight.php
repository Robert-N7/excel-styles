<?php

namespace RobertN7\ExcelStyles;

class AlignRight implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getAlignment()->setHorizontal('right');
    }

    public function __toString(): string
    {
        return 'align-right';
    }

}