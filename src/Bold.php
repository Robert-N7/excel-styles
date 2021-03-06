<?php

namespace RobertN7\ExcelStyles;

class Bold extends CascadingStyle implements ExcelStyle
{

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setBold(true);
    }

    public function __toString(): string
    {
        return 'bold';
    }
}