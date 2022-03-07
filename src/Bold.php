<?php

namespace Robert-N7\ExcelStyles;

class Bold extends CascadingStyle implements ExcelStyle
{

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setBold(true);
    }
}