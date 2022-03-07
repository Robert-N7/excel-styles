<?php

namespace App\Reports\Excel;

class Bold extends CascadingStyle implements ExcelStyle
{

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setBold(true);
    }
}