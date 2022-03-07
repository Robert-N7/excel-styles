<?php

namespace App\Reports\Excel;

class CascadingStyle implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        foreach ($element->elements as $el) {
            if(is_subclass_of($el, ExcelElement::class))
               $this->applyTo($sheet, $el);
        }
    }
}