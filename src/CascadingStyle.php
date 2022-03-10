<?php

namespace RobertN7\ExcelStyles;

class CascadingStyle implements ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element)
    {
        foreach ($element->elements as $el) {
            if(is_subclass_of($el, ExcelElement::class))
               $this->applyTo($sheet, $el);
        }
    }

    public function __toString(): string
    {
        return 'cascading';
    }
}