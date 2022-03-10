<?php

namespace RobertN7\ExcelStyles;

class Color implements ExcelStyle
{
    private $color;

    public function __construct($color)
    {
        $this->color = $color;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->getColor()->setARGB(
            strlen($this->color) > 6 ? $this->color : 'FF' . $this->color
        );
    }

    public function __toString(): string
    {
        return 'color: ' . $this->color;
    }
}