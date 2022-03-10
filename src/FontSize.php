<?php

namespace RobertN7\ExcelStyles;

class FontSize implements ExcelStyle
{
    private int $size;

    public function __construct($size)
    {
        $this->size = $size;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setSize($this->size);
    }

    public function __toString(): string
    {
        return 'font-size: ' . $this->size;
    }
}