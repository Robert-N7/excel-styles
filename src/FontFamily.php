<?php

namespace RobertN7\ExcelStyles;

class FontFamily implements ExcelStyle
{
    private string $font;

    public function __construct($font)
    {
        $this->font = $font;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getFont()->setName($this->font);
    }

    public function __toString(): string
    {
        return 'font: ' . $this->font;
    }
}