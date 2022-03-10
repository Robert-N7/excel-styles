<?php

namespace RobertN7\ExcelStyles;

class Border implements ExcelStyle
{
    public function __construct($border_style = 'thick', $rgb = '000000')
    {
        $this->border = ['borders' => [
            'outline' => [
                'borderStyle' => $border_style,
                'color' => ['argb' => (strlen($rgb) > 6 ? $rgb : 'FF' . $rgb)],
            ]
        ]];
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->applyFromArray($this->border);
    }

    public function __toString(): string
    {
        $outline = $this->border['borders']['outline'];
        return 'border: ' . $outline['borderStyle'] . ' ' . $outline['color']['argb'];
    }

}