<?php

namespace RobertN7\ExcelStyles;

class Border implements ExcelStyle
{
    /*
    *   $border_style: none, dashDot, dashDotDot, dashed,
            dotted, double, hair, medium, mediumDashDot,
            mediumDashDotDot, mediumDashed, slantDashDot,
            thick, thin;
    *   $rgb: RGB color
    *   $direction: bottom, diagonal, diagonalDirection, left, right, top,
    *                    outline, inside, vertical, horizontal, allBorders
    */
    public function __construct($border_style = 'thick', $rgb = '000000', $direction='outline')
    {
        $this->border = ['borders' => [
            $direction => [
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
        $s = 'border: [';
        foreach($this->border['borders'] as $style => $border) {
            $s .= $style . ' => ' . $border['borderStyle']
                . ' rgb('.substr($border['color']['argb'], 2).'),';
        }
        return $s . ']';
    }

}