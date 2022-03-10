<?php

namespace RobertN7\ExcelStyles;

use PhpOffice\PhpSpreadsheet\Style\NumberFormat;


class Format implements ExcelStyle
{
    private $format;
    // Custom formats - for others see NumberFormat
    private static array $formats = [
        '$' => NumberFormat::FORMAT_CURRENCY_USD,
        '%' => NumberFormat::FORMAT_PERCENTAGE,
        '€' => NumberFormat::FORMAT_CURRENCY_EUR,
        'float' => NumberFormat::FORMAT_NUMBER_00,
        'int' => NumberFormat::FORMAT_NUMBER,
    ];

    public function __construct($format='')
    {
        if(array_key_exists($format, self::$formats))
            $format = self::$formats[$format];
        $this->format = $format;
    }

    public function applyTo($sheet, ExcelElement $element)
    {
        $element->getStyle($sheet)->getNumberFormat()->setFormatCode($this->format);
    }

    public function toString(): string
    {
        return 'format: ' . $this->format;
    }
}