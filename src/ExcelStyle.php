<?php

namespace RobertN7\ExcelStyles;

use PhpOffice\PhpSpreadsheet\Style\Fill;

interface ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element);
}


