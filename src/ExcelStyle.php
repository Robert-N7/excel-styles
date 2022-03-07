<?php

namespace Robert-N7\ExcelStyles;

use PhpOffice\PhpSpreadsheet\Style\Fill;

interface ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element);
}


