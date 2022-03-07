<?php

namespace App\Reports\Excel;

use PhpOffice\PhpSpreadsheet\Style\Fill;

interface ExcelStyle
{
    public function applyTo($sheet, ExcelElement $element);
}


