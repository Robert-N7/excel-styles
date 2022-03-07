<?php

namespace App\Reports\Excel;

use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Excel;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

abstract class ExcelElement
{
    public $elements;
    protected $styles;
    protected $start;
    protected $end;
    private $_style;

    private static function normalizeElement($element) {

    }

    protected function _layout(&$arr, $excel)
    {
        // normalize
        $normalized = [];
        foreach($this->elements as $element) {
            if(!is_subclass_of($element, ExcelElement::class)) {
                if(is_iterable($element) && !is_string($element) && !$excel->_c_row) {
                    // typical when using a collection to get an array of arrays
                    // convert it to rows
                    foreach($element as $row) {
                        $generated = [];
                        foreach($row as $c) {
                            $cell = new ExCell([], $c, 1);
                            array_push($generated, $cell);
                        }
                        array_push($normalized, new ExRow([], $generated, 1));
                    }
                } else {
                    array_push($normalized, new ExCell([], $element, 1));
                }
            } else {
                array_push($normalized, $element);
            }
        }
        $this->elements = $normalized;
        $current_max = 0;
        foreach($normalized as $element) {
            if(($max = $element->layout($arr, $excel)) > $current_max)
                $current_max = $max;
        }
        return $current_max;
    }

    public function __construct($styles, $elements)
    {
        $this->elements = $elements;
        $this->styles = $styles;
    }

    public function layout(&$arr, $excel)
    {
        if($excel->_advance_row) {
            $excel->c_col = 0;
            $excel->c_row = $excel->c_row + 1;
            $excel->_advance_row = false;
        }
        $this->start = chr(65 + $excel->c_col) . $excel->c_row + 1;
        $max = $this->_layout($arr, $excel);
        $this->end = self::getColumn($max) . $excel->c_row + 1;
        return $max;
    }

    private static function getColumn($max)
    {
        $c = intdiv($max, 26);
        if($c)
            return self::getColumn($c - 1) . chr($max % 26 + 65);
        return chr($max + 65);
    }

    private function __applyStyle($style, $sheet, $styles)
    {
        if(is_string($style)) {
            $this->__applyStyle($styles[$style], $sheet, $styles);
        }
        else if(is_array($style)) {
            foreach ($style as $s)
                $this->__applyStyle($s, $sheet, $styles);
        } else {
            $style->applyTo($sheet, $this);
        }
    }

    public function render($sheet, $styles)
    {
        // Apply styles
        $this->__applyStyle($this->styles, $sheet, $styles);
        if(is_array($this->elements)) {
            foreach($this->elements as $element) {
                if(is_subclass_of($element, ExcelElement::class))
                    $element->render($sheet, $styles);
            }
        }
    }

    public function getStyle($sheet)
    {
        if(!$this->_style)
            $this->_style = $sheet->getStyle($this->start == $this->end ? $this->start
                : $this->start . ':' . $this->end);
        return $this->_style;
    }

}

class ExRow extends ExcelElement
{
    private int $n;
    private array $row_data;

    public function __construct($styles, $elements, $n)
    {
        parent::__construct($styles, $elements);
        $this->n=$n;
        $this->row_data = [[]];
    }

    public function _layout(&$arr, $excel)
    {
        $excel->_c_row = $this;
        $max = parent::_layout($this->row_data[0], $excel);
        for($i=1; $i < $this->n; $i++)
            array_push($this->row_data, []);
        array_push($arr, $this->row_data);
        $excel->c_row = $this->n - 1 + $excel->c_row;
        $excel->_advance_row = true;
        $excel->_c_row = null;
        return $max;
    }
}

class ExCell extends ExcelElement
{
    private int $n;

    public function __construct($styles, $elements, $n)
    {
        parent::__construct($styles, $elements);
        $this->n = $n;
    }

    public function layout(&$arr, $excel) {
        $max = parent::layout($arr, $excel);
        $excel->c_col = $this->n + $excel->c_col;
        return $max;
    }

    public function _layout(&$arr, $excel)
    {
        array_push($arr, $this->elements);
        for($i=1; $i<$this->n; $i++)
            array_push($arr, null);
        $excel->c_col = $this->n - 1 + $excel->c_col;
        return $excel->c_col;
    }
}

class ExDiv extends ExcelElement
{

}

abstract class ExcelView implements WithStyles, FromArray
{
    public int $c_row;
    public int $c_col;
    private $_layout;
    private array $_arrayed;
    private array $_styles;

    public function __construct()
    {
        $this->c_row = 0;
        $this->c_col = 0;
        $this->_layout = null;
        $this->_arrayed = [];
    }

    abstract protected function layout();
    abstract protected function init_styles();

    protected function tr($styles=[], $data=[], $n=1)
    {
        return new ExRow($styles, $data, $n);
    }

    protected function td($styles=[], $data=[], $n=1)
    {
        return new Excell($styles, $data, $n);
    }

    protected function table($styles=[], $data=[])
    {
        return new ExDiv($styles, $data);
    }

    protected function br($n=1)
    {
        return new ExRow([], [], $n);
    }

    protected function div($styles=[], $data=[])
    {
        if($data && is_string($data[0]))
            return new ExRow($styles, $data, 1);
        return new ExDiv($styles, $data);
    }

    public function styles(Worksheet $sheet)
    {
        $this->_layout->render($sheet, $this->_styles);
    }

    public function array(): array
    {
        $this->_styles = $this->init_styles();
        $this->_advance_row = false;
        $this->_c_row = null;
        $this->_layout = $this->layout();
        $this->_layout->layout($this->_arrayed, $this);
        return $this->_arrayed;
    }
}