<?php

namespace RobertN7\ExcelStyles;

use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Excel;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Exception;

class ExcelViewException extends Exception
{
}

abstract class ExcelElement
{
    public $elements;
    public $styles;
    public $start;
    public $end;
    private $_style;

    public function end_row()
    {
        preg_match('/\d+/', $this->end, $matches);
        return $matches[0];
    }

    public function end_col()
    {
        preg_match('/\D+/', $this->end, $matches);
        return $matches[0];
    }

    public function start_row()
    {
        preg_match('/\d+/', $this->start, $matches);
        return $matches[0];
    }

    public function start_col()
    {
        preg_match('/\D+/', $this->start, $matches);
        return $matches[0];
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
            if(array_key_exists($style, $styles))
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
        if(is_array($elements) and !count($elements)) {
            $elements = [new ExCell([], '', 1)];
        }
        parent::__construct($styles, $elements);
        $this->n=$n;
        $this->row_data = [[]];
    }

    public function _layout(&$arr, $excel)
    {
        $excel->_c_row = $this;
        if(!count($this->row_data[0]))
            parent::_layout($this->row_data[0], $excel);
        array_push($arr, $this->row_data);
        $excel->_advance_row = true;
        $excel->_c_row = null;

        $this->n -= 1;
        if($this->n)
            $this->layout($arr, $excel);
        return count($this->row_data[0]) - 1;
    }
}

class ExCell extends ExcelElement
{
    private int $n;
    private $next;

    public function __construct($styles, $elements, $n)
    {
        if(is_array($elements) && !count($elements)) {
            $elements = '';
        }
        if(is_subclass_of($elements, ExcelElement::class) || is_array($elements)) {
            print_r($elements);
            throw new ExcelViewException('Cell value must be simple type or function');
        }
        parent::__construct($styles, $elements);
        $this->n = $n;
        $this->next = null;
    }

    public function layout(&$arr, $excel) {
        $max = parent::layout($arr, $excel);
        $excel->c_col = 1 + $excel->c_col;
        if($this->n > 1) {
            $this->next = [];
            for($i=1; $i<$this->n; $i++) {
                $el = new ExCell($this->styles, $this->elements, 1);
                array_push($this->next, $el);
                $max = $el->layout($arr, $excel);
            }
        }
        return $max;
    }

    public function _layout(&$arr, $excel)
    {
        array_push($arr, $this);
        return $excel->c_col;
    }
}

class ExDiv extends ExcelElement
{

}

class ExHead extends ExCell
{
    public function __construct($styles, $elements, $n)
    {
        if(!is_array($styles)) {
            $styles = $styles == null ? [new Bold()] : [new Bold(), $styles];
        }
        parent::__construct($styles, $elements, $n);
    }
}


abstract class ExcelView implements WithStyles, FromArray
{
    // These variables are used in layout, don't use them
    public int $c_row;  // Current row number
    public int $c_col;  // Current col number
    public $_c_row;  // Current ExRow
    public $__advance_row; // should advance?

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

    protected function th($styles=[], $data=[], $n=1)
    {
        return new ExHead($styles, $data, $n);
    }

    protected function sum($e, $rows=null, $cols=null)
    {
        return function($e) use ($rows, $cols) {
            if(!$rows) {
                $rows = [$e->start_row(), $e->end_row()];
            }
            if(!$cols) {
                $cols = [$e->start_col(), $e->end_col()];
            }
            return '=SUM(' . $cols[0] . $rows[0] . ':' . $cols[1] . $rows[1] . ')';
        };
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

    private function iter_elements(&$collect, $el, $cond)
    {
        if($cond($el)) {
            array_push($collect, $el);
        }
        if(is_subclass_of($el, ExcelElement::class) && is_array($el->elements)) {
            foreach($el->elements as $element) {
                $this->iter_elements($collect, $element, $cond);
            }
        }
    }

    protected function el($style=null, $type=null)
    {
        $r = $this->els($style, $type);
        if(count($r))
            return $r[0];
        return null;
    }

    protected function els($style=null, $type=null)
    {
        $arr = [];
        $this->iter_elements($arr, $this->_layout, function($e) use ($style, $type) {
            if($type && get_class($e) != $type) {
                return false;
            }
            if($style) {
                if(is_array($e->styles)) {
                    if(!is_array($style))
                        $style = [$style];
                    if(array_diff($style, $e->styles))
                        return false;
                } else if($style != $e->styles) {
                    return false;
                }
            }
            return true;
        });
        return $arr;
    }

    private function resolveArray($_arrayed)
    {
        $arr = [];
        foreach($_arrayed as $row) {
            $new_row = [];
            foreach($row as $col) {
                foreach($col as $v) {
                    if(is_callable($v->elements)) {
                        $el = call_user_func($v->elements, $v);
                        array_push($new_row, $el);
                    } else {
                        if(is_array($v->elements) && !count($v->elements))
                            array_push($new_row, null);
                        else
                            array_push($new_row, $v->elements);
                    }
                }
            }
            array_push($arr, [$new_row]);
        }
        return $arr;
    }

    public function array(): array
    {
        $this->_styles = $this->init_styles();
        $this->_advance_row = false;
        $this->_c_row = null;
        $this->_layout = $this->layout();
        $this->_layout->layout($this->_arrayed, $this);
        $this->_arrayed = $this->resolveArray($this->_arrayed);
        return $this->_arrayed;
    }
}