<?php declare(strict_types=1);

namespace Tests;

use PHPUnit\Framework\TestCase;
use RobertN7\ExcelStyles\Examples\GroceryList;
use RobertN7\ExcelStyles\ExcelView;
use Exception;
use Maatwebsite\Excel\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

final class ExcelViewTest extends TestCase
{
    private function getExpectedGrocery() {
     return [
               0 => [
                  0 => [
                     0 => '',
                     ],

                  ],

               1 => [
                  0 => [
                     0 => '',
                     ],

                  ],

               2 => [
                  0 => [
                     0 => '',
                     ],

                  ],

               3 => [
                  0 => [
                     0 => '',
                     ],

                  ],

               4 => [
                  0 => [
                     0 => '',
                     ],

                  ],

               5 => [
                  0 => [
                     0 => 'Name',
                     1 => 'Quantity',
                     2 => '',
                     3 => '',
                     4 => 'Price',
                     5 => 'Total',
                     6 => 'Percentage of total',
                     ],

                  ],

               6 => [
                  0 => [
                     0 => 'Nuts',
                     1 => 25,
                     2 => '',
                     3 => '',
                     4 => 2.25,
                     5 => '=B7*E7',
                     6 => '=F7/F9',
                     ],

                  ],

               7 => [
                  0 => [
                     0 => 'Coconuts',
                     1 => 3,
                     2 => '',
                     3 => '',
                     4 => 1,
                     5 => '=B8*E8',
                     6 => '=F8/F9',
                     ],

                  ],

               8 => [
                  0 => [
                     0 => 'Total:',
                     1 => '=SUM(B7:B8)',
                     2 => '=SUM(C7:C8)',
                     3 => '=SUM(D7:D8)',
                     4 => '=SUM(E7:E8)',
                     5 => '=SUM(F7:F8)',
                     ],

                  ],

               ];
    }

    public function testCreateLayout()
    {
        $grocery_list = new GroceryList();
        $arr = ($grocery_list)->array();
        foreach($arr as $r1 => $els) {
            if(!count($els)) {
                throw new \Exception('Empty row at ' . $r1);
            }
            foreach($els as $r2 => $row) {
                if(!count($row)) {
                    throw new \Exception('Empty row at ' . $r1 . ':' . $r2);
                }
                foreach($row as $c => $col) {
                    if(is_object($col)) {
                        throw new \Exception('Invalid type of cell at ' . $r1 . ':' . $r2 . ':' .$c);
                    }
                }
            }
        }
        $this->assertEquals($arr, $this->getExpectedGrocery());

        // Check out the ranges layout is correct
        $this->assertEquals($grocery_list->header->start, 'A6');
        $this->assertEquals($grocery_list->header->end, 'G6');
        $this->assertEquals($grocery_list->body->start, 'A7');
        $this->assertEquals($grocery_list->body->end, 'G8');
        $this->assertEquals($grocery_list->footer->start, 'A9');
        $this->assertEquals($grocery_list->footer->end, 'F9');
        $this->assertEquals($grocery_list->main->start, 'A1');
        $this->assertEquals($grocery_list->main->end, 'G9');

    }

    private function loadSheet($view) {
        $list = new GroceryList;
        $array = $list->array();
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $row = 1;
        foreach($array as $a) {
            foreach($a as $r) {
                $col = 0;
                foreach($r as $c) {
                    $sheet->setCellValue(chr($col + 65) . $row, $c);
                    $col++;
                }
                $row++;
            }
        }
        $list->styles($sheet);
        return $spreadsheet;
    }

    public function testRender() {
        $writer = new Xlsx($this->loadSheet(new GroceryList));
        $writer->save('hello world.xlsx');
        $this->assertTrue(true);
    }

    private function dd($item, $indent=''): string
    {
        if(is_array($item)) {
            $s = '[' . PHP_EOL;
            $indent .= '   ';
            foreach($item as $k => $v) {
                $s .= $indent . $k  . ' => ' . $this->dd($v, $indent) . PHP_EOL;
            }
            return $s . $indent . '],' . PHP_EOL;
        } else if(is_object($item)) {
            return get_class($item) . ',';
        }
        return "'" . strVal($item) ."'" . ',';
    }
}