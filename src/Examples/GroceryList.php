<?php

namespace RobertN7\ExcelStyles\Examples;

use RobertN7\ExcelStyles\ExcelView;
use RobertN7\ExcelStyles\Bold;
use RobertN7\ExcelStyles\Border;
use RobertN7\ExcelStyles\AlignLeft;
use RobertN7\ExcelStyles\AlignCenter;
use RobertN7\ExcelStyles\AlignRight;
use RobertN7\ExcelStyles\BackgroundColor;
use RobertN7\ExcelStyles\Color;
use RobertN7\ExcelStyles\FontFamily;
use RobertN7\ExcelStyles\FontSize;
use RobertN7\ExcelStyles\Format;
use RobertN7\ExcelStyles\Italic;
use RobertN7\ExcelStyles\Underline;

class GroceryList extends ExcelView
{
    public function __construct()
    {
        parent::__construct();
        $this->groceries = [
            [
                'name' => 'Nuts',
                'quantity' => 25,
                'price' => 2.25,
            ],
            [
                'name' => 'Coconuts',
                'quantity' => 3,
                'price' => 1.00,
            ],
        ];
    }


    private function layout_data($item, $total_func)
    {
        return $this->tr(data: [
            $this->th(data: $item['name']),
            $item['quantity'],
            $this->td(n: 2),
            $this->td('$', $item['price']),
            $total = $this->td(data: $total_func),
            $this->td('%', function($e) use($total) {
                return '=' .$total->start . '/' .  $this->el('grand_total')->start;
            })
        ]);
    }

    protected function layout()
    {
        $total_func = function($e) {
           return '=B' . $e->start_row() . '*E'.$e->start_row();
        };
        $total_cols = function($e) {
           return $this->sum($e, rows: [$this->body->start_row(), $this->body->end_row()])($e);
        };
        $body_layout = [];
        foreach($this->groceries as $item)
            array_push($body_layout, $this->layout_data($item, $total_func));

        return $this->main = $this->div(data: [
            $this->tr(n: 5),
            $this->header = $this->tr('Header', [
                'Name',
                'Quantity',
                $this->td(n: 2),
                'Price',
                $this->th('th-total', 'Total'),
                'Percentage of total',
            ]),
            $this->body = $this->div('Body', $body_layout),
            $this->footer = $this->tr('Footer', [
                'Total:',
                $total_cols,
                $total_cols,
                $total_cols,
                $total_cols,
                $this->td('grand_total', function($e) {
                    $body = $this->el('Body');
                    return $this->sum($e, [$body->start_row(), $body->end_row()])($e);
                })
            ]),
        ]);
    }

    protected function init_styles()
    {
        return [
            'Header' => [
                new Bold(),
                new Border(),
                new Underline(),
                new AlignCenter(),
                new FontFamily('comic-sans'),
                new FontSize(18),
            ],
            'Body' => [
                new Border('thin'),
                new AlignLeft(),
            ],
            '%' => [
                new Format('%'),
                new Color('FFFFFF'),
                new BackgroundColor('333333'),
                new AlignRight(),
                new Italic,
            ],
            'Footer' => [
                new Underline(),
                new Bold(),
            ],
            '$' => [
                new Format('$'),
            ]
        ];
    }

}
