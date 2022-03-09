<?php

namespace RobertN7\ExcelStyles\Examples;

use RobertN7\ExcelStyles\ExcelView;
use RobertN7\ExcelStyles\Bold;
use RobertN7\ExcelStyles\Border;


class GroceryList extends ExcelView
{
    protected function layout()
    {
        return $this->div(data: [
            $this->tr('Header', [
                'Name',
                'Quantity',
            ]),
            $body = $this->div('Body', [
                $this->tr(data: [
                    'Nuts',
                    25,
                ]),
                $this->tr(data: [
                    'Coconuts',
                     function($e) {
                        $body = $this->el('Body');
                        return '=B' . $body->start[1];
                     },
                ]),
            ]),
            $this->tr('Footer', [
                'Total:',
                function($e) use ($body) {
                    return '=SUM(B' . $body->start[1]
                     . ':' . $body->end . ')';
                },
            ]),
        ]);
    }

    protected function init_styles()
    {
        return [
            'Header' => [
                new Bold(),
                new Border(),
            ],
            'Body' => [
                new Border('thin'),
            ]
        ];
    }

}
