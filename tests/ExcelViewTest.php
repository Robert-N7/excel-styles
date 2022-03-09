<?php declare(strict_types=1);

namespace Tests;

use PHPUnit\Framework\TestCase;
use RobertN7\ExcelStyles\Examples\GroceryList;

final class ExcelViewTest extends TestCase
{
    public function testCreate()
    {
        $arr = (new GroceryList())->array();
        $this->assertEquals(count($arr), 4);
    }
}