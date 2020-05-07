<?php
namespace objui\export\export;

interface IExcel
{
    public function __construct(array $config = []);
    public function export();
    public function import();
}
