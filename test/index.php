<?php
use objui\export\Export;

require '../vendor/autoload.php';

#CSV导入
/*
$csv_import_config = ['file_path'=>'./test.csv'];
$csv_import_obj = new Export($csv_import_config);
$data = $csv_import_obj->import();
var_dump($data);
*/

#CSV导出
/*
$csv_export_config = [
    'table_header' =>['id','name'],
    'table_data'   => [
       ['id'=>1,'name'=>'objui'     ]
     ],
    'suffix' => 'csv'
];
$obj = new Export($csv_export_config);
$obj->export();
*/

#xls导入
/*
$config = [
    'file_path' => './test.xls'
];
$obj = new Export($config);
$result = $obj->import();
var_dump($result);
*/


#xls导出
/*
$config = [
     'table_header' =>['id','name'],
     'table_data'   => [        
       ['id'=>1,'name'=>'objui'     ]
     ],   
    'suffix' => 'xls'
];
$obj = new Export($config);
$obj->export();
*/

#xlsx导入
/*
$config = [
    'file_path' => './test.xlsx'
];
$obj = new Export($config);
$result = $obj->import();
var_dump($result);
*/

#xlsx导出
/*
$config = [
    'table_header' => ['id', 'name'],
    'table_data'   => [
        ['id'=>1, 'name'=>'objui']
    ],
    'suffix' => 'xlsx'
];
$obj = new Export($config);
$obj->export();
*/

#WORD导入
/*
$config = [
    'file_path' => 'test.docx'
];
$obj = new Export($config);
$result = $obj->import();
var_dump($result);
*/

#word 导出
$config = [
   'table_data' => '姓名：objui <br>
网址：https://www.objui.com',
  'suffix' => 'docx'
];

$obj = new Export($config);
$obj->export(); 



