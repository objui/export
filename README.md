### 说明
使用PHP开发，批量导入导出，支持csv、xls、xlsx、doc、docx格式文件批量处理

### 安装
```
composer require objui/export
```

使用示例
```
#CSV导入
$csv_import_config = ['file_path'=>'./test.csv'];
$csv_import_obj = new Export($csv_import_config);
$data = $csv_import_obj->import();
var_dump($data);


#CSV导出
$csv_export_config = [
    'table_header' =>['id','name'],
    'table_data'   => [
       ['id'=>1,'name'=>'objui'     ]
     ],
    'suffix' => 'csv'
];
$obj = new Export($csv_export_config);
$obj->export();

#xls导入
$config = [
    'file_path' => './test.xls'
];
$obj = new Export($config);
$result = $obj->import();
var_dump($result);


#xls导出
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
$config = [
    'file_path' => './test.xlsx'
];
$obj = new Export($config);
$result = $obj->import();
var_dump($result);


#xlsx导出
$config = [
    'table_header' => ['id', 'name'],
    'table_data'   => [
        ['id'=>1, 'name'=>'objui']
    ],
    'suffix' => 'xlsx'
];
$obj = new Export($config);
$obj->export();


```
