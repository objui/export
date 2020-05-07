<?php
namespace objui\export\export;

use objui\export\export\csv\CsvReader;
use objui\export\export\IExcel;

class Csv implements IExcel
{
    protected static $config = [
        'file_path'   => '',
        'sheets'      => 1,
        'table_data'  => [],
        'table_header'=> [],
        'table_name'  => ''
    ];

    public function __construct(array $config = [])
    {
       self::$config = array_merge(self::$config, array_change_key_case($config));
       //判断系统类型
        $system = php_uname('s');
        $file_path = self::$config['file_path'];
        if($system == 'Windows NT'){
            $file_path = str_replace('/','\\',$file_path);
        }

         self::$config['file_path'] = $file_path;
    }

    /**
     * 批量导出
     */
    public function export()
    {
        ini_set('max_execution_time', '0');
        ini_set('memory_limit','-1');
        $file_name=self::$config['table_name'].date("YmdHis",time()).".csv";
        header ( 'Content-Type: application/vnd.ms-excel' );
        header ( 'Content-Disposition: attachment;filename='.$file_name );
        header ( 'Cache-Control: max-age=0' );
        ob_clean();
        $fp = fopen('php://output',"a");
        $limit=1000;
        $num = 0;
        $headlist =[];
        foreach (self::$config['table_header'] as $key => $value) {
            $headlist[$key] = iconv('utf-8', 'gbk', $value);
        }
        fputcsv($fp, $headlist);
        foreach (self::$config['table_data'] as $v){
            $num++;
            if($limit==$num){
                ob_flush();
                flush();
                $num=0;
            }
            foreach ($v as $t){
                $tarr[]=iconv('UTF-8', 'gbk',$t);
            }
            fputcsv($fp,$tarr);
            unset($tarr);
        }
        unset($data);
        fclose($fp);
        return true;
    }
   
    /**
     * 批量导入
     */
    public function import()
    {
        ini_set("max_execution_time", "0");
        ini_set("memory_limit", "-1");
        $csv = new CsvReader(self::$config['file_path']); 
        $data['list'] = $csv->get_data();
        return $data;
    }
    
    public function __destruct()
    {
        self::$config = null;
    }
}
