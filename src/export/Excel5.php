<?php
/**
 * Excel 2003版本批量导入导出
 * Author: objui@qq.com
 * Date: 2020/05/06
 */
namespace objui\export\export;

use objui\export\export\IExcel;
require 'phpexcel/PHPExcel.php';

class Excel5 implements IExcel
{
    protected static $config = [
        'file_path'     =>  '',     //文件路径
        'sheets'        =>  1 ,     //数据表数
        'table_data'    => [],      //导出表格数据
        'table_header'  => [],      //导出表头
        'table_name'    => ''       //导出表格名称
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

    public function import()
    {
        ini_set('max_execution_time', '0');
        ini_set('memory_limit','-1');
        $data =[];
        if(self::$config['file_path'] == null || !file_exists(self::$config['file_path'])){
            //缺少参数
            $data['error'] = '缺少参数';
        }else{
            $PHPReader = new \PHPExcel_Reader_Excel5();
            $file_name = self::$config['file_path'];
            $Excel = $PHPReader->load($file_name);
            for($a=1;$a<=self::$config['sheets'];$a++){
                $sheet = "sheet".$a;
                $Column = "Column".$a;
                $Row = "Row".$a;
                $$sheet = $Excel->getSheet($a-1);           //选择第几个表
                $$Column = $$sheet->getHighestColumn();     //获取总列数
                $$Row = $$sheet->getHighestRow();           //获取总行数

                $data['list'][$a] = array();                //用于保存Excel中的数据
                for($i=1;$i<=$$Row;$i++){
                    //循环获取表中的数据，$i表示当前行,索引值从0开始
                    for($j='A';$j<=$$Column;$j++){          //从哪列开始，A表示第一列
                        $address=$j.$i;                     //数据坐标
                        $data['list'][$a][$i-1][$j]=trim((string)$$sheet->getCell($address)->getValue());//读取到的数据，保存到数组中
                    }
                }
            }

        }
        return $data;
    }

    public function export()
    {
        ini_set('max_execution_time', '0');
        ini_set('memory_limit','-1');

        //引入第三方类库
        $excel = new \PHPExcel();
        $letter = array(
            'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S',
            'T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ',
            'AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ'
        );

        $table_header = self::$config['table_header'];
        $table_data = self::$config['table_data'];
        $table_name = self::$config['table_name'];

        for($i = 0;$i < count($table_header);$i++) {
            $letter_name = $letter[$i] . '1';
            $header_name = $table_header[$i];
            $excel->getActiveSheet()->setCellValue("$letter_name","$header_name");
        }

        //填充表格信息
        for ($i = 2;$i <= count($table_data)+1;$i++) {
            $j = 0;
            foreach ($table_data[$i - 2] as $key=>$value) {
                $letter_tmp = $letter[$j] . $i;
                $excel->getActiveSheet()->setCellValue("$letter_tmp","$value");
                $j++;
            }
        }

        if($table_name==''){
            $table_name=date('YmdHis',time());
        }

        //创建Excel输入对象
        $write = new \PHPExcel_Writer_Excel5($excel);
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.ms-execl");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename="'.$table_name.'.xls"');
        header("Content-Transfer-Encoding:binary");
        ob_clean();
        $write->save('php://output');
    }
}
