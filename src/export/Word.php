<?php
/**
 * @desc Word文档处理
 * @author: objui
 * @version: v1.0
 */
namespace objui\export\export;

use objui\export\export\IExcel;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Element\TextBreak;
use PhpOffice\PhpWord\Style\Font;
use PhpOffice\PhpWord\Shared\ZipArchive;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\TablePosition;

class Word implements IExcel
{
    protected static $config = [
        'file_path'   => '',
        'is_tags'     => 1,
        'table_name'  => '',
        'suffix'      => 'docx',
        'table_data'  => '',
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

    //导出
    public function export()
    {
        ini_set('max_execution_time', '0');
        ini_set('memory_limit','-1');
        $filename=self::$config['table_name'].date("YmdHis",time()).'.'. self::$config['suffix'];
        $phpword = new PhpWord();
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $section->addText(self::$config['table_data']);
        header("Content-Description:File Transfer");
        header('Content-Disposition:attachment;filename='.$filename);
        header("Expires:0");
        ob_clean();
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
        Header("Content-type:application/octet-stream");
        $objWriter->save('php://output');

    }

    //导入
    public function import()
    {
        ini_set("max_execution_time", "0");
        ini_set("memory_limit", "-1");
        $filePath = self::$config['file_path'];
        $phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath);
        $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, "HTML");
        $arr = $xmlWriter->getContent(); 
        preg_match('/<body>[\s\S]*?<\/body>/iu', $arr, $content);   
        $result = rtrim(ltrim($content[0], '<body>'),'</body>');
        if(self::$config['is_tags'] == 1) {
            $result = strip_tags($result, '<p><a><img><table><tr><td>');
        }
        return $result;
    }

    public function __destruct()
    {
        self::$config = null;
    }
} 
