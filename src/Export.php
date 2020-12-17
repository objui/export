<?php
namespace objui\export;

class Export
{
    protected static $config = [
        'file_path'   => '',
        'sheets'      => 1,
        'table_data'  => '',
        'table_header'=> '',
        'table_name'  => '',
        'suffix'      => 'xls'
    ];

    static public $obj;

    public function __construct(array $config = [])
    {
        self::$config = array_merge(self::$config, array_change_key_case($config));
        $this->getInstance();
    }

    public function getInstance()
    {
        if (empty(self::$obj)) {
            if(self::$config['file_path'] == ''){
                $suffix = self::$config['suffix'];
            }else{
                $suffix = substr(strrchr(self::$config['file_path'], '.'), 1);
            }
            switch($suffix){
                case 'xlsx':
                    self::$obj = new export\Excel2007(self::$config);
                    break;
                case 'csv':
                    self::$obj = new export\Csv(self::$config);
                    break;
              
                case 'doc':
                case 'docx':
                    self::$obj = new export\Word(self::$config);
                    break;
                default:
                    //Excel2003
                    self::$obj = new export\Excel5(self::$config);
            }         
        } 
        return self::$obj;
    }
    
    /**
     * 导出
     */
    public function export()
    {
        $obj = self::$obj;
        $obj->export();
    }

    /**
     * 导入
     */
    public function import()
    {
         $obj = self::$obj;
         $data = $obj->import();
         return $data;
    }
}
