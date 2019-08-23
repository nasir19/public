<?php
namespace fast;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
/**
 * 该扩展依赖
 * 安装：composer require phpoffice/phpspreadsheet:1.8.*
 * Created by PhpStorm.
 * User: godluck
 * Date: 2019/8/11
 * Time: 17:48
 */
final class Excel
{
    /**
     * 支持文件转换为数组的格式
     * @var array
     */
    protected static $readers = ['xls','xlsx','xml','csv','html'];

    /**
     * 支持导出的格式
     * @var array
     */
    protected static $writers = ['xls','xlsx','csv','html'];

    /**
     * 文件转换为数组数据
     * @param String $filePath Excel文件全路径[文件名称]
     * @param string $exts  文件后缀 默认是 xlsx
     * @param bool $match 是否匹配转换 url和时间 默认不匹配
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function fileToArray($filePath, $exts = 'xlsx', $match = false)
    {
        // 检测格式
        $exts = strtolower($exts);
        if (!in_array($exts,self::$readers))
        {
            exception('文件转换为数组,目前只支持‘xls,xlsx,xml,csv,Html’类型的格式');
        }
        // 实例化对应的格式对象
        $inputFileType = ucfirst($exts);
        $reader  = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
        // 读取Excel文件
        $PHPExcel = $reader->load($filePath);
        // 获取表中的第一个工作表，如果要获取第二个，把0改为1，依次类推
        $currentSheet = $PHPExcel->getSheet(0);
        // 获取总列数
        $allColumn = $currentSheet->getHighestColumn();
        // 获取总行数
        $allRow = $currentSheet->getHighestRow();
        //循环读取数据，默认编码是utf8
        $data = [];
        // 循环获取表中的数据，$currentRow表示当前行，从哪行开始读取数据，索引值从0开始
        for($currentRow = 1;$currentRow<=$allRow;$currentRow++)
        {
            // 从哪列开始，A表示第一列
            for($currentColumn='A';$currentColumn<=$allColumn;$currentColumn++)
            {
                // 数据坐标
                $address = $currentColumn.$currentRow;
                // 获取单元格的值
                $v = $currentSheet->getCell($address)->getValue();
                if ($match == true) {
                    // 匹配http/https
                    $preg_url = "/^((http|https):\/\/)+[\w-_.]+(\/[\w-_]+)*\/?$/";
                    // 过滤时间
                    if (preg_match("/^[0-9]{5}.[0-9]{1,20}$/",$v) && strrpos($v,'.') && \PhpOffice\PhpSpreadsheet\Shared\Date::isDateTime($currentSheet->getCell($address)))
                    {
                        // 匹配时间并且进行格式化
                        $data[$currentRow-1][$currentColumn] = gmdate("Y/m/d H:i", \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($v));
                    } else if (preg_match($preg_url,$v)) {
                        // 匹配url
                        $data[$currentRow-1][$currentColumn] = "<a href='$v'>$v</a>";
                    } else {
                        $data[$currentRow-1][$currentColumn] = $v;
                    }
                } else {
                    $data[$currentRow-1][$currentColumn] = $v;
                }
            }
        }
        //处理空白数组
        $res = [];
        if ($data)
        {
            foreach($data as $k => $v)
            {
                $res[] = array_values($v);
            }
        }
        return $res;
    }

    /**
     * 导出Excel文件
     * @param string $fileName 文件名称
     * @param array $excelColumnItem 设置Excel表格第一行的数据
     * @param array $data 需要处理的数据
     * @param string $exts 导出的格式 [xls,xlsx,csv,html]
     * @param string|bool $width 设置列的宽度 默认自适应
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * 处理表格中 银行卡号或者数字长度问题 可用 chunk_split()函数进行处理
     * chunk_split(string 要分割的字符 , int 分割位数, string 分割字符);
     * 使用案例 chunk_split('201701061051',4," ");已4位数字分割,中间已空格隔开。
     * 导出csv时,如果用Excel打开 可能会出现乱码 用txt打开不会出现 出现乱码时,需要转换一下内部编码将UTF-8转换为ASCII
     * 如果转换 点击文件用txt打开 之后另存为 底部有个编码选择 选择ASCII 保存之后 打开！
     */
     public static function export($fileName = '', $excelColumnItem = [], $data = [], $exts = 'xlsx', $width = true)
     {
         // 初始化数据 并且检测格式
         $exts = strtolower($exts);
         $date = date("Ymd", time());
         if (!in_array($exts,self::$writers)) {
             exception('文件转换为数组,目前只支持‘xls,xlsx,xml,csv’类型的格式');
         }
         $inputFileType = ucfirst($exts);
         $fileName .= "_{$date}." . $exts;
         // 获取对象
         $spreadsheet = new Spreadsheet();
         // 获取列
         $sheet = $spreadsheet->getActiveSheet();
         // 设置表头
         $key = ord("A");
         foreach ($excelColumnItem as $v)
         {
            $colum = chr($key);
            // 设置列的宽度
            if (is_bool($width) || is_numeric($width))
            {
                if (is_numeric($width)) {
                    $sheet->getColumnDimension($colum)->setWidth($width);
                } else {
                    $sheet->getColumnDimension($colum)->setAutoSize($width);
                }
            } else {
                $sheet->getColumnDimension($colum)->setAutoSize(true);
            }
            // 给具体列 赋值
            $sheet->setCellValue($colum . '1', $v);
            $key += 1;
         }
         $column = 2;
         // 循环数据 将具体值赋值具体列之中
         foreach ($data as $key => $rows)
         {
            $span = ord("A");
            foreach ($rows as $keyName => $value)
            {
                $j = chr($span);
                $sheet->setCellValue($j . $column, strip_tags($value));
                $span ++;
            }
            $column ++;
         }
         // 处理完毕 导出数据文件到浏览器
         header('Content-Type: application/vnd.ms-excel;charset=UTF-8');
         header("Content-Disposition: attachment;filename=\"$fileName\"");
         header('Cache-Control: max-age=0');
         $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, $inputFileType);
         $writer->save('php://output');//文件通过浏览器下载
         return true;
     }
}