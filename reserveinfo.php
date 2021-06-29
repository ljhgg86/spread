<?php
require 'vendor/autoload.php';
require 'conf/conn.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load('sql1.xlsx'); //载入excel表格
$worksheet = $spreadsheet->getActiveSheet();
$highestRow = $worksheet->getHighestRow(); // 总行数
$highestColumn = $worksheet->getHighestColumn(); // 总列数
$sql = "";
for($row=2;$row<=$highestRow;++$row){
    $name = strval($worksheet->getCellByColumnAndRow(1, $row)->getValue());
    $realName = $worksheet->getCellByColumnAndRow(2, $row)->getValue();
    $openId = $worksheet->getCellByColumnAndRow(4, $row)->getValue();
    $nickName = $worksheet->getCellByColumnAndRow(5, $row)->getValue();
    $cellphone = $worksheet->getCellByColumnAndRow(7, $row)->getValue();
    $officephone = $worksheet->getCellByColumnAndRow(8, $row)->getValue() ?:"(NULL)";
    if($row == $highestRow){
        $sql.="('".$name."','".$realName."','".$cellphone."','".$openId."','".$nickName."','".$cellphone."','".$officephone."')";
    }
    else{
        $sql.="('".$name."','".$realName."','".$cellphone."','".$openId."','".$nickName."','".$cellphone."','".$officephone."'),";
    }
    
}

// $sql = "insert into users (name, realName, password, openId, nickName, cellphone, officephone) values ".$sql;
// if($stmt = $mysqli->prepare($sql))
// {
//     $stmt->execute();
//     $stmt->close();
// }
// $mysqli->close();
exit;
