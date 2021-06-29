<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load('sheet.xlsx'); //载入excel表格
$worksheet = $spreadsheet->getActiveSheet();
$highestRow = $worksheet->getHighestRow(); // 总行数
$highestColumn = $worksheet->getHighestColumn(); // 总列数
//$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 5
for($row=2;$row<=$highestRow;++$row){
    //提交时间
    $checkTime = $worksheet->getCellByColumnAndRow(1, $row)->getValue();
    $name = $worksheet->getCellByColumnAndRow(5, $row)->getValue();
    $birth = $worksheet->getCellByColumnAndRow(6, $row)->getValue();
    $phone = $worksheet->getCellByColumnAndRow(7, $row)->getValue();
    //是否独生
    $isOnly = $worksheet->getCellByColumnAndRow(8, $row)->getValue();
    //是否留守
    $isLS = $worksheet->getCellByColumnAndRow(9, $row)->getValue();
    //出生地
    $birthAdd = $worksheet->getCellByColumnAndRow(10, $row)->getValue();
    //户口地
    $hukouAdd = $worksheet->getCellByColumnAndRow(11, $row)->getValue();
    //是否港澳台侨
    $isGAT = $worksheet->getCellByColumnAndRow(12, $row)->getValue();
    $email = $worksheet->getCellByColumnAndRow(13, $row)->getValue();
    $sex = $worksheet->getCellByColumnAndRow(14, $row)->getValue();
    $address = $worksheet->getCellByColumnAndRow(15, $row)->getValue();
    //证件类型
    $credentials = $worksheet->getCellByColumnAndRow(16, $row)->getValue();
    //户口性质
    $hukou = $worksheet->getCellByColumnAndRow(17, $row)->getValue();
    //家庭人数
    $familyPersons = strval($worksheet->getCellByColumnAndRow(18, $row)->getValue());
    $fatherName = $worksheet->getCellByColumnAndRow(19, $row)->getValue();
    //父亲工作单位
    $fatherEmployer = $worksheet->getCellByColumnAndRow(20, $row)->getValue();
    //父亲职务
    $fatherPost = $worksheet->getCellByColumnAndRow(21, $row)->getValue();
    $fatherPhone = $worksheet->getCellByColumnAndRow(22, $row)->getValue();
    $motherName = $worksheet->getCellByColumnAndRow(23, $row)->getValue();
    $motherEmployer = $worksheet->getCellByColumnAndRow(24, $row)->getValue();
    $motherPost = $worksheet->getCellByColumnAndRow(25, $row)->getValue();
    $motherPhone = $worksheet->getCellByColumnAndRow(26, $row)->getValue();
    //证件号码
    $identity = strval($worksheet->getCellByColumnAndRow(27, $row)->getValue());
    //籍贯
    $jiguan = $worksheet->getCellByColumnAndRow(28, $row)->getValue(); 
    
    
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle($name);
    //$sheet->getDefaultRowDimension()->setRowHeight(50);
    $sheet->getDefaultColumnDimension()->setWidth(20);
    $sheet->getRowDimension('1')->setRowHeight(40);
    $sheet->getRowDimension('2')->setRowHeight(40);
    $sheet->getRowDimension('3')->setRowHeight(40);
    $sheet->getRowDimension('4')->setRowHeight(40);
    $sheet->getRowDimension('5')->setRowHeight(40);
    $sheet->getRowDimension('6')->setRowHeight(40);
    $sheet->getRowDimension('7')->setRowHeight(40);
    $sheet->getRowDimension('8')->setRowHeight(40);
    $sheet->getRowDimension('9')->setRowHeight(40);
    $sheet->getRowDimension('10')->setRowHeight(40);
    $sheet->setCellValue('A1','姓名');
    $sheet->setCellValue('B1',$name);
    $sheet->setCellValue('C1', '性别');
    $sheet->setCellValue('D1',$sex);
    // $sheet->setCellValue('E1', '户口性质');
    // $sheet->setCellValue('F1',$hukou);
    $sheet->setCellValue('E1', '籍贯');
    $sheet->setCellValue('F1',$jiguan);
    $sheet->mergeCells('F1:H1');

    $sheet->setCellValue('A2','出生日期');
    $sheet->setCellValue('B2',$birth);
    $sheet->mergeCells('B2:C2');
    // $sheet->setCellValue('D2','港澳台侨');
    // $sheet->setCellValue('E2',$isGAT);
    $sheet->setCellValue('D2', '户口性质');
    $sheet->setCellValue('E2',$hukou);
    $sheet->setCellValue('F2',$credentials);
    $sheet->getCell('G2')->setDataType('inlineStr');
    $sheet->getStyle('G2')->getNumberFormat()->setFormatCode('0');
    $sheet->setCellValue('G2',$identity.' ');
    $sheet->mergeCells('G2:H2');

    $sheet->setCellValue('A3','电话号码');
    $sheet->setCellValue('B3',$phone);
    $sheet->mergeCells('B3:C3');
    $sheet->setCellValue('D3','邮箱地址');
    $sheet->setCellValue('E3',$email);
    $sheet->mergeCells('E3:F3');
    $sheet->setCellValue('G3','港澳台侨');
    $sheet->setCellValue('H3',$isGAT);

    $sheet->setCellValue('A4','家庭住址');
    $sheet->setCellValue('B4',$address);
    $sheet->mergeCells('B4:H4');

    $sheet->setCellValue('A5','出生地');
    $sheet->setCellValue('B5',$birthAdd);
    $sheet->mergeCells('B5:D5');
    $sheet->setCellValue('E5','独生子女');
    $sheet->setCellValue('F5',$isOnly);
    $sheet->setCellValue('G5','留守儿童');
    $sheet->setCellValue('H5',$isLS);

    $sheet->setCellValue('A6','户口地址');
    $sheet->setCellValue('B6',$hukouAdd);
    $sheet->mergeCells('B6:E6');
    $sheet->setCellValue('F6','家庭成员（人数）');
    $sheet->mergeCells('F6:G6');
    $sheet->setCellValue('H6',$familyPersons);

    $sheet->setCellValue('A7','父亲姓名');
    $sheet->setCellValue('B7',$fatherName);
    $sheet->mergeCells('B7:D7');
    $sheet->setCellValue('E7','父亲电话');
    $sheet->setCellValue('F7',$fatherPhone);
    $sheet->mergeCells('F7:H7');
    $sheet->setCellValue('A8','父亲单位');
    $sheet->setCellValue('B8',$fatherEmployer);
    $sheet->mergeCells('B8:D8');
    $sheet->setCellValue('E8','父亲职务');
    $sheet->setCellValue('F8',$fatherPost);
    $sheet->mergeCells('F8:H8');

    $sheet->setCellValue('A9','母亲姓名');
    $sheet->setCellValue('B9',$motherName);
    $sheet->mergeCells('B9:D9');
    $sheet->setCellValue('E9','母亲电话');
    $sheet->setCellValue('F9',$motherPhone);
    $sheet->mergeCells('F9:H9');
    $sheet->setCellValue('A10','母亲单位');
    $sheet->setCellValue('B10',$motherEmployer);
    $sheet->mergeCells('B10:D10');
    $sheet->setCellValue('E10','母亲职务');
    $sheet->setCellValue('F10',$motherPost);
    $sheet->mergeCells('F10:H10');

    // $sheet->getColumnDimension('A')->setWidth(30);
    // $sheet->getColumnDimension('B')->setWidth(30);
    // $sheet->getColumnDimension('C')->setWidth(30);
    //$sheet->getColumnDimension('B')->setAutoSize(true);
    $sheet->getStyle('A1:H10')->getFont()->setName('Arial')
    ->setSize(16);
    $styleArray = [
        'borders' => [
            // 'outline' => [
            //     'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            //     'color' => ['argb' => '00000000'],
            // ],
            'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]
        ],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],
    ];
    $sheet->getStyle('A1:H10')->applyFromArray($styleArray);

    $title = ($row-1).'-'.$name.'.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save('outexcel/'.($row-1).'-'.$name.'.xlsx');
    echo("<a href='http://gl.strtv.cn/spread/outexcel/".$title."'>".$checkTime."-".$title."</a><br>");
}
echo("分解成功！");
exit;
