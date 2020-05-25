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
    $checkTime = $worksheet->getCellByColumnAndRow(1, $row)->getValue();
    $name = $worksheet->getCellByColumnAndRow(4, $row)->getValue();
    $sex = $worksheet->getCellByColumnAndRow(5, $row)->getValue();
    $nation = $worksheet->getCellByColumnAndRow(6, $row)->getValue();
    $birth = $worksheet->getCellByColumnAndRow(7, $row)->getValue();
    $phone = $worksheet->getCellByColumnAndRow(8, $row)->getValue();
    $address = $worksheet->getCellByColumnAndRow(9, $row)->getValue();
    $jiguan = $worksheet->getCellByColumnAndRow(10, $row)->getValue();
    $isOnly = $worksheet->getCellByColumnAndRow(11, $row)->getValue();
    $isLS = $worksheet->getCellByColumnAndRow(12, $row)->getValue();
    $birthAdd = $worksheet->getCellByColumnAndRow(13, $row)->getValue();
    $hukouAdd = $worksheet->getCellByColumnAndRow(14, $row)->getValue();
    $hukouPCS = $worksheet->getCellByColumnAndRow(15, $row)->getValue();
    $isCHN = $worksheet->getCellByColumnAndRow(16, $row)->getValue();
    $identity = strval($worksheet->getCellByColumnAndRow(17, $row)->getValue());
    $email = $worksheet->getCellByColumnAndRow(18, $row)->getValue();
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
    $sheet->setCellValue('A1','姓名');
    $sheet->setCellValue('B1',$name);
    $sheet->setCellValue('C1', '性别');
    $sheet->setCellValue('D1',$sex);
    $sheet->setCellValue('E1', '民族');
    $sheet->setCellValue('F1',$nation);
    $sheet->setCellValue('G1', '籍贯');
    $sheet->setCellValue('H1',$jiguan);

    $sheet->setCellValue('A2','出生日期');
    $sheet->setCellValue('B2',$birth);
    $sheet->mergeCells('B2:C2');
    $sheet->setCellValue('D2','身份证号码');
    $sheet->mergeCells('D2:E2');
    $sheet->getCell('F2')->setDataType('inlineStr');
    $sheet->getStyle('F2')->getNumberFormat()->setFormatCode('0');
    $sheet->setCellValue('F2',$identity.' ');
    $sheet->mergeCells('F2:H2');

    $sheet->setCellValue('A3','手机号码');
    $sheet->setCellValue('B3',$phone);
    $sheet->mergeCells('B3:D3');
    $sheet->setCellValue('E3','邮箱地址');
    $sheet->setCellValue('F3',$email);
    $sheet->mergeCells('F3:H3');

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

    $sheet->setCellValue('A6','中国国籍');
    $sheet->setCellValue('B6',$isCHN);
    $sheet->setCellValue('c6','户口地址');
    $sheet->setCellValue('D6',$hukouAdd);
    $sheet->mergeCells('D6:E6');
    $sheet->setCellValue('F6','户口派出所');
    $sheet->mergeCells('G6:H6');
    $sheet->setCellValue('G6',$hukouPCS);

   

    // $sheet->getColumnDimension('A')->setWidth(30);
    // $sheet->getColumnDimension('B')->setWidth(30);
    // $sheet->getColumnDimension('C')->setWidth(30);
    //$sheet->getColumnDimension('B')->setAutoSize(true);
    $sheet->getStyle('A1:H6')->getFont()->setName('Arial')
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
    $sheet->getStyle('A1:H6')->applyFromArray($styleArray);

    $title = ($row-1).'-'.$name.'.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save('outexcel/'.($row-1).'-'.$name.'.xlsx');
    echo("<a href='http://gl.strtv.cn/spread/outexcel/".$title."'>".$checkTime."-".$title."</a><br>");
}
echo("分解成功！");
exit;
