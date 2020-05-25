<?php
require '../vendor/autoload.php';
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
    $name = $worksheet->getCellByColumnAndRow(5, $row)->getValue();
    $sex = $worksheet->getCellByColumnAndRow(6, $row)->getValue();
    $nation = $worksheet->getCellByColumnAndRow(7, $row)->getValue();
    $birth = $worksheet->getCellByColumnAndRow(8, $row)->getValue();
    $phone = $worksheet->getCellByColumnAndRow(9, $row)->getValue();
    $yzName = $worksheet->getCellByColumnAndRow(11, $row)->getValue();
    $yzRelate = $worksheet->getCellByColumnAndRow(14, $row)->getValue();
    $hukouAdd = $worksheet->getCellByColumnAndRow(15, $row)->getValue();
    $isCHN = $worksheet->getCellByColumnAndRow(17, $row)->getValue();
    $email = $worksheet->getCellByColumnAndRow(19, $row)->getValue();
    $address = $worksheet->getCellByColumnAndRow(20, $row)->getValue();
    $parentName = $worksheet->getCellByColumnAndRow(21, $row)->getValue();
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle($name);
    //$sheet->getDefaultRowDimension()->setRowHeight(50);
    $sheet->getDefaultColumnDimension()->setWidth(20);
    $sheet->getRowDimension('1')->setRowHeight(40);
    $sheet->getRowDimension('2')->setRowHeight(40);
    $sheet->getRowDimension('3')->setRowHeight(40);
    $sheet->getRowDimension('4')->setRowHeight(40);

    $sheet->setCellValue('A1','姓名');
    $sheet->setCellValue('B1',$name);
    $sheet->setCellValue('C1', '性别');
    $sheet->setCellValue('D1',$sex);
    $sheet->setCellValue('E1', '民族');
    $sheet->setCellValue('F1',$nation);
    $sheet->setCellValue('G1', '中国国籍');
    $sheet->setCellValue('H1',$isCHN);

    $sheet->setCellValue('A2','出生日期');
    $sheet->setCellValue('B2',$birth);
    $sheet->mergeCells('B2:D2');
    $sheet->setCellValue('E2','户口地址');
    $sheet->setCellValue('F2',$hukouAdd);
    $sheet->mergeCells('F2:H2');

    $sheet->setCellValue('A3','手机号码');
    $sheet->setCellValue('B3',$phone);
    $sheet->mergeCells('B3:D3');
    $sheet->setCellValue('E3','业主姓名');
    $sheet->setCellValue('F3',$yzName);
    $sheet->setCellValue('G3','与业主关系');
    $sheet->setCellValue('H3',$yzRelate);

    $sheet->setCellValue('A4','父母姓名');
    $sheet->setCellValue('B4',$parentName);
    $sheet->mergeCells('B4:C4');
    $sheet->setCellValue('D4','春泽庄南');
    $sheet->setCellValue('E4',$address);
    $sheet->setCellValue('F4','邮箱');
    $sheet->setCellValue('G4',$email);
    $sheet->mergeCells('G4:H4');

   

    // $sheet->getColumnDimension('A')->setWidth(30);
    // $sheet->getColumnDimension('B')->setWidth(30);
    // $sheet->getColumnDimension('C')->setWidth(30);
    //$sheet->getColumnDimension('B')->setAutoSize(true);
    $sheet->getStyle('A1:H4')->getFont()->setName('Arial')
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
    $sheet->getStyle('A1:H4')->applyFromArray($styleArray);

    $title = ($row-1).'-'.$name.'.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save('outexcel/'.($row-1).'-'.$name.'.xlsx');
    echo("<a href='http://gl.strtv.cn/spread/ef/outexcel/".$title."'>".$checkTime."-".$title."</a><br>");
}
echo("分解成功！");
exit;
