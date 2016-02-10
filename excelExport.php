<?php
public static function createExcel(/*$title, $header, $data*/){
    /** Пример входных данных **/
    $data   = array(
        array(
            'Адреналин',
            '03.04.2014',
            '35 бал. + 1 амп.',
            '54 руб.',
        ),
        array(
            'jirtgoo',
            '28.05.2014',
            '3 бал. + 1 амп.',
            '6 руб.',
        ),
    );
    // storeresidualsController: count($title) == 1
    // expenseController:        count($title) == 4
    $title  = array('Отчет «Общее списание»',
                    '14.08.2012',
                    '06.08.2014',
                    'Все группы');
    $header = array('Наименование',
                    'Дата накладной',
                    'Остаток',
                    'Цена');
    /*************  ************/
    $columns = count($header);//количество столбцов
    $indexes = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J');
    Yii::import('ext.phpexcel.XPHPExcel');
    $objPHPExcel= XPHPExcel::createPHPExcel();
    $objPHPExcel->getProperties()
        ->setCreator        (Yii::app()->user->name)
        ->setLastModifiedBy (Yii::app()->user->name);
    /** Автоматическая ширина ячеек **/
    for($i = 0; $i < $columns; $i++)
        $objPHPExcel->getActiveSheet()->getColumnDimension($indexes[$i])->setAutoSize(true);
    /** Title style **/
    $objPHPExcel->getActiveSheet()->mergeCells('A1:'.$indexes[$columns-1].'1');
    $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(
        PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle("A1")->getFont()->setSize(14);
    /** Header style **/
    $objPHPExcel->getActiveSheet()->getStyle('A2:'.$indexes[$columns-1].'2')->getFill()
        ->applyFromArray(array('type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array('rgb' => 'c3d9ff')
        ));
    $objPHPExcel->getActiveSheet()->getStyle('A2:'.$indexes[$columns-1].'2')->getFont()->setBold(true);
    //$i = ($titleCnt == 1) ? 3 : 4;
    $i = 3;
    foreach($data as $r){
        /** Автоматическое заполнение title **/
        switch(count($title)){
            case 1:
                $objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow(0, 1, $title[0]);
                /*$dataStart = 2;*/ break;
            case 4:
                $objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow(0, 1, $title[0].'_'.$title[1].' - '.$title[2].'('.$title[3].')');
                /*$dataStart = 3;*/ break;
            default: throw new CHttpException(400, '$title array is invalid'); break;
        }
    //TODO: Пропал header
        /** Автоматическое заполнение header **/
        for($j = 0; $j < $columns; $j++)
            $objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow($j, 2, $header[$j]);
        /** Автоматическое заполнение data **/
        for($j = 0; $j < $columns; $j++)
            $objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow($j, $i, $r[$j]);
        $i++;
    }
    $objPHPExcel->getActiveSheet()->setTitle($title[0]);
    $objPHPExcel->setActiveSheetIndex(0);
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename = '.$title[0].'.xls');
    header('Cache-Control: max-age=0');
    header('Cache-Control: max-age=1');
    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
    header ('Cache-Control: cache, must-revalidate');
    header ('Pragma: public');
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');
    Yii::app()->end();
}