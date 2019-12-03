<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class MergeCells
{
    /**
     * @param array $columnList [$columnIndex=>['isimg'=>true|false,'width'=>-1|int,'referenceColumnIndex'=>1,'startRowIndex'=>0|int,'rowNumber'=>0|int,'imgRoot'=>''],...]
     */
    static function MergeColumnCells(Worksheet $sheet,Array $columnList){
        foreach ($columnList as $columnIndex=>$column){
            $value = '';
            $b_row = $column['startRowIndex'];
            $len = $column['rowNumber'];
            $colString = Coordinate::stringFromColumnIndex($columnIndex);
            $referenceColumnIndex = $column['referenceColumnIndex'];
            $referenceColumnString = Coordinate::stringFromColumnIndex($referenceColumnIndex);
            for ($i = 0; $i < $len; $i++) {
                $rowIndex = $column['startRowIndex'] + $i;
                
                $coordinate = $colString . $rowIndex;
                $cur_value = $sheet->getCell($coordinate)->getValue();
                $cur_value = $cur_value=='' ? null : $cur_value;
                $referenceCoordinate = $referenceColumnString . $rowIndex;
                $cell = $sheet->getCell($referenceCoordinate);
                $needMerge = false;
                if($referenceColumnIndex != $columnIndex){
                    $range = $cell->getMergeRange();
                    if($range){
                        // $cellsRangeArr = Coordinate::splitRange($range);
                        if($rowIndex > $column['startRowIndex']){
                            if(!$sheet->getCell($referenceColumnString . ($rowIndex - 1))->isInRange($range)){
                                $needMerge = true;
                            }
                        }
                    }
                    else{
                        $needMerge = true;
                    }
                }
                if($value == '' || $value != $cur_value || $needMerge == true){
    
                    if($b_row < $rowIndex - 1){
                        $sheet->mergeCells($colString . $b_row . ':' . $colString . ($rowIndex - 1));
                    }
    
                    $value = $cur_value;

                    $b_row = $rowIndex;

                    if($column['isimg'] === true && ($value != '' || $value != null))
                    {
                        // $width = $sheet->getColumnDimension($colString)->getWidth();
                        $drawing = new Drawing();
                        $drawing->setPath($column['imgRoot'] . $value);
                        $drawing->setCoordinates($coordinate);
                        $drawing->setResizeProportional(true);
                        $drawing->setWidth($column['width']);
                        $drawing->setOffsetX(10);
                        $drawing->setOffsetY(10);
                        $drawing->setWorksheet($sheet);
                        // $defaultFont = $sheet->getStyle($coordinate)->getFont();
                        // $columnWidth = SharedDrawing::pixelsToPoints(80);
                        // $columnWidth = SharedDrawing::pixelsToCellDimension(80,$defaultFont);
                        // The templator copy style of cell, but not width. Let's make it manually
                        // $sheet->getColumnDimension($colString)->setWidth($columnWidth);


                    }
                }
                else if($i == $len-1){
                    if($b_row < $rowIndex){
                        $sheet->mergeCells($colString . $b_row . ':' . $colString . $rowIndex);
                    }
                }
                if($column['isimg'] === true)
                {
                    // Clear cell, which must contain just an image
                    $sheet->getCell($coordinate)->setValue(null);
                }
                // $mergeRange =  $sheet->getCell($colString . ($rowIndex-1))->getMergeRange();
                // $rowNumber = $sheet->rangeToArray($mergeRange);
                // $rowHeight = $drawing->getHeight();
                // $rowHeight = $rowHeight/$subCounts;
                //     // 计算行高，如果太小则不改变，避免挤压
                //     if($rowHeight > 20) {
                //         for($i = $startRow; $i < $curRow; $i++) {
                //             $sheet->getRowDimension($i)->setRowHeight($rowHeight);
                //         }
                //     }
                
            }
        }
        
    }
}