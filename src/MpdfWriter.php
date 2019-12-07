<?php

namespace alhimik1986\PhpExcelTemplator;

use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;

class MpdfWriter extends Mpdf
{
    /**
     * Gets the implementation of external PDF library that should be used.
     *
     * @param array $config Configuration array
     *
     * @return \Mpdf\Mpdf implementation
     */
    protected function createExternalWriterInstance($config)
    {
        $config = array_merge($config,[
            'autoScriptToLang' => true,
			'autoLangToFont' => true,
			'useSubstitutions' => true,
			'ignore_table_widths' => true,
        ]);
        $mpdf = new \Mpdf\Mpdf($config);
        $mpdf->baseScript = \Mpdf\Ucdn::SCRIPT_HAN;
        $mpdf->SetFooter('<div style="text-align: center; font-weight: bold;">Page {PAGENO} / {nbpg}</div>');
        // $mpdf->WriteHTML('	
		// 	<style>
		// 		@page {
		// 		footer: html_myFooter;
		// 		}
				
				
		// 		#footer {
		// 			position: absolute;
		// 			bottom: 15mm;
		// 			left: 30mm;
				
		// 			width: 50mm;
		// 			height: 8mm;
				
		// 			/* background: red;*/
		// 		}
		// 	</style>
			
			
		// 	<htmlpagefooter name="myFooter">
		// 		<div id="footer">
		// 			Page {PAGENO} / {nbpg}
		// 		</div>
        // 	</htmlpagefooter>');
        
        
        return $mpdf;
    }

        /**
     * Save Spreadsheet to file.
     *
     * @param string $pFilename Name of the file to save as
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @throws PhpSpreadsheetException
     */
    public function save($pFilename)
    {
        $fileHandle = parent::prepareForSave($pFilename);

        //  Default PDF paper size
        $paperSize = 'A4'; //    Letter    (8.5 in. by 11 in.)

        //  Check for paper size and page orientation
        if (null === $this->getSheetIndex()) {
            $orientation = ($this->spreadsheet->getSheet(0)->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet(0)->getPageSetup()->getPaperSize();
        } else {
            $orientation = ($this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getOrientation()
                == PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
            $printPaperSize = $this->spreadsheet->getSheet($this->getSheetIndex())->getPageSetup()->getPaperSize();
        }
        $this->setOrientation($orientation);

        //  Override Page Orientation
        if (null !== $this->getOrientation()) {
            $orientation = ($this->getOrientation() == PageSetup::ORIENTATION_DEFAULT)
                ? PageSetup::ORIENTATION_PORTRAIT
                : $this->getOrientation();
        }
        $orientation = strtoupper($orientation);

        //  Override Paper Size
        if (null !== $this->getPaperSize()) {
            $printPaperSize = $this->getPaperSize();
        }

        if (isset(self::$paperSizes[$printPaperSize])) {
            $paperSize = self::$paperSizes[$printPaperSize];
        }

        //  Create PDF
        $config = ['tempDir' => $this->tempDir];
        $pdf = $this->createExternalWriterInstance($config);
        $ortmp = $orientation;
        $pdf->_setPageSize(strtoupper($paperSize), $ortmp);
        $pdf->DefOrientation = $orientation;
        $pdf->AddPageByArray([
            'orientation' => $orientation,
            'margin-left' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getLeft()),
            // 'margin-right' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getRight()),
            'margin-top' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getTop()),
            'margin-bottom' => $this->inchesToMm($this->spreadsheet->getActiveSheet()->getPageMargins()->getBottom()),
            'pagenumstyle' => '1',
        ]);

        //  Document info
        $pdf->SetTitle($this->spreadsheet->getProperties()->getTitle());
        $pdf->SetAuthor($this->spreadsheet->getProperties()->getCreator());
        $pdf->SetSubject($this->spreadsheet->getProperties()->getSubject());
        $pdf->SetKeywords($this->spreadsheet->getProperties()->getKeywords());
        $pdf->SetCreator($this->spreadsheet->getProperties()->getCreator());

        

        $pdf->WriteHTML($this->generateHTMLHeader(false));
        $html = $this->generateSheetData();

        // $myfile = fopen("/Users/ponyhu/SourceCode/nmerp_fast/public/uploads/html.html", "w") or die("Unable to open file!");
        // fwrite($myfile, $html);
        // fclose($myfile);
        foreach (\array_chunk(\explode(PHP_EOL, $html), 1000) as $lines) {
            $pdf->WriteHTML(\implode(PHP_EOL, $lines));
            // fwrite($myfile, '---------------------------------');
            // fwrite($myfile, \implode(PHP_EOL, $lines));
        }

        // fwrite($myfile, $html);
        // fclose($myfile);
        
        $pdf->WriteHTML($this->generateHTMLFooter());

        //  Write to file
        fwrite($fileHandle, $pdf->Output('', 'S'));

        parent::restoreStateAfterSave($fileHandle);
    }

    /**
     * Convert inches to mm.
     *
     * @param float $inches
     *
     * @return float
     */
    private function inchesToMm($inches)
    {
        return $inches * 25.4;
    }
}
