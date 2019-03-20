<?php
include 'vendor/autoload.php';
include 'main.php';

if (empty($_SERVER['argv']) || empty($_SERVER['argv'][1])) {
    echo 'run like "php run.php filename.xlsx output.xlsx" or "php run.php filename.xlsx"' . "\n";
    exit;
}

$outputFile = (empty($_SERVER['argv'][2]) ? 'output_' . basename($_SERVER['argv'][1]) : $_SERVER['argv'][2]);

$path = dirname(__FILE__) . DIRECTORY_SEPARATOR . $_SERVER['argv'][1];
try {
    $document = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo 'Error during loading file: '.$e->getMessage()."\n";
    exit;
}

// get all sheets
foreach ($document->getAllSheets() as $sheet) {
    // get all rows
    foreach ($sheet->getRowIterator() AS $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);

        // get all cells
        foreach ($cellIterator as $cell) {

            /** @var  $cellValue */
            $cellValue = $cell->getValue();
            $needSetValue = false;

            // if is richText
            if ($cellValue instanceof PhpOffice\PhpSpreadsheet\RichText\RichText) {
                foreach ($cellValue->getRichTextElements() as $richTextElement) {
                    $richTextElementText = $richTextElement->getText();

                    $needSetTextElement = main::replaceStyleToTags($richTextElementText, $richTextElement);

                    // if need to replace richText element
                    if ($needSetTextElement) {
                        $richTextElement->setText($richTextElementText);
                        $needSetValue = true;
                    }
                }
            }

            // if cell is string
            if (is_string($cellValue)) {
                $needSetValue = main::replaceStyleToTags($cellValue, $cell->getStyle());
            }

            // if need to set cell value
            if ($needSetValue) {
                $cell->setValue($cellValue);
            }
        }
    }
}

// save
$writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($document);
$writer->save($outputFile);