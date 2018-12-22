<?php namespace NZTim\Excel;

use Carbon\Carbon;
use DateTime;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Symfony\Component\HttpFoundation\StreamedResponse;

class ExcelSpreadsheet
{
    protected $title;
    protected $data;
    protected $creator;
    protected $phpExcel;

    public function __construct(string $title, array $data, string $creator = '')
    {
        $this->title = $title;
        $this->data = $data;
        $this->creator = $creator;
        $this->create();
    }

    protected function create()
    {
        $this->phpExcel = new Spreadsheet;
        $this->phpExcel->setActiveSheetIndex(0);
        $this->phpExcel->getProperties()
            ->setCreator($this->creator)
            ->setTitle($this->title);
        $this->phpExcel->setActiveSheetIndex(0);
        $sheet = $this->phpExcel->getActiveSheet()->setTitle($this->title);
        $this->addHeadings($sheet);
        $this->addContent($sheet);
    }

    /** Save the sheet to $file (full path and filename) */
    public function save(string $file)
    {
        $writer = IOFactory::createWriter($this->phpExcel, 'Xlsx');
        $writer->save($file);
    }

    /** Download the file as $filename */
    // https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#redirect-output-to-a-clients-web-browser
    // https://medium.com/@barryvdh/streaming-large-csv-files-with-laravel-chunked-queries-4158e484a5a2
    public function download($filename)
    {
        $response = new StreamedResponse(function () {
            $writer = IOFactory::createWriter($this->phpExcel, 'Xlsx');
            $writer->save('php://output');
        });
        $response->headers->set('Content-Type', 'application/vnd.ms-excel');
        $response->headers->set('Content-Disposition', 'attachment;filename="' . $filename . '"');
        $response->headers->set('Cache-Control','max-age=0');
        return $response;
    }

    protected function headers($filename)
    {
        return [
            'Content-type'        => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition' => 'attachment; filename=' . $filename,
        ];
    }

    protected function addHeadings(Worksheet $sheet)
    {
        // Get key for the first data element
        $firstKey = array_keys($this->data)[0];
        // Get the keys for the first data element, these are the headings
        $keys = array_keys($this->data[$firstKey]);
        $row = 1;
        $column = 'A';
        foreach ($keys as $key) {
            // Skip relationship fields
            if (is_array($this->data[$firstKey][$key])) {
                continue;
            }
            $sheet->setCellValue($column . $row, strval($key));
            $column++;
        }
    }

    protected function addContent(Worksheet $sheet)
    {
        $row = 2;
        $column = 'A';
        foreach ($this->data as $rowData) {
            foreach ($rowData as $value) {
                // Skip relationship fields
                if (is_array($value)) {
                    continue;
                }
                if (is_integer($value)) {
                    $sheet->setCellValueExplicit($column . $row, $value, DataType::TYPE_NUMERIC);
                } elseif ($value instanceof Carbon) {
                    // Set value to Excel-specific timestamp
                    $sheet->setCellValue($column . $row, Date::PHPToExcel($value));
                    // Set display mask to appropriate format
                    $sheet->getStyle($column . $row)
                        ->getNumberFormat()
                        ->setFormatCode('dd-mm-yyyy'); // See \PhpOffice\PhpSpreadsheet\Style\NumberFormat
                } else {
                    $sheet->setCellValueExplicit($column . $row, strval($value), DataType::TYPE_STRING);
                }
                $column++;
            }
            $column = 'A';
            $row++;
        }
    }
}
