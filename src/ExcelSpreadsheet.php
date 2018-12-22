<?php namespace NZTim\Excel;

use Carbon\Carbon;
use DateTime;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

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

    public function save(string $file)
    {
        $writer = IOFactory::createWriter($this->phpExcel, 'Xlsx');
        $writer->save($file);
    }

    public function download($file, $filename = null)
    {
        $filename = $filename ?: pathinfo($file, PATHINFO_BASENAME);
        $response = response()->download($file, $filename, $this->headers($filename));
        ob_end_clean(); // https://github.com/laravel/framework/issues/2892
        return $response;
    }

    protected function headers($filename)
    {
        return [
            'Content-type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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
                $sheet->setCellValue($column . $row, strval($value));
                // Handle dates
                if ($this->dateTime($value)) {
                    // Set value to Excel-specific timestamp
                    $sheet->setCellValue($column . $row, $this->dateTime($value));
                    // Set display mask to appropriate format
                    $sheet->getStyle($column . $row)
                        ->getNumberFormat()
                        ->setFormatCode('dd-mm-yyyy'); // See \PhpOffice\PhpSpreadsheet\Style\NumberFormat
                }
                $column++;
            }
            $column = 'A';
            $row++;
        }
    }

    protected function dateTime($value): ?string
    {
        if ($value instanceof Carbon) {
            return Date::PHPToExcel($value);
        }
        try {
            return Date::PHPToExcel(new DateTime($value));
        } catch (\Throwable $e) {
            return null;
        }
    }
}
