<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Carbon\Carbon;
use Illuminate\Database\Eloquent\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ListOfTablesSheet implements WithTitle, ShouldAutoSize, WithMapping, FromCollection, WithStyles
{
    private $tableName;
    private $tableDescription;

    public function __construct($tableName, $tableDescription)
    {
        $this->tableName = $tableName;
        $this->tableDescription = $tableDescription;
    }

    /**
     * @return array
     */
    public function collection()
    {
        return new Collection([
            $this->tableName
        ]);
    }

    /**
     * @var $field
     * @return array
     */
    public function map($field): array
    {
        $rows = [
            [],
            [
                '',
                'No',
                'Table Name',
                'Updated at',
                'Table Description',
            ]
        ];
        foreach ($this->tableDescription as $key => $f) {
            $rows[] = [
                '',
                $key + 1,
                $field[$key],
                Carbon::now()->format('m/d/Y'),
                $f,
            ];
        }

        return $rows;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'List Of Tables';
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->getStyle('B2:E2')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('92d050');

        $sheet->getStyle('A1:E100')->getFont()->setName('Arial');
        $sheet->getStyle('A1:E100')->getAlignment()->setIndent(1);
        $sheet->getStyle('B2:E2')->getFont()->setBold(true);

        $sheet->getStyle('B2:E2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B2:B100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D3:D100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
        
        for ($i = 1; $i <= 100; $i++) {
            $sheet->getRowDimension($i)->setRowHeight(25);
        }

        $styleArray = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ],
        ];

        $sheet->getStyle('B2:E100')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    }
}