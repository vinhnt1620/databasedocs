<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Carbon\Carbon;
use Illuminate\Database\Eloquent\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ListOfTablesSheet implements WithTitle, ShouldAutoSize, WithMapping, FromCollection, WithStyles, WithEvents
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
        $key_explains = [
            [
                'id' => 1,
                'item' => 'PK',
                'explain' => 'Primary Key'
            ],
            [
                'id' => 2,
                'item' => 'FK',
                'explain' => 'Foreign Key'
            ],
            [
                'id' => 3,
                'item' => 'AI',
                'explain' => 'Auto Increment'
            ],
            [
                'id' => 4,
                'item' => 'NN',
                'explain' => 'Not NULL'
            ],
            [
                'id' => 5,
                'item' => 'UQ',
                'explain' => 'Unique'
            ]
        ];

        $rows = [
            [],
            [
                '',
                'No',
                'Table Name',
                'Updated at',
                'Table Description',
                '',
                '',
                '#',
                'Item',
                'Explain',
            ]
        ];
        foreach ($this->tableDescription as $key => $f) {
            $rows[] = [
                '',
                $key + 1,
                strtoupper($field[$key]),
                Carbon::now()->format('m/d/Y'),
                $f,
            ];
        }

        foreach ($key_explains as $key => $key_explain) {
            if (array_key_exists($key + 2, $rows)) {
                $rows[$key + 2] = array_merge($rows[$key + 2], [
                    '',
                    '',
                    $key_explain['id'],
                    $key_explain['item'],
                    $key_explain['explain']
                ]);
            } else {
                $rows[] = [
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    $key_explain['id'],
                    $key_explain['item'],
                    $key_explain['explain']
                ];
            }
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

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $columnIndex = 3;
                
                foreach ($this->tableName as $key => $tableName) {
                    // Generate the hyperlink URL based on the cell text
                    $hyperlinkURL = "sheet://'$tableName'!A1";

                    // Set the hyperlink in the cell
                    $event->sheet->getCellByColumnAndRow($columnIndex, $key + 3)
                        ->getHyperlink()
                        ->setUrl($hyperlinkURL);

                    $event->sheet->getCellByColumnAndRow($columnIndex, $key + 3)
                        ->getStyle()
                        ->getFont()
                        ->setUnderline(Font::UNDERLINE_SINGLE);
                }
            },
        ];
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->getStyle('B2:E2')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('92d050');

        $sheet->getStyle('H2:J2')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('92d050');

        $sheet->getStyle('A1:J100')->getFont()->setName('Arial');
        $sheet->getStyle('A1:J100')->getAlignment()->setIndent(1);
        $sheet->getStyle('B2:E2')->getFont()->setBold(true);

        $sheet->getStyle('H2:J2')->getFont()->setBold(true);

        $sheet->getStyle('B2:E2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B2:B100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D3:D100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

        $sheet->getStyle('H2:J2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
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
        $sheet->getStyle('H2:J7')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER)->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    }
}