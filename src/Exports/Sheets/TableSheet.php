<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

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

class TableSheet implements FromCollection, ShouldAutoSize, WithMapping, WithTitle, WithStyles, WithEvents
{
    private $table_name;
    private $fields;
    private $fileds_info;

    public function __construct($table_name, $fields, $fileds_info)
    {
        $this->table_name = $table_name;
        $this->fields = $fields;
        $this->fileds_info = $fileds_info;
    }

    /**
     * @return array
     */
    public function collection()
    {
        return new Collection([
            $this->fields
        ]);
    }

    /**
     * @var $field
     * @return array
     */
    public function map($field): array
    {
        $rows = [
            [
                '<<'
            ],
            [
                '',
                strtoupper($this->table_name),
            ],
            [
                '',
                'No',
                'Column Name',
                'PK',
                'UK',
                'NN',
                'AI',
                'FK',
                'Data Type',
                'Length',
                'Default',
                'Description',
            ]
        ];
        foreach ($this->fileds_info as $key => $f) {
            $rows[] = [
                '',
                $key + 1,
                $field[$key],
                $f['primary_key'] ? 'v' : '',
                $f['unique'] ? 'v' : '',
                $f['notnull'] ? 'v' : '',
                $f['auto_increment'] ? 'v' : '',
                $f['foreign_key'] ? 'v' : '',
                $f['type'] == 'string' ? 'VARCHAR' : strtoupper($f['type']),
                $f['length'],
                $f['default'],
                $f['description'],
            ];
        }

        return $rows;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return strtoupper($this->table_name);
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->mergeCells('B2:L2');
        $sheet->getStyle('B2')
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B2')->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('92d050');

        $sheet->getStyle('A1:L100')->getFont()->setName('Arial');
        $sheet->getStyle('A1:L100')->getAlignment()->setIndent(1);
        $sheet->getStyle('B2:L3')->getFont()->setBold(true);

        $sheet->getStyle('B3:L3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B3:B100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D3:H100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('I3:I100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('J3:J100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('K3:K100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
        for ($i = 1; $i <= 100; $i++) {
            $sheet->getRowDimension($i)->setRowHeight(25);
        }

        $sheet->getStyle('B3:L3')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('fde9d9');

        $styleArray = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ],
        ];

        $sheet->getStyle('B2:L100')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getCell('A1')->getHyperlink()->setUrl("sheet://'List Of Tables'!A1");
        $sheet->getCell('A1')->getStyle()->getFont()->setUnderline(Font::UNDERLINE_SINGLE);
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                foreach ($this->fields as $key => $field) {
                    for ($col_index = 4; $col_index <= 8; $col_index++) {
                        $cell_value = $event->sheet->getCellByColumnAndRow($col_index, $key + 4)
                            ->getValue();

                        if ($cell_value == "v") {
                            $event->sheet->getCellByColumnAndRow($col_index, $key + 4)
                                ->getStyle()
                                ->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()
                                ->setARGB('ff0000');

                            $event->sheet->setCellValueByColumnAndRow($col_index, $key + 4, "");
                        }
                    }
                }
            },
        ];
    }
}