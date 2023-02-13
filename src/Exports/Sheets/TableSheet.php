<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Illuminate\Database\Eloquent\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class TableSheet implements FromCollection, ShouldAutoSize, WithMapping, WithTitle, WithStyles
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
            [],
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

        $sheet->getStyle('A1:L25')->getFont()->setName('Arial');
        $sheet->getStyle('A1:L25')->getAlignment()->setIndent(1);
        $sheet->getStyle('B2:L3')->getFont()->setBold(true);

        $sheet->getStyle('B3:L3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('B3:B25')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('D3:H25')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('I3:I25')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('J3:J25')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('K3:K25')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
        for ($i = 1; $i <= 25; $i++) {
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

        $sheet->getStyle('B2:L25')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    }
}