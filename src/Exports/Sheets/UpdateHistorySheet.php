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

class UpdateHistorySheet implements WithTitle, ShouldAutoSize, WithStyles, WithMapping, FromCollection
{
    /**
     * @return array
     */
    public function collection()
    {
        $datas = [
            [
                'id' => '1',
                'update_date' => str_replace('-', '', Carbon::now()->toDateString()),
                'update_person' => 'vinh.nguyen',
                'sheet_name' => '',
                'update_subject' => '',
                'update_content' => '',
                'notes' => '',
            ]
        ];

        return new Collection([$datas]);
    }

    /**
     * @var $field
     * @return array
     */
    public function map($field): array
    {
        $rows = [
            [
                'Update History Record',
            ],
            [],
            [],
            [
                '#',
                'Update Date',
                'Update Person',
                'Sheet Name',
                'Update Subject',
                'Update Content - Update Reason',
                'Notes',
            ]
        ];
        foreach ($field as $key => $data) {
            $rows[] = [
                $data['id'],
                $data['update_date'],
                $data['update_person'],
                $data['sheet_name'],
                $data['update_subject'],
                $data['update_content'],
                $data['notes'],
            ];
        }

        return $rows;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'Update History';
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->mergeCells('A1:G3');
        $sheet->getStyle('A1')
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1')->getFont()->setSize(15);

        $sheet->getStyle('A1:G100')->getFont()->setName('Arial');
        $sheet->getStyle('A1:G100')->getAlignment()->setIndent(1);
        $sheet->getStyle('A1:G4')->getFont()->setBold(true);

        $sheet->getStyle('A4:G4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A4:G100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
        for ($i = 1; $i <= 100; $i++) {
            $sheet->getRowDimension($i)->setRowHeight(25);
        }

        $sheet->getStyle('A4:G4')
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('92d050');

        $styleArray = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ],
        ];

        $sheet->getStyle('A1:G100')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    }
}