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

class TableOfContentsSheet implements WithTitle, ShouldAutoSize, WithStyles, WithMapping, FromCollection
{
    /**
     * @return array
     */
    public function collection()
    {
        $datas = [
            [
                'id' => '1',
                'sheet_name' => 'Cover',
                'description' => 'Trang bìa',
                'note' => '',
            ],
            [
                'id' => '2',
                'sheet_name' => 'Update History',
                'description' => 'Lưu lại lịch sử update tài liệu',
                'note' => '',
            ],
            [
                'id' => '3',
                'sheet_name' => 'Table Of Contents',
                'description' => 'Định nghĩa các sheet và nội dung chi tiết',
                'note' => '',
            ],
            [
                'id' => '4',
                'sheet_name' => 'ERD',
                'description' => 'Sơ đồ quan hệ thực thể',
                'note' => '',
            ],
            [
                'id' => '5',
                'sheet_name' => 'List Of Tables',
                'description' => 'Danh sách các bảng',
                'note' => '',
            ],
            [
                'id' => '6',
                'sheet_name' => 'Details',
                'description' => 'Chi tiết các bảng',
                'note' => '',
            ],
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
                'Sheet Name',
                'Description',
                'Note',
            ]
        ];
        foreach ($field as $key => $data) {
            $rows[] = [
                $data['id'],
                $data['sheet_name'],
                $data['description'],
                $data['note'],
            ];
        }

        return $rows;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'Table Of Contents';
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->mergeCells('A1:D3');
        $sheet->getStyle('A1')
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A1')->getFont()->setSize(15);

        $sheet->getStyle('A1:D100')->getFont()->setName('Arial');
        $sheet->getStyle('A1:D100')->getAlignment()->setIndent(1);
        $sheet->getStyle('A1:D4')->getFont()->setBold(true);

        $sheet->getStyle('A4:D4')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A4:A100')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        
        for ($i = 1; $i <= 100; $i++) {
            $sheet->getRowDimension($i)->setRowHeight(100);
        }

        $sheet->getStyle('A4:D4')
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

        $sheet->getStyle('A1:D100')->applyFromArray($styleArray)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    }
}