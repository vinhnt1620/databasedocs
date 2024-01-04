<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Illuminate\Database\Eloquent\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ERDSheet implements WithTitle, WithDrawings, ShouldAutoSize, WithStyles, WithMapping, FromCollection
{
    /**
     * @return array
     */
    public function collection()
    {
        $datas = [
        ];

        return new Collection([$datas]);
    }

    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Database Schema');
        $drawing->setPath(storage_path('app').'/erd.png');
        $drawing->setHeight(500);
        $drawing->setCoordinates('B4');

        return $drawing;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'Database Schema';
    }

    /**
     * @var $field
     * @return array
     */
    public function map($field): array
    {
        $rows = [
            [
                'Database Schema',
            ],
        ];

        return $rows;
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->mergeCells('A1:M2');
        $sheet->getStyle('A1')
            ->getAlignment()
            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
            ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle('A1')->getFont()->setSize(15);

        $sheet->getStyle('A1:M2')->getFont()->setName('Arial');
        $sheet->getStyle('A1:M2')->getAlignment()->setIndent(1);
        $sheet->getStyle('A1:M2')->getFont()->setBold(true);
    }
}