<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithDrawings;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class ERDSheet implements WithTitle, WithDrawings
{
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('ERD');
        $drawing->setPath(storage_path('app').'/databasedocs/erd.png');
        $drawing->setHeight(500);
        $drawing->setCoordinates('B4');

        return $drawing;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'ERD';
    }
}