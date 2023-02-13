<?php
namespace Vinhnt\Databasedocs\Exports\Sheets;

use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithDrawings;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class CoverSheet implements WithTitle, WithDrawings
{
    public function drawings()
    {
        $drawing = new Drawing();
        $drawing->setName('Cover');
        $drawing->setPath(__DIR__.'/../../public/img/cover.png');
        $drawing->setHeight(600);
        $drawing->setCoordinates('B2');

        return $drawing;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'Cover';
    }
}