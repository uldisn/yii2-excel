<?php

namespace arogachev\excel\helpers;

class PHPExcelHelper
{
    /**
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Row $row
     * @return boolean
     */
    public static function isRowEmpty($row)
    {
        foreach ($row->getCellIterator() as $cell) {
            if ($cell->getValue()) {
                return false;
            }
        }

        return true;
    }
}
