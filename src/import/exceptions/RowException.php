<?php

namespace arogachev\excel\import\exceptions;

use PhpOffice\PhpSpreadsheet\Worksheet\Row;

class RowException extends ImportException
{
    /**
     * @param Row $row
     * {@inheritdoc}
     */
    public function __construct(Row $row, $message = "", $code = 0, \Exception $previous = null)
    {
        $sheetTitle = $row->getCellIterator()->current()->getWorksheet()->getTitle();
        $message = "Import failed at sheet \"$sheetTitle\", row \"{$row->getRowIndex()}\". $message";

        parent::__construct($message, $code, $previous);
    }

    /**
     * @inheritdoc
     */
    public function getName()
    {
        return 'Row Exception';
    }
}
