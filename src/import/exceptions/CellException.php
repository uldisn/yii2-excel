<?php

namespace arogachev\excel\import\exceptions;

use PhpOffice\PhpSpreadsheet\Cell\Cell;

class CellException extends ImportException
{
    /**
     * @param Cell $cell
     * {@inheritdoc}
     */
    public function __construct(Cell $cell, $message = "", $code = 0, \Exception $previous = null)
    {
        $sheetTitle = $cell->getWorksheet()->getTitle();
        $cellCoordinate = $cell->getCoordinate();
        $message = "Error when preparing data for import: sheet \"$sheetTitle\", cell \"$cellCoordinate\". $message";

        parent::__construct($message, $code, $previous);
    }

    /**
     * @inheritdoc
     */
    public function getName()
    {
        return 'Cell Exception';
    }
}
