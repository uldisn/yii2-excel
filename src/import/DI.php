<?php

namespace arogachev\excel\import;

use arogachev\excel\import\advanced\Importer as AdvancedImporter;
use arogachev\excel\import\basic\Importer as BasicImporter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Yii;

class DI
{
    /**
     * @return BasicImporter|AdvancedImporter
     * @throws \yii\base\InvalidConfigException
     */
    public static function getImporter()
    {
        return Yii::$container->get('importer');
    }

    /**
     * @param BaseImporter $value
     */
    public static function setImporter($value)
    {
        Yii::$container->set('importer', $value);
    }

    /**
     * @return Spreadsheet
     * @throws \yii\base\InvalidConfigException
     */
    public static function getPHPExcel()
    {
        return static::getImporter()->phpExcel;
    }

    /**
     * @return CellParser
     * @throws \yii\base\InvalidConfigException
     */
    public static function getCellParser()
    {
        return static::getImporter()->cellParser;
    }
}
