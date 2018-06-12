<?php

namespace arogachev\excel\import;

use arogachev\excel\import\basic\Model;
use arogachev\excel\import\basic\StandardModel;
use arogachev\excel\import\exceptions\ImportException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Yii;
use yii\base\Component;
use yii\base\Event;
use yii\base\InvalidArgumentException;

/**
 * @property Spreadsheet $phpExcel
 * @property string $error
 * @property \yii\db\ActiveRecord $wrongModel
 * @property Model[] $models
 */
abstract class BaseImporter extends Component
{
    const EVENT_RUN = 'run';

    /**
     * @var string
     */
    public $filePath;

    /**
     * @var array
     */
    public $standardModelsConfig;

    /**
     * @var Spreadsheet
     */
    protected $_phpExcel;

    /**
     * @var StandardModel[]
     */
    protected $_standardModels;

    /**
     * @var Model[]
     */
    protected $_models = [];

    /**
     * @var string
     */
    protected $_error;

    /**
     * @var \yii\db\ActiveRecord
     */
    protected $_wrongModel;


    /**
     * @inheritdoc
     */
    public function init()
    {
        if (!$this->filePath) {
            throw new InvalidArgumentException('File path not specified or file not uploaded.');
        }

        if (!file_exists($this->filePath)) {
            throw new InvalidArgumentException("File not exist in path \"$this->filePath\".");
        }

        foreach ($this->standardModelsConfig as $config) {
            $this->initStandardModel($config);
        }

        $this->configureEventHandlers();
        DI::setImporter($this);
    }

    /**
     * @param array $config
     */
    protected function initStandardModel($config)
    {
        $this->_standardModels[] = new StandardModel($config);
    }

    protected function configureEventHandlers()
    {
        Event::on(Model::className(), Model::EVENT_INIT, function ($event) {
            /* @var $model Model */
            $model = $event->sender;
            $model->instance = new $model->standardModel->className;
        });

        $this->on(self::EVENT_RUN, function () {
            if (!$this->_models) {
                throw new ImportException('No models for import.');
            }
        });
    }

    /**
     * @return boolean
     */
    public function run()
    {
        try {
            $this->safeRun();
        } catch (ImportException $e) {
            $this->_error = $e->getMessage();

            return false;
        }

        return true;
    }

    protected function safeRun()
    {
        $this->_phpExcel = IOFactory::load($this->filePath);
    }

    /**
     * @return Spreadsheet
     */
    public function getPhpExcel()
    {
        return $this->_phpExcel;
    }

    /**
     * @return string
     */
    public function getError()
    {
        return $this->_error;
    }

    /**
     * @return \yii\db\ActiveRecord
     */
    public function getWrongModel()
    {
        return $this->_wrongModel;
    }

    /**
     * @param \yii\db\ActiveRecord $value
     */
    public function setWrongModel($value)
    {
        $this->_wrongModel = $value;
    }

    /**
     * @return Model[]
     */
    public function getModels()
    {
        return $this->_models;
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\RowIterator|\PhpOffice\PhpSpreadsheet\Worksheet\Row[] $rows
     */
    abstract protected function fillModels($rows);
}
