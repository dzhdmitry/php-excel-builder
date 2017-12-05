<?php

namespace ExcelBuilder;

class ExcelBuilder
{
    /**
     * @var \PHPExcel
     */
    private $excel;

    /**
     * @var \PHPExcel_Writer_Abstract
     */
    private $writer;

    /**
     * @var SheetBuilder[]
     */
    private $sheets = [];

    /**
     * @var bool
     */
    private $_isBuilt = false;

    public function __construct(\PHPExcel $excel = null, $removeFirstSheet = true)
    {
        if ($excel === null) {
            $excel = new \PHPExcel();

            if ($removeFirstSheet) {
                $excel->removeSheetByIndex(0);
            }
        }

        $this->excel = $excel;

        $this->setWriterType(\PHPExcel_Writer_Excel2007::class);
    }

    /**
     * @param \PHPExcel|null $excel
     * @param bool $removeFirstSheet
     * @return ExcelBuilder
     */
    public static function create(\PHPExcel $excel = null, $removeFirstSheet = true)
    {
        return new self($excel, $removeFirstSheet);
    }

    /**
     * @return \PHPExcel
     */
    public function getExcel()
    {
        return $this->excel;
    }

    /**
     * @param string $type
     * @return ExcelBuilder
     */
    public function setWriterType($type)
    {
        $this->writer = new $type($this->excel);

        return $this;
    }

    /**
     * @return \PHPExcel_Writer_Abstract
     */
    public function getWriter()
    {
        return $this->writer;
    }

    /**
     * @param SheetBuilder $sheet
     * @param int|null $index
     * @return ExcelBuilder
     */
    public function addSheet(SheetBuilder $sheet, $index = null)
    {
        if ($index !== null) {
            $this->sheets[$index] = $sheet;
        } else {
            $this->sheets[] = $sheet;
        }

        return $this;
    }

    /**
     * @param SheetBuilder[] $sheets
     * @return ExcelBuilder
     */
    public function setSheets($sheets)
    {
        $this->sheets = $sheets;

        return $this;
    }

    /**
     * @return ExcelBuilder
     * @throws \PHPExcel_Exception
     */
    public function build()
    {
        if ($this->_isBuilt) {
            return $this;
        }

        foreach ($this->sheets as $index => $sheetBuilder) {
            $sheet = $sheetBuilder->build()->getSheet();

            $this->excel->addSheet($sheet, $index);
        }

        $this->_isBuilt = true;

        return $this;
    }

    /**
     * @param string $filename
     * @throws \PHPExcel_Exception
     */
    public function save($filename)
    {
        if (!$this->_isBuilt) {
            $this->build();
        }

        $this->writer->save($filename);
    }
}
