<?php

namespace ExcelBuilder;

class SheetBuilder
{
    /**
     * @var \PHPExcel_Worksheet
     */
    private $sheet;

    /**
     * @var string[]|null
     */
    private $header = null;

    /**
     * @var array
     */
    private $data = [];

    /**
     * @var int[]
     */
    private $urlColumns = [];

    /**
     * @var array
     */
    private $columnTypes = [];

    public function __construct(\PHPExcel_Worksheet $sheet)
    {
        $this->sheet = $sheet;
    }

    /**
     * @param string $title
     * @return SheetBuilder
     */
    public static function create($title)
    {
        $sheet = new \PHPExcel_Worksheet(null, $title);

        return new self($sheet);
    }

    /**
     * @return \PHPExcel_Worksheet
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * @param string[] $header
     * @return SheetBuilder
     */
    public function setHeader($header = [])
    {
        $this->header = $header;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasHeader()
    {
        return $this->header !== null;
    }

    /**
     * @param array[] $data
     * @return SheetBuilder
     */
    public function setData($data = [])
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param int[] $urlColumns
     * @return SheetBuilder
     */
    public function setUrlColumns($urlColumns = [])
    {
        $this->urlColumns = $urlColumns;

        return $this;
    }

    /**
     * @param int $urlColumn
     * @return SheetBuilder
     */
    public function setUrlColumn($urlColumn)
    {
        $this->urlColumns[] = $urlColumn;

        return $this;
    }

    /**
     * @param array $columnTypes
     * @return SheetBuilder
     */
    public function setColumnsTypes($columnTypes = [])
    {
        $this->columnTypes = $columnTypes;

        return $this;
    }

    /**
     * @param int $column
     * @param string $type
     * @return SheetBuilder
     */
    public function setColumnType($column, $type)
    {
        $this->columnTypes[$column] = $type;

        return $this;
    }

    /**
     * @param string $column
     * @param int $width
     * @return SheetBuilder
     */
    public function setColumnWidth($column, $width)
    {
        $this->sheet->getColumnDimension($column)->setWidth($width);

        return $this;
    }

    /**
     * @param array $widths
     * @throws \PHPExcel_Exception
     * @return SheetBuilder
     */
    public function setColumnsWidths($widths = [])
    {
        foreach ($widths as $column => $width) {
            $this->setColumnWidth($column, $width);
        }

        return $this;
    }

    /**
     * @return SheetBuilder
     */
    public function build()
    {
        $this->handleData();
        $this->handleColumns();

        return $this;
    }

    private function handleData()
    {
        $data = $this->data;

        if ($this->header !== null) {
            array_unshift($data, $this->header);
        }

        $this->sheet->fromArray($data);
    }

    /**
     * @throws \PHPExcel_Exception
     */
    private function handleColumns()
    {
        if (count($this->urlColumns) == 0 && count($this->columnTypes)) {
            return;
        }

        $begin = 1;
        $end = count($this->data) + 1;

        if ($this->hasHeader()) {
            $begin++;
            $end++;
        }

        for ($i=$begin; $i<$end; $i++) {
            foreach ($this->urlColumns as $column) {
                $cell = $this->sheet->getCellByColumnAndRow($column, $i);

                $cell->getHyperlink()->setUrl($cell->getValue());
            }

            foreach ($this->columnTypes as $column => $style) {
                $cell = $this->sheet->getCellByColumnAndRow($column, $i);

                $cell->setDataType($style);
            }
        }
    }
}
