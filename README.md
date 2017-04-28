# php-excel-builder

Wrapper for PHPExcel library.
Helps to create PHPExcel documents.

# How to use

<table>
<thead>
<th>Pure PHPExcel</th>
<th>With php-excel-builder</th>
</thead>
<tbody>
<tr>
<td>

```php
$sheet = new \PHPExcel_Worksheet(null, 'New xls list');

$sheet->fromArray([
    ['ID', 'Name', 'Numeric text field', 'Link'],
    [1, 'First', '1234', 'http://domain.com'],
    [2, 'Second', '555', 'https://example.com']
]);

$sheet->getColumnDimension('D')->setWidth(70);

for ($i=2; $i<4; $i++) {
    $cell = $sheet->getCellByColumnAndRow(3, $i);

    $cell->getHyperlink()->setUrl($cell->getValue());
}

$excel = new \PHPExcel();

$excel->removeSheetByIndex(0);
$excel->addSheet($sheet);

$writer = new \PHPExcel_Writer_Excel2007($excel);

$writer->save('document.xlsx');
```

</td>
<td>

```php
$sheet = ExcelFacade\SheetBuilder::create('New xls list')
    ->setHeader(['ID', 'Name', 'Numeric text field', 'Link'])
    ->setData([
        [1, 'First', '1234', 'http://domain.com'],
        [2, 'Second', '555', 'https://example.com']
    ])
    ->setColumnWidth('D', 70)
    ->setUrlColumn(3)
    ->setColumnType(2, \PHPExcel_Cell_DataType::TYPE_STRING2);

ExcelFacade\ExcelBuilder::create()
    ->addSheet($sheet)
    ->save('document.xlsx');
```

</td>
</tr>
</tbody>
</table>

# Reference

Wrapper is about 2 classes,
SheetBuilder - constructs `\PHPExcel_Worksheet object`,
and ExcelBuilder - constructs `\PHPExcel` object and contain collection of SheetBuilder`s.

## SheetBuilder

### .getSheet()

Return wrapped `\PHPExcel_Worksheet` object

### .setHeader($header = [])

Define first line of sheet.
Header will not be toched by methods such as `setUrlColumns()`.
Sheet can have no header

### .setData($data = [])

Define actual rows and columns of sheet

### .setUrlColumns($urlColumns = [])

Provide indexes of columns to define which columns will be converted as hyperlink

### .setUrlColumn($urlColumn)

Same as `setUrlColumns()` but for single column

### .setColumnTypes($columnTypes = [])

Provide indexes of columns columns to define which of `PHPExcel_Cell_DataType::TYPE_*` data-type will be used by each column

### .setColumnType($column, $type)

Same as `setColumnTypes()` but for single column

### .setColumnWidth($column, $width)

Wrapper for `\PHPExcel_Worksheet_ColumnDimension::setWidth()`. Set width of column by index

### .setColumnWidths($widths = [])

Same as `setColumnWidth()` but for many columns

## ExcelBuilder

### .getExcel()

Return wrapped `\PHPExcel` object

### .setWriterType($type)

Select class of writer will be used to save the document.
Only `\PHPExcel_Writer_Abstract` subclasses are allowed.
`\PHPExcel_Writer_Excel2007` is used by default

### .setSheets($sheets)

Set collection of `SheetBuilder` objects to current builder.

### .addSheet(SheetBuilder $sheet, $index = null)

Add sheet to collection at provided index

### .save($filename)

Build and save the document as provided filename
