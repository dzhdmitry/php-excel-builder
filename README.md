# php-excel-builder

Wrapper for [PHPExcel](https://github.com/PHPOffice/PHPExcel) library.
Helps to create simple PHPExcel documents easier.

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
// Initializing new sheet
$sheet = new \PHPExcel_Worksheet(null, 'New list');

// Set sheet content
$sheet->fromArray([
    ['ID', 'Name', 'Link'],
    [1, 'First', 'http://domain.com'],
    [2, 'Second', 'https://example.com']
]);

// Set width to a column
$sheet->getColumnDimension('C')->setWidth(70);

// Convert columns to hyperlinks
for ($i=2; $i<4; $i++) {
    $cell = $sheet
        ->getCellByColumnAndRow(2, $i);

    $cell
        ->getHyperlink()
        ->setUrl($cell->getValue());
}

// Initializing excel document
$excel = new \PHPExcel();

$excel->removeSheetByIndex(0);
$excel->addSheet($sheet);

// Saving excel document
$writer = new \PHPExcel_Writer_Excel2007($excel);

$writer->save('document.xlsx');
```

</td>
<td>

```php
// Initializing new sheet
$sheet = SheetBuilder::create('New list')

    // Set sheet content
    ->setHeader(['ID', 'Name', 'Link'])
    ->setData([
        [1, 'First', 'http://domain.com'],
        [2, 'Second', 'https://example.com']
    ])

    // Set width to a column
    ->setColumnWidth('C', 70)

    // Define columns to be converted to hyperlinks
    ->setUrlColumn(2);








// Initializing excel document
ExcelBuilder::create()
    ->addSheet($sheet)



    // ...and saving it
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

Return wrapped `\PHPExcel_Worksheet` object.

### .setHeader($header = [])

Define first line of sheet.
Header will not be toched by methods such as `setUrlColumns()` and `setColumnsTypes()`.
Sheet has no header by default.

### .setData($data = [])

Define actual rows and columns of sheet.

### .setUrlColumns($urlColumns = [])

Provide indexes of columns to define which columns will be converted as hyperlink.
All cells (except header) in provided columns will be converted.
Columns indexes begin with zero.

```php
SheetBuilder::create('List')
    ->setData([
        ['http://domain.com', 'Content', 'http://example.com']
    ])
    ->setUrlColumns([0, 2]);
```

### .setUrlColumn($urlColumn)

Same as `setUrlColumns()` but for single column.

### .setColumnsTypes($columnTypes = [])

Provide indexes of columns columns to define which of `PHPExcel_Cell_DataType::TYPE_*` data-type will be used by each column.
Columns indexes begin with zero.

```php
SheetBuilder::create('List')
    ->setData([
        ['100', 'Numbers will be shown as strings', '350']
    ])
    ->setColumnsTypes([0, 2]);
```

### .setColumnType($column, $type)

Same as `setColumnTypes()` but for single column.

### .setColumnWidth($column, $width)

Wrapper for `\PHPExcel_Worksheet_ColumnDimension::setWidth()`. Set width of column by index.

### .setColumnsWidths($widths = [])

Same as `setColumnWidth()` but for many columns.

## ExcelBuilder

### .getExcel()

Return wrapped `\PHPExcel` object.

### .setWriterType($type)

Select class of writer will be used to save the document.
Any of `\PHPExcel_Writer_Abstract` subclasses can be used.
Default writer is `\PHPExcel_Writer_Excel2007`.

```php
// Use CSV writer
ExcelBuilder::create()
    ->setWriterType(PHPExcel_Writer_CSV::class);
```

### .setSheets($sheets)

Set collection of `SheetBuilder` objects to current builder.
Before builder is saved, it builds all previously added sheets and composes them into document.

### .addSheet(SheetBuilder $sheet, $index = null)

Add sheet to collection at provided index.
If index is null, just add provided sheet to collection.

### .save($filename)

Build and save the document as provided filename.

```php
ExcelBuilder::create()
    ->addSheet($sheet)
    ->save(__DIR__ . '/../document.xlsx');
```

# Licence

Licensed under the [MIT license](https://github.com/dzhdmitry/php-excel-builder/blob/master/LICENSE.txt).
