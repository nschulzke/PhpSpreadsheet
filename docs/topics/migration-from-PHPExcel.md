# Migration from PHPExcel

PhpSpreadsheet introduced many breaking changes by introducing
namespaces and renaming some classes. To help you migrate existing
project, a tool was written to replace all references to PHPExcel
classes to their new names. But there are also manual changes that
need to be done.

## Automated tool

[RectorPHP](https://github.com/rectorphp/rector) can be used to migrate
automatically your codebase. Assuming your files to be migrated lives
in `src/`, you can run the migration like so:

```sh
composer require rector/rector --dev
vendor/bin/rector process src --set phpexcel-to-phpspreadsheet
composer remove rector/rector
```

For more details, see
[RectorPHP blog post](https://getrector.org/blog/2020/04/16/how-to-migrate-from-phpexcel-to-phpspreadsheet-with-rector-in-30-minutes).

## Column index based on 1

Column indexes are now based on 1. So column `A` is the index `1`. This is consistent
with rows starting at 1 and Excel function `COLUMN()` that returns `1` for column `A`.
So the code must be adapted with something like:

```php
// Before
$cell = $worksheet->getCellByColumnAndRow($column, $row);

for ($column = 0; $column < $max; $column++) {
    $worksheet->setCellValueByColumnAndRow($column, $row, 'value ' . $column);
}

// After
$cell = $worksheet->getCellByColumnAndRow($column + 1, $row);

for ($column = 1; $column <= $max; $column++) {
    $worksheet->setCellValueByColumnAndRow($column, $row, 'value ' . $column);
}
```

All the following methods are affected:

- `PHPExcel_Worksheet::cellExistsByColumnAndRow()`
- `PHPExcel_Worksheet::freezePaneByColumnAndRow()`
- `PHPExcel_Worksheet::getCellByColumnAndRow()`
- `PHPExcel_Worksheet::getColumnDimensionByColumn()`
- `PHPExcel_Worksheet::getCommentByColumnAndRow()`
- `PHPExcel_Worksheet::getStyleByColumnAndRow()`
- `PHPExcel_Worksheet::insertNewColumnBeforeByIndex()`
- `PHPExcel_Worksheet::mergeCellsByColumnAndRow()`
- `PHPExcel_Worksheet::protectCellsByColumnAndRow()`
- `PHPExcel_Worksheet::removeColumnByIndex()`
- `PHPExcel_Worksheet::setAutoFilterByColumnAndRow()`
- `PHPExcel_Worksheet::setBreakByColumnAndRow()`
- `PHPExcel_Worksheet::setCellValueByColumnAndRow()`
- `PHPExcel_Worksheet::setCellValueExplicitByColumnAndRow()`
- `PHPExcel_Worksheet::setSelectedCellByColumnAndRow()`
- `PHPExcel_Worksheet::stringFromColumnIndex()`
- `PHPExcel_Worksheet::unmergeCellsByColumnAndRow()`
- `PHPExcel_Worksheet::unprotectCellsByColumnAndRow()`
- `PHPExcel_Worksheet_PageSetup::addPrintAreaByColumnAndRow()`
- `PHPExcel_Worksheet_PageSetup::setPrintAreaByColumnAndRow()`
