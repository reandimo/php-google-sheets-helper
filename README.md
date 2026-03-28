<div align="center">

# PHP Google Sheets Helper

### A powerful, elegant wrapper around Google Sheets API for PHP

[![Latest Stable Version](https://img.shields.io/packagist/v/reandimo/google-sheets-helper?style=for-the-badge&color=blue)](https://packagist.org/packages/reandimo/google-sheets-helper)
[![License](https://img.shields.io/packagist/l/reandimo/google-sheets-helper?style=for-the-badge&color=green)](https://packagist.org/packages/reandimo/google-sheets-helper)
[![PHP Version](https://img.shields.io/packagist/dependency-v/reandimo/google-sheets-helper/php?style=for-the-badge&color=purple)](https://packagist.org/packages/reandimo/google-sheets-helper)
[![Downloads](https://img.shields.io/packagist/dt/reandimo/google-sheets-helper?style=for-the-badge&color=orange)](https://packagist.org/packages/reandimo/google-sheets-helper)

[Installation](#-installation) &bull; [Quick Start](#-quick-start) &bull; [API Reference](#-api-reference) &bull; [Tips](#-tips) &bull; [License](#-license)

</div>

---

Stop wrestling with the verbose Google Sheets API. This library wraps the official [Google APIs Client Library for PHP](https://github.com/googleapis/google-api-php-client) into a clean, fluent interface so you can **read, write, and manage spreadsheets** in just a few lines of code.

## Features

| | Feature | Description |
|---|---|---|
| **Read** | `get()` `getSingleCellValue()` `findCellByValue()` | Fetch ranges, single cells, or search by value |
| **Write** | `appendSingleRow()` `append()` `updateSingleCell()` `update()` | Append rows, update cells or entire ranges |
| **Sheets** | `addWorksheet()` `duplicateWorksheet()` `renameWorksheet()` `deleteWorksheet()` | Full worksheet lifecycle management |
| **Style** | `colorRange()` `clearRange()` | Color cells and clear data |
| **Manage** | `create()` `getSpreadsheetWorksheets()` | Create spreadsheets and list worksheets |

---

## Requirements

- PHP >= 5.4 with CLI and JSON extension
- A [Google Cloud Platform](https://console.cloud.google.com/) project with the **Sheets API** enabled
- Credentials (`credentials.json`) from the Google Cloud Console

---

## Installation

```bash
composer require reandimo/google-sheets-helper
```

---

## Credentials Setup

The library authenticates using two files: `credentials.json` (from Google Cloud Console) and `token.json` (generated on first auth).

Generate `token.json` by running from your project root:

```bash
php ./vendor/reandimo/google-sheets-helper/firstauth
```

Follow the interactive steps — this only needs to be done once.

<details>
<summary>See it in action</summary>

<img src="https://i.imgur.com/CSWXMrq.gif" width="100%"/>

</details>

---

## Quick Start

```php
use reandimo\GoogleSheetsApi\Helper;

// Set credentials via environment (recommended)
putenv('credentialFilePath=path/to/credentials.json');
putenv('tokenPath=path/to/token.json');

// Create instance and configure
$sheet = new Helper();
$sheet->setSpreadsheetId('your-spreadsheet-id');
$sheet->setWorksheetName('Sheet1');

// Read data
$sheet->setSpreadsheetRange('A1:D10');
$data = $sheet->get();

// Write data
$sheet->appendSingleRow(['Name', 'Email', 'Role']);

// Update a cell
$sheet->updateSingleCell('B2', 'john@example.com');
```

> You can also pass credential paths directly to the constructor:
> ```php
> $sheet = new Helper('path/to/credentials.json', 'path/to/token.json');
> ```

---

## API Reference

### Reading Data

#### `get()` — Get values from a range

```php
$sheet->setSpreadsheetRange('A1:C10');
$data = $sheet->get();
```

#### `getSingleCellValue(string $cell)` — Get a single cell value

```php
$value = $sheet->getSingleCellValue('B2');
echo "Value: $value\n";
```

#### `findCellByValue(string $searchValue)` — Search for a value

```php
$sheet->setSpreadsheetRange('A1:Z100');
$result = $sheet->findCellByValue('searchValue');

if ($result) {
    echo "Found at {$result['cell']} (row {$result['row']}, col {$result['column']})\n";
}
```

---

### Writing Data

#### `appendSingleRow(array $row)` — Append one row

```php
$sheet->setSpreadsheetRange('A1:C1');
$inserted = $sheet->appendSingleRow(['John', 'john@doe.com', 'Admin']);

if ($inserted >= 1) {
    echo 'Row inserted.';
}
```

#### `append(array $rows)` — Append multiple rows

```php
$sheet->setSpreadsheetRange('A1:C1');
$sheet->append([
    ['Alice', 'alice@example.com', 'Editor'],
    ['Bob',   'bob@example.com',   'Viewer'],
    ['Carol', 'carol@example.com', 'Admin'],
]);
```

#### `updateSingleCell(string $cell, mixed $value)` — Update one cell

```php
$update = $sheet->updateSingleCell('B5', 'Updated value');

if ($update->getUpdatedCells() >= 1) {
    echo 'Cell updated.';
}
```

#### `update(array $values)` — Update a range

```php
$sheet->setSpreadsheetRange('A1:C3');
$update = $sheet->update([
    ['val1', 'val2', 'val3'],
    ['val4', 'val5', 'val6'],
    ['val7', 'val8', 'val9'],
]);

if ($update->getUpdatedCells() >= 1) {
    echo 'Range updated.';
}
```

---

### Worksheet Management

#### `getSpreadsheetWorksheets()` — List all worksheets

```php
$worksheets = $sheet->getSpreadsheetWorksheets();

foreach ($worksheets as $ws) {
    echo "ID: {$ws['id']}, Title: {$ws['title']}\n";
}
```

#### `addWorksheet(string $title, int $rows, int $cols)` — Create a new worksheet

```php
$newSheetId = $sheet->addWorksheet('NewSheet', 100, 10);
echo "Created worksheet ID: $newSheetId\n";
```

#### `duplicateWorksheet(string $newName)` — Duplicate a worksheet

```php
$sheet->setWorksheetName('Sheet1');
$sheetId = $sheet->duplicateWorksheet('Sheet1 - Copy');

if ($sheetId) {
    echo 'Worksheet duplicated.';
}
```

#### `renameWorksheet(string $oldName, string $newName)` — Rename a worksheet

```php
$sheet->renameWorksheet('OldName', 'NewName');
```

#### `deleteWorksheet(string $name)` — Delete a worksheet

```php
$deleted = $sheet->deleteWorksheet('SheetToDelete');

if ($deleted) {
    echo 'Worksheet deleted.';
}
```

---

### Styling & Utilities

#### `colorRange(array $rgb)` — Set background color

```php
$sheet->setSpreadsheetRange('A1:Z10');
$sheet->colorRange([142, 68, 173]); // Purple background
```

#### `clearRange()` — Clear all values in a range

```php
$sheet->setSpreadsheetRange('A1:Z100');
$sheet->clearRange();
```

#### `create(string $title)` — Create a new spreadsheet

```php
$newId = $sheet->create('My New Spreadsheet');
echo "Spreadsheet ID: $newId\n";
```

#### `Helper::getColumnLettersIndex(string $letters)` — Column letter to index

```php
Helper::getColumnLettersIndex('AZ'); // Returns 52
```

---

## Tips

> **Blank cells on insert/update:** Use the constant `Google_Model::NULL_VALUE` to represent an empty cell.
>
> ```php
> $sheet->appendSingleRow([
>     'John Doe',
>     'john@doe.com',
>     Google_Model::NULL_VALUE, // skip this cell
>     'Sagittarius',
> ]);
> ```

> **Multiple sheet instances:** Create as many `Helper` instances as you need to work with different spreadsheets or worksheets simultaneously.
>
> ```php
> $orders = new Helper();
> $orders->setSpreadsheetId('spreadsheet-a');
> $orders->setWorksheetName('Orders');
>
> $inventory = new Helper();
> $inventory->setSpreadsheetId('spreadsheet-b');
> $inventory->setWorksheetName('Stock');
> ```

---

## License

MIT License. See [LICENSE](LICENSE) for details.

## Questions & Issues

Found a bug or have a suggestion? [Open an issue](https://github.com/reandimo/google-sheets-helper/issues).

## Author

**Renan Diaz** — Working with PHP since 2017 & Google's API since 2019.

