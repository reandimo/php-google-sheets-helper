PHP Google Spreadsheets API Helper
======================

[![Latest Stable Version](http://poser.pugx.org/reandimo/google-sheets-helper/v)](https://packagist.org/packages/reandimo/google-sheets-helper) [![License](http://poser.pugx.org/reandimo/google-sheets-helper/license)](https://packagist.org/packages/reandimo/google-sheets-helper) [![PHP Version Require](http://poser.pugx.org/reandimo/google-sheets-helper/require/php)](https://packagist.org/packages/reandimo/google-sheets-helper)

Google Spreadsheets API Helper - A bunch of functions to work easily with Google Sheets API

This library is a helper that encapsulate [Google APIs Client Library for PHP](https://github.com/googleapis/google-api-php-client) ([Documentation](https://developers.google.com/sheets/api/quickstart/php)) for simple usage.

--- 

Outline
-------

* [Requirements](#requirements)
* [Installation](#installation)
* [Credentials](#credentials)
* [Usage](#usage)
  - [Create Instance](#create-instance) 
    - [Set sheets props](#set-sheet-props) 
  - [Get Values](#get-values) 
    - [Get values from range](#get-values-from-range) 
    - [Get the value of a single cell](#get-the-value-of-a-single-cell)
    - [Find a cell by its value](#find-a-cell-by-its-value)
  - [Append Data](#append-data) 
    - [Append a single row](#append-a-single-row) 
    - [Append a range](#append-a-range) 
  - [Update Data](#update-data) 
    - [Update single cell](#update-single-cell) 
    - [Update a range](#update-a-range)
  - [Worksheets](#worksheets) 
    - [Get all worksheets of current spreadsheet](#get-all-worksheets-of-current-spreadsheet)
    - [Duplicate Worksheet](#duplicate-worksheet) 
    - [Delete a worksheet by name](#delete-a-worksheet-by-name)
    - [Rename a worksheet](#rename-a-worksheet)
    - [Add a new worksheet](#add-a-new-worksheet)
    - [Change background color of a range](#change-background-color-of-a-range) 
    - [Clear all values in a range](#clear-all-values-in-a-range)
  - [Misc](#misc) 
    - [Create a new spreadsheet](#create-a-new-spreadsheet)
    - [Calculate column index by the letters](#calculate-column-index-by-the-letters)  
* [License](#license) 
* [Question? Issues?](#questions-issues) 
* [Who's Behind](#whos-behind) 

---

Requirements
------------

This library requires the following:

- Dependent on [Google Client API](https://developers.google.com/sheets/api/quickstart/php)
    - PHP 5.4 or greater with the command-line interface (CLI) and 
    - PHP extension JSON installed
    - A Google Cloud Platform project with the API enabled. To create a project and enable an API, refer to Create a project and enable the API

---

Installation
------------

Run Composer in your project:

    composer require reandimo/google-sheets-helper
    
Then you could call it after Composer is loaded depending on your PHP framework:

```php
  require __DIR__ . '/vendor/autoload.php';

  $credentialFilePath = 'path/to/credentials.json';
  $tokenPath = 'path/to/token.json';
  $sheet1 = new \reandimo\GoogleSheetsApi\Helper($credentialFilePath, $tokenPath);
```
    
Also, you can use putenv() to set credentials.json and token.json like this:

```php
  require __DIR__ . '/vendor/autoload.php';

  putenv('credentialFilePath=path/to/credentials.json');
  putenv('tokenPath=path/to/token.json');
  $sheet1 = new \reandimo\GoogleSheetsApi\Helper();
```

Now when you create a new instance, the class automatically detects the paths. This is the recommended way to do it.

---

Credentials
------------

Google API Client needs to validate with 2 files credentials.json and token.json, the last one can be generated with a script included in the package called firstauth. You can use it to generate this file for the first time only to grant access to the API.

Execute in your project's root folder with `php ./vendor/reandimo/google-sheets-helper/firstauth` and follow the steps.

This is a 3 step script based on the quickstart.php mentioned in google's documentation (https://developers.google.com/sheets/api/quickstart/php).

<img src="https://i.imgur.com/CSWXMrq.gif" width="100%"/>

---

Usage
-----

### Create Instance

#### Set sheet props

You can have multiple sheet instances just invoke the Helper as many times you want:

```php
  use reandimo\GoogleSheetsApi\Helper;

  putenv('credentialFilePath=path/to/credentials.json');
  putenv('tokenPath=path/to/token.json');

  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somespreadsheetid');
  $sheet1->setWorksheetName('Sheet1');
  $sheet1->setSpreadsheetRange('A1:A20');

  $sheet2 = new Helper();
  $sheet2->setSpreadsheetId('somespreadsheetid');
  $sheet2->setWorksheetName('Sheet2');
  $sheet2->setSpreadsheetRange('B1:B20');
```
### Get Values

#### Get values from range
```php
  $sheet1->setSpreadsheetRange('A1:A3');
  $insert = $sheet1->get();
```

#### Get the value of a single cell
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$sheet1->setWorksheetName('Sheet1');
$value = $sheet1->getSingleCellValue('B2');
echo "Value in B2: $value\n";
```

#### Find a cell by its value
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$sheet1->setWorksheetName('Sheet1');
$sheet1->setSpreadsheetRange('A1:Z100');
$result = $sheet1->findCellByValue('searchValue');
if ($result) {
    echo "Found at cell: {$result['cell']} (row {$result['row']}, column {$result['column']})\n";
} else {
    echo 'Value not found.';
}
```

### Append Data

#### Append a single row
```php
  $sheet1->setSpreadsheetRange('A1:A3');
  $insert = $sheet1->appendSingleRow([
    'some',
    'useful',
    'data',
  ]);
```

The function will return a number of rows updated as int. So you can check if it's done like this:

```php
  if($insert >= 1){
    echo 'Insert done Hackerman.';
  }
```

#### Append a range
```php
  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somespreadsheetid');
  $sheet1->setWorksheetName('Sheet1');
  $sheet1->setSpreadsheetRange('A1:A');
  $i = $sheet1->append([
      ['test4', 'this4', 'shit4'],
      ['test2', 'this2', 'shit2'],
      ['test3', 'this3', 'shit3'],
  ]);
```

### Update Data

#### Update single cell
```php
  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somespreadsheetid');
  $sheet1->setWorksheetName('Sheet1');
  $update = $sheet1->updateSingleCell('B5', "Hi i'm a test!");
  if($update->getUpdatedCells() >= 1){
    echo 'Cell updated.';
  }
```

#### Update a range
```php
  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somesheetid');
  $sheet1->setWorksheetName('Sheet1');
  $sheet1->setSpreadsheetRange('A1:F5');
  $update = $sheet1->update([
      ['val1', 'test2', 'int3', 'four', '5', 'six6'],
      ['val1', 'test2', 'int3', 'four', '5', 'six6'],
      ['val1', 'test2', 'int3', 'four', '5', 'six6'],
      ['val1', 'test2', 'int3', 'four', '5', 'six6'],
      ['val1', 'test2', 'int3', 'four', '5', 'six6'],
  ]);

  // Get updated cells
  if($update->getUpdatedCells() >= 1){
    echo 'Range updated.';
  }
```

### Worksheets

#### Get all worksheets of current spreadsheet
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$worksheets = $sheet1->getSpreadsheetWorksheets();
foreach ($worksheets as $ws) {
    echo "Sheet ID: {$ws['id']}, Title: {$ws['title']}\n";
}
```

#### Duplicate Worksheet
```php
  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somesheetid');
  $sheet1->setWorksheetName('Sheet1'); // select the Worksheet you want to duplicate
  $newWorksheetName = 'New Duplicated Sheet'; // The name of the new sheet
  $sheet_id = $sheets->duplicateWorksheet($newWorksheetName);
  
  // Get updated cells
  if($sheet_id){
    echo 'The sheet was duplicated B)';
  }
```


#### Delete a worksheet by name
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$deleted = $sheet1->deleteWorksheet('SheetToDelete');
if ($deleted) {
    echo 'Worksheet deleted.';
}
```

#### Rename a worksheet
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$renamed = $sheet1->renameWorksheet('OldSheetName', 'NewSheetName');
if ($renamed) {
    echo 'Worksheet renamed.';
}
```

#### Add a new worksheet
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$newSheetId = $sheet1->addWorksheet('NewSheet', 100, 10);
echo "New worksheet ID: $newSheetId\n";
```

#### Change background color of a range
```php
  $sheet1 = new Helper();
  $sheet1->setSpreadsheetId('somespreadsheetid');
  $sheet1->setWorksheetName('Sheet1');
  $sheet1->setSpreadsheetRange('A1:Z10');
  $sheet1->colorRange([142, 68, 173]);
```

#### Clear all values in a range
```php
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('somespreadsheetid');
$sheet1->setWorksheetName('Sheet1');
$sheet1->setSpreadsheetRange('A1:Z100');
$cleared = $sheet1->clearRange();
if ($cleared) {
    echo 'Range cleared.';
}
```

### Misc

#### Create a new spreadsheet
```php
$sheet1 = new Helper();
$newSpreadsheetId = $sheet1->create('My New Spreadsheet');
echo "Created spreadsheet with ID: $newSpreadsheetId\n";
```

#### Calculate column index by the letters
If for some reason you need to calculate the column positions of a column by its letters, this is the way:

```php 
  Helper::getColumnLettersIndex('AZ'); // this will return 52
```

Tips
------------

Some things aren't very clear in Google's documentation without diggin a lot so i'll be leaving tips here:

- To leave blank a cell when you do an insert or update, you have to use this const: ``` Google_Model::NULL_VALUE ```.

  Example: 

  ```php 
    $sheet1->appendSingleRow([
      'John Doe',
      'jhon@doe.com',
      Google_Model::NULL_VALUE,
      'Sagittarius',
    ]);
  ``` 

## License

This Package is open source and released under MIT license. See the LICENSE file for more info.

## Question? Issues?

Feel free to open issues for suggestions, questions, and issues.

## Who's Behind

Renan Diaz, i'm dealing with PHP since 2017 & Google's API since 2019. Feel free to write me to my email (Please don't send any multi-level crap).


