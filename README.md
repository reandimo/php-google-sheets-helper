# google-sheets-helper
PHP Google Spreadsheets Helper
======================

PHP Google Spreadsheets - A bunch of functions to work easily with Google Sheets API

This library is a helper that encapsulate [Google APIs Client Library for PHP](https://github.com/googleapis/google-api-php-client) ([Documentation](https://developers.google.com/sheets/api/quickstart/php)) for simple usage.

--- 

OUTLINE
-------

* [Installation](#installation)
* [Requirements](#requirements)
* [Usage](#usage)
  - [Import & Export](#import--export) 

---

REQUIREMENTS
------------

This library requires the following:

- Dependent on [Google Client API](https://developers.google.com/sheets/api/quickstart/php)
    - PHP 7.4 or greater with the command-line interface (CLI) and 
    - PHP extension JSON installed
    - A Google Cloud Platform project with the API enabled. To create a project and enable an API, refer to Create a project and enable the API

---

INSTALLATION
------------

Run Composer in your project:

    composer require reandimo/google-sheets-helper
    
Then you could call it after Composer is loaded depended on your PHP framework:

```php
require __DIR__ . '/vendor/autoload.php';

$credentialFilePath = 'path/to/credentials.json';
$tokenPath = 'path/to/token.json';
$sheet1 = new \reandimo\GoogleSheetsAPI\Helper($credentialFilePath, $tokenPath);
```
    
---

CREDENTIALS
------------

Google API Client needs to validate with 2 files credentials.json and token.json, the last one can be generated with a script included in the package called firstauth. You can use to generate this file for first time only to grant access to the API.

Execute in your project's root folder with `php ./vendor/reandimo/google-sheets-helper/firstauth` and follow the steps.

This is a 3 step script based on the quickstart.php mentioned in google's documentation (https://developers.google.com/sheets/api/quickstart/php).

---

USAGE
-----

### Set sheets props

You can have multiple sheet instances just invoke the Helper as many times you want:

```php
$sheet1->setSpreadsheetId('somespreadsheetid');
$sheet1->setWorksheetName('Sheet1');
$sheet1->setSpreadsheetRange('A1:A20');
```
