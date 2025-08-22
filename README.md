# PHP Google Spreadsheets API Helper

[![Latest Stable Version](http://poser.pugx.org/reandimo/google-sheets-helper/v)](https://packagist.org/packages/reandimo/google-sheets-helper) [![Total Downloads](http://poser.pugx.org/reandimo/google-sheets-helper/downloads)](https://packagist.org/packages/reandimo/google-sheets-helper) [![License](http://poser.pugx.org/reandimo/google-sheets-helper/license)](https://packagist.org/packages/reandimo/google-sheets-helper) [![PHP Version Require](http://poser.pugx.org/reandimo/google-sheets-helper/require/php)](https://packagist.org/packages/reandimo/google-sheets-helper)

A comprehensive helper library that encapsulates the [Google APIs Client Library for PHP](https://github.com/googleapis/google-api-php-client) for easy Google Sheets API integration. This library simplifies common operations and provides a clean interface for working with Google Spreadsheets.

## Table of Contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Credentials Setup](#credentials-setup)
- [Usage](#usage)
  - [Creating Instances](#creating-instances)
  - [Reading Data](#reading-data)
  - [Writing Data](#writing-data)
  - [Updating Data](#updating-data)
  - [Worksheet Management](#worksheet-management)
  - [Utility Functions](#utility-functions)
- [Error Handling](#error-handling)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Support](#support)
- [Contributing](#contributing)

## Requirements

- PHP 7.4 or greater
- PHP JSON extension
- Composer
- Google Cloud Platform project with Google Sheets API enabled
- [Google APIs Client Library for PHP](https://github.com/googleapis/google-api-php-client)

## Installation

Install via Composer:

```bash
composer require reandimo/google-sheets-helper
```

## Credentials Setup

### 1. Google Cloud Platform Setup

1. Create a new project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the Google Sheets API
3. Create credentials (Service Account or OAuth 2.0)
4. Download the `credentials.json` file

### 2. Generate Token File

Use the included authentication script:

```bash
php ./vendor/reandimo/google-sheets-helper/firstauth
```

Follow the interactive prompts to generate your `token.json` file.

![Authentication Process](https://i.imgur.com/CSWXMrq.gif)

## Usage

### Creating Instances

#### Method 1: Environment Variables (Recommended)

```php
<?php
require __DIR__ . '/vendor/autoload.php';

use reandimo\GoogleSheetsApi\Helper;

// Set environment variables
putenv('credentialFilePath=path/to/credentials.json');
putenv('tokenPath=path/to/token.json');

$sheets = new Helper();
$sheets->setSpreadsheetId('your-spreadsheet-id');
$sheets->setWorksheetName('Sheet1');
$sheets->setSpreadsheetRange('A1:Z1000');
```

#### Method 2: Direct File Paths

```php
<?php
require __DIR__ . '/vendor/autoload.php';

use reandimo\GoogleSheetsApi\Helper;

$credentialFilePath = 'path/to/credentials.json';
$tokenPath = 'path/to/token.json';

$sheets = new Helper($credentialFilePath, $tokenPath);
$sheets->setSpreadsheetId('your-spreadsheet-id');
$sheets->setWorksheetName('Sheet1');
```

#### Multiple Instances

```php
<?php
use reandimo\GoogleSheetsApi\Helper;

// Configure environment variables once
putenv('credentialFilePath=path/to/credentials.json');
putenv('tokenPath=path/to/token.json');

// Create multiple instances for different sheets
$sheet1 = new Helper();
$sheet1->setSpreadsheetId('spreadsheet-1-id');
$sheet1->setWorksheetName('Sheet1');
$sheet1->setSpreadsheetRange('A1:A20');

$sheet2 = new Helper();
$sheet2->setSpreadsheetId('spreadsheet-2-id');
$sheet2->setWorksheetName('Sheet2');
$sheet2->setSpreadsheetRange('B1:B20');
```

### Reading Data

#### Get Values from a Range

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:C10');
    $values = $sheets->get();
    
    if ($values && !empty($values)) {
        foreach ($values as $row) {
            echo implode(', ', $row) . "\n";
        }
    } else {
        echo "No data found in range.\n";
    }
} catch (Exception $e) {
    echo "Error reading data: " . $e->getMessage() . "\n";
}
```

#### Get Single Cell Value

```php
<?php
try {
    $value = $sheets->getSingleCellValue('B2');
    echo "Value in B2: " . ($value ?: 'empty') . "\n";
} catch (Exception $e) {
    echo "Error reading cell: " . $e->getMessage() . "\n";
}
```

#### Find Cell by Value

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:Z100');
    $result = $sheets->findCellByValue('search term');
    
    if ($result) {
        echo "Found at cell: {$result['cell']} (row {$result['row']}, column {$result['column']})\n";
    } else {
        echo "Value not found.\n";
    }
} catch (Exception $e) {
    echo "Error searching: " . $e->getMessage() . "\n";
}
```

### Writing Data

#### Append Single Row

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:C1');
    $rowsUpdated = $sheets->appendSingleRow([
        'John Doe',
        'john@example.com',
        'Developer'
    ]);
    
    if ($rowsUpdated >= 1) {
        echo "Row appended successfully.\n";
    } else {
        echo "Failed to append row.\n";
    }
} catch (Exception $e) {
    echo "Error appending row: " . $e->getMessage() . "\n";
}
```

#### Append Multiple Rows

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:C');
    $rowsUpdated = $sheets->append([
        ['Jane Smith', 'jane@example.com', 'Designer'],
        ['Bob Johnson', 'bob@example.com', 'Manager'],
        ['Alice Brown', 'alice@example.com', 'Analyst']
    ]);
    
    echo "Appended {$rowsUpdated} rows.\n";
} catch (Exception $e) {
    echo "Error appending rows: " . $e->getMessage() . "\n";
}
```

#### Handle Empty Cells

```php
<?php
use Google_Model;

$sheets->appendSingleRow([
    'John Doe',
    'john@example.com',
    Google_Model::NULL_VALUE, // Leave cell empty
    'Active'
]);
```

### Updating Data

#### Update Single Cell

```php
<?php
try {
    $update = $sheets->updateSingleCell('B5', 'Updated value');
    
    if ($update->getUpdatedCells() >= 1) {
        echo "Cell updated successfully.\n";
    } else {
        echo "No cells were updated.\n";
    }
} catch (Exception $e) {
    echo "Error updating cell: " . $e->getMessage() . "\n";
}
```

#### Update Range

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:F5');
    $update = $sheets->update([
        ['Header1', 'Header2', 'Header3', 'Header4', 'Header5', 'Header6'],
        ['Data1', 'Data2', 'Data3', 'Data4', 'Data5', 'Data6'],
        ['Data7', 'Data8', 'Data9', 'Data10', 'Data11', 'Data12'],
        ['Data13', 'Data14', 'Data15', 'Data16', 'Data17', 'Data18'],
        ['Data19', 'Data20', 'Data21', 'Data22', 'Data23', 'Data24']
    ]);
    
    echo "Updated {$update->getUpdatedCells()} cells.\n";
} catch (Exception $e) {
    echo "Error updating range: " . $e->getMessage() . "\n";
}
```

### Worksheet Management

#### Get All Worksheets

```php
<?php
try {
    $worksheets = $sheets->getSpreadsheetWorksheets();
    
    foreach ($worksheets as $worksheet) {
        echo "Sheet ID: {$worksheet['id']}, Title: {$worksheet['title']}\n";
    }
} catch (Exception $e) {
    echo "Error getting worksheets: " . $e->getMessage() . "\n";
}
```

#### Duplicate Worksheet

```php
<?php
try {
    $sheets->setWorksheetName('Sheet1');
    $newSheetId = $sheets->duplicateWorksheet('Copy of Sheet1');
    
    if ($newSheetId) {
        echo "Worksheet duplicated successfully. New ID: {$newSheetId}\n";
    } else {
        echo "Failed to duplicate worksheet.\n";
    }
} catch (Exception $e) {
    echo "Error duplicating worksheet: " . $e->getMessage() . "\n";
}
```

#### Delete Worksheet

```php
<?php
try {
    $deleted = $sheets->deleteWorksheet('SheetToDelete');
    
    if ($deleted) {
        echo "Worksheet deleted successfully.\n";
    } else {
        echo "Failed to delete worksheet.\n";
    }
} catch (Exception $e) {
    echo "Error deleting worksheet: " . $e->getMessage() . "\n";
}
```

#### Rename Worksheet

```php
<?php
try {
    $renamed = $sheets->renameWorksheet('OldName', 'NewName');
    
    if ($renamed) {
        echo "Worksheet renamed successfully.\n";
    } else {
        echo "Failed to rename worksheet.\n";
    }
} catch (Exception $e) {
    echo "Error renaming worksheet: " . $e->getMessage() . "\n";
}
```

#### Add New Worksheet

```php
<?php
try {
    $newSheetId = $sheets->addWorksheet('NewSheet', 100, 10);
    echo "New worksheet created with ID: {$newSheetId}\n";
} catch (Exception $e) {
    echo "Error creating worksheet: " . $e->getMessage() . "\n";
}
```

#### Formatting

```php
<?php
try {
    // Change background color (RGB values)
    $sheets->setSpreadsheetRange('A1:Z10');
    $sheets->colorRange([142, 68, 173]); // Purple
    
    // Clear range values
    $cleared = $sheets->clearRange();
    if ($cleared) {
        echo "Range cleared successfully.\n";
    }
} catch (Exception $e) {
    echo "Error formatting: " . $e->getMessage() . "\n";
}
```

### Utility Functions

#### Create New Spreadsheet

```php
<?php
try {
    $newSpreadsheetId = $sheets->create('My New Spreadsheet');
    echo "Created new spreadsheet with ID: {$newSpreadsheetId}\n";
} catch (Exception $e) {
    echo "Error creating spreadsheet: " . $e->getMessage() . "\n";
}
```

#### Column Index Calculation

```php
<?php
use reandimo\GoogleSheetsApi\Helper;

$columnIndex = Helper::getColumnLettersIndex('AZ'); // Returns 52
echo "Column AZ is at index: {$columnIndex}\n";

$columnIndex = Helper::getColumnLettersIndex('AA'); // Returns 27
echo "Column AA is at index: {$columnIndex}\n";
```

## Error Handling

Always wrap your API calls in try-catch blocks to handle potential errors gracefully:

```php
<?php
try {
    $sheets->setSpreadsheetRange('A1:Z100');
    $values = $sheets->get();
    
    // Process data...
    
} catch (Google_Service_Exception $e) {
    // Handle Google API specific errors
    echo "Google API Error: " . $e->getMessage() . "\n";
    echo "Error Code: " . $e->getCode() . "\n";
    
} catch (Exception $e) {
    // Handle general errors
    echo "General Error: " . $e->getMessage() . "\n";
}
```

## Best Practices

1. **Use Environment Variables**: Store credentials securely using environment variables
2. **Error Handling**: Always implement proper error handling
3. **Batch Operations**: Use batch operations when updating multiple cells
4. **Rate Limiting**: Be mindful of Google's API rate limits
5. **Caching**: Cache frequently accessed data when possible
6. **Validation**: Validate data before sending to the API

## Troubleshooting

### Common Issues

#### Authentication Errors
- Ensure `credentials.json` and `token.json` files exist and are readable
- Check that the Google Sheets API is enabled in your project
- Verify the service account has proper permissions

#### Permission Errors
- Ensure the service account has access to the target spreadsheet
- Check that the spreadsheet is shared with the service account email

#### Rate Limiting
- Implement exponential backoff for failed requests
- Use batch operations to reduce API calls

#### Data Format Issues
- Use `Google_Model::NULL_VALUE` for empty cells
- Ensure data arrays match the specified range dimensions


## License

This package is open source and released under the MIT License. See the [LICENSE](LICENSE) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/reandimo/google-sheets-helper/issues)
- **Documentation**: [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- **Stack Overflow**: Tag questions with `google-sheets-api` and `php`

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Who's Behind

**Renan Diaz** - PHP developer since 2017, Google API specialist since 2019.

Renan Diaz, i'm dealing with PHP since 2017 & Google's API since 2019. Feel free to write me to my email (Please don't send any multi-level crap).
