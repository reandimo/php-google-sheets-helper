<?php

use \PHPUnit\Framework\TestCase;
use reandimo\GoogleSheetsApi\Helper;
use Exception;

class SheetsTests extends TestCase
{

    const GOOGLE_SHEET_TEST_ID = '1-bRRbBCS36VaO7_ewToX7UZVy0JXWR2K97SG8EfZhPk';

    public function testPutEnvCredentials()
    {

        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $this->assertEquals($sheet1->credentialFilePath, getenv('credentialFilePath'));
        $this->assertEquals($sheet1->tokenPath, getenv('tokenPath'));
    }

    public function testPutEnvCredentialsWithWrongPaths()
    {
        $this->expectException(Exception::class);
        putenv('credentialFilePath=doesnotexist/credentials.json');
        putenv('tokenPath=doesnotexist/token.json');
        $sheet1 = new Helper();
    }

    public function testGetSheetsId()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $this->assertEquals($sheet1->getSpreadsheetId(), self::GOOGLE_SHEET_TEST_ID);
    }

    public function testGetSheetsWorkname()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $this->assertEquals($sheet1->getWorksheetName(), 'Sheet1');
    }

    public function testGetSheetsRange()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:A');
        $this->assertEquals($sheet1->getSpreadsheetRange(), 'A1:A');
    }

    public function testGet()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:E5');
        print_r($sheet1->get());
    }

    public function testSingleInsert()
    {
        $credentialFilePath =  dirname(__DIR__, 2) . '/credentials.json';
        $tokenPath = dirname(__DIR__, 2) . '/token.json';
        $sheet1 = new Helper($credentialFilePath, $tokenPath);
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A:C');
        $i = $sheet1->appendSingleRow(['test', 'this', 'shit']);
        $this->assertIsObject($i);
    }

    public function testMultipleInserts()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:A');
        $i = $sheet1->append([
            ['test4', 'this4', 'shit4'],
            ['test2', 'this2', 'shit2'],
            ['test3', 'this3', 'shit3'],
        ]);
        $this->assertIsObject($i);
    }

    public function testColorRange()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:Z10');
        $i = $sheet1->colorRange([142, 68, 173]);
        $this->expectNotToPerformAssertions();
    }

    public function testColumnIndexesFunction()
    {
        $this->assertEquals(Helper::getColumnLettersIndex('AZ'), 52);
        $this->assertEquals(Helper::getColumnLettersIndex('CZ'), 104);
    }

    public function testUpdateSingleCell()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $update = $sheet1->updateSingleCell('B5', "Hi i'm a test!");
        $this->assertIsObject($update);
        $this->assertEquals($update->getUpdatedCells(), 1);
    }

    public function testUpdate()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:F5');
        $update = $sheet1->update([
            ['val1', 'test2', 'int3', 'four', '5', 'six6'],
            ['val1', 'test2', 'int3', 'four', '5', 'six6'],
            ['val1', 'test2', 'int3', 'four', '5', 'six6'],
            ['val1', 'test2', 'int3', 'four', '5', 'six6'],
            ['val1', 'test2', 'int3', 'four', '5', 'six6'],
        ]);
        $this->assertIsObject($update);
        $this->assertEquals($update->getUpdatedCells(), 30);
    }

    public function testGetSpreadsheetWorksheets()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        var_dump($sheet1->getSpreadsheetWorksheets());
    }

    public function testDeleteWorksheet()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        // Add a worksheet to delete
        $newSheetId = $sheet1->addWorksheet('TempSheetToDelete', 10, 5);
        $deleted = $sheet1->deleteWorksheet('TempSheetToDelete');
        $this->assertTrue($deleted);
    }

    public function testRenameWorksheet()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        // Add a worksheet to rename
        $newSheetId = $sheet1->addWorksheet('TempSheetToRename', 10, 5);
        $renamed = $sheet1->renameWorksheet('TempSheetToRename', 'TempSheetRenamed');
        $this->assertTrue($renamed);
        // Clean up
        $sheet1->deleteWorksheet('TempSheetRenamed');
    }

    public function testClearRange()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:Z10');
        $cleared = $sheet1->clearRange();
        $this->assertTrue($cleared);
    }

    public function testAddWorksheet()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $newSheetId = $sheet1->addWorksheet('TempSheetAdd', 10, 5);
        $this->assertIsInt($newSheetId);
        // Clean up
        $sheet1->deleteWorksheet('TempSheetAdd');
    }

    public function testGetSingleCellValue()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $value = $sheet1->getSingleCellValue('A1');
        $this->assertTrue($value === null || is_scalar($value));
    }

    public function testFindCellByValue()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $sheet1->setSpreadsheetId(self::GOOGLE_SHEET_TEST_ID);
        $sheet1->setWorksheetName('Sheet1');
        $sheet1->setSpreadsheetRange('A1:Z100');
        $result = $sheet1->findCellByValue('test');
        $this->assertTrue($result === null || (isset($result['cell']) && isset($result['row']) && isset($result['column'])));
    }

    public function testCreateSpreadsheet()
    {
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $title = 'Test Created Spreadsheet';
        $spreadsheetId = $sheet1->create($title);
        $this->assertIsString($spreadsheetId);
        // Optionally, clean up by deleting the spreadsheet via API if needed
    }
}
