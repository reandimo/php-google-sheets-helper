<?php

require dirname(__DIR__, 2) . '/vendor/autoload.php';
require dirname(__DIR__, 2) . '/src/apiclient/Helper.php';

use \PHPUnit\Framework\TestCase;
use reandimo\GoogleSheetsApi\Helper;
use Exception;

class SheetsTests extends TestCase
{

    const GOOGLE_SHEET_TEST_ID = 'XXXXXXXX';

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
        putenv('credentialFilePath=' . dirname(__DIR__, 2) . '/credentials.json');
        putenv('tokenPath=' . dirname(__DIR__, 2) . '/token.json');
        $sheet1 = new Helper();
        $this->assertEquals($sheet1->getColumnLettersIndex('AZ'), 52);
        $this->assertEquals($sheet1->getColumnLettersIndex('CZ'), 104);
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
}
