<?php

namespace reandimo\GoogleSheetsApi;

use Exception;
use \Google_Client;
use \Google_Service_Sheets;
use \Google_Service_Sheets_NamedRange;

use \Google_Service_Sheets_Request;
use \Google_Service_Sheets_BatchUpdateSpreadsheetRequest;
use \Google_Service_Sheets_ValueRange;
use \Google_Service_Sheets_CopySheetToAnotherSpreadsheetRequest;

/**
 * Google Spreadsheet API Helper
 * 
 * @author      Renan Diaz <reandimo23@gmail.com>
 * @version     1.1.2
 * @filesource 	Google APIs Client Library for PHP <https://github.com/googleapis/google-api-php-client>
 * @see         https://github.com/reandimo/google-sheets-helper 
 * 
 */

class Helper
{

    const LETTERS = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
    ];

    /**
     * @var Google_Client the authorized client object
     */
    public $client;

    /**
     * @var Google_Service_Sheets sheets service class
     */
    public $service;

    /**
     * @var string the ID of the sheet in use
     */
    public $spreadsheetId;

    /**
     * @var string the current name of the sheet in use
     */
    public $worksheetName;

    /**
     * @var string the current range of the sheet in use
     */
    public $range;

    /**
     * @var string the current input option. RAW = insert data entered as it. USER_ENTERED = The values will be parsed as if the user typed them into the UI. <https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption>
     */
    public $valueInputOption = "RAW";

    /**
     * @var string absolute path of credential file location
     */
    public $credentialFilePath;

    /**
     * @var string absolute path of token file location for auth, if not exist you need to follow the CLI steps for first time auth: <https://developers.google.com/sheets/api/quickstart/php#step_2_set_up_the_sample>
     */
    public $tokenPath;

    /**
     * @var string custom app name for auth in google: <https://developers.google.com/sheets/api/quickstart/php#step_2_set_up_the_sample>
     */
    public $appName;

    public function __construct(?string $credentialFilePath = null, ?string $tokenPath = null)
    {

        ## ENV SETUP
        if( getenv('credentialFilePath') && getenv('tokenPath') ){
            if (!file_exists(getenv('credentialFilePath'))) {
                throw new Exception("No credential file in: ".getenv('credentialFilePath'));
            }
            $this->tokenPath = getenv('tokenPath');
            $this->credentialFilePath = getenv('credentialFilePath');
        }

        ##PARAMS SETUP
        else{
            if (!file_exists($credentialFilePath)) {
                throw new Exception("No credential file in: {$credentialFilePath}");
            }
            $this->tokenPath = $tokenPath;
            $this->credentialFilePath = $credentialFilePath;
        }

        ## SETUP CLIENT
        if (!empty($this->tokenPath) && file_exists($this->tokenPath)) {
            $this->client = $this->getClient();
            $this->service = new \Google_Service_Sheets($this->client);
        }

    }

    /**
     * Set the current spreadsheet ID.
     * The spreadsheet ID is the unique identifier of the spreadsheet.
     * @param string|null $spreadsheetId
     * @return void
     */
    public function setSpreadsheetId(?string $spreadsheetId)
    {
        $this->spreadsheetId = $spreadsheetId;
    }

    /**
     * Get the current spreadsheet ID.
     * The spreadsheet ID is the unique identifier of the spreadsheet.
     * 
     * @return string|null
     */
    public function getSpreadsheetId()
    {
        return $this->spreadsheetId;
    }

    public function getService()
    {
        return $this->service;
    }

    /**
    * Set the range of the spreadsheet to work with.
    * The range should be in A1 notation, e.g. "A1:B10".
    * If you want to set a whole column, use "A:A" or "B:B".
    * If you want to set a whole row, use "1:1" or "2:2".
    * 
    * @param string|null $range
    */
    public function setSpreadsheetRange(?string $range)
    {
        $this->range = $range;
    }

    /**
    * Get the current range of the spreadsheet.
    * The range is in A1 notation, e.g. "A1:B10".
    */
    public function getSpreadsheetRange()
    {
        return $this->range;
    }

    /**
     * Set the name of the worksheet to work with.
     * The worksheet name should be the exact name of the sheet in the spreadsheet.
     * 
     * @param string|null $worksheetName
     */
    public function setWorksheetName(?string $worksheetName)
    {
        $this->worksheetName = $worksheetName;
    }

    /**
     * Get the current name of the worksheet.
     * The worksheet name is the exact name of the sheet in the spreadsheet.
     * 
     * @return string|null
     */
    public function getWorksheetName()
    {
        return $this->worksheetName;
    } 

    /**
     * Set the value input option for the API.
     * The value input option determines how the values are interpreted when written to the sheet.
     * Possible values are "RAW" or "USER_ENTERED".
     * 
     * @param string|null $valueInputOption
     */
    public function setValueInputOption(?string $valueInputOption)
    {
        $this->valueInputOption = $valueInputOption;
    }

    /**
     * Get the current value input option.
     * The value input option determines how the values are interpreted when written to the sheet.
     * Possible values are "RAW" or "USER_ENTERED".
     * 
     * @return string|null
     */
    public function getValueInputOption()
    {
        return $this->valueInputOption;
    } 

    /**
     * Set the application name for the Google Sheets API.
     * This is used for identification purposes in the Google Cloud Console.
     * 
     * @param string|null $appName
     */
    public function firstAuth(?string $tokenPath = null)
    {
        if ($tokenPath == null) {
            throw new Exception("token.json destination filepath not set");
        }

        $this->tokenPath = $tokenPath;
        $client = new \Google_Client();
        !empty($this->appName) ? $client->setApplicationName('Google Sheets API PHP') : $client->setApplicationName($this->appName);
        $client->setScopes(\Google_Service_Sheets::SPREADSHEETS);
        $client->setAuthConfig($this->credentialFilePath);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        // Request authorization from the user.
        $authUrl = $client->createAuthUrl();
        printf("Open the following link in your browser:\n%s\n", $authUrl);
        print 'Enter verification code: ';
        $authCode = trim(fgets(STDIN));

        // Exchange authorization code for an access token.
        $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
        $client->setAccessToken($accessToken);

        // Check to see if there was an error.
        if (array_key_exists('error', $accessToken)) {
            throw new Exception(join(', ', $accessToken));
        } 
        @mkdir(dirname($this->tokenPath), 0700, true);
        @file_put_contents($this->tokenPath, json_encode($client->getAccessToken()));

        if (file_exists($this->tokenPath)) {
            print "Token file successfully created in: {$this->tokenPath}. Congrats now you can access the API.";
        } else {
            throw new Exception("Token file could not be created. Try again or check your log.");
        }
    }

    /**
     * Returns an authorized API client.
     * @return Google_Client the authorized client object
     */
    public function getClient()
    {

        $client = new \Google_Client();
        !empty($this->appName) ? $client->setApplicationName('Google Sheets API PHP') : $client->setApplicationName($this->appName);
        $client->setScopes(\Google_Service_Sheets::SPREADSHEETS);
        $client->setAuthConfig($this->credentialFilePath);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        // Load previously authorized token from a file, if it exists.
        // The file token.json stores the user's access and refresh tokens, and is
        // created automatically when the authorization flow completes for the first
        // time. 

        if (file_exists($this->tokenPath)) {
            $accessToken = json_decode(file_get_contents($this->tokenPath), true);
            $client->setAccessToken($accessToken);
        } else {
            throw new Exception("Token file does not exist in {$this->tokenPath}");
        }

        // If there is no previous token or it's expired.
        if ($client->isAccessTokenExpired()) {
            // Refresh the token if possible, else fetch a new one.
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            } else {
                throw new Exception("You have to run 'firstauth' in your CLI to get a new token.json");
            }
            // Save the token to a file.
            if (!file_exists(dirname($this->tokenPath))) {
                @mkdir(dirname($this->tokenPath), 0700, true);
            }
            @file_put_contents($this->tokenPath, json_encode($client->getAccessToken()));
        }

        return $client;
    }

    /**
     * Creates a new spreadsheet.
     * @param string $title Title of the new spreadsheet.
     * @return string spreadsheet ID
     */
    public function create(?string $title): string
    {
        if (empty($title)) {
            throw new Exception("You have to set a title for the new Spreadsheet.");
        }
        $spreadsheet = new \Google_Service_Sheets_Spreadsheet([
            'properties' => [
                'title' => $title
            ]
        ]);
        $spreadsheet = $this->service->spreadsheets->create($spreadsheet, [
            'fields' => 'spreadsheetId'
        ]);
        return $spreadsheet->spreadsheetId;
    }

    /**
     * Get rows from a range 
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return array values of the range provided.
     * 
     */
    public function get()
    {
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }

        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        $range = "{$this->getWorksheetName()}!{$this->getSpreadsheetRange()}";
        
        $result = $this->service->spreadsheets_values->get(
            $this->getSpreadsheetId(),
            $range
        );
        return $result->getValues();
    }

    /**
     * Append rows after last row of the current spreadsheet
     * @param array $rowsData Array of values to insert in sheets. Must be a multi-dimensional array.
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return int The number of updated rows.
     * 
     */
    public function append(?array $rowsData = [])
    {
        
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }

        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $rowsData]);
        $append = $this->service->spreadsheets_values->append(
            $this->getSpreadsheetId(),
            "{$this->getWorksheetName()}!{$this->getSpreadsheetRange()}",
            $valueRange,
            ["valueInputOption" => $this->getValueInputOption()],
            ["insertDataOption" => "INSERT_ROWS"]
        );
        return $append->getUpdates();

    } 

    /**
     * Append a single row after last row of the current spreadsheet
     * @param array $row Array of values to append in sheets
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return Google\Service\Sheets\UpdateValuesResponse Object containing data of recent updates.
     * 
     */
    public function appendSingleRow(?array $row = []): object
    {
        
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }

        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => [$row]]);
        $insert = $this->service->spreadsheets_values->append(
            $this->getSpreadsheetId(),
            "{$this->getWorksheetName()}!{$this->getSpreadsheetRange()}",
            $valueRange,
            ["valueInputOption" => $this->getValueInputOption()],
            ["insertDataOption" => "INSERT_ROWS"]
        );
        return $insert->getUpdates();
    }
    
    /**
     * Update a range of the current spreadsheet
     * @param array $newValues The ID of spreadsheet to insert. Must be a multi-dimensional array.
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return Google\Service\Sheets\UpdateValuesResponse The response object.
     * 
     */
    public function update(?array $newValues = []): object
    {

        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }

        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $newValues]);
        $updateSheet = $this->service->spreadsheets_values->update(
            $this->getSpreadsheetId(),
            "{$this->getWorksheetName()}!{$this->getSpreadsheetRange()}",
            $valueRange,
            ["valueInputOption" => $this->getValueInputOption()]
        );

        return $updateSheet;
    }

    /**
     * Quick function to update a single cell in current worksheet
     * @param string $cell Column letter and row number of cell to update. Example: A1
     * @param string $value New value to set. 
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return object The number of updated rows.
     * 
     */
    public function updateSingleCell(?string $cell = null, ?string $value = null): object
    {
        
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if ($value == null) {
            throw new Exception("There's no value to set.");
        }

        if ($cell == null) {
            throw new Exception("There's no cell to set a range.");
        }

        $range = "{$this->worksheetName}!{$cell}:{$cell[0]}";
        $wrappedValue = [[$value]];
        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $wrappedValue]);
        $updateSheet = $this->service->spreadsheets_values->update(
            $this->getSpreadsheetId(), 
            $range, 
            $valueRange, 
            ["valueInputOption" => $this->getValueInputOption()]
        );
        return $updateSheet;
    }

    /**
     * Calculates the index of column letter IDs given
     * @param string $letters columns letters IDs to calculate. Example: BZ
     * @return int
     * 
     */
    public static function getColumnLettersIndex(?string $letters = null): int
    {
        $letterCount = strlen($letters);
        if($letterCount == 1){
            return array_search($letters, self::LETTERS) + 1;
        }else{
            $index = 0;
            $lastElementIndex = $letterCount-1;
            for ($i=0; $i < $letterCount; $i++) { 
                if( $letters[$i] == $letters[$lastElementIndex] ){
                    $index = $index + Helper::getColumnLettersIndex($letters[$lastElementIndex]);
                }else{
                    $index = $index + (count(self::LETTERS) * Helper::getColumnLettersIndex($letters[0]));
                }
            }
            return $index;
        }
    } 

    /**
     * Change background of a given range
     * @param array $rgb RGB color code. Example: [142, 68, 173, 1.0]
     * @return void
     * 
     */
    public function colorRange(?array $rgb): void
    { 

        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }
        
        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }
        
        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        if (count($rgb) < 3) {
            throw new Exception("RGB not correctly configured", 1);
        }

        $config['r'] = ((int)$rgb[0] / 255);
        $config['g'] = ((int)$rgb[1] / 255);
        $config['b'] = ((int)$rgb[2] / 255);
        $config['a'] = isset($rgb[3]) ? (float)$config['rgb'][3] : 1.0;

        $sheetId = $this->service->spreadsheets->get($this->getSpreadsheetId(), ['ranges' => $this->getWorksheetName()]);
        $ranges = explode(':', $this->getSpreadsheetRange());
        $columnStartLetters = preg_replace('/[0-9]+/', '', $ranges[0]);
        $columnEndLetters = preg_replace('/[0-9]+/', '', $ranges[1]);
        $rowStart = preg_replace("/[^0-9]/", "", $ranges[0]);
        $rowEnd = (!empty($ranges[1])) ? preg_replace("/[^0-9]/", "", $ranges[1]) : $rowStart;
        $myRange = [
            'sheetId' => $sheetId->sheets[0]->properties->sheetId,
            'startRowIndex' => $rowStart - 1,
            'endRowIndex' => $rowEnd,
            'startColumnIndex' => $this->getColumnLettersIndex($columnStartLetters) - 1,
            'endColumnIndex' => $this->getColumnLettersIndex($columnEndLetters),
        ];

        $format = [
            "backgroundColor" => [
                "red" => $config['r'],
                "green" => $config['g'],
                "blue" => $config['b'],
                "alpha" => $config['a'],
            ],
        ];

        $requests = [
            new \Google_Service_Sheets_Request([
                'repeatCell' => [
                    'fields' => 'userEnteredFormat.backgroundColor',
                    'range' => $myRange,
                    'cell' => [
                        'userEnteredFormat' => $format,
                    ],
                ],
            ])
        ];

        $batchUpdateRequest = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => $requests
        ]);

        $this->service->spreadsheets->batchUpdate(
            $this->getSpreadsheetId(),
            $batchUpdateRequest
        );

    } 

    /**
     * Get all worksheets of current spreadsheet
     * @return array Returns an array of worksheets with their IDs and titles.
     */
	public function getSpreadsheetWorksheets() : array
    {
        
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        $spreadSheet = $this->service->spreadsheets->get($this->getSpreadsheetId());
        $sheets = $spreadSheet->getSheets();
        $formattedSheet = [];
        foreach($sheets as $sheet) {
            $formattedSheet[] = [
                'id' => $sheet->properties->sheetId,
                'title' => $sheet->properties->title
            ];
        }   
        return $formattedSheet;
	}  

    /**
     * duplicate current worksheet with a new name into the same spreadsheet
     * @param string $newWorksheetName name of the new Worksheet 
     * @see https://developers.google.com/resources/api-libraries/documentation/sheets/v4/php/latest/class-Google_Service_Sheets_CopySheetToAnotherSpreadsheetRequest.html
     * @return int The ID of the updated spreadsheet
     */
    public function duplicateWorksheet(?string $newWorksheetName) : int
    {
        
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }

        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }
        
        if (empty($newWorksheetName)) {
            throw new Exception("You should set the new Worksheet name.");
        }
        
        $spreadsheet = $this->service->spreadsheets->get($this->getSpreadsheetId());
        $sheetId = null;
        foreach ($spreadsheet->getSheets() as $sheet) {
            if ($sheet->properties->title == $this->getWorksheetName()) {
                $sheetId = $sheet->properties->sheetId;
                break;
            }
        }

        if(empty($sheetId)){
            throw new Exception("Worksheet with name '{$this->getWorksheetName()}' was not found.");
        }

        ## Copy data to new Worksheet
        $request = new Google_Service_Sheets_CopySheetToAnotherSpreadsheetRequest([
            'destinationSpreadsheetId' => $this->getSpreadsheetId(),
        ]);
        $duplicatedWorksheet = $this->service->spreadsheets_sheets->copyTo($this->getSpreadsheetId(), $sheetId, $request);

        ## Change name of the new Worksheet 
        $duplicatedWorksheet->setTitle($newWorksheetName);

        ## Save changes
        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [
                [
                    'updateSheetProperties' => [
                        'properties' => $duplicatedWorksheet,
                        'fields' => 'title',
                    ],
                ],
            ],
        ]);
        
        $newSheet = $this->service->spreadsheets->batchUpdate($this->getSpreadsheetId(), $batchUpdateRequest);
        if(!$newSheet->spreadsheetId){
            throw new Exception("Could not create the spreadsheet.");
        }

        return (int)$newSheet->spreadsheetId;

    }

    /**
     * Find a cell by its value in the current worksheet.
     * This function searches through all cells in the specified range of the current worksheet.
     * @param string $value value to search in the current worksheet
     * @param int $limit The maximum number of cells to find. Default is 1. If set to 0, it will return all cells found.
     * @return array|null Returns an array with the cell coordinates or null if not found.
     */
    public function findCellByValue(?string $value, int $limit = 1): ?array
    {
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set. Use: 'setSpreadsheetId' before and try again.");
        }
        if (empty($this->getWorksheetName())) {
            throw new Exception("There's no worksheet range set. Use: 'setWorksheetName' before and try again.");
        }
        if (empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set. Use: 'setSpreadsheetRange' before and try again.");
        }

        $range = "{$this->getWorksheetName()}!{$this->getSpreadsheetRange()}";
        $values = $this->service->spreadsheets_values->get($this->getSpreadsheetId(), $range)->getValues();

        $rangeParts = explode(':', $this->getSpreadsheetRange());
        $startColLetters = preg_replace('/[0-9]+/', '', $rangeParts[0]);
        $startColIndex = max(self::getColumnLettersIndex($startColLetters) - 1, 0);
        $startRow = max((int)preg_replace('/[^0-9]/', '', $rangeParts[0]), 1);

        $results = [];
        foreach ($values as $rowIndex => $row) {
            foreach ($row as $colIndex => $cellValue) {
                if ($cellValue === $value) {
                    $realColIndex = $startColIndex + $colIndex;
                    $realRowIndex = $startRow + $rowIndex;
                    
                    $results[] = [
                        'row' => $realRowIndex,
                        'column' => $realColIndex,
                        'cell' => (self::LETTERS[$realColIndex] ?? '') . $realRowIndex,
                        // DEBUG INFO
                        'context' => [
                            'startLetter' => $startColLetters,
                            'startColumnIndex' => $startColIndex,
                            'startRow' => $startRow,
                            'range' => $this->getSpreadsheetRange(),
                            'value' => $cellValue,
                        ]
                    ];
                    if ($limit > 0 && count($results) >= $limit) {
                        return $limit === 1 ? $results[0] : $results;
                    }
                }
            }
        }
        return $limit === 1 ? null : $results;
    } 

}