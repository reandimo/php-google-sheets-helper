<?php

namespace reandimo\GoogleSheetsApi;

use Exception;
use \Google_Client;
use \Google_Service_Sheets;
use \Google_Service_Sheets_NamedRange;

/**
 * Google Spreadsheet API Helper
 * 
 * @author      Renan Diaz <reandimo23@gmail.com>
 * @version     1.0.0
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

    public function setSpreadsheetId(?string $spreadsheetId)
    {
        $this->spreadsheetId = $spreadsheetId;
    }

    public function getSpreadsheetId()
    {
        return $this->spreadsheetId;
    }

    public function getService()
    {
        return $this->service;
    }

    public function setSpreadsheetRange(?string $range)
    {
        $this->range = $range;
    }

    public function getSpreadsheetRange()
    {
        return $this->range;
    }

    public function setWorksheetName(?string $worksheetName)
    {
        $this->worksheetName = $worksheetName;
    }

    public function getWorksheetName()
    {
        return $this->worksheetName;
    } 

    public function setValueInputOption(?string $valueInputOption)
    {
        $this->valueInputOption = $valueInputOption;
    }

    public function getValueInputOption()
    {
        return $this->valueInputOption;
    } 

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
            throw new Exception("There's no ID spreadsheet set.");
        }
        if (empty($this->getSpreadsheetRange()) && empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set.");
        }
        $result = $this->service->spreadsheets_values->get(
            $this->getSpreadsheetId(),
            $this->getSpreadsheetRange()
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

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $rowsData]);
        $append = $this->service->spreadsheets_values->append(
            $this->getSpreadsheetId(),
            $this->getSpreadsheetRange(),
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

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => [$row]]);
        $insert = $this->service->spreadsheets_values->append(
            $this->getSpreadsheetId(),
            $this->getSpreadsheetRange(),
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
            throw new Exception("There's no ID spreadsheet set.");
        }

        if (empty($this->getSpreadsheetRange()) && empty($this->getSpreadsheetRange())) {
            throw new Exception("There's no spreadsheet range set.");
        }

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $newValues]);
        $updateSheet = $this->service->spreadsheets_values->update(
            $this->getSpreadsheetId(),
            $this->getSpreadsheetRange(),
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
     * @return object
     * 
     */
	public function getSpreadsheetWorksheets() : object
    {
        if (empty($this->getSpreadsheetId())) {
            throw new Exception("There's no ID spreadsheet set.");
        }
        $spreadSheet = $this->service->spreadsheets->get($this->getSpreadsheetId());
        $sheets = $spreadSheet->getSheets();
        foreach($sheets as $sheet) {
            $sheets[] = $sheet->properties->sheetId;
        }   
        return (object)$sheets;
	} 

}
