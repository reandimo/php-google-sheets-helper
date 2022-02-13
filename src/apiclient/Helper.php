<?php


namespace reandimo\GoogleSheetsApi;

use Exception;

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
     * @var string absolute path of credential file location
     */
    private $credentialFilePath;

    /**
     * @var string absolute path of token file location for auth, if not exist you need to follow the CLI steps for first time auth: <https://developers.google.com/sheets/api/quickstart/php#step_2_set_up_the_sample>
     */
    private $tokenPath;

    /**
     * @var string custom app name for auth in google: <https://developers.google.com/sheets/api/quickstart/php#step_2_set_up_the_sample>
     */
    private $appName;

    public function __construct(?string $credentialFilePath = null, ?string $tokenPath = null)
    {
        if (!file_exists($credentialFilePath)) {
            throw new Exception("No credential file in: {$credentialFilePath}");
        }
        $this->tokenPath = $tokenPath;
        $this->credentialFilePath = $credentialFilePath;
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
        return $this->service;
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
        mkdir(dirname($this->tokenPath), 0700, true);
        file_put_contents($this->tokenPath, json_encode($client->getAccessToken()));

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
     * Inserts a single row after last row of the current spreadsheet
     * @param string $spreadsheetId The ID of spreadsheet to insert
     * @param string $range Range of columns for insert. Example: Sheet1!A:D 
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return int The number of updated rows.
     * 
     */
    public function insertSingleRow(?array $row = [])
    {

        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $row]);
        $this->service->spreadsheets_values->append(
            $this->getSpreadsheetId(),
            $this->range,
            $valueRange,
            ["valueInputOption" => "RAW"],
            ["insertDataOption" => "INSERT_ROWS"]
        );
    }

    /**
     * Update a single row after last row of the current spreadsheet
     * @param string $spreadsheetId The ID of spreadsheet to insert
     * @param string $range Range of columns for insert. Example: Sheet1!A:D 
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return int The number of updated rows.
     * 
     */
    public function updateSingleRow(?array $row = []): int
    {

        if (empty($this->getSpreadsheetId()) && empty($spreadsheetId)) {
            throw new Exception("There's no ID spreadsheet set.");
        }

        if (empty($this->getSpreadsheetRange()) && empty($this->range)) {
            throw new Exception("There's no spreadsheet range set.");
        }

        $newCellValues = [$row];
        $valueRange = new \Google_Service_Sheets_ValueRange(["values" => $newCellValues]);
        $update_sheet = $this->service->spreadsheets_values->update(
            $this->getSpreadsheetId(),
            $this->range,
            $valueRange,
            ["valueInputOption" => "RAW"]
        );

        return $update_sheet->getUpdatedCells();
    }
    /**
     * Quick function to update a single cell in current worksheet
     * @param string $cell Column letter and row number of cell to update. Example: A1
     * @param string $value New value to set. 
     * @see https://developers.google.com/sheets/api/guides/concepts
     * @return int The number of updated rows.
     * 
     */
    public function updateSingleCell(?string $cell = null, ?string $value = null): int
    {

        if ($value == null) {
            throw new Exception("There's no value to set.");
        }

        if ($cell == null) {
            throw new Exception("There's no cell to set a range.");
        }

        $range = "{$this->worksheetName}!{$cell}:{$cell[0]}";
        $valueRange = new Google_Service_Sheets_ValueRange(["values" => [$value]]);
        $update_sheet = $service->spreadsheets_values->update($sheets::COURSEA_SHEET_ID, $range, $valueRange, ["valueInputOption" => "RAW"]);
        return $update_sheet->getUpdatedCells();
    }

    /**
     * Get Row Number Map by key or all from the actived sheet
     * 
     * @param array $config Required config to execute
     * @param string|int $config[line] Key set by addRow()
     * @param string|int $config[startColumnIndex] Key set by addRow()
     * @param string|int $config[endColumnIndex] Key set by addRow()
     * @param int        $config[r] R color code. [0,255]
     * @param int        $config[g] G color code. [0,255]
     * @param int        $config[b] B color code. [0,255]
     * 
     */
    public function colorLine(?array $config): void
    {

        if (empty($config)) {
            throw new Exception("No configuration set to execute colorLine", 1);
        }

        $config['a'] = $config['a'] ?? 1.0;
        $config['r'] = ((int)$config['r'] / 255);
        $config['g'] = ((int)$config['g'] / 255);
        $config['b'] = ((int)$config['b'] / 255);

        $sheetId = $this->service->spreadsheets->get($this->getSpreadsheetId(), ['ranges' => $this->worksheetName]);

        $myRange = [
            'sheetId' => $sheetId->sheets[0]->properties->sheetId,
            'startRowIndex' => $config['row'] - 1,
            'endRowIndex' => $config['row'],
            'startColumnIndex' => $config['startColumnIndex'],
            'endColumnIndex' => $config['endColumnIndex'],
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

        $response = $this->service->spreadsheets->batchUpdate(
            $this->getSpreadsheetId(),
            $batchUpdateRequest
        );

    }
}
