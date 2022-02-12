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
     * @var string absolute path of credential file location
     */
    private $credentialFilePath;

    /**
     * @var string absolute path of token file location for auth, if not exist you need to follow the CLI steps for first time auth: <https://developers.google.com/sheets/api/quickstart/php#step_2_set_up_the_sample>
     */
    private $tokenPath;

    public function __construct(?string $credentialFilePath = null, ?string $tokenPath = null)
    { 
        if (!file_exists($credentialFilePath)) {
            throw new Exception("No credential file in: {$credentialFilePath}");
        }
        $this->tokenPath = $tokenPath;
        $this->credentialFilePath = $credentialFilePath;
        $this->client = $this->getClient();
        $this->service = new \Google_Service_Sheets($this->client);
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

    public function firstAuth(?string $tokenPath = null)
    {

        if( $tokenPath == null ){
            throw new Exception("token.json destination filepath not set");
        }

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

        if(file_exists($this->tokenPath)){
            return true;
        }else{
            return false;
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
        }

        // If there is no previous token or it's expired.
        if ($client->isAccessTokenExpired()) {
            // Refresh the token if possible, else fetch a new one.
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            } else {
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
            }
            // Save the token to a file.
            if (!file_exists(dirname($this->tokenPath))) {
                mkdir(dirname($this->tokenPath), 0700, true);
            }
            file_put_contents($this->tokenPath, json_encode($client->getAccessToken()));
        }
        return $client;
    }

    /**
     * Get Row Number Map by key or all from the actived sheet
     * 
     * @param string|int $key Key set by addRow()
     * @return int|array Row number | Key-Coordinate array
     * 
     */
    public function colorLine(?array $config)
    {

        if (empty($config)) {
            throw new Exception("No configuration set to execute colorLine", 1);
        }

        $config['a'] = $config['a'] ?? 1.0;
        $config['r'] = ((int)$config['r'] / 255);
        $config['g'] = ((int)$config['g'] / 255);
        $config['b'] = ((int)$config['b'] / 255);

        $sheetId = $this->service->spreadsheets->get($this->getSpreadsheetId(), ['ranges' => $config['worksheetName']]);

        $myRange = [
            'sheetId' => $sheetId->sheets[0]->properties->sheetId,
            'startRowIndex' => $config['line'] - 1,
            'endRowIndex' => $config['line'],
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

        return $response;
    }
}
