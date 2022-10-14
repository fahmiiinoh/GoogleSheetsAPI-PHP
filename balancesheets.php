<?php
/**
 * Copyright 2018 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

require __DIR__ . '/vendor/autoload.php';

if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}

use Google\Client;
use Google\Service\Sheets\ValueRange;
use Google\Service\Sheets\BatchUpdateSpreadsheetRequest;

/**
 * Returns an authorized API client.
 * @return Client the authorized client object
 */
function getClient()
{
    $client = new Google\Client();
    $client->setApplicationName('Google Sheets API PHP Tester');
    $client->setScopes('https://www.googleapis.com/auth/spreadsheets');
    $client->setAuthConfig('credentials.json');
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');

    // Load previously authorized token from a file, if it exists.
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    $tokenPath = 'token.json';
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
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
        if (!file_exists(dirname($tokenPath))) {
            mkdir(dirname($tokenPath), 0700, true);
        }
        file_put_contents($tokenPath, json_encode($client->getAccessToken()));
    }
    return $client;
}


function writeValues($spreadsheetId, $range)
    {
        /* Load pre-authorized user credentials from the environment.
           TODO(developer) - See https://developers.google.com/identity for
            guides on implementing OAuth2 for your application. */
        $client = getClient();
        $service = new Google_Service_Sheets($client);
        try{

        $values = [["",'Balance Sheet', ''],
        ['', 'SIMIT QC SDN BHD',''],
        ["Code", 'Accounts','Amount(MYR)'],
        ["10-0000", '[Assets]',''],
        ["11-0000", '[Current Assets]',''],
        ["11-1000", 'Trade Debtors','10,955.00'],
        ["11-3000", '[Cash and Banks]',''],
        ["11-3001", 'Petty Cash','8,444.00'],
        ["11-3011", 'Bank Account A','18,425.00'],
        ["11-3012", 'UOB','1,810.00'],
        ["12-190", 'RevPay','106.00'],
        ["", 'Total[Cash and Banks]','28,785.00'],
        ["", 'Total[Current Assets]','39,740.00'],
        ["12-0000", '[Non-Current Assets]',''],
        ["12-1000", '[Vehicle]',''],
        ["", 'Total[Vehicle]','0.00'],
        ["12-2000", '[Equipment]',''],
        ["", 'Total[Equipment]','0.00'],
        ["12-3000", '[Land and Building]',''],
        ["", 'Total[Land and Building]','0.00'],
        ["", 'Total[Non-Current Assets]','0.00'],
        ["", 'Total[Assets]','39,740.00'],
        ["", '',''],
        ["", 'TOTAL ASSET','39,740.00'],
        ["", '',''],
        ["20-0000", '[Liabilities]',''],
        ["21-0000", '[Current Liabilities]',''],
        ["2-SST", 'Sales & Service Tax','200.00'],
        ["21-1000", 'Trade Creditors','71,907.00'],
        ["21-2000", 'Others Creditor','5,440.00'],
        ["21-3000", 'Output Tax','6.00'],
        ["21-4000", 'Prepayment From Debtor','825.00'],
        ["", 'Total[Current Liabilities]','78,378.00'],
        ["22-0000", '[Non-Current Liabilities]',''],
        ["22-1000", '[Long Term Loan]',''],
        ["", 'Total[Long Term Loan]','0.00'],
        ["", 'Total[Non-Current Liabilities]','0.00'],
        ["", 'Total[Liabilities]','78,378.00'],
        ["30-0000", '[Equity]',''],
        ["31-0000", '[Capitals]',''],
        ["31-0001", 'Share Holder 1','1,000.00'],
        ["", 'Total[Capitals]','1,000.00'],
        ["31-1000", 'Dividend','20.00'],
        ["33-0000", 'Current Year Earning','(48,811.00)'],
        ["", 'Total[Equity]','(47,791.00)'],
        ["", '',''],
        ["", 'TOTAL LIABILITY & SHAREHOLDER EQUITY','30,587.00']];

        $body = new Google_Service_Sheets_ValueRange([
            'values' => $values
        ]);
        $params = [
            'valueInputOption' => 'USER_ENTERED'
        ];
        //executing the request
        $result = $service->spreadsheets_values->update($spreadsheetId, $range,
        $body, $params);
        printf("%d data inserted.\n", $result->getUpdatedCells());
        return $result->getUpdatedCells();
    }
    catch(Exception $e) {
            // TODO(developer) - handle error appropriately
            echo 'Message: ' .$e->getMessage();
          }
    }
    writeValues('1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0', 'BSSheets!B2');

    function batchUpdate($spreadsheetId, $sheetId, $range)
    {   
        $client = getClient(); // This is from your script.

        // $spreadsheet_id = "1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0"; // please set Spreadsheet ID.
        //  $sheet_id = "635779689"; // please set Sheet ID.
        $column_width1 = 300; // Please set the column width you want.
        $column_width2 = 50;
        $service = new Google_Service_Sheets($client);
        try{
        $requests = [
            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startRowIndex" => 1,
                        "endRowIndex" => 3,
                    ],
                    "cell" => [
                        "userEnteredFormat" => ["horizontalAlignment" => "CENTER"],
                    ],
                    "fields" => "userEnteredFormat.horizontalAlignment",
                ],
            ]),

            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startColumnIndex" => 1,
                        "endColumnIndex" => 2,
                    ],
                    "cell" => [
                        "userEnteredFormat" => ["horizontalAlignment" => "LEFT"],
                    ],
                    "fields" => "userEnteredFormat.horizontalAlignment",
                ],
            ]),
            
            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startRowIndex" => 1,
                        "endRowIndex" => 4,
                    ],
                    "cell" => [
                        "userEnteredFormat" => [
                            "textFormat" => [
                                                // "fontSize" => 30,
                                                "bold" => true,
                                ],
                        ]
                    ],
                    "fields" => "userEnteredFormat.textFormat.bold",
                ],
            ]),

            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startRowIndex" => 1,
                        "endRowIndex" => 4,
                        "startColumnIndex"=> 1,
                        "endColumnIndex"=> 4,
                    ],
                    "cell" => [
                        "userEnteredFormat" => [
                            "backgroundColor"=> [
                                "red"=> 0.7,
                                "green"=> 0.7,
                                "blue"=> 0.7,
                                ]
                            ]
                        ],
                    "fields" => "userEnteredFormat.backgroundColor",
                ],
            ]),

            new \Google\Service\Sheets\Request([
                "updateDimensionProperties" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startIndex" => 2,
                        "endIndex" =>3,
                        "dimension" => "COLUMNS",
                    ],
                    "properties" => ["pixelSize" => $column_width1],
                    "fields" => "pixelSize",
                ],
            ]),
            new \Google\Service\Sheets\Request([
                "updateDimensionProperties" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startIndex" => 0,
                        "endIndex" =>1,
                        "dimension" => "COLUMNS",
                    ],
                    "properties" => ["pixelSize" => $column_width2],
                    "fields" => "pixelSize",
                ],
            ]),
        ];
        $batchUpdate = new \Google\Service\Sheets\BatchUpdateSpreadsheetRequest(["requests" => $requests]);
        $service->spreadsheets->batchUpdate($spreadsheetId, $batchUpdate);

        return $service;
    }
    catch(Exception $e) {
        // TODO(developer) - handle error appropriately
        echo 'Message: ' .$e->getMessage();
    }
}
batchUpdate('1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0', '722765238', 'BSSheets!A1');