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

        $values = [["SIMIT QC SDN BHD HQ",'', '','','','','','','','','','','','',''],
        ["",'', '','','','','Sales History Between 11-10-2022 To 12-10-2022','','','','','','','',''],
        ["Branch Code",'Doc. Date', 'Doc Type','Doc. No','Ref. No','Terms Days','Bussiness Partner Name','Item Code','Item Name','Alternate Item Name','Quantity','U.O.M Code','Unit Price','Amount(MYR)','Sales Agent Name'],
        ["Doc. No : ARDN 0006",'', '','','','','','','','','','','','',''],
        ["Bussiness Partner Name : Twitch",'', '','','','','','','','','','','','',''],
        ["HQ",'11-10-2022', 'ARDN','ARDN 0006','','0','Twitch','GS001','Gum Strawberry','','1.00','Unit','125.00','125.00',''],
        ["",'', '','','','','Twitch Total:','','','','1.00','','0.00','125.00',''],
        ["",'', '','','','','','','','','','','','',''],
        ["Doc. No : ARDN 0007",'', '','','','','','','','','','','','',''],
        ["Bussiness Partner Name : Twitter",'', '','','','','','','','','','','','',''],
        ["HQ",'12-10-2022', 'ARDN','ARDN 0007','','0','Twitter','GS001','Gum Strawberry','','5.00','Unit','125.00','625.00',''],
        ["",'', '','','','','Twitter Total:','','','','5.00','','0.00','625.00',''],
        ["",'', '','','','','','','','','','','','',''],
        ["Doc. No : ARDN 0008",'', '','','','','','','','','','','','',''],
        ["Bussiness Partner Name : Dingtalk",'', '','','','','','','','','','','','',''],
        ["HQ",'12-10-2022', 'ARDN','ARDN 0008','','0','Dingtalk','GS001','Gum Strawberry','','10.00','Unit','125.00','1250.00',''],
        ["",'', '','','','','Dingtalk Total:','','','','10.00','','0.00','1250.00',''],
        ["",'', '','','','','','','','','','','','',''],
        ["Generate Date: 13-10-2022 02:23:10",'', '','','','','','','','','','','','','']
        ];


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
    writeValues('1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0', 'SHSheets!A1');
    function batchUpdate($spreadsheetId, $sheetId, $range)
    {   
        $client = getClient(); // This is from your script.

        // $spreadsheet_id = "1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0"; // please set Spreadsheet ID.
        //  $sheet_id = "635779689"; // please set Sheet ID.
        $column_width1 = 250; // Please set the column width you want.
        $column_width2 = 50;
        $service = new Google_Service_Sheets($client);
        try{
        $requests = [
            // new \Google\Service\Sheets\Request([
            //     "repeatCell" => [
            //         "range" => [
            //             "sheetId" => $sheetId,
            //             "startRowIndex" => 1,
            //             "endRowIndex" => 3,
            //         ],
            //         "cell" => [
            //             "userEnteredFormat" => ["horizontalAlignment" => "CENTER"],
            //         ],
            //         "fields" => "userEnteredFormat.horizontalAlignment",
            //     ],
            // ]),

            // new \Google\Service\Sheets\Request([
            //     "repeatCell" => [
            //         "range" => [
            //             "sheetId" => $sheetId,
            //             "startColumnIndex" => 0,
            //             "endColumnIndex" => 1,
            //         ],
            //         "cell" => [
            //             "userEnteredFormat" => ["horizontalAlignment" => "LEFT"],
            //         ],
            //         "fields" => "userEnteredFormat.horizontalAlignment",
            //     ],
            // ]),
            
            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheetId,
                        "startRowIndex" => 0,
                        "endRowIndex" => 3,
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
                        "endRowIndex" => 3,
                        "startColumnIndex"=> 0,
                        "endColumnIndex"=> 15,
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
                        "startIndex" => 0,
                        "endIndex" =>1,
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
                        "startIndex" => 6,
                        "endIndex" =>7,
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
                        "startIndex" => 9,
                        "endIndex" =>10,
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
                        "startIndex" => 14,
                        "endIndex" =>15,
                        "dimension" => "COLUMNS",
                    ],
                    "properties" => ["pixelSize" => $column_width1],
                    "fields" => "pixelSize",
                ],
            ]),
            // new \Google\Service\Sheets\Request([
            //     "updateDimensionProperties" => [
            //         "range" => [
            //             "sheetId" => $sheetId,
            //             "startIndex" => 0,
            //             "endIndex" =>1,
            //             "dimension" => "COLUMNS",
            //         ],
            //         "properties" => ["pixelSize" => $column_width2],
            //         "fields" => "pixelSize",
            //     ],
            // ]),
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
batchUpdate('1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0', '597252610', 'SHSheets!A1');