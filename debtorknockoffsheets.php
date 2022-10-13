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
        ["",'', '','','','','Debtor Knockoff Details Report Between 12-10-2022 To 12-10-2022','','','','','','','',''],
        ["Doc. No",'Ref. No', 'Doc. Date','Terms','Business Partner Code','Bussiness Partner Name','Currency','Amount(MYR)','Knock Off Amount','Knock Off Amount(Local)','Balance Amount','Balance Amount (Local)','Knock Off Document','Account Name','Item Name'],
        ["Bussiness Partner Name: Riot Vanguard Inc",'', '','','','','','','','','','','','',''],
        ["RC 0001",'DKD001', '07-06-2022','','005','Riot Vanguard Inc','MYR','-106.00','-106','-106','0.00','0.00','SI 0001 | 106.00','Trade Debtors','Payment for SI 0001'],
        ["SI 0001",'', '07-06-2022','','005','Riot Vanguard Inc','MYR','106.00','106','106','0.00','0.00','RC 0001 | 106.00','Trade Debtors',''],
        ["",'', '','','','Riot Vanguard Inc Total:','','0.00','','','0.00','0.00','','',''],
        ["",'', '','','','','','','','','','','','',''],
        ["Bussiness Partner Name: Valve Inc",'', '','','','','','','','','','','','',''],
        ["RC 0004",'DKD004', '08-06-2022','','001','Steam Inc','MYR','-250.00','-250','-250','0.00','0.00','SI 0004 | 250.00','UOB','Payment for SI 0004'],
        ["SI 0004",'', '08-06-2022','','001','Steam Inc','MYR','250.00','250','250','0.00','0.00','RC 0004 | 250.00','UOB',''],
        ["RC 0008",'DKD008', '10-06-2022','','001','Steam Inc','MYR','-1500.00','-1500','-1500','0.00','0.00','SI 0008 | 1500.00','UOB','Payment for SI 0008'],
        ["SI 0008",'', '10-06-2022','','001','Steam Inc','MYR','1600.00','1600','1600','0.00','0.00','RC 0008 | 1600.00','UOB',''],
        ["JE 0005",'', '09-06-2022','','001','Steam Inc','MYR','100.00','0','0','100.00','100.00','Prepayment From Debtor','UOB','Gain'],
        ["JE 0006",'', '09-06-2022','','001','Steam Inc','MYR','50.00','0','0','50.00','50.00','Prepayment From Debtor','UOB','Gain'],
        ["",'', '','','','Steam Inc Total:','','250.00','','','150.00','150.00','','','']
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
    writeValues('1eTJJpYCMQq9EPmmcu4lf123KQfi_6QEhjfqScUBRfJ0', 'DKDRBSheets!A1');