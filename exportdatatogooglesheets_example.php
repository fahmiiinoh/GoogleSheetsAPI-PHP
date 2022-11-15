<?php

namespace Export;

session_start();

$client = require('auth.php');

use Google\Service\Exception;
use Google\Service\Sheets;
use Google\Service\Sheets\Spreadsheet;
use Google\Service\Sheets\ValueRange;

class ExportGoogleSheet extends Files
{


    private $spreadsheet;

    private $service;

    public function __construct($type = 'xls', $exportDataParam = '', $settingArr = [], $directdownload = 0)
    {
        /**
         * google_auth_token get from cookie from export.php
         * Verify the google auth token, $settingArr['google_auth_token'] 
         * 1. Able to give error when google_auth_token invalid
         */
        $this->accessToken = $settingArr['google_auth_token'];
        unset($settingArr['google_auth_token']);

        parent::__construct($type, $exportDataParam, $settingArr, $directdownload);
    }


    protected function draw()
    {
                //.1
                $this->prepareSheetHeadData();

                // 2.
                // Usually if got sub group root row will empty.
                foreach ($this->rootDataGroup->getDataRows() as $row)
                {
                    $this->prepareSheetRowLineData($row);
                }
                // Draw child (if any)
                foreach ($this->rootDataGroup->getChildren() as $child)
                {
                    $this->prepareSheetGroupData($child, 0);
                }
        
                $this->prepareSheetSummaryData();
                // Footer
                $this->sheetData[] = [$this->generatedDateString];
        
        
                // Remove all grouped data since ady use.
                $this->rootDataGroup->__destruct();
                unset($this->rootDataGroup);
                gc_collect_cycles();
        
                // 3.
                $this->createSheet();

                // 4.
                $this->drawSheet();

                // 5.
                $this->initSheet();
                
                // 6.
                $this->sheetstylingFormat();
    }

    /** 
     * Create a new sheets.
    *Load user credentials.
    *Contain title for sheets
    *Contain Spreadsheets ID
    */
    protected function createSheet()
    {       
        // try
        // {
            $client = getClientAuth();
            
            $this->service = new Sheets($client);
            $this->sheetData[] = [$this->title];
            $this->spreadsheet = new Spreadsheet([
                'properties' => [
                     'title' => "$this->title"
                    ]
                ]);
                $this->spreadsheet = $this->service->spreadsheets->create($this->spreadsheet, [
                    'fields' => 'spreadsheetId'
                ]);
                return $this->spreadsheet->spreadsheetId;

            // }catch(Exception $e)
            // {
            //     $e->getMessage();
            //     // header("Location: http://localhost/simbiz/auth-google-sheets.php");
            // }
    }

    /** 
    *Put $this->sheetData in ROWS Dimension into actives sheets.
    */
    protected function drawSheet()
    {
        /* Load pre-authorized user credentials from the environment.*/
        $client = getClientAuth();

        $spreadsheet_ID = $this->spreadsheet->spreadsheetId; 
        
        $range = "Sheet1";

        $values = $this->sheetData;

        $this->service = new Sheets($client);


        $body = new ValueRange([
            'majorDimension' => 'ROWS',
            'values' => array_map(function($row) {
              return array_map(function($col) {
                return (is_null($col)) ? "" : $col;
              }, $row);
            }, $values)
        ]);

        $params = [
            'valueInputOption' => 'RAW',
            'insertDataOption' => 'INSERT_ROWS'
        
        ];

        //executing the request
        $data = $this->service->spreadsheets_values->append($spreadsheet_ID, $range,
        $body, $params);
    
        return $data;

    }

    /** 
    *Initialize or Redirect to active sheets
    */
    protected function initSheet()
    {
            header("Location:https://docs.google.com/spreadsheets/d/".$this->spreadsheet->spreadsheetId);
    }

    /** 
    * Contain automated Styling Format for sheets.
    */
    protected function sheetstylingFormat()
    {   
        $spreadsheetID = $this->spreadsheet->spreadsheetId; 
        
        $sheet_id = "0";

        $requests = [

            new \Google\Service\Sheets\Request([
                "mergeCells" => [
                    "mergeType" => "MERGE_ALL",
                    
                    "range" => [
                        "sheetId" => $sheet_id,
                        "startRowIndex" => 1,
                        "endRowIndex" => 2,
                        "startColumnIndex" => 0,
                        "endColumnIndex" => (sizeof($this->dataCols->all())),
                    ],
                ],
                ]),

            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheet_id,
                        "startRowIndex" => 0,
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
                        "sheetId" => $sheet_id,
                        "startColumnIndex" => 0,
                        "endColumnIndex" => 1,
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
                        "sheetId" => $sheet_id,
                        "startRowIndex" => 0,
                        "endRowIndex" => 3,
                    ],
                    "cell" => [
                        "userEnteredFormat" => [
                            "textFormat" => [
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
                        "sheetId" => $sheet_id,
                        "startRowIndex" => 1,
                        "endRowIndex" => 3,
                        "startColumnIndex"=> 0,
                        "endColumnIndex"=> (sizeof($this->dataCols->all())),
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
                "autoResizeDimensions" => [
                    "dimensions" => [
                        "dimension" => "COLUMNS",
                        "startIndex" => 0,
                    ]
                ]
            ]),

            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheet_id,
                        "startRowIndex" => 1,
                        "endRowIndex" => 2,
                    ],
                    "cell" => [
                        "userEnteredFormat" => ["horizontalAlignment" => "CENTER"],
                    ],
                    "fields" => "userEnteredFormat.horizontalAlignment",
                ],
            ]),
        ];
        $batchUpdate = new \Google\Service\Sheets\BatchUpdateSpreadsheetRequest(["requests" => $requests]);
        $this->service->spreadsheets->batchUpdate($spreadsheetID, $batchUpdate);

        return $this->service;
    }
}
