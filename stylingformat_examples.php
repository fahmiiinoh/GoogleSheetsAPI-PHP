<?php
    //THIS EXAMPLE USED OOP STRUCTURE

    /** 
    * Contain automated Styling Format for sheets.
    */
    function sheetstylingFormat()
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

            new \Google\Service\Sheets\Request([
                "repeatCell" => [
                    "range" => [
                        "sheetId" => $sheet_id,
                        "startColumnIndex" => 1,
                        "endColumnIndex" => 3,
                    ],
                    "cell" => [
                        "userEnteredFormat" => [
                            "numberFormat" => [
                            "type" => "DATE",
                            "pattern" => "dd-mm-yyyy"
                            ]
                        ]
                    ],
                    "fields" => "userEnteredFormat.numberFormat",
                ],
            ]),
        ];
        $batchUpdate = new \Google\Service\Sheets\BatchUpdateSpreadsheetRequest(["requests" => $requests]);
        $this->service->spreadsheets->batchUpdate($spreadsheetID, $batchUpdate);

        return $this->service;
    }