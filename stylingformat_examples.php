<?php
    //THIS EXAMPLE USED OOP STRUCTURE

    /** 
    * Contain automated Styling Format for sheets.
    */
    function stylingformat_examples(){
    $spreadsheetID = $this->spreadsheet->spreadsheetId; 
        
    $sheet_name = "Sheet1";

    $sheet_id = "0";

    $obj = $this->getDatatype();

    // Retrieve the header title.
    $retrieve = $this->service->spreadsheets_values->get($spreadsheetID, "'" . $sheet_name . "'!A3:3");
    $header = $retrieve["values"][0];

    $requests = [

        //Merge Company Info
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

        //Allign Company Info & Column Title 
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
        
        //Bold Text Company Info & Column Title
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

        //Column Title and Company Info Column Background Color
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

        //Column Sizing
        new \Google\Service\Sheets\Request([
            "autoResizeDimensions" => [
                "dimensions" => [
                    "dimension" => "COLUMNS",
                    "startIndex" => 2,
                ]
            ]
        ]),
        
    ];

    //Date & Decimal Format
    foreach ($header as $i => $h) {
        if (array_key_exists($h, $obj)) {
            array_push($requests, 
                new \Google\Service\Sheets\Request([
                    "repeatCell" => [
                        "range" => [
                            "sheetId" => $sheet_id,
                            "startColumnIndex" => $i,
                            "endColumnIndex" => $i + 1,
                            "startRowIndex" => 3,
                            "endRowIndex" => 500,
                        ],
                        "cell" => ["userEnteredFormat" => $obj[$h]],
                        "fields" => "userEnteredFormat.numberFormat",
                    ],
                ]),
            );
        };
    }
        $batchUpdate = new \Google\Service\Sheets\BatchUpdateSpreadsheetRequest(["requests" => $requests]);
        $this->service->spreadsheets->batchUpdate($spreadsheetID, $batchUpdate);

        return $this->service;
    }

function getDatatype()
{
        //             //format type & pattern 
    //             $obj = [
   //                 // "" => ["numberFormat" => ["type" => "DATE", "pattern" => "dd-mm-yyyy"]],
    //                 // "" => ["numberFormat" => ["type" => "NUMBER", "pattern" => "#,##0.00"]],
    //                 // "" => ["numberFormat" => ["type" => "NUMBER", "pattern" => "#,##0.0000"]],
    //                 // "" => ["numberFormat" => ["type" => "NUMBER", "pattern" => "#,##0.000000"]],
    //             ];
    //             return $obj;
}