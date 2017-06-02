var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
var util = require('util');

// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'sheets.googleapis.mthor_report.json';

// Load client secrets from a local file.
var auth;
var oauth2Client;
var service;

var SheetsHelper = function () {

    fs.readFile('client_secret.json', function processClientSecrets(err, content) {
        if (err) {
            console.log('Error loading client secret file: ' + err);
            return;
        }
        // Authorize a client with the loaded credentials, then call the
        // Google Sheets API.
        var credentials = JSON.parse(content);
        var clientSecret = credentials.installed.client_secret;
        var clientId = credentials.installed.client_id;
        var redirectUrl = credentials.installed.redirect_uris[0];
        var auth = new googleAuth();
        var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

        // Check if we have previously stored a token.
        fs.readFile(TOKEN_PATH, function (err, token) {
            if (err) {
                //getNewToken(oauth2Client, callback);
                console.log("Token not exist: " + TOKEN_PATH);
            } else {
                oauth2Client.credentials = JSON.parse(token);
                service = google.sheets({version: 'v4', auth: oauth2Client});
            }
        });
    });


};

var COLUMNS = [
    {field: 'id', header: 'ID'},
    {field: 'description', header: 'Description'},
    {field: 'status', header: 'Status'},
    {field: 'contactName', header: 'Contact Name'},
    {field: 'occurTime', header: 'Occur Time'},
    {field: 'version', header: 'Version'},
    {field: 'osVersion', header: 'OS Version'},
    {field: 'mailboxID', header: 'mailboxID'},
    {field: 'device', header: 'Device'}
];

var REPORT_COLUMNS = [
    {field: 'id', header: 'ID'},
    {field: 'title', header: 'title'},
    {field: 'date_time', header: 'date_time'},
    {field: 'run', header: 'run'},
    {field: 'failed', header: 'failed'},
    {field: 'passed', header: 'passed'},
    {field: 'skipped', header: 'skipped'},
];

var SUMMARY_COLUMNS = [
    {field: 'date', header: 'Date'},
    {field: 'total', header: 'Total'},
    {field: 'pass', header: 'Pass'},
    {field: 'fail', header: 'Failed'},
    {field: 'skip', header: 'Skip'}
];

var DETAILS_COLUMNS = [
    {field: 'tc_name', header: 'TestCase'},
    {field: 'te_result', header: 'Result'}
]

SheetsHelper.prototype.createSpreadsheet = function (title, callback) {
    var self = this;
    var request = {
        resource: {
            properties: {
                title: title
            },
            sheets: [
                {
                    properties: {
                        title: 'Feedback',
                        gridProperties: {
                            columnCount: COLUMNS.length + 1,
                            frozenRowCount: 1
                        }
                    }
                }
                // TODO: Add more sheets.
                /*{
                 properties: {
                 title: 'Pivot',
                 gridProperties: {
                 hideGridlines: true
                 }
                 }
                 }*/
            ]
        }
    };
    self.service.spreadsheets.create(request, function (err, spreadsheet) {
        if (err) {
            return callback(err);
        }
        // TODO: Add header rows.
        var dataSheetId = spreadsheet.sheets[0].properties.sheetId;
        var requests = [
            buildHeaderRowRequest(dataSheetId),
        ];
// TODO: Add pivot table and chart.
        /*var pivotSheetId = spreadsheet.sheets[1].properties.sheetId;
         requests = requests.concat([
         buildPivotTableRequest(dataSheetId, pivotSheetId),
         buildFormatPivotTableRequest(pivotSheetId),
         buildAddChartRequest(pivotSheetId)
         ]);*/
        var request = {
            spreadsheetId: spreadsheet.spreadsheetId,
            resource: {
                requests: requests
            }
        };
        self.service.spreadsheets.batchUpdate(request, function (err, response) {
            if (err) {
                return callback(err);
            }
            return callback(null, spreadsheet);
        });
    });
};

SheetsHelper.prototype.sync = function (spreadsheetId, sheetId, feedbacks, callback) {
    var requests = [];
    // Resize the sheet.
    requests.push({
        updateSheetProperties: {
            properties: {
                sheetId: sheetId,
                gridProperties: {
                    rowCount: feedbacks.length + 1,
                    columnCount: COLUMNS.length
                }
            },
            fields: 'gridProperties(rowCount,columnCount)'
        }
    });
    // Set the cell values.
    switch (spreadsheetId) {
        //feedback
        case "1pm20jrnKlSlo4f5oz5qQYUclI1lam4ht9d5dp6niA2k":
            requests.push({
                updateCells: {
                    start: {
                        sheetId: sheetId,
                        rowIndex: 1,
                        columnIndex: 0
                    },
                    rows: buildRowsForFeedBacks(feedbacks),
                    fields: '*'
                }
            })
            break;
        //ucc report
        case  "1MZRJpHVixCy13tdZkkI0POcU3bLuO7Dj91DqTiKkN-A":
            requests.push({
                updateCells: {
                    start: {
                        sheetId: sheetId,
                        rowIndex: 1,
                        columnIndex: 0
                    },
                    rows: buildRowsForReports(feedbacks),
                    fields: '*'
                }
            });
            break;
    }

    // Send the batchUpdate request.
    var request = {
        spreadsheetId: spreadsheetId,
        resource: {
            requests: requests
        }
    };
    service.spreadsheets.batchUpdate(request, function (err) {
        if (err) {
            return callback(err);
        }
        return callback();
    });
};


SheetsHelper.prototype.generateTestReport = function (spreadsheetId, summary, details, callback) {

    var requestForGet = {
        spreadsheetId: spreadsheetId,
        includeGridData: true
    }
    service.spreadsheets.get(requestForGet, function (err, response) {
        if (err) {
            console.log(err);
            return;
        }

        var res_json = JSON.stringify(response);
        var res = JSON.parse(res_json);
        var summary_sheetID = res.sheets[0].properties.sheetId;
        var summary_sheet_rows = res.sheets[0].data[0].rowData.length;

        var detail_sheetID = res.sheets[1].properties.sheetId;
        var detail_sheet_rows = res.sheets[1].data[0].rowData.length;
        var detail_sheet_columns = res.sheets[1].data[0].rowData[0].values.length;

        var requestForUpdateSummary = [];
        requestForUpdateSummary.push({
            updateCells: {
                start: {
                    sheetId: summary_sheetID,
                    rowIndex: summary_sheet_rows,
                    columnIndex: 0,
                },
                rows: buildRowsForTestSummary(summary),
                fields: '*'
            }
        });

        var requestForUpdateDetails = [];
        requestForUpdateDetails.push({
            updateCells: {
                start: {
                    sheetId: detail_sheetID,
                    rowIndex: 1,
                    columnIndex: 0,
                },
                rows: buildRowsForTestDetails(details),
                fields: '*'
            }
        });

        // Send the batchUpdate request.
        var updateSummaryRequest = {
            spreadsheetId: spreadsheetId,
            resource: {
                requests: requestForUpdateSummary
            }
        };
        var updateDetailsRequest = {
            spreadsheetId: spreadsheetId,
            resource: {
                requests: requestForUpdateDetails
            }
        };

        var clearDetailsSheetRequest = {
            spreadsheetId: spreadsheetId,
            resource: {
                requests: {
                    "deleteRange": {
                        "range": {
                            "sheetId": detail_sheetID,
                            "startRowIndex": 1,
                            "endRowIndex": detail_sheet_rows,
                        },
                        "shiftDimension": "ROWS"
                    }
                }
            }
        };

        service.spreadsheets.batchUpdate(updateSummaryRequest, function (err, res) {
            if (err) {
                callback(err);
            }
            service.spreadsheets.batchUpdate(clearDetailsSheetRequest, function (err, res) {
                if (err) {
                    callback(err)
                }
                service.spreadsheets.batchUpdate(updateDetailsRequest, function (err, res) {
                    if (err) {
                        callback(err)
                    }
                    return callback();
                })
            })
        })
        //JSON.stringify(response,null,2);
    })
}

SheetsHelper.prototype.syncReport = function (spreadsheetId, sheetId, reports, callback) {
    var requests = [];
    // Resize the sheet.
    requests.push({
        updateSheetProperties: {
            properties: {
                sheetId: sheetId,
                gridProperties: {
                    rowCount: reports.length + 1,
                    columnCount: REPORT_COLUMNS.length
                }
            },
            fields: 'gridProperties(rowCount,columnCount)'
        }
    });
    // Set the cell values.
    requests.push({
        updateCells: {
            start: {
                sheetId: sheetId,
                rowIndex: 1,
                columnIndex: 0
            },
            rows: buildRowsForReports(reports),
            fields: '*'
        }
    });
    // Send the batchUpdate request.
    var request = {
        spreadsheetId: spreadsheetId,
        resource: {
            requests: requests
        }
    };
    this.service.spreadsheets.batchUpdate(request, function (err) {
        if (err) {
            return callback(err);
        }
        return callback();
    });
};


function buildHeaderRowRequest(sheetId) {
    var cells = COLUMNS.map(function (column) {
        return {
            userEnteredValue: {
                stringValue: column.header
            },
            userEnteredFormat: {
                textFormat: {
                    bold: true
                }
            }
        }
    });
    return {
        updateCells: {
            start: {
                sheetId: sheetId,
                rowIndex: 0,
                columnIndex: 0
            },
            rows: [
                {
                    values: cells
                }
            ],
            fields: 'userEnteredValue,userEnteredFormat.textFormat.bold'
        }
    };
}

function buildRowsForTestSummary(summary) {

    return summary.map(function (summary) {
        var cells = SUMMARY_COLUMNS.map(function (column) {
            if (column.field == "date") {
                return {
                    userEnteredValue: {
                        stringValue: summary[column.field].toString()
                    }
                }
            } else {
                return {
                    userEnteredValue: {
                        numberValue: summary[column.field]
                    }
                };
            }

        });

        console.log(JSON.stringify(cells));
        return {
            values: cells
        };
    });
}
function buildRowsForTestDetails(details) {

    return details.map(function (detail) {
        var cells = DETAILS_COLUMNS.map(function (column) {
            return {
                userEnteredValue: {
                    stringValue: detail[column.field]
                }
            }
        });

        console.log(JSON.stringify(cells));
        return {
            values: cells
        };
    });
}

function buildRowsForFeedBacks(feedbacks) {
    return feedbacks.map(function (feedback) {
        var cells = COLUMNS.map(function (column) {
            switch (column.field) {
                case 'status':
                    return {
                        userEnteredValue: {
                            stringValue: feedback.status
                        },
                        dataValidation: {
                            condition: {
                                type: 'ONE_OF_LIST',
                                values: [
                                    {userEnteredValue: 'new'},
                                    {userEnteredValue: 'pending'},
                                    {userEnteredValue: 'closed'}
                                ]
                            },
                            strict: true,
                            showCustomUi: true
                        }
                    };
                    break;
                default:
                    return {
                        userEnteredValue: {
                            stringValue: feedback[column.field].toString()
                        }
                    };
            }
        });

        return {
            values: cells
        };
    });
}

function buildRowsForReports(reports) {
    return reports.map(function (report) {
        var cells = REPORT_COLUMNS.map(function (column) {
            return {
                userEnteredValue: {
                    stringValue: report[column.field].toString()
                }
            };
        });


        return {
            values: cells
        };
    });
}

function buildPivotTableRequest(sourceSheetId, targetSheetId) {
    return {
        updateCells: {
            start: {sheetId: targetSheetId, rowIndex: 0, columnIndex: 0},
            rows: [
                {
                    values: [
                        {
                            pivotTable: {
                                source: {
                                    sheetId: sourceSheetId,
                                    startRowIndex: 0,
                                    startColumnIndex: 0,
                                    endColumnIndex: COLUMNS.length
                                },
                                rows: [
                                    {
                                        sourceColumnOffset: getColumnForField('productCode').index,
                                        showTotals: false,
                                        sortOrder: 'ASCENDING'
                                    }
                                ],
                                values: [
                                    {
                                        summarizeFunction: 'SUM',
                                        sourceColumnOffset: getColumnForField('unitsOrdered').index
                                    },
                                    {
                                        summarizeFunction: 'SUM',
                                        name: 'Revenue',
                                        formula: util.format("='%s' * '%s'",
                                            getColumnForField('unitsOrdered').header,
                                            getColumnForField('unitPrice').header)
                                    }
                                ]
                            }
                        }
                    ]
                }
            ],
            fields: '*'
        }
    };
}

function buildFormatPivotTableRequest(sheetId) {
    return {
        repeatCell: {
            range: {sheetId: sheetId, startRowIndex: 1, startColumnIndex: 2},
            cell: {
                userEnteredFormat: {
                    numberFormat: {type: 'CURRENCY', pattern: '"$"#,##0.00'}
                }
            },
            fields: 'userEnteredFormat.numberFormat'
        }
    };
}

function buildAddChartRequest(sheetId) {
    return {
        addChart: {
            chart: {
                spec: {
                    title: 'Revenue per Product',
                    basicChart: {
                        chartType: 'BAR',
                        legendPosition: 'RIGHT_LEGEND',
                        domains: [
                            // Show a bar for each product code in the pivot table.
                            {
                                domain: {
                                    sourceRange: {
                                        sources: [{
                                            sheetId: sheetId,
                                            startRowIndex: 0,
                                            startColumnIndex: 0,
                                            endColumnIndex: 1
                                        }]
                                    }
                                }
                            }
                        ],
                        series: [
                            // Set that bar's length based on the total revenue.
                            {
                                series: {
                                    sourceRange: {
                                        sources: [{
                                            sheetId: sheetId,
                                            startRowIndex: 0,
                                            startColumnIndex: 2,
                                            endColumnIndex: 3
                                        }]
                                    }
                                }
                            }
                        ]
                    }
                },
                position: {
                    overlayPosition: {
                        anchorCell: {sheetId: sheetId, rowIndex: 0, columnIndex: 3},
                        widthPixels: 600,
                        heightPixels: 400
                    }
                }
            }
        }
    };
}

function getColumnForField(field) {
    return COLUMNS.reduce(function (result, column, i) {
        if (column.field == field) {
            column.index = i;
            return column;
        }
        return result;
    });
}

module.exports = SheetsHelper;