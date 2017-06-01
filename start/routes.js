/*
 Copyright 2016 Google, Inc.

 Licensed to the Apache Software Foundation (ASF) under one or more contributor
 license agreements. See the NOTICE file distributed with this work for
 additional information regarding copyright ownership. The ASF licenses this
 file to you under the Apache License, Version 2.0 (the "License"); you may not
 use this file except in compliance with the License. You may obtain a copy of
 the License at

 http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
 License for the specific language governing permissions and limitations under
 the License.
 */

'use strict';

var express = require('express');
var router = express.Router();
var models = require('./models');
var Sequelize = require('sequelize');
var google = require('googleapis');
// Add route for creating spreadsheet.
var SheetsHelper = require('./sheets');

//update mThor iOS test report
router.post('/updateReport/ios', function (req, res, next) {
    var results;
    var date = new Date().toISOString().split("T")[0];

    var spreadsheetID = "1NGYcuLf88AYvf7HnMklLDq_jx4zYcjjgp4x1w73wYew";
    var totalNum = 0, passedNum = 0, failedNum = 0, skipNum = 0;
    var helper = new SheetsHelper();

    /*    //get all passed case number
     Sequelize.Promise.all(models.test_result.findAll(
     {
     attributes: ['tc_name', 'te_result'],
     where: {
     te_platform: 'IOS',
     te_result: 'PASS',
     te_start_time: {
     $gt: '2017-05-18 00:00:01'
     }
     }
     })).then(function (results) {
     passedNum = results.length;
     //get all failed case number
     Sequelize.Promise.all(models.test_result.findAll(
     {
     attributes: ['tc_name', 'te_result'],
     where: {
     te_platform: 'IOS',
     te_result: 'FAIL',
     te_start_time: {
     $gt: '2017-05-18 00:00:01'
     }
     }
     })).then(function (results) {
     failedNum = results.length;
     //get all skip case number
     Sequelize.Promise.all(models.test_result.findAll(
     {
     attributes: ['tc_name', 'te_result'],
     where: {
     te_platform: 'IOS',
     te_result: 'SKIP',
     te_start_time: {
     $gt: '2017-05-18 00:00:01'
     }
     }
     })).then(function (results) {
     skipNum = results.length;
     });
     });
     });*/
    var details = [];
    Sequelize.Promise.all(models.test_result.findAll(
        {
            attributes: ['tc_name', 'te_result'],
            where: {
                te_platform: 'IOS',
                te_start_time: {
                    $gt: '2017-05-18 00:00:01'
                }
            }
        })).then(function (results) {

        totalNum = results.length;
        var x;
        for (x in results) {
            var tc_name = results[x].get("tc_name");
            var te_result = results[x].get("te_result");
            details.push({tc_name, te_result});
            switch (results[x].get("te_result")) {
                case "PASS":
                    passedNum += 1;
                    break;

                case "FAIL":
                    failedNum += 1;
                    break;

                case "SKIP":
                    skipNum += 1;
                    break;

            }

        }
        var summary = [
            {date: date, total: totalNum, pass: passedNum, fail: failedNum, skip: skipNum}
        ]

        console.log(JSON.stringify(summary));
        helper.updateIOSReport(spreadsheetID, summary, details, function (err) {
            if (err) {
                return next(err);
            }
            return res.json(summary.concat(details));

        })
    });
});

module.exports = router;
