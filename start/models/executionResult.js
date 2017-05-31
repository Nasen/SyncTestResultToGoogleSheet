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


"use strict";

module.exports = function (sequelize, DataTypes) {
    var Result = sequelize.define('test_result', {
        te_failure_info: {type: DataTypes.STRING, allowNull: true},
        te_start_time: {type: DataTypes.DATE, allowNull: false},
        te_end_time: {type: DataTypes.DATE, allowNull: false},
        te_result: {type: DataTypes.STRING, allowNull: false},
        tc_name: {type: DataTypes.STRING, allowNull: false},
        te_platform: {type: DataTypes.STRING, allowNull: false},
    });

    return Result;
};
