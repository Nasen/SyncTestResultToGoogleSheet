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

// Bind handlers when the page loads.

$(function () {
    $('a.mdl-button').click(function () {
        setSpinnerActive(true);
    });
});

function setSpinnerActive(isActive) {
    if (isActive) {
        $('#spinner').addClass('is-active');
    } else {
        $('#spinner').removeClass('is-active');
    }
}

function showError(error) {
    console.log(error);
    var snackbar = $('#snackbar');
    snackbar.addClass('error');
    snackbar.get(0).MaterialSnackbar.showSnackbar(error);
}

function showMessage(message) {
    var snackbar = $('#snackbar');
    snackbar.removeClass('error');
    snackbar.get(0).MaterialSnackbar.showSnackbar({
        message: message
    });
}

// TODO: Add Google Sign-in.
function onSignIn(user) {
    var profile = user.getBasicProfile();
    $('#profile .name').text(profile.getName());
    $('#profile .email').text(profile.getEmail());
}

function onFail(e){
    console.log(e);
}


// TODO: Add spreadsheet control handlers.
$(function () {
    $('button[rel="create"]').click(function () {
        makeRequest('POST', '/spreadsheets', function (err, spreadsheet) {
            if (err) return showError(err);
            window.location.reload();
        });
    });
    $('button[rel="sync"]').click(function () {
        var spreadsheetId = $(this).data('spreadsheetid');
        var url = '/spreadsheets/' + spreadsheetId + '/sync';
        makeRequest('POST', url, function (err) {
            if (err) return showError(err);
            showMessage('Sync complete.');
            console.log("update sheet name with latest sync time");
            switch (spreadsheetId){
                case "1pm20jrnKlSlo4f5oz5qQYUclI1lam4ht9d5dp6niA2k":
                    var new_sheet_name='Feedback (' + new Date().toDateString() + ')';
                    break;
                case "1MZRJpHVixCy13tdZkkI0POcU3bLuO7Dj91DqTiKkN-A":
                    var new_sheet_name='UCC_Report (' + new Date().toDateString() + ')';
                    break;
                default :
                    var new_sheet_name='Feedback (' + new Date().toDateString() + ')';
            }

            var url='/spreadsheets/'+spreadsheetId+'/update/name/'+new_sheet_name;
            $.ajax(url);
            window.location.reload();
        });
    });
    $('button[rel="remove"]').click(function(){
        var spreadsheetId = $(this).data('spreadsheetid');
        var url = '/spreadsheets/' + spreadsheetId + '/remove';
        $.ajax(url);
        window.location.reload();
    });

     $('button[name="convert"]').click(function(){
        var todo_id = $(this).data('id');
        var url='/convert/'+todo_id;
        window.location=url;
        
    });
});

function makeRequest(method, url, callback) {

    //TODO：check access token expired or not

    var auth = gapi.auth2.getAuthInstance();
    //auth.disconnect();
    if (!auth.isSignedIn.get()) {
        //auth.signIn();
        return callback(new Error('Signin required.'));
    }
    //auth.signOut();
    var g_user = auth.currentUser.get();
    //g_user.grantOfflineAccess();
    var accessToken = g_user.getAuthResponse().access_token;
    //var expires_at = g_user.getAuthResponse().expires_in;
    console.log("access token: " + accessToken);
    //console.log("expire at: " + expires_at);
    setSpinnerActive(true);
    $.ajax(url, {
        method: method,
        headers: {
            'Authorization': 'Bearer ' + accessToken
        },
        success: function (response) {
            setSpinnerActive(false);
            return callback(null, response);
        },
        error: function (response) {
            setSpinnerActive(false);
            return callback(new Error(response.responseJSON.message));
        }
    });
}


