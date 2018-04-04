/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function() {

    // create
    angular
        .module('outlook-2-sp', [])
        .controller('HomeController', [HomeController])
        .config(['$logProvider', function($logProvider) {
            // set debug logging to on
            if ($logProvider.debugEnabled) {
                $logProvider.debugEnabled(true);
            }
        }]);

    /**
     * Home Controller
     */
    function HomeController() {
        var _this = this;
        var timesGetOneDriveFilesHasRun = 0;
        var triedWithoutForceConsent = false;
        _this.title = 'Home';
        _this.baseRestHost = Office.context.mailbox.restUrl;


        // Displays the data, assumed to be an array.
        function showResult(data) {
            _this.data = data;
        }

        function logError(error) {
            console.log(error);
        }

        function handleClientSideErrors(result) {

            switch (result.error.code) {

                // Handle the case where user is not logged in, or the user cancelled, without responding, a
                // prompt to provide a 2nd authentication factor. 
                case 13001:
                    getDataWithToken({ forceAddAccount: true });
                    break;

                    // Handle the case where the user's sign-in or consent was aborted.
                case 13002:
                    if (timesGetOneDriveFilesHasRun < 2) {
                        showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
                    } else {
                        logError(result);
                    }
                    break;
                    // Handle the case where the user is logged in with an account that is neither work or school, 
                    // nor Micrososoft Account.
                case 13003:
                    showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
                    break;

                case 13005:
                    getDataWithToken({ forceConsent: true });
                    break;

                    // Handle an unspecified error from the Office host.
                case 13006:
                    showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
                    break;

                    // Handle the case where the Office host cannot get an access token to the add-ins 
                    // web service/application.
                case 13007:
                    showResult(['That operation cannot be done at this time. Please try again later.']);
                    break;

                    // Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
                    // before a previous call of it completed.
                case 13008:
                    showResult(['Please try that operation again after the current operation has finished.']);
                    break;

                    // Handle the case where the add-in does not support forcing consent.
                case 13009:
                    if (triedWithoutForceConsent) {
                        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
                    } else {
                        getDataWithToken({ forceConsent: false });
                    }
                    break;

                    // Log all other client errors.
                default:
                    logError(result);
                    break;
            }
        }

        // function handleServerSideErrors(result) {

        //     // TODO10: Handle the case where AAD asks for an additional form of authentication.
        //     if (result.responseJSON.error.innerError &&
        //         result.responseJSON.error.innerError.error_codes &&
        //         result.responseJSON.error.innerError.error_codes[0] === 50076) {
        //         getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
        //     }
        //     // TODO11: Handle the case where consent has not been granted, or has been revoked.
        //     else if (result.responseJSON.error.innerError &&
        //         result.responseJSON.error.innerError.error_codes &&
        //         result.responseJSON.error.innerError.error_codes[0] === 65001) {
        //         getDataWithToken({ forceConsent: true });
        //     }
        //     // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow
        //     else if (result.responseJSON.error.innerError &&
        //         result.responseJSON.error.innerError.error_codes &&
        //         result.responseJSON.error.innerError.error_codes[0] === 70011) {
        //         showResult(['The add-in is asking for a type of permission that is not recognized.']);
        //     }
        //     // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //     //         server-side is not valid because it is missing `access_as_user` scope (permission).
        //     else if (result.responseJSON.error.name &&
        //         result.responseJSON.error.name.indexOf('expected access_as_user') !== -1) {
        //         showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
        //     }
        //     // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //     //         data is expired or invalid.
        //     else if (result.responseJSON.error.name &&
        //         result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        //         if (!timesMSGraphErrorReceived) {
        //             timesMSGraphErrorReceived = true;
        //             timesGetOneDriveFilesHasRun = 0;
        //             triedWithoutForceConsent = false;
        //             getOneDriveFiles();
        //         } else {
        //             logError(result);
        //         }
        //     }
        //     // TODO15: Log all other server errors.
        //     else {
        //         logError(result);
        //     }
        // }

        function getDataWithToken(options) {
            Office.context.auth.getAccessTokenAsync(options,
                function(result) {
                    if (result.status === "succeeded") {
                        console.log('accesstokenasync', result.value);
                        var getTokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token?grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&client_id=68ed1183-98b6-4eff-a9da-3ced59f61d23&client_secret=uaxITU03?{dzpxQRJF979:!&assertion=' + result.value + '&scope=user.read sites.read.all profile&requested_token_use=on_behalf_of'

                        $.ajax({
                            url: getTokenUrl,
                            method: 'POST'
                        }).done(function(item) {
                            console.log(item);
                        }).fail(function(error) {
                            console.log(error)
                        });

                    } else {
                        handleClientSideErrors(result);
                    }
                });
        }

        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
            if (result.status === "succeeded") {
                _this.accessToken = result.value;

                // Use the access token
                console.log(_this.accessToken);
            } else {
                // Handle the error
            }
        });

        _this.run = function() {
            function getItemRestId() {
                if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
                    // itemId is already REST-formatted
                    return Office.context.mailbox.item.itemId;
                } else {
                    // Convert to an item ID for API v2.0
                    return Office.context.mailbox.convertToRestId(
                        Office.context.mailbox.item.itemId,
                        Office.MailboxEnums.RestVersion.v2_0
                    );
                }
            }

            function getCurrentThread(accessToken, threadId) {
                // Construct the REST URL to the current item
                // Details for formatting the URL can be found at 
                // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest
                var getMessageUrl = _this.baseRestHost + '/v2.0/me/messages/?$filter=conversationId eq \'' + threadId + '\'&$top=20';

                $.ajax({
                    url: getMessageUrl,
                    dataType: 'json',
                    headers: { 'Authorization': 'Bearer ' + accessToken }
                }).done(function(item) {
                    console.log(item);
                    triedWithoutForceConsent = true;
                    getDataWithToken({ forceConsent: false });
                }).fail(function(error) {
                    // Handle error
                });
            }

            function getCurrentItem(accessToken) {
                // Get the item's REST ID
                var itemId = getItemRestId();

                // Construct the REST URL to the current item
                // Details for formatting the URL can be found at 
                // https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#get-a-message-rest
                var getMessageUrl = _this.baseRestHost + '/v2.0/me/messages/' + itemId;

                $.ajax({
                    url: getMessageUrl,
                    dataType: 'json',
                    headers: { 'Authorization': 'Bearer ' + accessToken }
                }).done(function(item) {
                    getCurrentThread(accessToken, item.ConversationId);
                }).fail(function(error) {
                    // Handle error
                });
            }

            getCurrentItem(_this.accessToken);
        }
    }

    // when Office has initalized, manually bootstrap the app
    Office.initialize = function() {
        angular.bootstrap(document.body, ['outlook-2-sp']);
    };

})();