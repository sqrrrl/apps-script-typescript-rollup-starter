/*
 * Copyright 2022 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
function getHostApp() {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
        return SpreadsheetApp;
    }
    if (DocumentApp.getActiveDocument()) {
        return DocumentApp;
    }
    if (SlidesApp.getActivePresentation()) {
        return SlidesApp;
    }
    if (FormApp.getActiveForm()) {
        return FormApp;
    }
    return undefined;
}

/*
 * Copyright 2022 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
// Cached instance of service for current script execution
let _service;
/**
 * Fetches the configured oauth service.
 *
 * @returns Oauth service instance
 */
function getOAuthService() {
    if (_service) {
        return _service;
    }
    _service = buildOAuthService();
    return _service;
}
/**
 * Handles the oauth callback.
 *
 * @param request
 * @returns HTML page to display to user
 */
function authCallback(request) {
    const service = getOAuthService();
    service.handleCallback(request);
    return HtmlService.createHtmlOutputFromFile('auth-complete');
}
/**
 * Fetch the current oauth state for the user/document.
 *
 * @returns Current oauth state.
 */
function getAuthorizationState() {
    const service = getOAuthService();
    let user = undefined;
    try {
        user = getAuthorizedUser();
    }
    catch (err) {
        // Ignore failures
        console.error(err);
    }
    return {
        user,
        authorized: service.hasAccess(),
        authorizationUrl: service.getAuthorizationUrl()
    };
}
/**
 * Removes saved oauth credentials.
 *
 * @returns Updated auth info.
 */
function disconnect() {
    const service = getOAuthService();
    service.reset();
    return getAuthorizationState();
}
/**
 * Fetches userinfo for the connected account.
 *
 * @returns user id
 */
function getAuthorizedUser() {
    const service = getOAuthService();
    if (!service.hasAccess()) {
        throw new Error('App is not authorized.');
    }
    const response = UrlFetchApp.fetch('https://oauth.mocklab.io/userinfo', {
        method: 'get',
        headers: {
            'Authorization': `Bearer ${service.getAccessToken()}`,
        },
        muteHttpExceptions: true,
        followRedirects: true,
    });
    const status = response.getResponseCode();
    if (status >= 400) {
        throw new Error(`Unable to fetch user, status code ${status} ${response.getContentText()}`);
    }
    const parsedResponse = JSON.parse(response.getContentText());
    return parsedResponse['email'];
}
/**
 * Constructs the oauth service. Callers should use
 * `getOAuthService` to get a cached instance.
 *
 * @returns Initialized oauth service
 */
function buildOAuthService() {
    // TODO(developer) - Configure appropriately and/or read from script properties
    const clientId = 'mock-client-id';
    const clientSecret = 'mock-client-secret';
    const scope = 'mock-scope';
    const authUrl = 'https://oauth.mocklab.io/oauth/authorize';
    const tokenUrl = 'https://oauth.mocklab.io/oauth/token';
    const service = OAuth2.createService('mock_oauth')
        .setClientId(clientId)
        .setClientSecret(clientSecret)
        .setScope(scope)
        .setCallbackFunction('authCallback')
        .setAuthorizationBaseUrl(authUrl)
        .setTokenUrl(tokenUrl);
    const cache = CacheService.getDocumentCache();
    if (cache) {
        service.setCache(cache);
    }
    const store = PropertiesService.getUserProperties();
    if (store) {
        service.setPropertyStore(store);
    }
    return service;
}

/*
 * Copyright 2022 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Adds a custom menu with items to show the sidebar and dialog.
 */
function onOpen() {
    const host = getHostApp();
    if (!host) {
        return;
    }
    host.getUi().createAddonMenu()
        .addItem('Sidebar Example', 'showSidebar')
        .addItem('OAuth Example', 'showAuthorizationExample')
        .addToUi();
}
/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 */
function onInstall() {
    onOpen();
}
/**
 * Displays authorization info for the sheet
 */
function showAuthorizationExample() {
    const host = getHostApp();
    if (!host) {
        return;
    }
    const html = HtmlService.createHtmlOutputFromFile('authorize')
        .setWidth(1200)
        .setHeight(800)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    host.getUi().showModalDialog(html, 'Authorize Add-on');
}
/**
 * Displays a sample sidebar
 */
function showSidebar() {
    const host = getHostApp();
    if (!host) {
        return;
    }
    const html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    host.getUi().showSidebar(html);
}
/**
 * Small demo function for sidebar example to show RPC to apps script
 */
function getActiveUser() {
    return Session.getActiveUser().getEmail();
}
