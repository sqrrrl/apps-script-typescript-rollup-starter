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

import { AuthInfo } from '../shared/types';

// Cached instance of service for current script execution
let _service: GoogleAppsScriptOAuth2.OAuth2Service | undefined;

/**
 * Fetches the configured oauth service.
 * 
 * @returns Oauth service instance
 */
export function getOAuthService(): GoogleAppsScriptOAuth2.OAuth2Service {
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
 export function authCallback(request: object): GoogleAppsScript.HTML.HtmlOutput {
  const service = getOAuthService();
  service.handleCallback(request);

  return HtmlService.createHtmlOutputFromFile('auth-complete')
}

/**
 * Fetch the current oauth state for the user/document.
 * 
 * @returns Current oauth state.
 */
 export function getAuthorizationState(): AuthInfo {
  const service = getOAuthService();
  let user = undefined;
  try {
    user = getAuthorizedUser();
  } catch (err) {
    // Ignore failures
    console.error(err);
  }
  return {
    user,
    authorized: service.hasAccess(),
    authorizationUrl: service.getAuthorizationUrl()
  }
}

/**
 * Removes saved oauth credentials.
 * 
 * @returns Updated auth info.
 */
 export function disconnect(): AuthInfo {
  const service = getOAuthService();
  service.reset();
  return getAuthorizationState();
}


/**
 * Fetches userinfo for the connected account.
 * 
 * @returns user id 
 */
export function getAuthorizedUser(): string {
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
function buildOAuthService(): GoogleAppsScriptOAuth2.OAuth2Service {
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