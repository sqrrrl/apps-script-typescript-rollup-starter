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

import { getHostApp } from './utils';

// Re-export exposed oauth methods
export { authCallback, disconnect, getAuthorizationState} from './oauth';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 */
export function onOpen() {
  const host = getHostApp();
  if(!host) {
    return;
  }
  host.getUi().createAddonMenu()
    .addItem('Sidebar Example', 'showSidebar')
    .addItem('OAuth Example', 'showAuthorizationExample')
    .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initialization work is done immediately.
 */
export function onInstall() {
  onOpen();
}

/**
 * Displays authorization info for the sheet
 */
 export function showAuthorizationExample() {
  const host = getHostApp();
  if(!host) {
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
export function showSidebar() {
  const host = getHostApp();
  if(!host) {
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  host.getUi().showSidebar(html);
}

/**
 * Small demo function for sidebar example to show RPC to apps script
 */
export function getActiveUser(): string {
  return Session.getActiveUser().getEmail();
}

