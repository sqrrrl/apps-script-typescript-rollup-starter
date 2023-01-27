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

import { LitElement, css, html } from 'lit'
import { customElement, state } from 'lit/decorators.js'
import { asyncAppsScriptFunction } from './script-utils';

const getActiveUser = asyncAppsScriptFunction<[],string> ('getActiveUser');

/**
 * An example element.
 *
 * @slot - This element has a slot
 * @csspart button - The button
 */
@customElement('app-active-user')
export class QueryEditorElement extends LitElement {
  @state()
  private _user: string | undefined;

  connectedCallback() {
    super.connectedCallback();
    this._fetchActiveUser();
  }

  render() {
    if (!this._user) {
      return html`<div>Loading...</div>`;
    }
    return html`
      <div>Hello ${this._user}</div>
    `;
  }

  private async _fetchActiveUser() {
    this._user = await getActiveUser();
  }

  static styles = css`
    :host {
      max-width: 1280px;
      margin: 0 auto;
      padding: var(--size-3);
      display: flex;
      flex-direction: column;
      padding: var(--size-3);
      gap: var(--size-3);
    }
  `
}

declare global {
  interface HTMLElementTagNameMap {
    'app-user-info': QueryEditorElement
  }
}
