/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/* eslint-disable @typescript-eslint/no-unused-vars */
import {
  onGmailMessageOpen as _onGmailMessageOpen,
  createEvent as _createEvent,
} from './gmail-to-calendar';

// Google Apps Scriptのグローバル関数として公開
declare global {
  function onGmailMessageOpen(
    e: GoogleAppsScript.Addons.EventObject
  ): Promise<GoogleAppsScript.Card_Service.Card>;
  function createEvent(
    e: GoogleAppsScript.Addons.EventObject
  ): GoogleAppsScript.Card_Service.ActionResponse;
}

// グローバル関数の実装
globalThis.onGmailMessageOpen = _onGmailMessageOpen;
globalThis.createEvent = _createEvent;
