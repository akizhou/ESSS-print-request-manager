// Copyright 2021 Aki Zhou
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


const REQUEST_QUEUE_ID: string = "1VaIgGmAjHij4lnaOpsfFHlfxuzJwyVUicb8dzskDZaY";

var spreadsheet: Spreadsheet = SpreadsheetApp.openById(REQUEST_QUEUE_ID);
var queueSheet: Sheet = spreadsheet.getSheetByName("Queue");
var statSheet: Sheet = spreadsheet.getSheetByName("Service Status");
var archiveSheet: Sheet = spreadsheet.getSheetByName("Archive");
var formSheet: Sheet = spreadsheet.getSheetByName("Form Responses");

const enum SERVICESTAT_COLS {
  serviceActive = 1,
  directorName = 6,
  directorPhone = 7
}

const enum FORM_COLS {
  timestamp = 1,
  requestType,
  name,
  email,
  academic,
  safeAccess,
  gcode,
  gcodeLink,
  gcodeColor,
  modelLink,
  modelColor,
  comments,
  receipt,
  cancelName,
  cancelEmail,
  cancelID,
  id
}

const enum QUEUE_COLS {
  id = 1,
  timestamp,
  name,
  email,
  academic,
  gcode,
  link,
  color,
  comments,
  status,
  printTime,
  materialUsage,
  price,
  pin,
  receipt
}
