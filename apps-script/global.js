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


const REQUEST_QUEUE_ID = "1VaIgGmAjHij4lnaOpsfFHlfxuzJwyVUicb8dzskDZaY";

var spreadsheet = SpreadsheetApp.openById(REQUEST_QUEUE_ID);
var queueSheet = spreadsheet.getSheetByName("Queue");
var statSheet = spreadsheet.getSheetByName("Service Status");
var archiveSheet = spreadsheet.getSheetByName("Archive");
var formSheet = spreadsheet.getSheetByName("Form Responses");

const SERVICESTAT_COLS = {
  serviceActive: 1,
  directorName: 6,
  directorPhone: 7
}

const FORM_COLS = {
  timestamp: 1,
  requestType: 2,
  name: 3,
  email: 4,
  academic: 5,
  safeAccess: 6,
  gcode: 7,
  comments: 12,
  receipt: 13,
  cancelName: 14,
  cancelEmail: 15,
  cancelID: 16,
  id: 17
}
Object.freeze(FORM_COLS);

const QUEUE_COLS = {
  id: 1,
  timestamp: 2,
  name: 3,
  email: 4,
  academic: 5,
  gcode: 6,
  link: 7,
  color: 8,
  comments: 9,
  status: 10,
  printTime: 11,
  materialUsage: 12,
  price: 13,
  pin: 14,
  receipt: 15
}
Object.freeze(QUEUE_COLS);