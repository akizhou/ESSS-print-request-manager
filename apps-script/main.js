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


// Run only once when triggers haven't been set up
function setUpTriggers() {
  ScriptApp.newTrigger("onRequestSubmission")
  .forSpreadsheet(REQUEST_QUEUE_ID)
  .onFormSubmit().create();

  ScriptApp.newTrigger("onEditRequest")
  .forSpreadsheet(REQUEST_QUEUE_ID)
  .onEdit().create();
}

function onRequestSubmission(e) {
  let rowOfCellEdited = e.range.getRow();
  let serviceActive = statSheet.getRange(2, SERVICESTAT_COLS.serviceActive).getValue() == "Active";
  let update = new QueueUpdater(rowOfCellEdited);
  let requestType = formSheet.getRange(rowOfCellEdited, FORM_COLS.requestType).getValue();

  if (requestType == "Cancel request") {
    let cancelID = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelID).getValue();
    let requestInfo = new Array(15).fill("");
    requestInfo[0] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelID).getValue();
    requestInfo[2] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelName).getValue();
    requestInfo[3] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelEmail).getValue();
    
    // Attempt to cancel the request with submitted ID from Queue and respond accordingly
    let cancellationStatus = update.cancelRequest(cancelID);
    let respond = new EmailResponder(requestInfo);

    // Send confirmation email depending on the cancellation status
    respond.confirmCancellation(cancellationStatus);
  }
  else if(requestType == "New request") {
    // Enqueue request
    let rowOfRequest = update.newRequest();
    let requestInfo = queueSheet.getRange(`A${ rowOfRequest }:O${ rowOfRequest }`).getValues()[0];
    let respond = new EmailResponder(requestInfo);

    // Send confirmation email and inform the ID
    respond.confirmReception();
    
    if (!serviceActive) {
      // Notify service is inactive
      respond.notifyServiceInactive();
    }
  }
}

function onEditRequest(e) {
  let columnOfCellEdited = e.range.getColumn();
  let rowOfCellEdited = e.range.getRow();
  let update = new QueueUpdater(rowOfCellEdited);
  
  // Status changes
  if (columnOfCellEdited == QUEUE_COLS.status) {
    let status = queueSheet.getRange(rowOfCellEdited, columnOfCellEdited).getValue();
    let requestInfo = queueSheet.getRange(`A${ rowOfCellEdited }:O${ rowOfCellEdited }`).getValues()[0];
    requestInfo[QUEUE_COLS.printTime - 1] = queueSheet.getRange(rowOfCellEdited, QUEUE_COLS.printTime).getDisplayValue();    
    let respond = new EmailResponder(requestInfo);

    if (status == "quoted") {
      respond.sendQuote();
    }
    else if (status == "printing") {
      respond.notifyPrintStarted();
    }
    else if (status == "awaiting pickup") {
      let pin = respond.notifyPickUpReady();
      update.setPin(pin);
    }
    else if (status == "completed") {
      update.archiveRequest(false);
    }
  }

  // Calculate price
  if (columnOfCellEdited == QUEUE_COLS.printTime || columnOfCellEdited == QUEUE_COLS.materialUsage) {
    update.setPrice();
  }
}
