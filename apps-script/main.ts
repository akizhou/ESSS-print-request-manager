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
  Logger.log("Creating triggers");

  // Creates a onFormSubmit trigger. Triggered by Request Queue spreadsheet
  ScriptApp.newTrigger("onRequestSubmission")
  .forSpreadsheet(REQUEST_QUEUE_ID)
  .onFormSubmit().create();

  // Creates a onEdit trigger. Triggered Request Queue spreadsheet
  ScriptApp.newTrigger("onEditRequest")
  .forSpreadsheet(REQUEST_QUEUE_ID)
  .onEdit().create();

  Logger.log("Triggers created")
}

// Executed onFormSubmit
function onRequestSubmission(e) {
  let rowOfCellEdited: bigint = e.range.getRow();  // Get the row of new form submission. This is from the From responses sheet.
  let serviceActive: boolean = statSheet.getRange(2, SERVICESTAT_COLS.serviceActive).getValue() == "Active";
  let update: QueueUpdater = new QueueUpdater(rowOfCellEdited);

  // There are two types of requests: Cancel request and New request
  let requestType: string = formSheet.getRange(rowOfCellEdited, FORM_COLS.requestType).getValue();

  if (requestType == "Cancel request") {
    Logger.log("Received cancel request");

    // Get request info from From responses sheet
    let cancelID: string = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelID).getValue();
    // For cancel requests the requestInfo passed to EmailResponder is special.
    // Since the only useful info are ID, name, and email everything else is initialized as empty strings.
    let requestInfo: Array<string> = new Array(15).fill("");
    requestInfo[0] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelID).getValue();
    requestInfo[2] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelName).getValue();
    requestInfo[3] = formSheet.getRange(rowOfCellEdited, FORM_COLS.cancelEmail).getValue();
    
    // Attempt to cancel the request with submitted ID from Queue sheet
    let cancellationStatus: bigint = update.cancelRequest(cancelID);

    Logger.log("Responding by email");
    // Send confirmation email depending on the cancellation status
    let respond: EmailResponder = new EmailResponder(requestInfo);
    respond.confirmCancellation(cancellationStatus);

    Logger.log("Processed cancel request");
  }
  else if(requestType == "New request") {
    Logger.log("Received new request");

    // Enqueue request to Queue sheet
    let rowOfRequest: bigint = update.newRequest();
    let requestInfo: Array<string> = queueSheet.getRange(`A${ rowOfRequest }:O${ rowOfRequest }`).getValues()[0];

    Logger.log("Responding by email");
    // Send confirmation email and inform the ID
    let respond: EmailResponder = new EmailResponder(requestInfo);
    respond.confirmReception();
    
    if (!serviceActive) {
      // Notify service is inactive
      respond.notifyServiceInactive();
    }

    Logger.log("Enqueued new request");
  }
}

// Executed onEdit
function onEditRequest(e) {
  let columnOfCellEdited: bigint = e.range.getColumn();
  let rowOfCellEdited: bigint = e.range.getRow();
  let update: QueueUpdater = new QueueUpdater(rowOfCellEdited);
  
  // Status changes
  if (columnOfCellEdited == QUEUE_COLS.status) {
    Logger.log("Status of request updated");

    let status: string = queueSheet.getRange(rowOfCellEdited, columnOfCellEdited).getValue();
    let requestInfo: Array<string> = queueSheet.getRange(`A${ rowOfCellEdited }:O${ rowOfCellEdited }`).getValues()[0];
    requestInfo[QUEUE_COLS.printTime - 1] = queueSheet.getRange(rowOfCellEdited, QUEUE_COLS.printTime).getDisplayValue();    
    let respond: EmailResponder = new EmailResponder(requestInfo);

    if (status == "quoted") {
      respond.sendQuote();
      Logger.log("Request quoted");
    }
    else if (status == "printing") {
      respond.notifyPrintStarted();
      Logger.log("Request print initiation notified");
    }
    else if (status == "awaiting pickup") {
      let pin: string = respond.notifyPickUpReady();
      update.setPin(pin);
      Logger.log("Request pickup notified");
    }
    else if (status == "completed") {
      update.archiveRequest(0);  // 0 passed to archiveRequest() indicates the request is completed
      Logger.log("Request completed and archived");
    }
  }

  // Calculate price
  if (columnOfCellEdited == QUEUE_COLS.printTime || columnOfCellEdited == QUEUE_COLS.materialUsage) {
    update.setPrice();
  }
}
