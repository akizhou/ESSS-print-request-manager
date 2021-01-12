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


class QueueUpdater {
  row: bigint;  // Row of the request
  newRow: bigint;  // Row of which a new request is going to be inserted to in Queue sheet
  ;

  constructor(rowOfRequest: bigint) {
    this.row = rowOfRequest;
    this.newRow = queueSheet.getLastRow() + 1;
  }

  /**
   * Parses request info from From responses sheet and enqueue it to Queue sheet
   */
  newRequest() {
    Logger.log("Processing new request");

    // Get request info from Form responses sheet
    let id: string = this.generateID();
    let timestamp: string = formSheet.getRange(this.row, FORM_COLS.timestamp).getValue();
    let name: string = formSheet.getRange(this.row, FORM_COLS.name).getValue();
    let email: string = formSheet.getRange(this.row, FORM_COLS.email).getValue();
    let academic: string = formSheet.getRange(this.row, FORM_COLS.academic).getValue();
    let safeAccess: string = formSheet.getRange(this.row, FORM_COLS.safeAccess).getValue();
    let gcode: string = formSheet.getRange(this.row, FORM_COLS.gcode).getValue();
    let link: string = (gcode == "Yes") ? formSheet.getRange(this.row, 8).getValue() : formSheet.getRange(this.row, 10).getValue();
    let color: string = (gcode == "Yes") ? formSheet.getRange(this.row, 9).getValue() : formSheet.getRange(this.row, 11).getValue();
    let comments: string = formSheet.getRange(this.row, FORM_COLS.comments).getValue();
    let receipt: string = formSheet.getRange(this.row, FORM_COLS.receipt).getValue();

    // Assign new ID
    formSheet.getRange(this.row, FORM_COLS.id).setValue(id)
    .copyTo(queueSheet.getRange(this.newRow, QUEUE_COLS.id), {contentOnly: true});

    // Enqueue request by putting request information into corresponding cell in Queue
    queueSheet.getRange(this.newRow, QUEUE_COLS.timestamp).setValue(timestamp);
    queueSheet.getRange(this.newRow, QUEUE_COLS.name).setValue(name);
    queueSheet.getRange(this.newRow, QUEUE_COLS.email).setValue(email);
    queueSheet.getRange(this.newRow, QUEUE_COLS.academic).setValue(academic);
    queueSheet.getRange(this.newRow, QUEUE_COLS.gcode).setValue(gcode);
    queueSheet.getRange(this.newRow, QUEUE_COLS.link).setValue(link);
    queueSheet.getRange(this.newRow, QUEUE_COLS.color).setValue(color);
    queueSheet.getRange(this.newRow, QUEUE_COLS.comments).setValue(comments);
    queueSheet.getRange(this.newRow, QUEUE_COLS.receipt).setValue(receipt);
    queueSheet.getRange(this.newRow, QUEUE_COLS.pin).setValue(safeAccess);

    // Change status to received
    queueSheet.getRange(this.newRow, QUEUE_COLS.status).setValue("received");

    Logger.log("Processed");
    
    return this.newRow;
  }

  /**
   * Verifies if a request can be cancelled and archive it as cancelled if eligible.
   * Returns -1 if request with matching ID DNE, 1 if request cannot be cancelled, otherwise a integer > 1.
   */
  cancelRequest(targetID: string) {
    Logger.log("Processing cancel request");

    // If no request in Queue
    if (this.newRow == 2) {
      // No request with matching ID found
      Logger.log("No request found");

      return -1;
    }

    // Match with IDs in Queue
    let scanOffset: bigint = 2;
    let scans: Array<string> = queueSheet.getRange(scanOffset, QUEUE_COLS.id, (this.newRow - scanOffset)).getValues().map(x=>x[0]);

    for (let i = 0; i < scans.length; i++) {
      if (scans[i] == targetID) {
        let targetRow: bigint = i + scanOffset;
        let requestStatus: string = queueSheet.getRange(targetRow, QUEUE_COLS.status).getValue();
        if (requestStatus == "received" || requestStatus == "quoted") {
          // Request is present and allowed to be cancelled
          // Set status to cancelled
          queueSheet.getRange(targetRow, QUEUE_COLS.status).setValue("cancelled");

          // Move request to Archive
          this.archiveRequest(targetRow);
          Logger.log("Request cancelled");

          return targetRow;
        }
        // Request with matching ID is found but too late to be cancelled
        Logger.log("Request cannot be cancelled");

        return 1;
      }
    }

    // No request with matching ID found
    Logger.log("No request found");

    return -1;
  }

  /**
   * Archive request. Request is either completed or cancelled.
   */
  archiveRequest(cancel: bigint) {
    Logger.log("Archiving request");

    // cancel is false if request is completed, in that case this.row is the target row
    // since this.row is the row of request whose status got changed to complete
    let targetRow: bigint = cancel ? cancel : this.row;

    // Move request from queue to archive
    // moveTo() will remove the data validation from Queue for the moved row
    // for this reason copy and then clear the content instead
    let target: Range = queueSheet.getRange(`${ targetRow }:${ targetRow }`)
    target.copyTo(archiveSheet.getRange(`A${ archiveSheet.getLastRow() + 1 }`))
    target.clearContent();

    // Refresh Queue if there are any requests
    if (queueSheet.getLastRow() > 1) {
      queueSheet.getRange(2, QUEUE_COLS.id, (queueSheet.getLastRow() - 1), queueSheet.getLastColumn())
      .sort(QUEUE_COLS.timestamp);
    }
    
    // Sort Archive
    archiveSheet.getRange(2, QUEUE_COLS.id, (archiveSheet.getLastRow() - 1), archiveSheet.getLastColumn())
    .sort(QUEUE_COLS.timestamp);

    Logger.log("Request archived");
  }

  setPin(pin: string) {
    queueSheet.getRange(this.row, QUEUE_COLS.pin).setValue(pin);
  }

  /**
   * Calculates and sets the price based on print time and material usage.
   * A discount is applied if the request is for academic projects.
   */
  setPrice() {
    const PLA_COST_PER_GRAM: number = 0.05;
    const CHARGE_PER_MIN: number = 0.017;
    const CHARGE_PER_MIN_ACADEMIC: number = 0.013;

    let timeStr: string = queueSheet.getRange(this.row, QUEUE_COLS.printTime).getDisplayValue();
    // Convert time from hh:mm to all minutes
    let timeParse: Array<string> = timeStr.split(":")
    let time: bigint = parseInt(timeParse[0]) * 60 + parseInt(timeParse[1]);
    let material: number = parseInt(queueSheet.getRange(this.row, QUEUE_COLS.materialUsage).getValue());

    if (!isNaN(time) && !isNaN(material)) {
      Logger.log("Calculating request price");

      let academic: string = queueSheet.getRange(this.row, FORM_COLS.academic).getValue();
      let rate: number = (academic == "Yes") ? CHARGE_PER_MIN_ACADEMIC : CHARGE_PER_MIN;

      let price: bigint = Math.ceil(time * rate + material * PLA_COST_PER_GRAM);

      queueSheet.getRange(this.row, QUEUE_COLS.price).setValue(price);

      Logger.log("Price calculated");
    }
  }

  /**
   * Generates an unique ID of length 5, consisting of 4 digits followed by one upper case alphabet.
   */
  generateID() {
    const ALPHABET: string = "abcdefghijklmnopqrstuvwxyz";
    let issuedIDs: Array<string> = formSheet.getRange(`O2:O${ formSheet.getLastRow() }`).getValues().map(x=>x[0]);
    let newID: string;

    Logger.log("Generating a new ID");

    do {
      // ID shouldn't be predictable in order to prevent people deleting other people's request
      let randomCharacter: string = ALPHABET[Math.floor(Math.random() * ALPHABET.length)].toUpperCase();
      newID = Math.random().toString().slice(2, 6) + randomCharacter;
    }
    while (issuedIDs.indexOf(newID) != -1);

    Logger.log("ID generated");

    return newID;
  }
}
