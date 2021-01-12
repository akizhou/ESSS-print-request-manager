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
  constructor(rowOfRequest) {
    this.row = rowOfRequest;
    this.newRow = queueSheet.getLastRow() + 1;  //this.getLastRowIndex(queueSheet);
  }

  // Only used for new requests
  init() {
    this.id = this.generateID();
    this.timestamp = formSheet.getRange(this.row, FORM_COLS.timestamp).getValue();
    this.name = formSheet.getRange(this.row, FORM_COLS.name).getValue();
    this.email = formSheet.getRange(this.row, FORM_COLS.email).getValue();
    this.academic = formSheet.getRange(this.row, FORM_COLS.academic).getValue();
    this.safeAccess = formSheet.getRange(this.row, FORM_COLS.safeAccess).getValue();
    this.gcode = formSheet.getRange(this.row, FORM_COLS.gcode).getValue();
    this.link = (this.gcode == "Yes") ? formSheet.getRange(this.row, 8).getValue() : formSheet.getRange(this.row, 10).getValue();
    this.color = (this.gcode == "Yes") ? formSheet.getRange(this.row, 9).getValue() : formSheet.getRange(this.row, 11).getValue();
    this.comments = formSheet.getRange(this.row, FORM_COLS.comments).getValue();
    this.receipt = formSheet.getRange(this.row, FORM_COLS.receipt).getValue();
  }

  newRequest() {
    this.init();
    // Assign new ID
    formSheet.getRange(this.row, FORM_COLS.id).setValue(this.id)
    .copyTo(queueSheet.getRange(this.newRow, QUEUE_COLS.id), {contentOnly: true});

    // Enqueue request by putting request information into corresponding cell in Queue
    queueSheet.getRange(this.newRow, QUEUE_COLS.timestamp).setValue(this.timestamp);
    queueSheet.getRange(this.newRow, QUEUE_COLS.name).setValue(this.name);
    queueSheet.getRange(this.newRow, QUEUE_COLS.email).setValue(this.email);
    queueSheet.getRange(this.newRow, QUEUE_COLS.academic).setValue(this.academic);
    queueSheet.getRange(this.newRow, QUEUE_COLS.gcode).setValue(this.gcode);
    queueSheet.getRange(this.newRow, QUEUE_COLS.link).setValue(this.link);
    queueSheet.getRange(this.newRow, QUEUE_COLS.color).setValue(this.color);
    queueSheet.getRange(this.newRow, QUEUE_COLS.comments).setValue(this.comments);
    queueSheet.getRange(this.newRow, QUEUE_COLS.receipt).setValue(this.receipt);
    queueSheet.getRange(this.newRow, QUEUE_COLS.pin).setValue(this.safeAccess);

    // Change status to received
    queueSheet.getRange(this.newRow, QUEUE_COLS.status).setValue("received");
    
    return this.newRow;
  }

  cancelRequest(targetID) {
    // If no request in Queue
    if (this.newRow == 2) {
      // No request with matching ID found
      return -1;
    }

    // Match with IDs in Queue
    let scanOffset = 2;
    let scans = queueSheet.getRange(scanOffset, QUEUE_COLS.id, (this.newRow - scanOffset)).getValues().map(x=>x[0]);

    for (let i = 0; i < scans.length; i++) {
      if (scans[i] == targetID) {
        let targetRow = i + scanOffset;
        let requestStatus = queueSheet.getRange(targetRow, QUEUE_COLS.status).getValue();
        if (requestStatus == "received" || requestStatus == "quoted") {
          // Request is present and allowed to be cancelled
          // Set status to cancelled
          queueSheet.getRange(targetRow, QUEUE_COLS.status).setValue("cancelled");

          // Move request to Archive
          this.archiveRequest(targetRow);
          return targetRow;
        }
        // Request with matching ID is found but too late to be cancelled
        return 1;
      }
    }
    // No request with matching ID found
    return -1;
  }

  archiveRequest(cancel) {
    // cancel is false if request is completed, in that case this.row is the target row
    // because in that case this.row is the row of request whose status got changed to complete
    let targetRow = cancel ? cancel : this.row;

    // Move request from queue to archive
    // moveTo() will remove the data validation from Queue for the moved row
    // for this reason copy and then clear the content instead
    let target = queueSheet.getRange(`${ targetRow }:${ targetRow }`)
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
  }

  setPin(pin) {
    queueSheet.getRange(this.row, QUEUE_COLS.pin).setValue(pin);
  }

  setPrice() {
    const PLA_COST_PER_GRAM = 0.05;
    const CHARGE_PER_MIN = 0.017;
    const CHARGE_PER_MIN_ACADEMIC = 0.013;

    let time = queueSheet.getRange(this.row, QUEUE_COLS.printTime).getDisplayValue();
    // Convert time from hh:mm to all minutes
    time = time.split(":")
    time = parseInt(time[0]) * 60 + parseInt(time[1]);
    let material = parseInt(queueSheet.getRange(this.row, QUEUE_COLS.materialUsage).getValue());

    if (!isNaN(time) && !isNaN(material)) {
      this.academic = queueSheet.getRange(this.row, FORM_COLS.academic).getValue();
      let rate = (this.academic == "Yes") ? CHARGE_PER_MIN_ACADEMIC : CHARGE_PER_MIN;

      let price = Math.ceil(time * rate + material * PLA_COST_PER_GRAM);

      queueSheet.getRange(this.row, QUEUE_COLS.price).setValue(price);
    }
  }

  generateID() {
    const ALPHABET = "abcdefghijklmnopqrstuvwxyz";
    let issuedIDs = formSheet.getRange(`O2:O${ formSheet.getLastRow() }`).getValues().map(x=>x[0]);
    let newID;

    do {
      // ID shouldn't be predictable in order to prevent people deleting other people's request
      let randomCharacter = ALPHABET[Math.floor(Math.random() * ALPHABET.length)].toUpperCase();
      newID = Math.random().toString().slice(2, 6) + randomCharacter;
    }
    while (issuedIDs.indexOf(newID) != -1);

    return newID;
  }
}
