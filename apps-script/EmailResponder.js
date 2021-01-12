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


class EmailResponder {
  constructor(info) {
    this.id = info[QUEUE_COLS.id - 1];
    this.name = info[QUEUE_COLS.name - 1];
    this.email = info[QUEUE_COLS.email - 1];
    this.academic = info[QUEUE_COLS.academic - 1];
    this.color = info[QUEUE_COLS.color - 1];
    this.printTime = info[QUEUE_COLS.printTime - 1];
    this.materialUsage = info[QUEUE_COLS.materialUsage - 1];
    this.price = info[QUEUE_COLS.price - 1];
    this.safeAccess = info[QUEUE_COLS.pin - 1];  // Using pin cell to determine if client has safe access
    this.receipt = info[QUEUE_COLS.receipt - 1];

    this.DIRECTOR_NAME = statSheet.getRange(2, SERVICESTAT_COLS.directorName).getValue();
    this.DIRECTOR_PHONE = statSheet.getRange(2, SERVICESTAT_COLS.directorPhone).getValue();
    this.message = "Please enable html";
  }

  notifyServiceInactive() {
    let subject = "Service inactive";

    let emailTemp = HtmlService.createTemplateFromFile("serviceInactive");
    emailTemp.name = this.name;
    emailTemp.directorName = this.DIRECTOR_NAME;
    emailTemp.directorPhone = this.DIRECTOR_PHONE;
    let htmlMessage = emailTemp.evaluate().getContent();

    GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing"
    });
  }

  confirmReception() {
    let subject = `Print request received (ID: ${ this.id })`;

    let emailTemp = HtmlService.createTemplateFromFile("received");
    emailTemp.name = this.name;
    emailTemp.id = this.id;
    emailTemp.directorName = this.DIRECTOR_NAME;
    emailTemp.directorPhone = this.DIRECTOR_PHONE;
    let htmlMessage = emailTemp.evaluate().getContent();
    
    GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing"
    });
  }

  confirmCancellation(cancellationStatus) {
    let subject = `Request couldn't be cancelled (ID: ${ this.id })`;

    let emailTemp = HtmlService.createTemplateFromFile("cancel");
    emailTemp.name = this.name;
    if (cancellationStatus > 1) {
      // Successfully cancelled
      subject = `Request cancelled (ID: ${ this.id })`;
      emailTemp.status = "successfully completed";
    }
    else if (cancellationStatus == 1) {
      // Request exists but is too late to be cancelled
      emailTemp.status = "failed because the request is already printing or printed";
    }
    else {
      // No matching request was found to be cancelled
      emailTemp.status = "failed because ID did not match any requests in our queue";
    }
    emailTemp.directorName = this.DIRECTOR_NAME;
    emailTemp.directorPhone = this.DIRECTOR_PHONE;
    let htmlMessage = emailTemp.evaluate().getContent();
    
    GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing"
    });
  }

  sendQuote() {
    let subject = `Here is your quote (ID: ${ this.id })`;

    let emailTemp = HtmlService.createTemplateFromFile("quote");
    emailTemp.name = this.name;
    emailTemp.printTime = this.printTime;
    emailTemp.materialUsage = this.materialUsage;
    emailTemp.price = this.price;
    emailTemp.directorName = this.DIRECTOR_NAME;
    emailTemp.directorPhone = this.DIRECTOR_PHONE;
    let htmlMessage = emailTemp.evaluate().getContent();
    
    GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing"
    });
  }

  notifyPrintStarted() {
    let subject = `We started to print your request!! (ID: ${ this.id })`;

    if (this.receipt == "Yes") {
      let emailTemp = HtmlService.createTemplateFromFile("printStarted");
      emailTemp.name = this.name;
      emailTemp.directorName = this.DIRECTOR_NAME;
      emailTemp.directorPhone = this.DIRECTOR_PHONE;
      let htmlMessage = emailTemp.evaluate().getContent();
      let requestReceipt = this.generateReceipt();

      GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing",
      attachments: [requestReceipt]
      });
    }
    else {
      let emailTemp = HtmlService.createTemplateFromFile("printStarted_noReceipt");
      emailTemp.name = this.name;
      emailTemp.directorName = this.DIRECTOR_NAME;
      emailTemp.directorPhone = this.DIRECTOR_PHONE;
      let htmlMessage = emailTemp.evaluate().getContent();

      GmailApp.sendEmail(this.email, subject, this.message, {
      htmlBody: htmlMessage,
      name: "ESSS 3D Printing"
      });
    }
  }

  notifyPickUpReady() {
    let subject = `Your print is ready for pickup!! (ID: ${ this.id })`;

    if (this.safeAccess == "Yes") {
      let emailTemp = HtmlService.createTemplateFromFile("pickup");
      emailTemp.name = this.name;
      let newPin = this.generatePin();
      emailTemp.pin = newPin;
      emailTemp.directorName = this.DIRECTOR_NAME;
      emailTemp.directorPhone = this.DIRECTOR_PHONE;
      let htmlMessage = emailTemp.evaluate().getContent();

      GmailApp.sendEmail(this.email, subject, this.message, {
        htmlBody: htmlMessage,
        name: "ESSS 3D Printing"
      });

      return newPin;
    }
    else {
      let emailTemp = HtmlService.createTemplateFromFile("arrangePickup");
      emailTemp.name = this.name;
      emailTemp.directorName = this.DIRECTOR_NAME;
      emailTemp.directorPhone = this.DIRECTOR_PHONE;
      let htmlMessage = emailTemp.evaluate().getContent();

      GmailApp.sendEmail(this.email, subject, this.message, {
        htmlBody: htmlMessage,
        name: "ESSS 3D Printing"
      });

      return "-";
    }
    
  }

  generatePin() {
    // Get pins that have been issued
    let pins = new Set();
    let pastPins = archiveSheet.getRange("N2:N").getValues().map(x=>x[0]);
    let currentPins = queueSheet.getRange("N2:N").getValues().map(x=>x[0]);
    for (let i = 0; i < pastPins.length; i++) {
      if (pastPins[i] == "") {
        break;
      }
      pins.add(pastPins[i]); 
    }
    for (let i = 0; i < currentPins.length; i++) {
      if (currentPins[i] == "") {
        break;
      }
      pins.add(currentPins[i]); 
    }

    // Generate a pin that hasn't been issued ever
    const MIN_PINLENGTH = 4;
    const MAX_PINLENGTH = 8;
    let newPin = undefined;
    do {
      // Get a random length
      let pinLength = Math.random() * (MAX_PINLENGTH - MIN_PINLENGTH) + MIN_PINLENGTH;
      // Generate a random number according to pinLength
      newPin = Math.random().toString().slice(2, 3 + pinLength);
    }
    while (pins.has(newPin))

    return newPin;
  }

  generateReceipt() {
    const RECEIPT_TEMPLATE_ID = "1tGHNmwnDmMDuyayAJeUhPA9iOoRzMTSqP3fAKaaMS1I";
    const RECEIPT_FOLDER_ID = "1hBP0ahflX8xWtvsrUmHLDEQY9ih7zjRN";

    let receiptName = `Receipt (${ this.id })`;
    let destination = DriveApp.getFolderById(RECEIPT_FOLDER_ID);

    // Make a copy of the template in Drive and save to Issued Receipts folder  
    let receiptDocFile = DriveApp.getFileById(RECEIPT_TEMPLATE_ID).makeCopy(receiptName, destination);

    // Change the contents of the doc
    let receiptDoc = DocumentApp.openById(receiptDocFile.getId());
    let textBody = receiptDoc.getBody();

    let today = new Date();
    let dd = String(today.getDate()).padStart(2, '0');
    let mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
    let yyyy = today.getFullYear();
    today = mm + '/' + dd + '/' + yyyy;

    let color = (this.color == "Don't care") ? "Random color" : this.color;
    let description = `${ color } PLA ${ this.materialUsage } grams, ${ this.printTime } print`;

    textBody.replaceText("!date issued!", today)
        .replaceText("!id!", this.id)
        .replaceText("!client name!", this.name)
        .replaceText("!price!", this.price.toFixed(2))
        .replaceText("!description!", description);

    // Apply changes
    receiptDoc.saveAndClose();

    return receiptDocFile.getAs("application/pdf");
  }

  sleep(milliseconds) {
    const date = Date.now();
    let currentDate = null;
    do {
      currentDate = Date.now();
    } while (currentDate - date < milliseconds);
  }
}
