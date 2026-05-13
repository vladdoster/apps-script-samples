/**
 * Copyright Google LLC
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

// [START apps_script_triggers_status_change]
/**
 * @typedef {Object} ChangeEvent
 * @property {GoogleAppsScript.Spreadsheet.Range} range The range that was edited.
 * @property {GoogleAppsScript.Spreadsheet.Spreadsheet} source The spreadsheet.
 */

/**
 * @typedef {Object} EditEvent
 * @property {GoogleAppsScript.Spreadsheet.Range} range The range that was edited.
 * @property {string} oldValue The old value of the cell.
 * @property {string} value The new value of the cell.
 * @property {GoogleAppsScript.Spreadsheet.Spreadsheet} source The spreadsheet.
 */

/**
 * Trigger 1: Auto-updates status from "pending" to "started" when a new row is added.
 * This function should be set as an installable onChange trigger.
 * It checks if a new row has "status" = "pending" and if the reviewer_email
 * matches a row in the "Reviewers" sheet with "autoMessage" = true.
 *
 * @param {ChangeEvent} e The onChange event object.
 * @see https://developers.google.com/apps-script/guides/triggers/events#change
 */
function onRowAdded(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // Only process if this is the main data sheet (not "Reviewers")
  if (sheet.getName() === "Reviewers") {
    return;
  }

  // Get the edited row
  const row = range.getRow();
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  // Find column indices
  const statusColIndex = headers.indexOf("status");
  const reviewerEmailColIndex = headers.indexOf("reviewer_email");

  if (statusColIndex === -1 || reviewerEmailColIndex === -1) {
    console.log("Required columns not found: status or reviewer_email");
    return;
  }

  // Get values from the row
  const rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  const status = rowData[statusColIndex];
  const reviewerEmail = rowData[reviewerEmailColIndex];

  // Check if status is "pending"
  if (status !== "pending") {
    return;
  }

  // Look up reviewer in the "Reviewers" sheet
  const reviewersSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reviewers");
  if (!reviewersSheet) {
    console.log("Reviewers sheet not found");
    return;
  }

  const reviewersData = reviewersSheet.getDataRange().getValues();
  const reviewersHeaders = reviewersData[0];
  const emailColIndex = reviewersHeaders.indexOf("email");
  const autoMessageColIndex = reviewersHeaders.indexOf("autoMessage");

  if (emailColIndex === -1 || autoMessageColIndex === -1) {
    console.log("Required columns not found in Reviewers sheet");
    return;
  }

  // Find matching reviewer
  for (let i = 1; i < reviewersData.length; i++) {
    const reviewerRow = reviewersData[i];
    if (
      reviewerRow[emailColIndex] === reviewerEmail &&
      reviewerRow[autoMessageColIndex] === true
    ) {
      // Update status to "started"
      sheet.getRange(row, statusColIndex + 1).setValue("started");
      console.log(
        `Auto-updated status to "started" for reviewer: ${reviewerEmail}`,
      );
      return;
    }
  }
}
// [END apps_script_triggers_status_change]

// [START apps_script_triggers_status_transition]
/**
 * Trigger 2: Handles status changes from "pending" to "started".
 * This function should be set as an installable onEdit trigger.
 * It extracts data from the row and sends a POST request with the information.
 *
 * @param {EditEvent} e The onEdit event object.
 * @see https://developers.google.com/apps-script/guides/triggers/events#edit
 */
function onStatusChange(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // Only process if this is the main data sheet (not "Reviewers")
  if (sheet.getName() === "Reviewers") {
    return;
  }

  // Only process single cell edits
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    return;
  }

  const row = range.getRow();
  const col = range.getColumn();
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  // Check if the edited column is "status"
  const editedColumnName = headers[col - 1];
  if (editedColumnName !== "status") {
    return;
  }

  // Check if the value changed from "pending" to "started"
  const oldValue = e.oldValue;
  const newValue = e.value;

  if (oldValue !== "pending" || newValue !== "started") {
    return;
  }

  // Extract data from the row
  const rowData = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  const reviewerEmailIndex = headers.indexOf("reviewer_email");
  const taskUrlIndex = headers.indexOf("task_url");
  const createdAtIndex = headers.indexOf("created_at");
  const l0EmailIndex = headers.indexOf("L0_email");

  if (
    reviewerEmailIndex === -1 ||
    taskUrlIndex === -1 ||
    createdAtIndex === -1 ||
    l0EmailIndex === -1
  ) {
    console.log("Required columns not found in data sheet");
    return;
  }

  const reviewerEmail = rowData[reviewerEmailIndex];
  const taskUrl = rowData[taskUrlIndex];
  const createdAt = rowData[createdAtIndex];
  const l0Email = rowData[l0EmailIndex];

  // Log extracted values
  console.log('Status changed from "pending" to "started"');
  console.log(`Reviewer Email: ${reviewerEmail}`);
  console.log(`Task URL: ${taskUrl}`);
  console.log(`Created At: ${createdAt}`);
  console.log(`L0 Email: ${l0Email}`);

  // Look up meet_link in the "Reviewers" sheet
  const reviewersSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reviewers");
  let meetLink = null;

  if (reviewersSheet) {
    const reviewersData = reviewersSheet.getDataRange().getValues();
    const reviewersHeaders = reviewersData[0];
    const emailColIndex = reviewersHeaders.indexOf("email");
    const meetLinkColIndex = reviewersHeaders.indexOf("meet_link");

    if (emailColIndex !== -1 && meetLinkColIndex !== -1) {
      for (let i = 1; i < reviewersData.length; i++) {
        const reviewerRow = reviewersData[i];
        if (reviewerRow[emailColIndex] === reviewerEmail) {
          const meetLinkValue = reviewerRow[meetLinkColIndex];
          if (meetLinkValue && meetLinkValue !== "") {
            meetLink = meetLinkValue;
            console.log(`Meet Link: ${meetLink}`);
          }
          break;
        }
      }
    }
  }

  if (!meetLink) {
    console.log("No meeting link is set");
  }

  // Construct and send POST request
  /** @type {{reviewer_email: any, task_url: any, created_at: any, L0_email: any, meet_link?: any}} */
  const payload = {
    reviewer_email: reviewerEmail,
    task_url: taskUrl,
    created_at: createdAt,
    L0_email: l0Email,
  };

  if (meetLink) {
    payload.meet_link = meetLink;
  }

  // Replace with your actual webhook URL
  const webhookUrl = "https://example.com/webhook";

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  console.log("POST request sent with payload:", JSON.stringify(payload));
}
// [END apps_script_triggers_status_transition]

// [START apps_script_triggers_install_status_triggers]
/**
 * Creates installable triggers for status change monitoring.
 * Run this function once to set up the triggers.
 *
 * @see https://developers.google.com/apps-script/guides/triggers/installable
 */
function createStatusChangeTriggers() {
  const ss = SpreadsheetApp.getActive();

  // Create onChange trigger for new rows (Trigger 1)
  ScriptApp.newTrigger("onRowAdded").forSpreadsheet(ss).onChange().create();

  // Create onEdit trigger for status transitions (Trigger 2)
  ScriptApp.newTrigger("onStatusChange").forSpreadsheet(ss).onEdit().create();

  console.log("Status change triggers created successfully");
}
// [END apps_script_triggers_install_status_triggers]
