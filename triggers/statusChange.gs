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
 * The event handler triggered when editing the spreadsheet.
 * Detects when a cell in the "status" column changes from "pending" to "started".
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEditStatusChange(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // Get the column headers
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColIndex = headers.indexOf("status") + 1;

  // Check if the edited cell is in the "status" column
  if (col !== statusColIndex || statusColIndex === 0) {
    return;
  }

  const newValue = range.getValue();
  const oldValue = e.oldValue;

  // Check if status changed from "pending" to "started"
  if (oldValue === "pending" && newValue === "started") {
    handleStatusChange(sheet, row, headers);
  }
}

/**
 * Handles the status change by logging relevant information.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet where the change occurred.
 * @param {number} row The row number that was edited.
 * @param {string[]} headers The column headers array.
 */
function handleStatusChange(sheet, row, headers) {
  // Find column indices
  const reviewerEmailCol = headers.indexOf("reviewer_email") + 1;
  const taskUrlCol = headers.indexOf("task_url") + 1;
  const l0EmailCol = headers.indexOf("L0_email") + 1;

  // Get values from the row
  const reviewerEmail =
    reviewerEmailCol > 0
      ? sheet.getRange(row, reviewerEmailCol).getValue()
      : "";
  const taskUrl =
    taskUrlCol > 0 ? sheet.getRange(row, taskUrlCol).getValue() : "";
  const l0Email =
    l0EmailCol > 0 ? sheet.getRange(row, l0EmailCol).getValue() : "";

  // Log the values
  console.log("Reviewer Email:", reviewerEmail);
  console.log("Task URL:", taskUrl);
  console.log("L0 Email:", l0Email);

  // Get the Reviewers sheet and look up the meeting link
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewersSheet = spreadsheet.getSheetByName("Reviewers");

  if (!reviewersSheet) {
    console.log("No meeting link is set (Reviewers sheet not found)");
    return;
  }

  // Get all data from Reviewers sheet
  const reviewersData = reviewersSheet.getDataRange().getValues();
  const reviewersHeaders = reviewersData[0];
  const emailColIndex = reviewersHeaders.indexOf("email");
  const meetLinkColIndex = reviewersHeaders.indexOf("meet_link");

  if (emailColIndex === -1 || meetLinkColIndex === -1) {
    console.log(
      "No meeting link is set (required columns not found in Reviewers sheet)",
    );
    return;
  }

  // Search for matching email
  let meetLink = null;
  for (let i = 1; i < reviewersData.length; i++) {
    if (reviewersData[i][emailColIndex] === reviewerEmail) {
      meetLink = reviewersData[i][meetLinkColIndex];
      break;
    }
  }

  if (meetLink) {
    console.log("Meeting Link:", meetLink);
  } else {
    console.log("No meeting link is set");
  }
}
// [END apps_script_triggers_status_change]
