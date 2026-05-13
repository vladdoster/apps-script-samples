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
 * Triggered when the spreadsheet is edited. Monitors status column changes
 * from "pending" to "started" and processes the workflow.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEditStatusChange(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // Get header row to find column positions
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("status") + 1;

  // Check if the edited cell is in the status column
  if (col !== statusCol || row === 1) {
    return;
  }

  const newValue = range.getValue();
  const oldValue = e.oldValue;

  // Check if status changed from "pending" to "started"
  if (oldValue !== "pending" || newValue !== "started") {
    return;
  }

  // Find column indices
  const reviewerEmailCol = headers.indexOf("reviewer_email") + 1;
  const taskUrlCol = headers.indexOf("task_url") + 1;
  const l0EmailCol = headers.indexOf("L0_email") + 1;

  // Extract values from the current row
  const reviewerEmail = sheet.getRange(row, reviewerEmailCol).getValue();
  const taskUrl = sheet.getRange(row, taskUrlCol).getValue();
  const l0Email = sheet.getRange(row, l0EmailCol).getValue();

  // Log extracted values
  console.log("Reviewer Email:", reviewerEmail);
  console.log("Task URL:", taskUrl);
  console.log("L0 Email:", l0Email);

  // Look up reviewer in the "Reviewers" sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewersSheet = spreadsheet.getSheetByName("Reviewers");

  let meetLink = null;

  if (reviewersSheet) {
    const reviewersData = reviewersSheet.getDataRange().getValues();
    const reviewersHeaders = reviewersData[0];

    const emailColIndex = reviewersHeaders.indexOf("email");
    const autoMessageColIndex = reviewersHeaders.indexOf("autoMessage");
    const meetLinkColIndex = reviewersHeaders.indexOf("meet_link");

    // Search for matching reviewer email
    for (let i = 1; i < reviewersData.length; i++) {
      if (reviewersData[i][emailColIndex] === reviewerEmail) {
        const autoMessage = reviewersData[i][autoMessageColIndex];

        if (autoMessage === "true" || autoMessage === true) {
          meetLink = reviewersData[i][meetLinkColIndex];
          console.log("Meet Link:", meetLink);
        } else {
          console.log("No meeting link is set");
        }
        break;
      }
    }

    // If no matching email found
    if (
      meetLink === null &&
      !reviewersData
        .slice(1)
        .some((row) => row[emailColIndex] === reviewerEmail)
    ) {
      console.log("No meeting link is set");
    }
  } else {
    console.log("No meeting link is set");
  }

  // Construct and send POST request
  /** @type {{reviewer_email: any, task_url: any, L0_email: any, meet_link?: any}} */
  const payload = {
    reviewer_email: reviewerEmail,
    task_url: taskUrl,
    L0_email: l0Email,
  };

  // Add meet_link if available
  if (meetLink) {
    payload.meet_link = meetLink;
  }

  const url = "https://example.com/api/webhook";
  /** @type {GoogleAppsScript.URL_Fetch.URLFetchRequestOptions} */
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(url, options);
}
// [END apps_script_triggers_status_change]
