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

// [START apps_script_triggers_status_update]
/**
 * Triggers when a cell is edited in the spreadsheet.
 * Checks if the "status" column changed from 'pending' to 'started'.
 * If so, logs the reviewer_email, task_url, and L0_email for that row.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e) {
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

  const newValue = e.value;
  const oldValue = e.oldValue;

  // Check if status changed from 'pending' to 'started'
  if (oldValue === "pending" && newValue === "started") {
    // Find the column indices for the required fields
    const reviewerEmailCol = headers.indexOf("reviewer_email") + 1;
    const taskUrlCol = headers.indexOf("task_url") + 1;
    const l0EmailCol = headers.indexOf("L0_email") + 1;

    // Get the values from the row
    const reviewerEmail =
      reviewerEmailCol > 0
        ? sheet.getRange(row, reviewerEmailCol).getValue()
        : "Column not found";
    const taskUrl =
      taskUrlCol > 0
        ? sheet.getRange(row, taskUrlCol).getValue()
        : "Column not found";
    const l0Email =
      l0EmailCol > 0
        ? sheet.getRange(row, l0EmailCol).getValue()
        : "Column not found";

    // Log the values
    console.log("Status changed from pending to started:");
    console.log(`Reviewer Email: ${reviewerEmail}`);
    console.log(`Task URL: ${taskUrl}`);
    console.log(`L0 Email: ${l0Email}`);
  }
}
// [END apps_script_triggers_status_update]
