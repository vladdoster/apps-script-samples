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
 * Detects when the "status" column changes from "pending" to "started"
 * and logs related row data.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const column = range.getColumn();

  // Get the header row to find column positions
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find the column index for "status"
  const statusColumnIndex = headers.indexOf("status") + 1; // +1 because columns are 1-indexed

  // Check if the edited cell is in the status column
  if (column === statusColumnIndex && row > 1) {
    // row > 1 to skip header
    const newValue = range.getValue();
    const oldValue = e.oldValue;

    // Check if status changed from "pending" to "started"
    if (oldValue === "pending" && newValue === "started") {
      // Get column indices
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
      console.log("Status changed to started:");
      console.log("Reviewer Email:", reviewerEmail);
      console.log("Task URL:", taskUrl);
      console.log("L0 Email:", l0Email);
    }
  }
}
// [END apps_script_triggers_status_change]
