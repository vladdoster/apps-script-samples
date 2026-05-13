# Status Update Trigger

This sample demonstrates how to use the `onEdit` trigger to respond to changes in a Google Sheets spreadsheet. Specifically, it monitors the "status" column and triggers an action when the value changes from 'pending' to 'started'.

## Setup

1. Create a Google Sheets spreadsheet with the following columns:
   - `status`
   - `reviewer_email`
   - `task_url`
   - `L0_email`

2. Create a second sheet named "Reviewers" with the following columns:
   - `email`
   - `meet_link`

3. Open the script editor (Extensions > Apps Script)

4. Copy the code from `Code.gs` into the script editor

5. Save the project

## Usage

When you edit a cell in the "status" column and change it from 'pending' to 'started', the script will automatically:

1. Detect the change
2. Read the values from the same row's `reviewer_email`, `task_url`, and `L0_email` columns
3. Look up the reviewer's `meet_link` from the "Reviewers" sheet based on their email
4. Log all these values to the console

To view the logs:
1. In the Apps Script editor, go to Executions
2. Click on a completed execution to see the console output

## How it Works

The `onEdit` trigger is a simple trigger that runs automatically whenever a user edits a cell in the spreadsheet. The function:

1. Gets the edited range and checks if it's in the "status" column
2. Compares the old value (`e.oldValue`) with the new value (`e.value`)
3. If the change is from 'pending' to 'started', it retrieves the other column values
4. Calls `getMeetLink()` to look up the reviewer's meeting link in the "Reviewers" sheet
5. Logs the information to the console (or "No meeting link is set" if not found)

## Notes

- Simple triggers like `onEdit` run without requiring authorization for basic operations
- The trigger runs in the context of the user making the edit
- For more complex operations, you may need to use an installable trigger
- If the "Reviewers" sheet doesn't exist or the email isn't found, a message will be logged accordingly
