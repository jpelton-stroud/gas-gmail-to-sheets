# gas-gmail-to-sheets

A template script that pulls CSV or XLSX files from Gmail, then updates a Google Sheet

## Configuration

### Add your project id

Open `.clasp.json` and paste your script's `scriptId` in the field of the same name.

### The `CONFIG` object

Config happens primarily in the `CONFIG` object starting at `index.ts:10`:

- set `SHEET_NAME` to the name of the _target_ sheet that is to be updated;
- set `SENDER` to the email address from which the files come;
- set `FILE_TYPE` to either `.csv` (default) or `.xlsx`, depending on which file type is being sent to you;
- set `QUERY_SUBJECT` to the subject line for the emails (the more specific you can be here, the better. the goal is to pull only a single email at run time).

### Overwrite or Append?

Beyond the `CONFIG` object, you must tell the script whether to overwrite the data on the target sheet on each run, or append data to the sheet.
The default is to overwrite; if you wish to append, then comment out `index.ts:53` and uncomment `index.ts:54`.

### Set a trigger

The last config step is to go to script.google.com and set up a time-based trigger in your project that corresponds to how often you get the email.
