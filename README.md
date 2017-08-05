# fogbugz-importer

Manual import from https://code.google.com/archive/p/fogbugz-importer/

## Original Documentation

Automatically imports tickets (cases) and their history from Excel spreadsheet into FogBugz using FogBugz XML API

FogBugz Importer reads tickets (cases) data from an Excel spreadsheet and automatically imports that data into FogBugz using FogBugz XML API.

You can prepare a spreadsheet with data by hand or with some tool. If you migrating from Unfuddle then you may want to check Unfuddle Backup Parser. That parser will process your backup file and prepare spreadsheet for you.

Usage:
FogBugzImporter.exe FogBugzApiUrl "Your Name" "Your Password" "X:\Path\To\tickets.xlsx"

Example:
FogBugzImporter.exe https://your-account.fogbugz.com/api.asp "John Doe" "Strong password" "C:\Users\John Doe\Documents\tickets.xlsx"

NOTE:

You should create users, projects, areas, milestones in FogBugz before import.
You may want to tune up workflow in FogBugz before import if you need non-standard workflow.
There should be folder media next to tickets.xlsx even if you don't plan to import attachments to tickets.
The format of the tickets.xlsx is as follows (you may download sample tickets.xlsx using link on the right of this page or in Downloads section):

- number of rows is unrestricted.
- first row is reserved (used as header) - data in it won't be imported.
- data from first 12 columns (A to L, inclusive) will be processed.
- column A (cmd) contains one of the FogBugz commands (new, edit, assign, reactivate, reopen, resolve or close). An empty cell in this column will cause stop of processing.
- column K contains reporter (author) name. This column should not contain empty cells.
- each other cell (B, C, D, E, F, G, H, I, J and L) can be empty.
- DateTime values in columns B (dt) and J (dtDue) should be in UTC format.
- values in columns C (sProject), D (sArea) and E (sFixFor) should correspond to existing (created in FogBugz) project, area and milestone names.
- values in columns I (sPersonAssignedTo) and K (reporter) should correspond to existing (created in FogBugz) user names.
- attached files are presented as string with format attachment-name;attachment-file-name. attachment-file-name should point to a file in media folder next to tickets.xlsx.
