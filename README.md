## Domain Filtering Tools  - Google Sheets Apps Script
 
Custom script for Google Sheets. Filter domain-related cells based on colour, tld extension, etc.

Functions transform cell tables to filtered single-column lists. 

### Installation

In Google Sheets, navigate to *Tools > Script Editor*. Copy the *FilteringTools.js* script to a new file and save. 
Re-load your Sheet. Grant the script permission from your Google account, and a custom "Filtering Tools" menu will appear.

### Functions available

- `getDotComs`
- `getCoUks`
- `getColouredCells`
- `getBlueCells`
- `tableToHTML`
- `tableToList`

### Usage 

Execute functions via their buttons on the "Filtering Tools" menu. Script binds to the Google Sheet it is installed on.
