# Sheets to Calendar Integration

Uses Google Apps Script to sync Google Sheets with Google Calendar.
Refer to sample spreadsheet: https://docs.google.com/spreadsheets/d/17QKC5xU3nU6wocf1msv85kCoSHghepim7koUYiemxWk/edit?usp=sharing

Setting Up The Spreadsheet:
- Create a new sheets document using the format in the sample sheet above
- Go to "Tools" > "Script Editor" to replace the code with the Code.gs script in this repository
- Create a Google calendar and replace 'cal' in line 27 of the script with your calendar ID
- Fill out the name, start time, and end time columns (these are the only required columns)
- The start time and end time must be in proper datetime format

Adding/Updating/Deleting Events:
- When an event is added, an event ID should be created and added to column K
- To delete an event or leave it uncreated, set the "Leave off Calendar?" cell to "Yes". This will remove the event ID from the    spreadsheet and calendar.
- To update to event, simply make the updated change on the spreadsheet
- Run the code in "Tools" > "Script Editor" > "Run Function" to fire changes from the spreadsheet to the linked calendar

