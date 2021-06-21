# gsheet_notifier 1.0

## Notify users of changes to a G-Sheet

1. Add this script to your G-Sheet from the 'Tools>Script Editor' command.
2. Update the recipients list, and save.
3. Run the 'resetSheetUpdate' script to initialize the script properties.
4. Add two triggers to call the script:
  1. An On-change trigger, that alls the 'receiveUpdate' function, and
  2. An Time-based trigger, set for say 6 hourly, that calls the 'processUpdates' function.

## In Use

When a change is made a cell on the sheet to which this script is attached, 
the changed cell is added to a list of changes.

When the 'processUpdates' function runs, it amalgamates the changes, and sents a GMail
containing snippets of the sheet, with the changed cells highlighted in light green.
