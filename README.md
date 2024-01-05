# Automate reminder via mail using VBA

## REQUIREMENTS
- An xlsm workbook
- VBA add-in
- From VBA, add as a requirement Outlook 16.0


## OVERVIEW
This VBA script reads some excel sheets and, for every row, sends a reminder email in case the date contained in other columns is past and an email has not yet been sent.

## SPECIFICALLY
For every row,
There should be two date columns, if they are past it checks two columns in which the status ("sent"/"") is stored, one for each deadline. 

In case the status is blank, it sends the email to the script-sculpted address with:
  - subject: the header (first row) of the date column.
  - body: the one contained in a column.  
And updates the status to "sent".

Otherwise, it goes on.

## THE CODE
The project consists of two scripts:
- first_run.txt: Updates the cell containing the status with "sent", in order to set as not to send the reminders before today.
                 There should be one version per sheet. I gave a generic one that manages dates in three columns 'T', 'X' and 'Y'.
- wbOpen_script: Gets executed at the opening of the work book and sends the emails for two sheets 

I highlighted with the "' Edit" tag the points that may require adaptation.

Also, given the requirements of the project all the setups are static. With some adaptation and setup design can be made dynamic.
