# Automate reminder via mail using VBA

## OVERVIEW
This VBA script reads one excel sheet and, for every row, sends a reminder email in case the date contained in two columns is past and an email has not yet been sent.
The script in this version can only be manually executed, even though the purpose is to have it scheduled. This can be done using the VBA Workbook_Open() event, combined with the OS task scheduler.


## SPECIFICALLY
For every row,
There should be two date columns, if they are past it checks two columns in which the status ("sent"/"") is stored, one for each deadline. 

In case the status is blank, it sends the email to the script-sculpted address with:
  - subject: the header (first row) of the date column.
  - body: the one contained in a column.  
And updates the status to "sent".

Otherwise, it goes on.



## THE CODE
The script is contained in the file script.txt. I highlighted with the "' Edit" tag the points that may require adaptation.
