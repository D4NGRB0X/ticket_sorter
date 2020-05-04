# ticket_sorter
automated ticket sorting for support desk

This is to make my life easier for monthly service ticket reporting.

It is a cli tool that breaks down an xlsx file by client and outputs to .txt file.

user is asked for client name
new file name
file path to xlsx files

currently if the dir is empty or does not contain an xlsx the output txt will be empty

TODO: File pathing
    - check for xls files in dir
      - Return error if files not found
TODO: Output
    - total hours for day per client
    - label day and date
TODO: distribution
    - Package as executable
