# Description
This is Python 3 script that will fix the hyperlinks in Excel file (xlsx).

After unexpected shut down of Excel, if you use AutoSaved files to recover the previously saved file you will often find that all the previously working hyperlinks don't work. The AutoSave is changing the hyperlinks to reference the "../../../AppData/Roaming/Microsoft/Excel" folder. With this script you can easily fix them to their original state.

Before starting the script the user should set the limits of the Excel file to be searched for hyperlinks in order to reduce the unnecessary computation.  It is possible to limit number of sheets, number of columns and number of rows, for example:

sheets = 1

columns = 8

rows = 200

This will affect only Sheet1 from 1 to 8th column and from 1st to 200th row.

After starting the script the user will be asked to browse to the file to be changed.

The fixed file will be saved in the same directory as original one with the same name extended with “_fixed”.

