# TAIGridControlDLL 

## Lonnie Allen Watson
## May 6th 2020
 
 A Little Back Story from early 2000's

 Being sick and tired of the crappy implimentation of the grid control of VB.net
 I developed this grid to allow easier programatic access, The current databound grid
 works better if its used in a databound way, If you want to use the grid in a manner simillar
 to the way it was used under VB6 you are out of luck. This grid will even expose doubleclick
 events directly on the cell ( Wow what a concept )

### Version 2.0.9.0
- Holding the CONTROL key down while using the mousewheel will not adjust the font size of the
items in the grid as opposed to scrolling up and down or left and right through the grid

### Version 2.0.8.0
- Excel output now defaults to naming the file with date and time to avoid colissions
EXCELOUTPUT_MMDDYYYY_HHMMSS.xlsx is the default name now for the excel output files
Ths avoids colissions with file locks and prevents overwriting old exports accidentally

### Version 2.0.7.0
- Change the Excel output dialog to contain a new feature OPEN FILE WHEN SAVED option that is Checked by default
This will attempt to execute the file that was just saved to the users system. Resulting in that file being opened 
in the registered application for XLXS files. (Usually thats Excel but can be other things like Libra Office for example)

- Removed the 6 things in the SORT context menu and now just has ASCENDING and DESCENDING. Will attempt to Sort dates first
the numbers and finally ASCII in that order. So as to not be confusing as it was with seperate sort options for
dates and number as well as ASCII text..

### Version 2.0.6.0
- Augmented the new **TAIGridControl.TaiGridColContentTypes GetColumnType(int ColNumber)** to
differentiat Numbers between integer types and floating point types. 
Adding to the ENUM WholeNumber and FloatingPointNumber

- change the behavior of the columncheck to favor string if everything in the column is empty used to default to date.
  
### Version 2.0.5.0
- Added two Methods **TAIGridControl.TaiGridColContentTypes GetColumnType(int ColNumber)** and 
**TAIGridControl.TaiGridColContentTypes GetColumnType(String ColName)** Will return from the newly added
enum **TaiGridColContentTypes**  String, Number or Date

- more code refactoring


### Version 2.0.4.0 
- Change the internal handleing of menu button (Right mousebutton) tracking and cacheing

- Removed redundant handlers for Mouse MOVE,BUTTONDOWN, BUTTONUP handlers

- Fixed issues with Sorting ASCII, DATE and NUMERIC

- Improved Context menu descruction and rehydration

### Version 2.0.3.0 Minor change to Version number scheme
 
### Version 2.0.2.0 Fixed a wierd error with Cellfonts and autosizing cells to contents

### Version 2.0.0.1
- Implemented excel file output using ClosedXML (Which itself wraps DocumentFormat.OpenXML)
 grid no longer COM interops with Excel for EXCEL output

- Added dialog to on Exporting to excel for selection of filename and worksheet name
 as well as selection of Null omission on resulting excel file.

### Version 2.0.0.0 First Version in C#
