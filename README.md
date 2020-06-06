# TAIGridControlDLL 

## Lonnie Allen Watson
## May 6th 2020
 
 A Little Back Story from early 2000's

 Being sick and tired of the crappy implimentation of the grid control of VB.net
 I developed this grid to allow easier programatic access, The current databound grid
 works better if its used in a databound way, If you want to use the grid in a manner simillar
 to the way it was used under VB6 you are out of luck. This grid will even expose doubleclick
 events directly on the cell ( Wow what a concept )

### Version 2.0.6.0
- Augmented the new **TAIGridControl.TaiGridColContentTypes GetColumnType(int ColNumber)** to
differentiat Numbers between integer types and floating point types. 
Adding to the ENUM WholeNumber and FloatingPointNumber
  
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
