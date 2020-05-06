Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.Collections
Imports System.Collections.Generic

' 
' TAIGRIDcontrol.vb
' Lonnie Allen Watson
' Jan 17, 2003
'
' Being sick and tired of the crappy implimentation of the grid control of VB.net
' I developed this grid to allow easier programatic access, The current databound grid
' works better if its used in a databound way, If you want to use the grid in a manner simillar
' to the way it was used under VB6 you are out of luck. This grid will even expose doubleclick
' events directly on the cell ( Wow what a concept )
' 
' Jan 23, 2003  Added functionality to take a sqlconnection and a sqlstring and autopopulate the grid
'               with the contents of a recordset of that data
'               Overloaded with three formats 
'               PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Gridfont As Font)
'               Public Sub PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String)
'               Public Sub PopulateGridWithData(ByVal cn As SqlClient.SqlConnection, ByVal sql As String)
'
' Jan 23, 2003  Added functionality to auto size a grid to its contents
'
' Jan 23, 2003  Added functionality to take a cell or a column and format its contents with
'               a special case being for money
'
' Feb 28, 2003  Added an overload for PopulateGridWData to accept an actual Connection that is already open
'
' Feb 28, 2003  Added an overload for PopulateGridWData to accept  an open connection and a color to place on
'               the rendered text
'
' Mar 15, 2003  Added Some overloads for PopulateGridWData to accept a font for the grid, and the color for the font
'
' Mar 17, 2003  Added a public method to sum up the contents of a column and return it as a numeric
'
' Mar 23, 2003  Added some error precautions to all PopulateGridWdata methods to handle timeouts better
'               Increased Timeouts to 500 seconds. Need to do a parameter for this that can be read
'               and written to for the class, allowing the external program to configure.
'
' Mar 26, 2003  Added the ability to key up and down on a grid to move the mouse selected row
'               added the ability to pageup and pagedown on a grid to move the mouseselected row
'               Added the ability to hit enter or return to fire a doubleclicked event on the 
'               currently selected mouserow
'               Added two events RowSelected and RowDeSelected to be fired whenever the mouserow
'               is altered via keyboard or the mouse
'               Added a DataBaseTimeOut property to tied its value to all database access functions default to 500 sec
'               Added a paginationsize property for paging up and down on the grid default to 10 lines
'               Added code to handle resizeing the grid more elegantly it should suffice for most uses
'
' Apr 03, 2003  Added SelectedRow property that tracks Mouserow DOH! 
'               Added event calls to the populate methods so that the alteration of the grid would
'               be notified to external clients.
'
' May 07, 2003  Fixed a stupid bug in sum up column that made it handle negative numbers improperly
'
' May 12, 2003  Added new event PartialSelection and a new property MaxRowsSelected
'               This will allow programatic setting of a maximum number of items to select with
'               any of the PopulateGridWData calls. Default is 0 = disabled
'               on list truncation the event PartialSelection is fired for notiofication purposes
'
' May 13, 2003  Various bug fixes around yesterdays additional functionality
'
' May 21, 2003  Fixed a state bug with the selected row not refreshing properly
'
' Jun 08, 2003  Fixed some internal rendering bugs
'               Header now properly shown on manual inserts of data as well as populategridwithdata calls
'               Radical adjustments in fontsize on allcellsusethisfont will now scroll to visible region properly
'               added an overload for AlternateRowColoration() call to accept no parameters will use default or last
'               added new subroutine PopulateGridFromArray takes a 2 dimensional string array and populates the grid
'                   with it. The first row of the array is the header text
'
' Jun 09, 2003  Had an epiphany today "Why not range check the cell drawing operations and not render the ones
'               that are not visible?" ... DOH! its much faster now...
'               Added three overloads for PopulateGridFromArray() array can be Strings, Integers, Longs, and Doubles
'               Added 8 more overloads for all the PopulateGridFromArray() calls with font and forecolor like
'               the overloads for PopulateGridWithData() calls
'
' Jun 10, 2003  Optimized the range checking operations to more quickly find the cells that are to be drawn
'               Added two new functions GetRowAsString() hand it a rowid and you get a pipe delimited string
'               perfect for the split function to blow it into an array
'               GetGridAsArray() will return a two dimensional array of the size of the current grid filled with
'               the current grids contents. Hand it to another grid and you will get a copy operation.
'               Added a TAIGCanvas_MouseEnter() event handler to focus this control for MouseWheel Functionality
'               
' Jun 17, 2003  Fixed a simple bug in the keydown handler that was not fireing the rowdeslected event properly when the
'               user scrolled off the botom of the grid with the down arrow.
'
' Jun 18, 2003  Added a method to remove a column from the grid entirely
'               altered the COLWIDTH and ROWHEIGHT calls to accept a 0 value or any value on a write
'               this allows you to hide a column from display by setting its width to 0
'               or hide a row by setting its height to 0 (Note setting a colwidth or rowheight to anything
'               turns off _Autosizecellstocontents)
'               Implimented CellOutline property to either draw or not draw bounding boxes on individual cells
'
' Jun 30, 2003  After a number of fits and starts finally go the control to support moving the highlite row
'               programatically set SelectedRow to whatever row you want and call DoSelectedRowHighLite()
'               Fixed a bug where radically altering the grids dimensions via PopulateGridFromArray calls would
'               Leave the thing in a wacked out state with the keyboard handler object visible at the bottom of a 
'               large empty area where grid items used to be
'
' Jul 01, 2003  Added an overload for DoSelectedRowHighlite() to accept the row to highlite
'               Added functionality to keep the selected row active with global font changes to the grid via AllCellsUseThisFont()
'               Fixed a bug that caused hidden columns to render their contents even though the column was supposed to be hidden
'               this also extended to hidden rows. Would only do so on the last row or column but was a bug none the less
'               Fixed another bug where if the grid was not in autosize mode then it would not seemingly draw properly if there
'               were more than the height rows.
'
' Jul 02, 2003  Added 4 new populate calls OLEPopulateGridWithData() that mirror the 4 regulare POPULATEGRIDWITHDATA calls that work 
'               with SQL server. These new ones work on access data sources.. Jeff needed this for the Butler Project
'               Also added 4 new calls to the same PopulateGridWithData methods that are called SQLPopulateGridWithData for completion
'               of the syntax. The old PopulateGridWithData methods are still there for backwards compatability.
'
' Jul 14, 2003  Added 2 functions ( 1 overload ) GetColumnIDByName(name) and GetColumnIDByName(name,NoUpperCase) to return an integer
'               representing the column ID passed in as name string. If you pass True to the second version then the call with not
'               convert the arguments to uppercase on the compare otherwise it will. The first version always converts to uppercase
'               on the compare.
'
' Jul 17, 2003  At the insistance of our other ANAL developers here at Tidgewell Associates, Inc. All functions and
'               routine calls have been alphabetized in their respective sections now.
'
' Jul 18, 2003  At the request of Jeff added a new public function and one overload FindInCol. Function will return
'               RowID of a search from a stringvalue in a given column number. Overload also allows for case sensativity
'               or no case sensativity. And I even put them in Aplhabetical order...
'               Implemented a potential fix for an obscure bug issue where the grid was doing some drawing of cells that
'               were not yet fully populated resulting in trapped errors that left the drawing surface in a meseed up
'               Ghostbusters like Crossed state (Red X) Basically added some additional checks to ensure no attempt is made
'               to use a NULL brush or Font in a drawing operation
'
' Aug 01, 2003  Added some more command timeout functions to each of the PopulateGrid calls to accomodate slow links to data
'               Added some better error recovery in the PopulateGrid calls to better handle problems
'
' Aug 11, 2003  Added an overload for InitializeGrid that takes the number of rows and cols as arguments. This it is hoped will
'               facilitate the speedup of the populategridfromarray calls as they are seemingly slow on the draw now.
'               Changed the InitializeGrid call used by the sql populators to default all setups in an effort to avoid the
'               problem of the top row on a col from being centered across calls.
'               Added new property OmitNulls as a boolean. when tru will replace NULL values with and empty string into array
'               When false will place {NULL} into array. Default is False to mimic old behavior requested by J. Corrao...
'
' Aug 21, 2003  Added new property COLPASSWORD(idx) as string, use it to set a columns visualization characteristic from the 
'               contained text to whatever string value you assign here. The actual column contents are unaffected by this
'               it is simply a diaplsy trick. Used to shorten extremely long columns like password hashes and what not in an
'               effort to enhance readability.
'
' Sep 09, 2003  Added new Methods to populate grid from ODBC data sources. Will not work on 1.0 framework without 1.0 additions
'               of the ODBC data access mechanisms.
'
' Sep 22, 2003  Fixed a small problem with modality on setting the default forecolor for all subsequent additions to the
'               grid programatically. Stupid on my part has been fixed.
'
' Sep 29, 2003  Added new method PivotPopulate with a load of overloads for formatting, color and font.
'               PivotPopulate(sgrid,xcol,ycol,scol,xxx,xxx,xxx,xxx) takes a source grid and scans Xcol for unique values
'               these values are populated into ME across the top. Then it scans ycol for unique values and populates ME
'               down the side with them. Finally is scans scol for the cross section of the corresponding Xcol and Ycol value
'               Summing the result and populating ME with those results. Operates simillarly to Excels Pivot Table functionality.
'               Usefull in those times when non rectangular results are coming from SQL and you cannot get a simple sql query to
'               perform as expected. ( which is like always )
'
' Oct 21, 2003  Added new methods with overloads to take a specifically formatted string from any source ( Our webmethods for exmple )
'               this string contains as the firt two elements the row and column counts | {pipe} delimited. Then the First row names 
'               of the grids columns | {pipe} delimited. Followed by each row of data also | {pipe} delimited.
'               The standard assortment of overloads are provided with Color and font selection as well as both color and font select.
'
' Nov 07, 2003  Added new property AntiAlias as a boolean True and its on False and its off. Off by default
'               Applies to both Text rendered and Lines rendered ( though lines rendered are at right angles anyways thus show
'               no difference between on and off.
'               Added new Popupmenu on right mousebutton to copy entire grids contents to clipboard. Easily pasted into Excel
'               Added Font Smaller option to popup menu decreases defaultcellfont by 1 point and redisplays grid
'               Added Font Larger option to popul menu increases defaultcellfont by 1 nd redisplays grid
'               Added Smoothing boolean to popup menu Tracks current Antialias setting and allows users to toggle setting
'               Added new boolean property AllowPopupMenu defaults to true that will allow or disallow popup menu from coming up
'
' Nov 18, 2003  Fixed some minor internal stuff around rendering
'
' Dec 10, 2003  Fixed some minor internal stuff around rendering
'
' Jan 21, 2004  Added new methods to RollupColumn and RollupRow
'               Syntax RollupColumn(row,col) will rollup up col+1 to cols on row and return value
'                      RollupRow(row,col) will rollup row+1 to rows onn col and return value
'                      RollupCube(row,col) will rollup row+1 to rows and col+1 to cols and return value
'               Necessary for 24 month columer rollup on Lag tool for State of PA
'
' Feb 04, 2004  Added the export to text file functions
'
' Feb 12, 2004  Fixed a bug on the showing of the title area chopping off the lower portion of the grid
'               Fixed a bug where the PopulateGridWData call might bomb on there being a timestamp in the resulting selection
'               Timestamps convert to Byte arrays and Byte arrays don't cast to strings easily. Implimented the fix on all
'               PopulateGridWData calls uses new internal private function ReturnByteArrayAsHexString
'               Fixed a situation where the grid might try and allocate a bitmap of dimensions that were illegal. The Maximum
'               size of the scroll bitmap is now set to 16766 x 16766 pixels
'
' Feb 16, 2004  Added two new events and placed their raiseevent calls in the appropriate places
'               StartedDatabasePopulateOperation and FinishedDatabasePopulateOperation called when the system starts and finishes
'               the lengthy populategridwdata functions...
'               Fixed a small bug having to do with font size stuff on the font context menu
'               Fixed another bug where partial selections forced via the MaxRowsSelected Property were returning MaxRows-1
'               This affected all the PopulateGridWData calls
'
' Feb 17, 2004  Added new functionality to allow Users to resize column width on their own Requested by Lreedy.
'               Added new Property to allow/disallow this new functionality. Routines affected MouseDown, Mouseup, and Mousemoce 
'               for the canvas subobject. User just clicks and holds on header row and moves mouse side to side to adjust the colwidth.
'               Raises a new event ColumnResized(server,ColID)
'               
'
' Feb 18, 2004  Added new property UserColresizeMinimum. This property defaults to 5 pixels and set the smallest size that a user can 
'               resize a coloumn to. This prevents a user from making a column 0 pixels wide making that column to small to be selected
'               and resized to a larger width.
'
' Feb 18, 2004  Added net Public Method to CopyGridToClipboard.  Removed code from Menuitem1 and placed that code into the CopyGridToClipboard
'               Method.
'
' Mar 04, 2004  Fixed a bug in Columformat and ColumnformatasMoney when invalid colums were handed into them
'
' Mar 11, 2004  Lots of aditions to the included context menu
'               Export to excel, Fixed Alternate row colorization to only color populated columns
'               Fixed the wacky sizing of the rows in excel export
'               Added menu options to send to excel, and adjust alternatecolorizing, and column/row autofiting
'               Added formatters for the grid column contents Format as Money, Number, Text, Left Center and Right justification
'               These new formatters were added to the internal context menu
'               Fixed the cause for the menu to popup in a different place then where the mouse was on right mousebutton
'
' Mar 15, 2004  Added two new Properties LastConnectionString and LastSQLString each will return the last passed in connection or
'               SQL code used for any of the PopulateGridWithData calls.
'
' Mar 24, 2004 Added the ability to hand the PopulateViaWebServiceString method a delimiter of more then one character.  
'              Added a new Public Function called ReturnDelimitedStringAsArray to parse a delimited string 
'                       and return a string array(,) 
'
' Mar 28, 2004  Added some further modifications to what is drawn and when its drawn to try and always ensure that a row and col is
'               drawn at least in part.
'
' Mar 30, 2004  Some further refinements to the drawing process. Overrode the on paint background. Helps reduce flicker
'
' Apr 15, 2004  Added new events toomanyrecords and toomanyfields called when either the autosize check length and width of the resulting
'               bitmap would exceed 16735 pixels and the bitmap is truncated.
'               Also added DefaultForegroundColor property to allow the autopopulators to get the color right
'               Changed the default color for the grids text being added from system.drawing.color.blue to system.drawing.color.black
'
' Jun 07, 2004  Fixed a bug in the Item(x,y) call that resulted in exporting to excel bombing on having a programatic cell populated with
'               no data.
'               Added 4 new subs to POPULATE, ODBCPOPULATE, OLDPOPULATE and SQLPOPULATE w data calls (16 in all ) these calls now will
'               take a datareader of the appropriate type and will populate from the datareader. The call will not close that datareader
'               however since it is basically a firehose cursor you will not be able to reiterate over the data after the grid has
'               done its thing with it...
'               Added two new properties ExcelMatchGridColors and ExcelOutLineCells. These properties allow the export to excel to
'               Attempt to match the colorization of the grid itself. If you set up the grid control with custome colors in certain 
'               Cells the export function will attempt to match those color now. ( If you use any of the aplha blending that setting
'               Will not be carries over to excel as its not supported in excel. The OutlineCells boolean will attempt to outline the
'               exported cells in a solid line ( Usefull as colorizations will remove outlines and make the grid look flat and lifeless )
'               Added two new context menu items to support user mode setting of the above new properties and code to tract these 
'               new properties current settings when the menu pops up.
'
' Jun 08, 2004  Added new exporttoexcel function that takes the excel instance the workbook instance along with the worksheet name
'               this was added to support Beaver county CAP reporter function. Allowing multiple Reports to be sent to a single
'               excel instance.
'
' Jun 16, 2004  Complete rewrite from the ground up
'               Primarily to fix a problem of getting to much data on a populate.
'               Not really anything that went untouched. 
'
' Jun 17, 2004
' Jun 18, 2004
' Jun 19, 2004
' Jun 21, 2004  Continued rewrite and testing as replecement in our projects
'
' Jun 23, 2004  Added new property AllowColumnSelection defaults to true if you turn it off column selection is disabled
'               Added new property AllowControlKeyMenuPopup will disallow the ctrl key forcing of menu to popup even if a cistom
'               context menu has been attached to the control. Defaults to allowing ( TRUE )
'
' Jun 25, 2004  Added SelectedRowForeColor and SelectedRowBackColor and SelectedColForeColor and SelectedColBackColor using FLYWHEEL
'               Fixed a bug where the last row might not be visible using keyboard navigation
'               Added new methods to SetColBackColor and SetColForeColor to add to the ROW ones already in the control
'               Added new Methods GetColAsArrayList and GetRowAsArrayList
'               Fixed a bug where the elected row was being cleared on a mouse wheel
'               Fixed a bug where a populate call might cause the grid to bomb if there were less than the number of cols or rows
'               in the grid after the populate vs before the populate and the viewport was adjusted to these areas before the 
'               populate call was made. (Basically I reset the status of the scrollbars on all InitGrid calls )
'               Grid now gets Focus when the mouse enters it so that the mousewheel will be active in the grid as soon as you
'               hover over the grid
'               Added Default Value Attributes for a ton of the exposed properties using Flywheel
'
' Jul 02, 2004  Added some additional checks in MouseWheel event handler to prevent out of bounds conditions
'
' Jul 08, 2004  Added DeleteRow and DeleteColumn for compatability with CTINVOICER
'               Added AutoFocus property defaults to true
'               Added Enter and Return keyboard handlers
'               Added Left and Right keyboard Handlers
'               Added the tossing of events RowSelected and RowDeSelected in the appropriate places
'               Fixed some issues with Populate grid from array
'
' Jul 24, 2004
' Jul 25, 2004  Two main new things done to improve performance and reduce memory footprint
'               Created a semaphore to prevent automated grid sizing needlessly ( Vastly improved performance )
'               Created a level of redirection for the creation of the grids Fonts, Pens, Stringformats, and Backcolors
'               This has the effect of vastly reducing memory footprint. Example
'               5000 Rows of 58 colums ( Select top 5000 * from claims ) prodiced a grid quiet state of using about 900 megs of PF!
'               That Same setup now uses just < 100 megs
'               Also added a DefaultBackColor property that was logically missing
'               Now putting 100000 rows in the grid is doable and that uses only about 300 mgs of PF memory
'               Also moved scrollbar calculations before intitial render to reduce that pause between progress bar ending and initial
'               render
'
' Jul 26, 2004  Performance enhancements in colorizers to speed up the works
'               Export to excel progress messages implimented
'
'
' Oct 07, 2004  Re-added ColPasswords property aand implimented it back into the rendering engine as an oversite on my part when
'               we re-implimented the grid back in the summer. Larry needed this in his FTP processor
'
' Nov 05, 2004  Fixed two bugs The Default ForgroundColor and DefaultBackgroundColor were setting the wrong internal variables
'               The Col Passwords item was not being used for colum width calculations when AutoSizing the grid to cells contents
'               That has been fixed in this release.
'
' Nov 09, 2004  Added new method SetColorScheme that uses a builtin enumeration TaiGridColorSchemes to set the coloring used
'               in the grid. The schemes need some work yet but the ability is defined in the grid itself now.
'               Fixed a bug where in some cases the last column might not be visable because of the size in fonts increase.
'
' Nov 15, 2004  Added check for the visibility of the grid on raising ceryain mouseevents. CellClicked and CellDoubleClicked.
'               It appares that under some circumstances the grid was fireing a second doubleclick event with the cirrect parameters
'               but if its on a form that is handling the event that is also being hidden then the form will wire the event to the
'               wrong handler. It will also send the wrong handler the wrong parameters. This ofcourse will cause all maanner of
'               weirdness that proved difficult to track down. These changes should prevent that sort of thing from happening.
'
' Nov 18, 2004  Fixed a stupid bug with searching in the grid that jeff found yesterday
'
' Nov 23, 2004  Fixed a bug in sumupcolumn.
'
' Dec 01, 2004  Forced setting of selected row to jump to the top of the grid on a refresh.
'
' Dec 02, 2004  Fixed a search in column issue Jeff uncovered in conjunction with Care Tracker
'               Ctrl F fill now repeate the last search operation from the currently selected row and NOT Beep
'               Show now handle population from array and password cols correctly per Larry.
'
' Dec 03, 2004  Added two new functions and tied them to Menu items under the SORT menu to allow you to sort a column as dates
'               either descending or ascending. Old sort routine is still there but this treats cell contents as textual entities
'               and date sorts often are broken if the dates are not formatted properly first read as date sorts are always broken.
'               new routines handle date sorts differently after first ensuring that the entire column contains data that CAN be
'               converted to a date first. If a column has any data that caannot be converted to a date then a message box will
'               inform the user that date conversion fails for some of the columns data and the sort was aborted.
'
' Dec 06, 2004  Added aa semaphore on mousedoubleclick to prevent the fireing of a mouseup after the doubleclick event is fired. 
'               This was the cause of the grid not tracking the selected row after doubleclicking on that row and having the grid
'               hidden by another form and then reselected after showing that form again. The event model would fire
'               Mouse Down    Mouse DoubleClick     Mouse Up. CTinvoicer woulod select the row after the doubleclick and then
'               the mouse Up would reselect the row after the vert scrollbar was moved to the last selected row. Te effect was if
'               There was no scrollbar or the scrollbar was at the top of the screen the last selected row would re-select. If the
'               Scroll bar was scrolled down at all the selectedrown would jump down by the offset of the vert scrollbar because of
'               the phantom mouseclick. Since the System.Windows.Forms event model has no way to short circuit this characteristic
'               I had to do it in code.
'
' Dec 07, 2004  Massive performance enhancements using REDGATE ANTS profiler tools just purchased.
'               Changes in the AutosizeCheck routine, The Populate from data routines, implemented using the ANTs profile
'               runs as a guidline. Now an order of magnitude faster in typical use.
'
' Dec 29, 2004  Added a boat load of new context menu options to support math on columns and rows. 
'               MAX and MIN, SUM and AVERAGE on cols and rows selected with the mouse menubutton (right menu button)
'               Also added the copy cell to clipboard option to facilitate the grids use as a research tool when used in
'               conjunction with our other tools like Claimsexplorer.
'
' Jan 05, 2005  Set the default for the autofocus property to be False. For some CareTracker issues.
'
' Jan 13, 2005  Configured the two PopulateFromWebService base functions. Now will examine the string for {NULL} note the case
'               if and only if the OmitNulls property is true. If so it will replace these with empty strings to more closely
'               approximate the behaivior of the true populatefromdatabase calls. Implemented for the time tracker application
'               as jeff uncovered this anomaly.
'
' Jan 28, 2005  Larry re-wrote the ExportToExcel(ByVal _excel As Object, ByVal _WorkBook As Object, ByVal wsname As String) method
'               to increase performance.  Also added ReturnExcelColumn function to return the Excel column as a letter.
'
' Feb 17, 2005  Larry Modified export to excel to get around Excel 2003 specifics tested back to office 2000 version of excel.
'
' Mar 15, 2005  Added Populate Grid with Datatable support
'
' Mar 28, 2005  Fixed the ReturnExcelColumn call to properly handle columns over 104, uncovered in the HCQUreporter
'
' Mar 30, 2005  Fixed a issue in the pivot populate call that was mishandling the header labels
'
' Apr 01, 2005  Added the ability to Print the Grid for support of reporting output in THe TAIQuery Reporter and
'               The Report tools being developed for VBH datasets in Beaver County and Fayette County
'               PrintTheGrid(4 Overloads)
'               Print and Preview the grid output in the grids own context menu
'               Properties to control Landscape/Portrait, GridOutlining, Matching Coloration
'               Title and Page Numbering. 
'
' Apr 19, 2005  Fixed a problem with the FormatColumnAsMoney method.
'
' Apr 20, 2005  Fixed a big in the print grid procedure where multiheight rows were not coloring properly
'
' Jun 02, 2005  Added Multiple Selection Capabilities
'               CTRL and click selects SHIFT and click selects range
'               SelectedRows() returns an arraylist of integers for the rows that have been selected
'               Edit Capabilities 
'                   AllowInGridEdits allows or disallows editing globally
'                   ColEditable(colindex) = True or False allows or disallows editing the contents of a given columns data defaults
'                                           to false in all cases
'                   Raises events on changing any cell with CellEdited(sender,row,col,oldval,newval)
'                   The textbox that allows editing can be changed to a different backcolor with the CellBackcolorEdit property
'               New event GridResorted(sender,colindex) will fire whenever the sort operation is called Colindex is the sort column
'               New Method SortGridOnColumn(col,Descending) will sort of col id col, Descending if true Ascending otherwise
'               Will clear the SelectedRow property and the SelectedRows arraylist on sort as well
'               New Method PopulateFromWQL(WQLStatement) Like the other database calls will Populate the grid with the results
'               of the supplied WQL statement
'               Completion of Major re-work on printing. Can print to any installed printer, using any supported page size
'               can adjust the Orientation, print any range of pages or all pages
'
' Jun 03, 2005  Fixed some minor issues with MultiSelect
'               Added new Property ColMaxCharacters(colid) this will allow you to elpisis truncate the display of a given colid
'               without actually truncating the data being stored inside the grid. Its a display trick like the ColPassword(colid)
'               property that already exists
'               Added event calls to GridResorted on user date sorting of the grid as well
'               Added new method RemoveRowsFromGrid(listofrows as Arraylist) Send in an array list of row integers and the grid will
'               Remove them in a single shot
'               Added a new event KeyPressedInGrid(Sender as Object,keycode as Keys) will send in a keys object representing the
'               whenever the grid has focus and the grid is NOT in editmode on a cell. You will get thing like Keys.Delete and Keys.F1 
'               and what not in this event. used to determining keys being pressed not just letters on the keyboard.
'
' Jun 09, 2005  Added the ability to restrict edited cells to a list of values
'               RestrictColumnEditsTo(ByVal colid As Integer, ByVal CaretDelimitedString As String) allows you to send in
'               a ^ delimited array of strings to display in the drop down combo box on editing a cell rather than the standard
'               textbox for edits.
'               You can also use RestrictColumnEditsTo(ByVal colid As Integer, ByVal ArrayListOfStrings As ArrayList) to build an
'               arraylist of strings to send into the grid as a restrictor list (Internally it will use the supplied arraylist to
'               build a CaretDelimitedString for you. If your arraylist items contain a ^ character that character will be
'               converted to a + character prior to this operation.
'               ClearAllColumnEditRestrictionLists() will remove any current column restrictions
'               to clear a specific column restriction use the ClearSpecificColumnEditRestrictionList(ByVal colid As Integer)
'               Sending in the colid of the desired colum you wish to clear the list on
'               You still have to AllowInGridEdits and set the desired columid as editable...
'               Example:
'               TAIG.AllowInGridEdits = True
'               TAIG.ColEditable(2) = True
'               TAIG.ColEditable(4) = True
'               TAIG.RestrictColumnEditsTo(2, "123^456^789")
'               
'               Fixed a bug in the RemoveRowsFromGrid that was creating an additional column on repopulate. 
'               SetAllCellBackColors(color as Color) will set all cells to a certain background color 
'               SetAllCellForeColors(color as Color) will set all cells to a certain foreground color
'               New Property PageSettings  Gets or Sets the PageSettings object used to print the grid
'               Use this to configure the Landscape,Papersize, and other printer specific settings externally before printing
'               the grids contents
'               The Tab key inside the grid when you are editing a field will n ow tab to the next edit field on the grid
'               If you are at the end of the grids edit fields the Tab key will tab off the grid to the next tabstop on the
'               Containing form
'               Fixed a bug where the grid did not Scroll to where the user was editing a cell on Tabbing through the Grids
'               Edit fields. It will now scroll to keep the edit fields on the top of the grid itself
'               restricted columns will now NOT allow you to type anything into the combo text box you must either leave
'               It blank to select nothing or select something in the list.
'
' Jun 10, 2005  Added SetEditItemText(string) to allow you to set the contents of either the combobox text or a textbox text on edit
'               while you are in edit mode on a given cell
'               Fixed a bounds condition where ig a grid had 2 cells on a row it would not tab to the next row on editing and tabbing
'               properly.
'               Changed the way a doubleclick is handled internally to NOT deselect the row that wwas just cliccked on in selected
'               row to better accomodaare the way the grid was being used in some of our applications. Single click interfaces
'               remin unchanged. (The problem was introduced with the introduction of Multiple Selections)
'
' Jun 12, 2005  Added ReplaceColMonthNumericWithMonthName(ByVal columnid As Integer)
'                   Use this to rip through a columns contents aand replace all occurances of Month Numerics with
'                   The Long Gregorian name for the month. Usefull in ordering operations where you want the months to order
'                   Numerically but want to show Alphabetically. Use SSQL to get the numeric form and order then convert them
'                   all to show the name for the month.
'               Added DoControlBreakProcessing(ByVal BreakColIntArrayList As ArrayList, _
'                                        ByVal SumColumnIntegerArraylist As ArrayList, _
'                                        ByVal IgnoreCase As Boolean, _
'                                        ByVal ColumnToPlaceSubtotalTextIn As Integer, _
'                                        ByVal SubtotalText As String, _
'                                        ByVal RightAlignSubTotalText As Boolean, _
'                                        ByVal ColorForSubTotalRows As System.Drawing.Color, _
'                                        ByVal BlankSeperateBreaks As Boolean, _
'                                        ByVal EchoBreakFieldsOnSubtotalLines As Boolean)
'                   The work horse routine can help replace all that grid manipulation routines in some of our reports to get
'                   ordered and grouped output. Simply craft the SQL code as usual and tell the grid to do the break reporting
'                   for you. Hand it 2 arraylists of integers representing column IDs that you want to break and Sum on.
'                   Boolean to ignore or not ignore case on the breaks, An Integer indicating Where to put the subtotal text.
'                   What text to place on the subtotal row, What color to color the subtotal rows, 
'                   A Boolean to Right Justify the subtotal text, A boolean to seperate each break
'                   with a blank row or not, Finally a boolean to indicate if you want to echo the Break data in the subtotal rows
'                   With this all manner of Formatted output can be crafted with a few commands and the grids internal functionality.
'
' Jun 20, 2005  Added new Event GridHover(Sender,row,col,item)
'                   Sender is the Grid object itself
'                   Row and Col are integers representing the Row and the column that the user is hovering over
'                   Item is a string representing the actuaal value at that position in the grid itself
'                   Event is Not raised if the user is hovering over the Header or the Title or is outside the bounds
'                   of the grid itself.
'
' Jun 21, 2005  Added new event GridHoverleave(sender as object)
'                   will fire this event if you are hivering over the grid canvas but not on grid data itself
'                   use this event to remove any tooltip text associated with the grid object. To prevent floating unassociated
'                   tooltips from hanging on out there.
'               New Methods
'                   DoControlBreakSubTotals(ByVal BreakColArrayValues As ArrayList, _
'                                       ByVal ColToFindValues As Integer, _
'                                       ByVal IgnoreCase As Boolean, _
'                                       ByVal SumColumnIntegerArrayList As ArrayList, _
'                                       ByVal ColorForBreakSubtotals As System.Drawing.Color, _
'                                       ByVal CutoffRow As Integer)
'                   BreakColArrayValues is an arraylist of distinct valued to subtotal on (Actual textual values in the grid)
'                   ColToFindValues what column to look in for the values above
'                   IgnoreCase  Self explanitory
'                   SumColumnIntegerArrayList and arraylist of column IDs to sum on
'                   ColorForBreakSubtotals What color to put in the new rows that will be added to the grid
'                   CutoffRow What row to stop looking past. (Use this to preset the stoprow if you are using this call
'                   multiple times to subtotal different columns)
'
'                   Public Function GetDistinctColumnEntries(ByVal colid As Integer, _
'                                             ByVal exclusionlist As ArrayList, _
'                                             ByVal ignorecase As Boolean) As ArrayList
'                   Colid is the integer columnid to look into
'                   exclusionlist is an arraylist of strings to exclude from the results
'                   ignorecase self explanitory
'                   This will return an arraylist of descrete values contained in a column minus any in the exclusionlist
'                   use this to supply the BreakColArrayValues parameter to the above new method
'                   Two overloads with sane parameters for the above additional entries
'
' Jun 22, 2005  Added new functionality allowing users to elect to have a column displayed as a checkbox.
'               ColCheckBox(colid) as Boolean property set to True or False (Default is false)
'               Rendered checkbox uses the ControlPaint.Drawcheckbox method to simulate the checkbox drawing
'               System will interpret values in the _grid(x,y) array of text TRUE, FALSE, YES, NO, Y, N, 1, 0 will all render
'               Checked or unchecked boxes as appropriate.
'               If you sel a Checkbox column to be editable then clicking on that cell will intrepret the value contained in the
'               _grid(x,y) cell and flip its meaning TRUE becomes FALSE and visa versa, YES becomes NO aand visa versa, 
'               and so on and so forth.
'               The new gridHover events will still pass back the value contained inside the grid not a boolean status of the checkbox
'               think of the grid containing text still but the display being tricked to interpret that text and render either a 
'               checked or unchecked box.
'               TABBING around in the grids edit fields will bypass checkbox edits and skip to the next non checkbox.
'
'               ClearAllGridCheckboxStates() will reset all grid renderings back to normal mode, turning off all checkbox rendering
'
'               Bug Fixes....
'               Fixed a bug where the LANDSCAPEMODE parameter of PrintTheGrid was being ignored
'               Fixed a bug where the Grids additional staatus arrays for hidden columns and editable columns where not being
'               reinitialized properly after an alteration of the grids dimensions.
'               Fixed a bug where populating with an sqldatareader would give you 1 row even if the reader was empty. It now
'               will hive you just the field names in the header and no empty row
'               
' Jun 29, 2005  Added AllowRowSelection to the grid to Allow (Default) or Disallow row selection. Necessary for BHISAuths  
'
' Jun 30, 2005  Added the property set for SelectedRows(Arraylist) to allow programatic setting of multiple rows of grid data
'               this was necessary for Auths selection of multiple services attached to an auth
'               Sorting of the grid will now retain all column width specifications even hidden columns because they are width 0
'               across the sort call. Even the sortoncolumn methods in the grid itself because they really call the same functions
'               internally as the context menus. Added test functionality to test the selected rows call to the test harnes
'
' Jul 1, 2005   Made RenderGrid() method public. This enables the grid to be used with rendering itself onto alternate
'               graphics contexts.
'               Added SelectAllRows() to automatically select all rows in the grid. (For Jen)
'
' Jul 3, 2005   Made Render Grid Private Again. Exposed a new Method GetGridContentsAsImage
'               that returns an Bitmap (Image) of the grids rendered surface
'               Non clipped and no scrollbars. (even if the scrollers are visible for the onscreen version of the grid).
'               This will be useful for placing grids contents onto printer object surfaces.
'               or even Direct3D rotating Cubes (sic)
'               It is important to use caution on how you call this routine because the resulting Bitmap can get to be HUGE
'               While the in memory version of this bitmap is miniscule until the routine is called.
'               calling it on a large grid can create a massive in memory bitmap that can drag the system to it's knees while
'               persisted.
'               That is why there is also a FreeGridContentImage method that allows you to remove this in memory footprint
'               after you are finished with it.
'               Usage:
'               Assume you have the grid and it is called myGrid.
'               Assume you have a graphics object called GR
'
'               Dim w As Integer = myGrid.GetGridContentsAsImage.Width  ' width of the internal to the grid image
'               Dim h As Integer = myGrid.GetGridContentsAsImage.Height ' height of the internal to the grid image
'
'               Dim sf As Single = Me.Width / w ' scale factor
'
'               Dim ww As Single = w * sf
'               Dim hh As Single = h * sf
'
'               GR.DrawImage(myGrid.GetGridContentsAsImage, 0, 0, ww, hh)
'               myGrid.FreeGridContentImage()
'
'               Fixed a small bug in the DeSelectAllRows call that was leaving selected roww selected but clearing the selectedrows 
'               collection. The result waas whatever row was last selected remained selected after the call.
'
' Jul 11, 2005  Ver 1.1.0.15
'               Added two new methods to assist reporting efforts 
'               PlaceGridOnGraphicsContext(Graphics, Xloc,Yloc,Width,Height)
'                   This will render at native resolutions the grids contents offset at xloc,yloc and clipped at width and height
'                   onto the supplied graphics context. Print the grid inside a report now is as easy as...
'                   TAIG.PlaceGridOnGraphicsContext(gr, 100, 100, 600, 200)
'                       places grid named TAIG onto Graphics Context GR at coord 100,100 with a width of 600 and height of 200 px
'
'               WordWrapColumn(Colid, WrapLength)
'                   will force a wordwrap operation on supplied column id with a length of wraplength. Will strip off embedded
'                   System.Environment.Newlines and will insert its own trimming excess whitespace and will reinsert the data
'                   back into the grid. Note unlike some of the other methods that format output (Password columns and what not)
'                   this method will aalter the internal data stored inside the grid. If you are going to check against the orig
'                   data in a daatabase you will likely get inequalities on columns that you run this method against. Use for 
'                   display purposes only
'
'               Fixed a bug in the edit mode where a tab off the last edit field in a given grid would not save the edited data
'               and would blank the field/cell instead. 
'
' Jul 12, 2005  Fixed an off by one error in SelectAllRows
'
' Jul 14, 2005  Added 
'               SetEditItem(ByVal row As Integer, ByVal col As Integer)
'                   Allows you to programatically select a certain col and row as the items being edited in the grid.
'                   The column must allow edits and the row and col must be in the range of rows and columns in the grid itself
'                   row >-1 and < _rows col >-1 and < _cols.... 
'                   Implemented for Larrys messaging scheme in the BHIS Project.
'
' Aug 01, 2005  Ver 1.1.0.18
'               Added 4 new Methods and one Function
'               Functions
'               CreatePersistanceScript(ByVal tname As String) as String
'                   Will craft a syntactically correct Create Table statement using table name TNAME
'                   that will drop that table name first from a database. The create will then be
'                   Followed by a series of Inserts that will insert data into the database table created
'                   by the create table statement. The fiels added to the table will follow the
'                   Header names unless the field is named ID. In which case the field will be renamed
'                   ID_DATA. This is because the create table statement will craft an identity field
'                   called ID, seting its type to numeric and setting its autonumbering to start at
'                   1. All other fields will be VARCHARS that are large enough to hold the largest
'                   value currently in the grid, or 8000 characters whichever is smaller.
'                   If the field contains all empty strings then the resuklting Varchar field will be 
'                   set to 10 characters in length.
'                   The strings inserted will be ' escaped allowing
'                   values like O'Mally to be inserted into the result set.
'
'               Methods
'               FrequencyDistribution(ByVal sgrid As TAIGridControl, ByVal ColForFrequency As Integer)
'                   This populate grid call will take a source grid contents and rip through that grids
'                   data on colid ColForFrequency and count the number of times each distinct value
'                   appears in that column. The result will be a grid with 2 columns one the distinct
'                   values plucked from the source grid. The other a count of how many times each of
'                   those distinct values appeared in the source grid.
'                   Usefull in reporting apps where things like Diagnosis or Procedure code counts
'                   are desired.
'                   1 Overload where you pass in a SortDescending boolean and the resulting grid will
'                   be sorted numerically on the frequency eirther descending or ascending.
'
'               SortGridOnColumnDate(ByVal col As Integer, ByVal Descending As Boolean)
'               SortGridOnColumnNumeric(ByVal col As Integer, ByVal Descending As Boolean)
'                   will sort the colid COL as either a date or as a number Asc or Desc as defined by
'                   the boolean parameter Descending.
'               Added 2 new menu options on the context menu to call the SortGridOnColumnNumeric 
'               renamed the SortGrid menu options to read ASCII Sort Grid to better reflect the
'               function of the selection. The Date sorters were already there they have just been
'               exposed to the outside world programatically.
'
'
' Aug 02, 2005  Ver 1.1.0.19
'               Added 1 new function w 1 overload
'               CreateHTMLTableScript()
'                   Defaults to Border 1, Match colors and Omits Nulls see below...
'               CreateHTMLTableScript(ByVal BorderVal As Integer, _
'                                     ByVal MatchColors As Boolean, _
'                                     ByVal OmitNulls as Boolean)
'               Will return a string containing an HTML table format that represents the grids contents
'               BorderVal is the Border thickness ranges from 0 to whatever 
'               MatchColors will make the grids title, Header, and Cells match the origin grids colors
'               OmitNulls will make the resultset have blank cells where the {null} cell was in the source
'
' Aug 03, 2005  Ver 1.1.0.20
'               Added two new menu options to the context manu that Wrap the SQL and the HTML creators
'               that have been addded over the past few days
'
' Aug 04, 2005  Ver 1.1.0.21
'               Added a new item to the Context Menus under Math. Display Frequency Distribution
'               This will call the internal FrequencyDistribution method and display the results in
'               another grid contained inside a dialog window. Usefull in the Query reporter to do 
'               things like count the number of each procedure codes or Diag codes in a query like
'               Select * from claims where svcdate between '10/1/2001' and '12/31/2001'. For those
'               analysis freaks out there.
'
' Aug 04, 2005  Ver 1.1.0.22
'               Fixed a short sighted issue where creating an SQL persistance script if the Grid had
'               Header labled with Spaces in them the resulting SQLscript would have a problem on the
'               Inserts because the field names would not have been enclosed in []'s
'               DOH!
'
' Aug 05, 2005  Ver 1.1.0.23
'               Major performance issue with the HTML table creation functions fixed
'               String Manipulations are a Drag. They were making the thing run like it was in 
'               waist deep wet cement.
'
' Aug 25, 2005  Ver 1.1.0.24
'               Added a properties option to the context menu. For now brings up a limited dialog
'                   where the end user can adjust font sizes, and the visibility of he title and header
'               Fixed the notion of selecting with the mouse and then arrowing about leaving the
'                   clicked row selected as you cursored about. Now you have to shift click or 
'                   ctrl click to get those results
'               Made the rendering of checkbox columns render a ghosted checkbox if the cells contents
'                   are empty rather than a non checked checkbox. This was done for the TAIHCSISAugmentor
'                   project. If that cells contents are rendered ghosted then it cannot be change via
'                   the mouseclick even if the column editing is allowed with the other methods.
'
' Sep 05, 2005  Ver 1.1.0.26
'               Re-implemented the way the Selected Rows collection was being populated to more
'               closely approximate the way it used to work before the Aug 25 additions. This should
'               fix the breakage that some folks endured with multiple selection interpretation caused
'               by those changes.
'               This build is a prelude version to the new version of the grid with build in Fuzzy
'               algorythims for set analsys.
'
' Nov 10, 2005  Ver 1.1.0.27
'               Changed the way the printer object is being dealt with. Switched some of the printer
'               object that are internal to late binging to better catch exceptions around printers
'               and printer manipulation. (Like 'There aren't any installed printers')
'               If the initial printersetup is failing internally the contexts menus printer stuff
'               will be disabled. 
'               Also the printer setup dialog won't display if the grid itself is empty. 
'               All this should solve issue 352 on BHIS as well as Larrys issue with the HCSIS Robot
'
' Nov 28, 2005  Ver 1.1.0.30
'               Implementation of a new feature called tearaway columns
'               Rational:
'                   With some tools like the Query Reporter and the Livanta Explorer lots of columns
'                   in data grids proved cumbersome to message.
'                   One persons column layout was good for that person but sub optimal
'                   for the next bloke.
'                   Tearaway columns allow for user manipulation of column order in seperate
'                   floating tool windows.
'                   If the user repopulates the grid with data then the open floating tearaways
'                   will adjust to the new contents as follows
'                   1)  If the floating windows was on a column that no longer exists in the new grid
'                       It will be closed
'                   2)  If the floating tool window still exists in the new grid its title will change
'                       to refrlect the new grid column title and its contents will change to reflect the
'                       new column contents
'                   3)  If the Verticle scrollbar is visible in the main grid the toolwindows will also
'                       have the verticle scrollbar
'                   User interaction on any verticle scrollbar both in a floating toolwindow and
'                   the main grid itself will adjust all visible verticle scrollbars accordingly
'
'               The context menu now has three new options
'               Tear Column Away
'                   Will take the colum you were over when you pressedthe right mousebutton and tear
'                   it out into a floating window
'               Hide Tear Away Column
'                   Will take the column in the main grid and will cloas a tearaway for the column
'                   if one is already open otherwise it will do nothing
'               Hide All Tearaways
'                   This global option will toss away any tearaway floating windows that are open
'
' Nov 29, 2005  Ver 1.1.0.31
'               Added the ability for the grid to broadcast hovering events on tearawys as if it was
'               getting User Interaction on the main grid itself. The issue is that if you are using
'               a simple tooltip control to display end user information to the user based on these
'               GridHover events then the tooltip appears to be behind any of the tearaway windows.
'               To solve this I implemented a sort of builtintooltip on each tearaway as well as 
'               the main grid itself.
'               In the GridHover event you get the Sender, row, col and value hovering over
'               call DisplayGridToolTip(sender, TextVal) sending back the sender and whatever you
'               want to have displayed in the tool tip. The Grid will route the request to the correct
'               control or form based on who the initial sender was.
'               Call HideGridToolTip() to make any displayed tooltips go away.
'               Of course you can just respond to the gridhover event in your own way if you want to
'               Also fixed a small error in sizing with the grid and determining what column you are
'               over when you pull up a context menu.
'
'
'               Ver 1.1.0.32 
'               Simple Change to make Tearaways topmost on the screen
'               to help aid in its use with the Livanta Explorer
'
' Dec 01, 2005  Ver 1.1.0.33
'               Added new Tearawway option to the context menu. Tear Away Multiple Columns will bring
'               up a dialog allowwing you to click select more than one ccolumn in te current grid
'               to tear away. Ok will tear the selected ones away cancel will abort that process
'               Added the Arrange Tear Aways option to the context menu to arrange tearaway windows
'               across and then down the screen. Using system defined screen resolution as its
'               for placement. Windows will be arrange in the order they were opened and will be non
'               overlapping.
'               Tearawys are now embellished coloration wise to match grid defaults. They are outlined
'               per Ann's request with her use of the Livanta Explorer.
'
'               Ver 1.1.0.34
'               Added the ability to click on a row in a tear away and have that rows selection
'               be echoed vack to the parent grid and subsequently echoed to all other tearaway
'               windows automagically. Asked for by Ann in her use of the Livanta Explorer
'
' Dec 02, 2005  Ver 1.1.0.35
'               Added new functionality to the AutoArrange Tear Aways to automatically shrink
'               windows to match the height of the rendered content on small lists.
'               Added internal functionality to better inhibit redrawing on tearaway movement
'               and resizing. This will better remove the Windorms getting hosed bugs that have
'               plagued the autotrack window sizing feature I put in yesterday for the Livanta 
'               explorer.
'               Plugged a small memory hole that would ultimately result in a GC.Collect freeze
'               during use in some situations
'               Mousing over a tearawy will now focus that tearaway forcing displayed tooltips 
'               to render topmost preventing the tooltip layering issue that occured with multiple
'               tearaways being present.
'               Row selection in Any tearaway will track on all other tearaways and the mother grid
'               Arrow key movement of selected row in any tearaway will track in all other tearaways
'               as well as the mother grid.
'               General performance tweaks and internal optimizations on the tearaway feature in general
'
' Dec 05, 2005  Ver 1.1.0.36
'               Added the ability to click and double click on a tearaway and have the click or doubleclick
'               send the mother grid the event so it could raise the CellClicked and CellDoubleClicked 
'               events respectively
'               Added two new public Methods RaiseCellClickedEvent and RaiseCellDoubleClickedEvent
'               these public methods support the above functionality and allow a developer to 
'               message these functions of the grid externally
'               Changed the way a cell selection in a tear away occurs to clear the selectedrows collection
'               in the mothergrid on selection of a cell in a tearaway.
'
' Dec 19, 2005  Ver 1.1.0.37
'               Major refactoring effort around the storing and rendition of different backcolors and 
'               forecolors in cells of the grid. The old method attempted to determine if a desired
'               color was already in he list. The equality operator uses was flawed and this a new
'               color entry was always being made. Wasting memory and performance. Flaw was uncovered
'               during implementation of new reporting requirements for Beaver county. These changes 
'               were made to the PDF version of the grid and have been echoed back to the mainstream
'               version in 1.1.0.37.
'               Also in this version are a new Overload of DoControlBreakProcessing with an extra
'               parameter boolean TreatBlanksAsSame. This will allow and already controlbreaked grid
'               to be further refined with additional breaks. The blanked fields that are the same from
'               the prior control will be inferred thus treated as the same.
'               Two new public methods.
'               InsertRowsIntoGridAt(ByVal atrow As Integer, ByVal numrows As Integer)
'                   Will do as it seay and insert numrows of open space int the grid atrow
'               Public Sub PolulateGridWithDataAt(ByVal ConnectionString As String, _
'                                            ByVal Sql As String, _
'                                            ByVal Atrow As Integer, _
'                                            ByVal newbackcolor As Color, _
'                                            ByVal newheadercolor As Color, _
'                                            ByVal ColOffSet As Integer)
'                   Will insert the sql results of SQL against ConnecrionString into the grid
'                   atrow with the newbackcolor, and newheadercolor offseting by coloffset
'               There is one overlolad on this the takes no coloffset and assumes column 0 for
'               the starting column.
'               This is used in the beaver reporting application to insert subreports into the grid.
'
'               Ver 1.1.0.38
'               Added two new properties for Larry EditModeCol and EditModeRow
'               If the grid is in editmode and you examine these properties you should get the 
'               Col and Row of the edit operation. Usefull if you are processing dialog keys at
'               the form level in Lookups and what not.
'
' Dec 23, 2005  Ver 1.1.0.39
'               Fixed an issue where if a cell ws in edit mode and the use either scrolled up to down
'               and or left to right the edit boc (cmbo or textbox would no scroll to match the
'               underlieing grid movement. Code added to the VS amd HS scroll managers to account
'               for this. Bug discovered by Larry in his BHIS claims module.
'
' Jan 10, 2006  Ver 1.1.0.40
'               Fixed a stupid bug with currency conversion crap in the ColFormatAsMoney() call
'               
'               Ver 1.1.0.41
'               Fixed a different stupid bug in the ColFormatAsMoney method for Jeff
'
' Feb 17, 2006  Ver 1.1.0.42
'               Added a call to KillAllTearAways() on handledestroyed to help in leaving tearaways orphand
'               in BHIS. 
'               Moved KillAllTearAways to public methods allowing developers to call it externally
'               in cases where they might want to ensure the tearaways are removed AKA in BHIS
'               Cleared the SelectedRows Collecton on edit mode transitions from row to row this should fix
'               a problem Jeff Id'd with the grids usage inside of BHIS
'               Some printer output cleanup in an attempt to fix the BHIS 4039 lexmark printer issues.
'               I fear the issue still exists and that the 4039 printer driver in Lehigh is Flatlined.
'               All new functionality and actions tested in harness and passed.
'
'               Ver 1.1.0.43
'               Added new functionality to allow attachment of custom contextual menus to the grids
'               right mouse button functionality. This will replace the large menu that comes along for
'               the ride with the grid with a user supplied menu. Event handlers of this supplied menu
'               functionality remain external to the grid and are the responsibility of the developer 
'               Usage...
'               TAIG.ContextMenu = {some menu that you create}
'               In support for this there are two new Readonly Properties
'               ColOverOnMenuButton and RowOverOnMenuButton. These two integer properties will return
'               the zero based integer values the represent the row and col that the user was over when
'               they clicked the right mousebutton to bring up the menu in the first place. Even if the
'               menu selection itself had the user move over a new row or column. If these return -1's
'               then the user contexted menu over a blank part of the grid or the header or title
'               of the grid.
'               Functionality implemented for support of the usage of our grid control in ClaimsExplorer
'               rather than the Crappy Microsoft grid.
'
'               Ver 1.1.0.44
'               Added new functionality to allow or disallow functional groups from the builtin context
'               menu....
'               AllowTearAwayFunctionality(),AllowExcelFunctionality(),AllowTextFunctionality(),
'               AllowHTMLFunctionality(),AllowSQLFunctionality(),AllowMathFunctionality(),
'               AllowFormatFunctionality(),AllowSettingsFunctionality(),AllowSortFunctionality()
'               All are Boolean properties that default to True
'               Setting one of false will turn off that bit of functionality in the builtin context menu
'               for example:
'               AllowTearAwayFunctionality = False will disable the 5 tearaway menu items in the context
'               menu 
'               Implemented for BHIS selective disallowing of some items that cause modality issues
'               with other window functions contained inside the parent program. Because the all default
'               to true this has no effect on software that does not exercise these new properties.
'               It's a drop in replacement all all grid since 1.1.0.33
'
' Mar 03, 2006  Ver 1.1.0.45
'               Added new functionality to allow spread sheet output to roll over onto a new sheet when
'               the grid has more than X rows in it. New property of ExcelMaxRowsPerSheet which defaults
'               to 30000 and has a range of 100 - 65000 dictates where the split will occur on exporting
'               insanely large grid contents to excel.
'
' Mar 06, 2006  Ver 1.1.0.46
'               Fixed a rounding issue with breaking large sheets across multiple pages in excel export
'               Integer division working differently between c# and VB in the .Net framework
'
' Mar 08, 2006  Ver 1.1.0.47
'               Fixed a problem that eluded me earlier with the header of the grid appearing on
'               the newly created sheet in a multisheet excel export operation. Problem only happened
'               where there was some rows, columns or cells that were colored differently than the predominant
'               background color of the grid itself. Basically it breaks down into what I believe is a 
'               bug in the way non instance types of objects are created and scoped during a programs execution
'               It breaks down into this maxium....
'               If you Create it you must initialize it lest you get surprised by unexpected bahaviors 
'               later on in the codes execution. 12 Hrs blown on this simple to resolve issue but
'               maddingly elusive to track down.
'
' Mar 17, 2006  Ver 1.1.0.48
'               Added in System.DateTime and System.Single to list of DataTypes in PopulateGridWithADataTable
'
' Jul 06, 2006  Ver 1.1.0.49
'               Added an overload for the ITEM property to take a string as the column identifier
'               where it will do the GetColomnIDByName for you.
'
' Oct 20, 2006  Ver 1.1.0.51
'               Added code to shrink and wrap the gridreporttitle on printing grid contents
'
' Oct 27, 2006  Ver 1.1.0.52
'               Added new property Defaults to False ShowDatesWithTime. THis will force the population
'               of any SystemDatetime column to expand the the time portion to show HH:MM AM/PM even if
'               the time portion is 12 midnight ( system defaults to not showing the time by default )
'               If left at false the time portion is ignored in all cases. All populators from database
'               calls will honor this boolean property. Implemented for DAS issues Jen has in BHIS
'
'
' USAGE:
'
'   assume the grid control is named TAIG
'
'   To put some stuff in  the grid
'
'   Private _database as string = "Data Source={SERVER};User ID={SOMEUSER}; Password={SOMEPASSWORD};Initial Catalog={DATABASE};"
'   Private _gridfontsmall As New Font("ARIAL", 8, FontStyle.Bold, GraphicsUnit.Point)
'   Private _sql as string = "SELECT TOP 10 * FROM SOMETABLE"
'
'   TAIG.PopulateGridWithData(_database,_sql,_gridfontsmall)
'
'   ' there are over loads for this method that take a color also but the default works fine
'   ' if you want to know more... Use the Force, Read the Source...
'
'   
''' <summary>
''' The main grid control used in all our applications. Developed from scratch over the course of years it
''' has functionality to do all matter of things related to the gathering,display,editing,exporting tabular data.
''' <code>
''' <example>
''' Example Usage
''' Assume that the control is on a form and is named TAIG
''' the following code will populate the grid with some data taken from SQL server
''' </example>
''' Private _database as string = "Data Source={SERVER};User ID={SOMEUSER}; Password={SOMEPASSWORD};Initial Catalog={DATABASE};"
''' Private _gridfontsmall As New Font("ARIAL", 8, FontStyle.Bold, GraphicsUnit.Point)
''' Private _sql as string = "SELECT TOP 10 * FROM SOMETABLE"
''' TAIG.PopulateGridWithData(_database,_sql,_gridfontsmall)
''' </code>
''' </summary>
''' <remarks></remarks>
''' 
Public Class TAIGridControl
    Inherits System.Windows.Forms.UserControl

#Region " Internal Storage "

    Public _LoggingEnabled As Boolean = False

    ' Items for the grid Title
    Private _GridTitle As String = "Grid Title"
    Private _GridTitleVisible As Boolean = True
    Private _GridTitleFont As Font = New Font("Arial", 16, FontStyle.Regular, GraphicsUnit.Point)
    Private _GridTitleHeight As Integer = 16
    Private _GridTitleBackcolor As Color = Color.Blue
    Private _GridTitleForeColor As Color = Color.White
    Private _GridSize As Point

    ' Items for the grid Header
    Private _GridHeader(1) As String
    Private _GridHeaderFont As Font = New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
    Private _GridHeaderVisible As Boolean = True
    Private _GridHeaderBackcolor As Color = Color.LightBlue
    Private _GridHeaderForecolor As Color = Color.Black
    Private _GridHeaderHeight As Integer = 16
    Private _GridHeaderStringFormat As StringFormat = New StringFormat

    Private _grid(1, 1) As String
    Private _gridBackColor(1, 1) As Integer
    Private _gridBackColorList(1) As Brush
    Private _gridCellAlignment(1, 1) As Integer
    Private _gridCellAlignmentList(1) As StringFormat
    Private _gridCellFonts(1, 1) As Integer
    Private _gridCellFontsList(1) As Font
    Private _gridForeColor(1, 1) As Integer
    Private _gridForeColorList(1) As Pen
    Private _colPasswords(1) As String
    Private _colMaxCharacters(1) As Integer
    Private _CellOutlines As Boolean = True
    Private _CellOutlineColor As Color = Color.Black
    Private _OldContextMenu As ContextMenu

    Private _colwidths(1) As Integer
    Private _colEditable(1) As Boolean
    Private _rowEditable(1) As Boolean
    Private _colEditableTextBackColor As Color = Color.Yellow
    Private _colEditRestrictions As New ArrayList
    '' Private _rowEditRestrictions As New ArrayList
    Private _coloffsets(1) As Integer
    Private _colhidden(1) As Boolean
    Private _colboolean(1) As Boolean
    Private _rowheights(1) As Integer
    Private _rowoffsets(1) As Integer
    Private _rowhidden(1) As Boolean

    Private _ColWidthsBeforeAutoSize() As Integer
    Private _RowHeightsBeforeAutoSize() As Integer
    Private _ColWidthsAfterAutoSize() As Integer
    Private _RowHeightsAfterAutoSize() As Integer

    Private _rows As Integer = 0
    Private _cols As Integer = 0

    Private _LMouseX As Integer = -1
    Private _LMouseY As Integer = -1

    Private _AllowPopupMenu As Boolean = True
    Private _AllowInGridEdits As Boolean = False
    Private _AllowRowSelection As Boolean = True

    Private _AllowTearAwayFuncionality As Boolean = True
    Private _AllowExcelFunctionality As Boolean = True
    Private _AllowTextFunctionality As Boolean = True
    Private _AllowHTMLFunctionality As Boolean = True
    Private _AllowSQLScriptFunctionality As Boolean = True
    Private _AllowPrintFunctionality As Boolean = True
    Private _AllowMathFunctionality As Boolean = True
    Private _AllowSettingsFunctionality As Boolean = True
    Private _AllowSortFunctionality As Boolean = True
    Private _AllowFormatFunctionality As Boolean = True
    Private _AllowRowAndColumnFunctionality As Boolean = True

    Private _EditMode As Boolean = False    ' set whenever a edit session is possible on a textbox or combobox
    ' cleared when that textbox or combobox loses focus
    Private _EditModeRow As Integer = -1
    Private _EditModeCol As Integer = -1
    Private _AllowControlKeyMenuPopup As Boolean = True
    Private _AllowColumnSelection As Boolean = True
    Private _AllowMultipleRowSelections As Boolean = True
    Private _AllowWhiteSpaceInCells As Boolean = True
    Private _antialias As Boolean = False
    Private _alternateColorationALTColor As Color = Color.MediumSpringGreen
    Private _alternateColorationBaseColor As Color = Color.AntiqueWhite
    Private _alternateColorationMode As Boolean = False
    Private _AutoFocus As Boolean = False
    Private _DefaultColWidth As Integer = 50
    Private _DefaultRowHeight As Integer = 14
    Private _DefaultBackColor As Color = Color.AntiqueWhite
    Private _DefaultForeColor As Color = System.Drawing.Color.Black
    Private _DefaultCellFont As Font = New Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Point)
    Private _DefaultStringFormat As StringFormat = New StringFormat
    Private _AutosizeCellsToContents As Boolean = False
    Private _AutoSizeAlreadyCalculated As Boolean = False
    Private _AutoSizeSemaphore As Boolean = True
    Private _Painting As Boolean = False
    Private _TearAwayWork As Boolean = False
    Private _dataBaseTimeOut As Integer = 500
    Private _omitNulls As Boolean = False
    Private _MouseWheelScrollAmount As Integer = 10
    Private _RowClicked As Integer = -1
    Private _ColClicked As Integer = -1
    Private _SelectedRow As Integer = -1
    Private _ShiftMultiSelectSelectedRowCrap As Integer = -1
    Private _SelectedRows As New ArrayList
    Private _SelectedColumn As Integer = -1
    Private _ShowDatesWithTime As Boolean = False
    Private _ShowProgressBar As Boolean = True
    Private _ShowExcelExportMessage As Boolean = True
    Private _RowHighLiteBackColor As Color = System.Drawing.Color.Blue
    Private _RowHighLiteForeColor As Color = System.Drawing.Color.White
    Private _ColHighliteBackColor As Color = System.Drawing.Color.MediumSlateBlue
    Private _ColHighliteForeColor As Color = System.Drawing.Color.LightGray
    Private _BorderColor As Color = Color.Black
    Private _ScrollBarWeight As Integer = 14
    Private _BorderStyle As System.Windows.Forms.BorderStyle = BorderStyle.FixedSingle
    Private _MaxRowsSelected As Integer = 0
    Private _PaginationSize As Integer
    Private _scrollinterval As Integer = 5
    Private _LastSearchText As String = ""
    Private _LastSearchColumn As Integer = -1
    Private _DoubleClickSemaphore As Boolean = False

    ' items for the export to text
    Private _delimiter As String = ","
    Private _includeFieldNames As Boolean = True
    Private _includeLineTerminator As Boolean = True

    ' excel constants
    Public Const xlPortrait As Integer = 1
    Public Const xlLandscape As Integer = 2
    Private Const xlAutomatic As Integer = -4105
    Private Const xlContinuous As Integer = 1
    Private Const xlThin As Integer = 2
    Private Const xlEdgeLeft As Integer = 7
    Private Const xlEdgeTop As Integer = 8
    Private Const xlEdgeBottom As Integer = 9
    Private Const xlEdgeRight As Integer = 10
    Private Const xlInsideVertical As Integer = 11
    Private Const xlInsideHorizontal As Integer = 12
    Private Const xlCenter As Integer = -4108
    Private Const xlTop As Integer = -4160
    Private Const xlToRight As Integer = -4161
    Private Const xlNormalView As Integer = 1
    Private Const xlPageBreakPreview As Integer = 2
    Private Const xlMaximized As Integer = -4137

    ' items for the export to excel
    Private _excelFilename As String = ""
    Private _excelWorkSheetName As String = "Grid Output"
    Private _excelKeepAlive As Boolean = True
    Private _excelPageOrientation As Integer = xlPortrait
    Private _excelPageFit As Boolean = True
    Private _excelIncludeColumnHeaders As Boolean = True
    Private _excelShowBorders As Boolean = False
    Private _excelMaximized As Boolean = True
    Private _excelAutoFitRow As Boolean = True
    Private _excelAutoFitColumn As Boolean = True
    Private _excelAlternateRowColor As Color = Color.FromArgb(204, 255, 204)
    Private _excelUseAlternateRowColor As Boolean = True
    Private _excelMatchGridColorScheme As Boolean = True
    Private _excelOutlineCells As Boolean = True
    Private _excelMaxRowsPerSheet As Integer = 30000

    ' items to support user resizing of columns 
    Private _MouseDownOnHeader As Boolean = False
    Private _ColOverOnMouseDown As Integer = -1
    Private _RowOverOnMouseDown As Integer = -1
    Private _AllowUserColumnResizing As Boolean = True
    Private _LastMouseY As Integer = 0
    Private _LastMouseX As Integer = 0
    Private _UserColResizeMinimum As Integer = 5

    ' items fo the export to xml
    Private _xmlFilename As String = ""
    Private _xmlDataSetName As String = "Grid_Output"
    Private _xmlNameSpace As String = "TAI_Grid_Ouptut"
    Private _xmlTableName As String = "Table"
    Private _xmlIncludeSchema As Boolean = False

    ' Support menuing
    Private _ColOverOnMenuButton As Integer = -1
    Private _RowOverOnMenuButton As Integer = -1

    ' support for report printing

    Private _gridReportTitle As String = ""
    Private _gridReportMatchColors As Boolean = True
    Private _gridReportOutlineCells As Boolean = True
    Private _gridReportPreviewFirst As Boolean = True
    Private _gridReportNumberPages As Boolean = True
    Private _gridReportOrientLandscape As Boolean = False
    Private _gridReportScaleFactor As Single = 1.0

    Private _gridReportPageNumbers As Integer = -1
    Private _gridReportCurrentrow As Integer = -1
    Private _gridReportCurrentColumn As Integer = -1
    Private _gridReportPrintedOn As DateTime = Now

    Private _gridStartPage As Integer = -1
    Private _gridEndPage As Integer = -1
    Private _gridPrintingAllPages As Boolean = True
    Private _gridStartPageRow As Integer = -1

    'Private _psets As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings

    'Private _OriginalPrinterName As String = _psets.PrinterSettings.PrinterName

    'Private _image As System.Drawing.Bitmap

    'Private WithEvents _PageSetupForm As New frmPageSetup(_psets)

    Private _psets As Object

    Private _OriginalPrinterName As String = ""

    Private _image As System.Drawing.Bitmap

    Private WithEvents _PageSetupForm As New frmPageSetup

    'Private WithEvents _PageSetupForm As Object

    ''Private WithEvents TearItem As New frmColumnTearAway

    Private TearAways As ArrayList = New ArrayList
    Friend WithEvents miStats As System.Windows.Forms.MenuItem

    ''' <summary>
    ''' Denotes the form of action necessary to be taken to have a cell in editmode actually have its value 
    ''' change. Fireing the cell edited event. Either having the user press the enter/return key or having the
    ''' user shift focus to another control or cell in the grid itself.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum GridEditModes
        KeyReturn = 0
        LostFocus = 1
    End Enum

    Private _GridEditMode As GridEditModes = GridEditModes.KeyReturn


    <System.Runtime.InteropServices.DllImportAttribute("gdi32.dll")> _
        Private Shared Function BitBlt( _
                ByVal hdcDest As IntPtr, _
                ByVal nXDest As Integer, _
                ByVal nYDest As Integer, _
                ByVal nWidth As Integer, _
                ByVal nHeight As Integer, _
                ByVal hdcSrc As IntPtr, _
                ByVal nXSrc As Integer, _
                ByVal nYSrc As Integer, _
                ByVal dwRop As System.Int32) As Boolean
    End Function

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.txtHandler.Height = 1
        Me.txtHandler.Width = 1
        Me.txtHandler.Left = 0
        Me.txtHandler.Top = 0
        Me.txtHandler.BackColor = _BorderColor
        'Me.txtHandler.Visible = False

        ' Set the value of the double-buffering style bits to true.
        'Me.SetStyle(ControlStyles.DoubleBuffer _
        '  Or ControlStyles.UserPaint _
        '  Or ControlStyles.AllPaintingInWmPaint, _
        '  True)
        'Me.UpdateStyles()

        'Try
        '    _psets = New System.Drawing.Printing.PageSettings

        '    'MsgBox(_psets.ToString())

        '    _OriginalPrinterName = _psets.PrinterSettings.PrinterName

        '    _PageSetupForm = New frmPageSetup(_psets)

        'Catch ex As Exception

        '    miPreviewTheGrid.Enabled = False
        '    miPrintTheGrid.Enabled = False
        '    miPageSetup.Enabled = False

        'End Try

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents hs As System.Windows.Forms.HScrollBar
    Friend WithEvents vs As System.Windows.Forms.VScrollBar
    Friend WithEvents txtHandler As System.Windows.Forms.TextBox
    Friend WithEvents menu As System.Windows.Forms.ContextMenu
    Friend WithEvents miExportToExcelMenu As System.Windows.Forms.MenuItem
    Friend WithEvents miExportToExcel As System.Windows.Forms.MenuItem
    Friend WithEvents miAutoFitCols As System.Windows.Forms.MenuItem
    Friend WithEvents miAutoFitRows As System.Windows.Forms.MenuItem
    Friend WithEvents miALternateRowColors As System.Windows.Forms.MenuItem
    Friend WithEvents miMatchGridColors As System.Windows.Forms.MenuItem
    Friend WithEvents miOutlineExportedCells As System.Windows.Forms.MenuItem
    Friend WithEvents miFormatStuff As System.Windows.Forms.MenuItem
    Friend WithEvents miFormatAsMoney As System.Windows.Forms.MenuItem
    Friend WithEvents miFormatAsDecimal As System.Windows.Forms.MenuItem
    Friend WithEvents miFormatAsText As System.Windows.Forms.MenuItem
    Friend WithEvents miCenter As System.Windows.Forms.MenuItem
    Friend WithEvents miLeft As System.Windows.Forms.MenuItem
    Friend WithEvents miRight As System.Windows.Forms.MenuItem
    Friend WithEvents miFontsSmaller As System.Windows.Forms.MenuItem
    Friend WithEvents miFontsLarger As System.Windows.Forms.MenuItem
    Friend WithEvents miSmoothing As System.Windows.Forms.MenuItem
    Friend WithEvents pBar As System.Windows.Forms.ProgressBar
    Friend WithEvents miExportToTextFile As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents miHeaderFontLarger As System.Windows.Forms.MenuItem
    Friend WithEvents miHeaderFontSmaller As System.Windows.Forms.MenuItem
    Friend WithEvents miTitleFontLarger As System.Windows.Forms.MenuItem
    Friend WithEvents miTitleFontSmaller As System.Windows.Forms.MenuItem
    Friend WithEvents gb1 As System.Windows.Forms.GroupBox
    Friend WithEvents miSearchInColumn As System.Windows.Forms.MenuItem
    Friend WithEvents miAutoSizeToContents As System.Windows.Forms.MenuItem
    Friend WithEvents miAllowUserColumnResizing As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents miSortAscending As System.Windows.Forms.MenuItem
    Friend WithEvents miSortDescending As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents miHideRow As System.Windows.Forms.MenuItem
    Friend WithEvents miHideColumn As System.Windows.Forms.MenuItem
    Friend WithEvents miSetRowColor As System.Windows.Forms.MenuItem
    Friend WithEvents miSetColumnColor As System.Windows.Forms.MenuItem
    Friend WithEvents miSetCellColor As System.Windows.Forms.MenuItem
    Friend WithEvents miShowAllRowsAndColumns As System.Windows.Forms.MenuItem
    Friend WithEvents clrdlg As System.Windows.Forms.ColorDialog
    Friend WithEvents miDateAsc As System.Windows.Forms.MenuItem
    Friend WithEvents miDateDesc As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents miSumColumn As System.Windows.Forms.MenuItem
    Friend WithEvents miSumRow As System.Windows.Forms.MenuItem
    Friend WithEvents miMaxCol As System.Windows.Forms.MenuItem
    Friend WithEvents miMaxRow As System.Windows.Forms.MenuItem
    Friend WithEvents miMinCol As System.Windows.Forms.MenuItem
    Friend WithEvents miMinRow As System.Windows.Forms.MenuItem
    Friend WithEvents miColAverage As System.Windows.Forms.MenuItem
    Friend WithEvents miRowAverage As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents miCopyCellToClipboard As System.Windows.Forms.MenuItem
    Friend WithEvents pdoc As System.Drawing.Printing.PrintDocument
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents miPrintTheGrid As System.Windows.Forms.MenuItem
    Friend WithEvents miPreviewTheGrid As System.Windows.Forms.MenuItem
    Friend WithEvents miPageSetup As System.Windows.Forms.MenuItem
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
    Friend WithEvents cmboInput As System.Windows.Forms.ComboBox
    Friend WithEvents miSortNumericAsc As System.Windows.Forms.MenuItem
    Friend WithEvents miSortNumericDesc As System.Windows.Forms.MenuItem
    Friend WithEvents miExportToSQLScript As System.Windows.Forms.MenuItem
    Friend WithEvents miExportToHTMLTable As System.Windows.Forms.MenuItem
    Friend WithEvents miDisplayFrequencyDistribution As System.Windows.Forms.MenuItem
    Friend WithEvents miProperties As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents miTearColumnAway As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents miHideColumnTearAway As System.Windows.Forms.MenuItem
    Friend WithEvents miHideAllTearAwayColumns As System.Windows.Forms.MenuItem
    Friend WithEvents _TTip As System.Windows.Forms.ToolTip
    Friend WithEvents miMultipleColumnTearAway As System.Windows.Forms.MenuItem
    Friend WithEvents miArrangeTearAways As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.hs = New System.Windows.Forms.HScrollBar()
        Me.vs = New System.Windows.Forms.VScrollBar()
        Me.txtHandler = New System.Windows.Forms.TextBox()
        Me.menu = New System.Windows.Forms.ContextMenu()
        Me.miCopyCellToClipboard = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.miSumColumn = New System.Windows.Forms.MenuItem()
        Me.miSumRow = New System.Windows.Forms.MenuItem()
        Me.miMaxCol = New System.Windows.Forms.MenuItem()
        Me.miMaxRow = New System.Windows.Forms.MenuItem()
        Me.miMinCol = New System.Windows.Forms.MenuItem()
        Me.miMinRow = New System.Windows.Forms.MenuItem()
        Me.miColAverage = New System.Windows.Forms.MenuItem()
        Me.miRowAverage = New System.Windows.Forms.MenuItem()
        Me.miDisplayFrequencyDistribution = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.miExportToExcelMenu = New System.Windows.Forms.MenuItem()
        Me.miExportToExcel = New System.Windows.Forms.MenuItem()
        Me.miAutoFitCols = New System.Windows.Forms.MenuItem()
        Me.miAutoFitRows = New System.Windows.Forms.MenuItem()
        Me.miALternateRowColors = New System.Windows.Forms.MenuItem()
        Me.miMatchGridColors = New System.Windows.Forms.MenuItem()
        Me.miOutlineExportedCells = New System.Windows.Forms.MenuItem()
        Me.miExportToTextFile = New System.Windows.Forms.MenuItem()
        Me.miExportToHTMLTable = New System.Windows.Forms.MenuItem()
        Me.miExportToSQLScript = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.miFormatStuff = New System.Windows.Forms.MenuItem()
        Me.miFormatAsMoney = New System.Windows.Forms.MenuItem()
        Me.miFormatAsDecimal = New System.Windows.Forms.MenuItem()
        Me.miFormatAsText = New System.Windows.Forms.MenuItem()
        Me.miCenter = New System.Windows.Forms.MenuItem()
        Me.miLeft = New System.Windows.Forms.MenuItem()
        Me.miRight = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.miFontsSmaller = New System.Windows.Forms.MenuItem()
        Me.miFontsLarger = New System.Windows.Forms.MenuItem()
        Me.miHeaderFontSmaller = New System.Windows.Forms.MenuItem()
        Me.miHeaderFontLarger = New System.Windows.Forms.MenuItem()
        Me.miTitleFontSmaller = New System.Windows.Forms.MenuItem()
        Me.miTitleFontLarger = New System.Windows.Forms.MenuItem()
        Me.miSmoothing = New System.Windows.Forms.MenuItem()
        Me.miAutoSizeToContents = New System.Windows.Forms.MenuItem()
        Me.miAllowUserColumnResizing = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.miSortAscending = New System.Windows.Forms.MenuItem()
        Me.miSortDescending = New System.Windows.Forms.MenuItem()
        Me.miDateAsc = New System.Windows.Forms.MenuItem()
        Me.miDateDesc = New System.Windows.Forms.MenuItem()
        Me.miSortNumericAsc = New System.Windows.Forms.MenuItem()
        Me.miSortNumericDesc = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.miHideRow = New System.Windows.Forms.MenuItem()
        Me.miHideColumn = New System.Windows.Forms.MenuItem()
        Me.miShowAllRowsAndColumns = New System.Windows.Forms.MenuItem()
        Me.miSetRowColor = New System.Windows.Forms.MenuItem()
        Me.miSetColumnColor = New System.Windows.Forms.MenuItem()
        Me.miSetCellColor = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.miSearchInColumn = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.miPrintTheGrid = New System.Windows.Forms.MenuItem()
        Me.miPreviewTheGrid = New System.Windows.Forms.MenuItem()
        Me.miPageSetup = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.miTearColumnAway = New System.Windows.Forms.MenuItem()
        Me.miMultipleColumnTearAway = New System.Windows.Forms.MenuItem()
        Me.miArrangeTearAways = New System.Windows.Forms.MenuItem()
        Me.miHideColumnTearAway = New System.Windows.Forms.MenuItem()
        Me.miHideAllTearAwayColumns = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.miProperties = New System.Windows.Forms.MenuItem()
        Me.miStats = New System.Windows.Forms.MenuItem()
        Me.pBar = New System.Windows.Forms.ProgressBar()
        Me.gb1 = New System.Windows.Forms.GroupBox()
        Me.clrdlg = New System.Windows.Forms.ColorDialog()
        Me.pdoc = New System.Drawing.Printing.PrintDocument()
        Me.txtInput = New System.Windows.Forms.TextBox()
        Me.cmboInput = New System.Windows.Forms.ComboBox()
        Me._TTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.gb1.SuspendLayout()
        Me.SuspendLayout()
        '
        'hs
        '
        Me.hs.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.hs.Location = New System.Drawing.Point(0, 138)
        Me.hs.Name = "hs"
        Me.hs.Size = New System.Drawing.Size(528, 12)
        Me.hs.TabIndex = 1
        Me.hs.Visible = False
        '
        'vs
        '
        Me.vs.Dock = System.Windows.Forms.DockStyle.Right
        Me.vs.Location = New System.Drawing.Point(516, 0)
        Me.vs.Name = "vs"
        Me.vs.Size = New System.Drawing.Size(12, 138)
        Me.vs.TabIndex = 0
        Me.vs.Visible = False
        '
        'txtHandler
        '
        Me.txtHandler.Location = New System.Drawing.Point(0, 0)
        Me.txtHandler.Name = "txtHandler"
        Me.txtHandler.Size = New System.Drawing.Size(12, 20)
        Me.txtHandler.TabIndex = 2
        Me.txtHandler.Text = "TextBox1"
        '
        'menu
        '
        Me.menu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miCopyCellToClipboard, Me.MenuItem7, Me.MenuItem5, Me.MenuItem9, Me.miExportToExcelMenu, Me.miExportToTextFile, Me.miExportToHTMLTable, Me.miExportToSQLScript, Me.MenuItem8, Me.miFormatStuff, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem10, Me.miSearchInColumn, Me.MenuItem1, Me.miPrintTheGrid, Me.miPreviewTheGrid, Me.miPageSetup, Me.MenuItem12, Me.miTearColumnAway, Me.miMultipleColumnTearAway, Me.miArrangeTearAways, Me.miHideColumnTearAway, Me.miHideAllTearAwayColumns, Me.MenuItem11, Me.miProperties, Me.miStats})
        '
        'miCopyCellToClipboard
        '
        Me.miCopyCellToClipboard.Index = 0
        Me.miCopyCellToClipboard.Text = "Copy Cell To Clipboard"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miSumColumn, Me.miSumRow, Me.miMaxCol, Me.miMaxRow, Me.miMinCol, Me.miMinRow, Me.miColAverage, Me.miRowAverage, Me.miDisplayFrequencyDistribution})
        Me.MenuItem5.Text = "Math"
        '
        'miSumColumn
        '
        Me.miSumColumn.Index = 0
        Me.miSumColumn.Text = "Sum Column"
        '
        'miSumRow
        '
        Me.miSumRow.Index = 1
        Me.miSumRow.Text = "Sum Row"
        '
        'miMaxCol
        '
        Me.miMaxCol.Index = 2
        Me.miMaxCol.Text = "Max In Column"
        '
        'miMaxRow
        '
        Me.miMaxRow.Index = 3
        Me.miMaxRow.Text = "Max In Row"
        '
        'miMinCol
        '
        Me.miMinCol.Index = 4
        Me.miMinCol.Text = "Min In Column"
        '
        'miMinRow
        '
        Me.miMinRow.Index = 5
        Me.miMinRow.Text = "Min In Row"
        '
        'miColAverage
        '
        Me.miColAverage.Index = 6
        Me.miColAverage.Text = "Column Average"
        '
        'miRowAverage
        '
        Me.miRowAverage.Index = 7
        Me.miRowAverage.Text = "Row Average"
        '
        'miDisplayFrequencyDistribution
        '
        Me.miDisplayFrequencyDistribution.Index = 8
        Me.miDisplayFrequencyDistribution.Text = "Frequency Distribution"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 3
        Me.MenuItem9.Text = "-"
        '
        'miExportToExcelMenu
        '
        Me.miExportToExcelMenu.Index = 4
        Me.miExportToExcelMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miExportToExcel, Me.miAutoFitCols, Me.miAutoFitRows, Me.miALternateRowColors, Me.miMatchGridColors, Me.miOutlineExportedCells})
        Me.miExportToExcelMenu.Text = "Export To Excel"
        '
        'miExportToExcel
        '
        Me.miExportToExcel.Index = 0
        Me.miExportToExcel.Text = "Export To Excel"
        '
        'miAutoFitCols
        '
        Me.miAutoFitCols.Index = 1
        Me.miAutoFitCols.Text = "AutoFit Excel Columns"
        '
        'miAutoFitRows
        '
        Me.miAutoFitRows.Index = 2
        Me.miAutoFitRows.Text = "AutoFit Excel Rows"
        '
        'miALternateRowColors
        '
        Me.miALternateRowColors.Index = 3
        Me.miALternateRowColors.Text = "Alternate Row Colors"
        '
        'miMatchGridColors
        '
        Me.miMatchGridColors.Index = 4
        Me.miMatchGridColors.Text = "Match Grid Colors"
        '
        'miOutlineExportedCells
        '
        Me.miOutlineExportedCells.Index = 5
        Me.miOutlineExportedCells.Text = "Outline Exported Grid Cells"
        '
        'miExportToTextFile
        '
        Me.miExportToTextFile.Index = 5
        Me.miExportToTextFile.Text = "Export To a Text File"
        '
        'miExportToHTMLTable
        '
        Me.miExportToHTMLTable.Index = 6
        Me.miExportToHTMLTable.Text = "Export To an HTML Table"
        '
        'miExportToSQLScript
        '
        Me.miExportToSQLScript.Index = 7
        Me.miExportToSQLScript.Text = "Export To an SQL Script"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 8
        Me.MenuItem8.Text = "-"
        '
        'miFormatStuff
        '
        Me.miFormatStuff.Index = 9
        Me.miFormatStuff.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miFormatAsMoney, Me.miFormatAsDecimal, Me.miFormatAsText, Me.miCenter, Me.miLeft, Me.miRight})
        Me.miFormatStuff.Text = "Format Functions"
        '
        'miFormatAsMoney
        '
        Me.miFormatAsMoney.Index = 0
        Me.miFormatAsMoney.Text = "Format As Money"
        '
        'miFormatAsDecimal
        '
        Me.miFormatAsDecimal.Index = 1
        Me.miFormatAsDecimal.Text = "Format as Decimal"
        '
        'miFormatAsText
        '
        Me.miFormatAsText.Index = 2
        Me.miFormatAsText.Text = "Format as Text"
        '
        'miCenter
        '
        Me.miCenter.Index = 3
        Me.miCenter.Text = "Center"
        '
        'miLeft
        '
        Me.miLeft.Index = 4
        Me.miLeft.Text = "Left"
        '
        'miRight
        '
        Me.miRight.Index = 5
        Me.miRight.Text = "Right"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 10
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miFontsSmaller, Me.miFontsLarger, Me.miHeaderFontSmaller, Me.miHeaderFontLarger, Me.miTitleFontSmaller, Me.miTitleFontLarger, Me.miSmoothing, Me.miAutoSizeToContents, Me.miAllowUserColumnResizing})
        Me.MenuItem2.Text = "Settings Functions"
        '
        'miFontsSmaller
        '
        Me.miFontsSmaller.Index = 0
        Me.miFontsSmaller.Text = "Grid Fonts Smaller"
        '
        'miFontsLarger
        '
        Me.miFontsLarger.Index = 1
        Me.miFontsLarger.Text = "Grid Fonts Larger"
        '
        'miHeaderFontSmaller
        '
        Me.miHeaderFontSmaller.Index = 2
        Me.miHeaderFontSmaller.Text = "Header Fonts Smaller"
        '
        'miHeaderFontLarger
        '
        Me.miHeaderFontLarger.Index = 3
        Me.miHeaderFontLarger.Text = "Header Fonts Larger"
        '
        'miTitleFontSmaller
        '
        Me.miTitleFontSmaller.Index = 4
        Me.miTitleFontSmaller.Text = "Title Font Smaller"
        '
        'miTitleFontLarger
        '
        Me.miTitleFontLarger.Index = 5
        Me.miTitleFontLarger.Text = "Title Font Larger"
        '
        'miSmoothing
        '
        Me.miSmoothing.Index = 6
        Me.miSmoothing.Text = "Smoothing"
        '
        'miAutoSizeToContents
        '
        Me.miAutoSizeToContents.Index = 7
        Me.miAutoSizeToContents.Text = "Auto Size To Contents"
        '
        'miAllowUserColumnResizing
        '
        Me.miAllowUserColumnResizing.Index = 8
        Me.miAllowUserColumnResizing.Text = "Allow User Column Resizing"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 11
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miSortAscending, Me.miSortDescending, Me.miDateAsc, Me.miDateDesc, Me.miSortNumericAsc, Me.miSortNumericDesc})
        Me.MenuItem3.Text = "Sort"
        '
        'miSortAscending
        '
        Me.miSortAscending.Index = 0
        Me.miSortAscending.Text = "Ascii Ascending"
        '
        'miSortDescending
        '
        Me.miSortDescending.Index = 1
        Me.miSortDescending.Text = "Ascii Descending"
        '
        'miDateAsc
        '
        Me.miDateAsc.Index = 2
        Me.miDateAsc.Text = "Date Ascending"
        '
        'miDateDesc
        '
        Me.miDateDesc.Index = 3
        Me.miDateDesc.Text = "Date Descending"
        '
        'miSortNumericAsc
        '
        Me.miSortNumericAsc.Index = 4
        Me.miSortNumericAsc.Text = "Numeric Ascending"
        '
        'miSortNumericDesc
        '
        Me.miSortNumericDesc.Index = 5
        Me.miSortNumericDesc.Text = "Numeric Descending"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 12
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.miHideRow, Me.miHideColumn, Me.miShowAllRowsAndColumns, Me.miSetRowColor, Me.miSetColumnColor, Me.miSetCellColor})
        Me.MenuItem4.Text = "Row and Column Options"
        '
        'miHideRow
        '
        Me.miHideRow.Index = 0
        Me.miHideRow.Text = "Hide Row"
        '
        'miHideColumn
        '
        Me.miHideColumn.Index = 1
        Me.miHideColumn.Text = "Hide Column"
        '
        'miShowAllRowsAndColumns
        '
        Me.miShowAllRowsAndColumns.Index = 2
        Me.miShowAllRowsAndColumns.Text = "Show All Rows and Columns"
        '
        'miSetRowColor
        '
        Me.miSetRowColor.Index = 3
        Me.miSetRowColor.Text = "Color Row"
        '
        'miSetColumnColor
        '
        Me.miSetColumnColor.Index = 4
        Me.miSetColumnColor.Text = "Color Column"
        '
        'miSetCellColor
        '
        Me.miSetCellColor.Index = 5
        Me.miSetCellColor.Text = "Color Cell"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 13
        Me.MenuItem10.Text = "-"
        '
        'miSearchInColumn
        '
        Me.miSearchInColumn.Index = 14
        Me.miSearchInColumn.Text = "Find In Column"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 15
        Me.MenuItem1.Text = "-"
        '
        'miPrintTheGrid
        '
        Me.miPrintTheGrid.Index = 16
        Me.miPrintTheGrid.Text = "Print The Grids Contents"
        '
        'miPreviewTheGrid
        '
        Me.miPreviewTheGrid.Index = 17
        Me.miPreviewTheGrid.Text = "Preview The Grids Contents"
        '
        'miPageSetup
        '
        Me.miPageSetup.Index = 18
        Me.miPageSetup.Text = "Page and Printer Setup"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 19
        Me.MenuItem12.Text = "-"
        '
        'miTearColumnAway
        '
        Me.miTearColumnAway.Index = 20
        Me.miTearColumnAway.Text = "Tear Column Away"
        '
        'miMultipleColumnTearAway
        '
        Me.miMultipleColumnTearAway.Index = 21
        Me.miMultipleColumnTearAway.Text = "Tear Multiple Columns Away"
        '
        'miArrangeTearAways
        '
        Me.miArrangeTearAways.Index = 22
        Me.miArrangeTearAways.Text = "Arrange Open Tear Away Columns"
        '
        'miHideColumnTearAway
        '
        Me.miHideColumnTearAway.Index = 23
        Me.miHideColumnTearAway.Text = "Hide Column Tear Away"
        '
        'miHideAllTearAwayColumns
        '
        Me.miHideAllTearAwayColumns.Index = 24
        Me.miHideAllTearAwayColumns.Text = "Hide All Tear Away Columns"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 25
        Me.MenuItem11.Text = "-"
        '
        'miProperties
        '
        Me.miProperties.Index = 26
        Me.miProperties.Text = "Properties"
        '
        'miStats
        '
        Me.miStats.Index = 27
        Me.miStats.Text = "Stats"
        '
        'pBar
        '
        Me.pBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pBar.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.pBar.Location = New System.Drawing.Point(4, 16)
        Me.pBar.Name = "pBar"
        Me.pBar.Size = New System.Drawing.Size(392, 16)
        Me.pBar.TabIndex = 3
        Me.pBar.Visible = False
        '
        'gb1
        '
        Me.gb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gb1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.gb1.Controls.Add(Me.pBar)
        Me.gb1.Location = New System.Drawing.Point(56, 8)
        Me.gb1.Name = "gb1"
        Me.gb1.Size = New System.Drawing.Size(400, 48)
        Me.gb1.TabIndex = 4
        Me.gb1.TabStop = False
        Me.gb1.Text = "Progress..."
        Me.gb1.Visible = False
        '
        'pdoc
        '
        '
        'txtInput
        '
        Me.txtInput.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtInput.Location = New System.Drawing.Point(0, 20)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.Size = New System.Drawing.Size(12, 13)
        Me.txtInput.TabIndex = 6
        Me.txtInput.Text = "TextBox1"
        Me.txtInput.Visible = False
        '
        'cmboInput
        '
        Me.cmboInput.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboInput.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboInput.ItemHeight = 13
        Me.cmboInput.Location = New System.Drawing.Point(0, 36)
        Me.cmboInput.Name = "cmboInput"
        Me.cmboInput.Size = New System.Drawing.Size(36, 21)
        Me.cmboInput.TabIndex = 7
        Me.cmboInput.Visible = False
        '
        'TAIGridControl
        '
        Me.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.Controls.Add(Me.cmboInput)
        Me.Controls.Add(Me.txtInput)
        Me.Controls.Add(Me.gb1)
        Me.Controls.Add(Me.txtHandler)
        Me.Controls.Add(Me.vs)
        Me.Controls.Add(Me.hs)
        Me.Name = "TAIGridControl"
        Me.Size = New System.Drawing.Size(528, 150)
        Me.gb1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region " Event Templates "

    '	CellClicked
    ''' <summary>
    ''' Raised whenever a cell is clicked. Coordinates designated by RowClicked/ColumnClicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="RowClicked"></param>
    ''' <param name="ColumnClicked"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a cell is clicked designated by RowClicked, ColumnClicked")> _
    Public Event CellClicked(ByVal sender As Object, ByVal RowClicked As Integer, ByVal ColumnClicked As Integer)

    '	CellDoubleClicked
    ''' <summary>
    ''' Raised whenever a cell is doubleclicked. Coordinates designated by RowClicked/ColumnClicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="RowClicked"></param>
    ''' <param name="ColumnClicked"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a cell is doubleclicked designated by RowClicked, ColumnClicked")> _
    Public Event CellDoubleClicked(ByVal sender As Object, ByVal RowClicked As Integer, ByVal ColumnClicked As Integer)

    '	CellEdited
    ''' <summary>
    ''' Raised whenever a cell is edited by the user, if cell editing is turned on correctly.
    ''' RowClicked/ColumnClicked designated which cell was edited. oldval/newval designated the previous contents and
    ''' the new contents respectively
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="RowClicked"></param>
    ''' <param name="ColumnClicked"></param>
    ''' <param name="oldval"></param>
    ''' <param name="newval"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a cell Edited by the user")> _
    Public Event CellEdited(ByVal sender As Object, ByVal RowClicked As Integer, ByVal ColumnClicked As Integer, ByVal oldval As String, ByVal newval As String)

    '	RowSelected
    ''' <summary>
    ''' Raised whenever a row is selected with the mouse or the keyboard. RowSelected designated which row.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="RowSelected"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a row is selected with mouse or keyboard. Rowselected is returned")> _
    Public Event RowSelected(ByVal sender As Object, ByVal RowSelected As Integer)

    '	RowDeSelected
    ''' <summary>
    ''' Raised whenever a row is deselected with the mouse or the kayboard. RowDeselected designated which row
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="RowDeselected"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a row is deselected with mouse or keyboard. RowSelected is returned")> _
       Public Event RowDeSelected(ByVal sender As Object, ByVal RowDeselected As Integer)

    '	PartialSelection
    ''' <summary>
    ''' Raised whenever a populategrid from database call exceeded the set threshold of records
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a populategrid from database exceeds a set threshold of records")> _
    Public Event PartialSelection(ByVal sender As Object)

    '   TooManyRecords
    ''' <summary>
    ''' Raised whenever a populategrid from database call gets to many records that the grid cannot handle. After the rewrite
    ''' in 2005 this event is exceedintgly difficult to fire as the grid can now handle millions of records at a time.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a populategrid from database gets so many records that the bitmap becomes to big")> _
    Public Event TooManyRecords(ByVal sender As Object)

    '   TooManyFields
    ''' <summary>
    ''' Raised whenever a populategrid from database call gets to many records that the grid cannot handle. After the rewrite
    ''' in 2005 this event is exceedintgly difficult to fire as the grid can now handle millions of records at a time.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever a populategrid from database gets so many records that the bitmap becomes to big")> _
    Public Event TooManyFields(ByVal sender As Object)

    '	StartedDatabasePopulateOperation
    ''' <summary>
    ''' Raised when the grid starts a PopulateGridWData call from a supported data source (SQL,OLE,ODBC,DATATABLE etc.)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised whenever the grid is starting to PopulateGridWData")> _
    Public Event StartedDatabasePopulateOperation(ByVal sender As Object)

    '	FinishedDatabasePopulateOperation
    ''' <summary>
    ''' Raised when the grid finishes a PopulateGridWData call from a supported data source (SQL,OLE,ODBC,DATATABLE etc.)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the grid is finished PopulatingGridWData")> _
    Public Event FinishedDatabasePopulateOperation(ByVal sender As Object)

    '	Column Resized
    ''' <summary>
    ''' Raised when the user resizes a column using the mouse. ColumnIndex designated the column being resized
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ColumnIndex"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user Resizes as column using the mouse")> _
    Public Event ColumnResized(ByVal sender As Object, ByVal ColumnIndex As Integer)

    '	Column Selected
    ''' <summary>
    ''' Raised when the user selectes a column using the mouse. ColumnIndex designated the column being resized
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ColumnIndex"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user selects as column using the mouse")> _
    Public Event ColumnSelected(ByVal sender As Object, ByVal ColumnIndex As Integer)

    '	Column DeSelected
    ''' <summary>
    ''' Raised when the user selectes a column using the mouse. ColumnIndex designated the column being resized
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="OldColumnIndex"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user deselects as column using the mouse")> _
    Public Event ColumnDeSelected(ByVal sender As Object, ByVal OldColumnIndex As Integer)


    '   GridResorted
    ''' <summary>
    ''' Raised when the user resorts the grids contents on a chosen column index. ColumnIndex is the chosen column.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ColumnIndex"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user Sorts the grid on a given column ColumnIndex is that column")> _
    Public Event GridResorted(ByVal sender As Object, ByVal ColumnIndex As Integer)

    '   KeypressedInGrid
    ''' <summary>
    ''' Raised when the user presses a key on the keyboard while the grid has focus and a cell is not being edited.
    ''' The Keycode parameter is of the type Keys
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="keyCode"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user Presses any keyboard key returns a type of Keys")> _
    Public Event KeyPressedInGrid(ByVal sender As Object, ByVal keyCode As Keys)

    '   RightMouseButtonInGrid
    ''' <summary>
    ''' Raised when the user presses the rightmousebutton in the grid and the grid is not doing any of its Popup context
    ''' menus.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user selects the rightmousebutton in a grid and the grid is NOT doing POPUP menus")> _
    Public Event RightMouseButtonInGrid(ByVal sender As Object)

    '   GridHover
    ''' <summary>
    ''' Raised as the user loiters over the rendered grid contents with the mouse. Row/Col esignated the cell being
    ''' hovered over, Item indicates the contents of 6the cell being hovered over. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <param name="Item"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user is hovering over the grid itself Not the grid container just the rendered grid")> _
    Public Event GridHover(ByVal sender As Object, ByVal row As Integer, ByVal col As Integer, ByVal Item As String)

    '   GridHoverLeave
    ''' <summary>
    ''' Raised as the user moved the mouse off of the grid rendered contens after hovering over those contents previously
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Raised when the user is hovering over the grid itself Not the grid container just the rendered grid")> _
    Public Event GridHoverleave(ByVal sender As Object)

#End Region

#Region " Public Types "
    ''' <summary>
    ''' Enmeration for selecting preset color schemes used to configures the theme of the grids display
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TaiGridColorSchemes As Integer
        _Default = 0
        _Business
        _Technical
        _Fancy
        _Colorful1
        _Colorful2
    End Enum

#End Region

#Region " Properties "

    '	AllowTearAwayFunctionality
    ''' <summary>
    ''' Allow or Disallow the tearaway a column functionality within the grid itself.
    ''' Column tearaways allow for removing a columns contents to a seperate window that floats outside the
    ''' boundarys of the grids containers. This functionality might prove useful in some circumstances but may
    ''' also confuse the display for some users. This setting will turn on or off the availability of this function.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Tearaway menu items of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowTearAwayFunctionality() As Boolean
        Get
            Return _AllowTearAwayFuncionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowTearAwayFuncionality = Value
        End Set
    End Property

    '	AllowExcelFunctionality
    ''' <summary>
    ''' Allow or disallow the ability to export the grids contents to excel via the Context menu. 
    ''' Heavily used with some reporting applications where numerics are displayed in aggregate, 
    ''' other uses of the grid in items where personal data are displayed might
    ''' necessitate turning the functionality off for privacy/Hipaa reasons. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Excel menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowExcelFunctionality() As Boolean
        Get
            Return _AllowExcelFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowExcelFunctionality = Value
        End Set
    End Property

    '	AllowTextFunctionality
    ''' <summary>
    ''' Allows or disallows the text menu functionality of the grids context menu. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Text menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowTextFunctionality() As Boolean
        Get
            Return _AllowTextFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowTextFunctionality = Value
        End Set
    End Property

    '	AllowHTMLFunctionality
    ''' <summary>
    ''' Allows or disallows the HTML menu on the grids context menu.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the HTML menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowHTMLFunctionality() As Boolean
        Get
            Return _AllowHTMLFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowHTMLFunctionality = Value
        End Set
    End Property

    '	AllowSQLFunctionality
    ''' <summary>
    ''' Allows or disallows the SQL menu off of the grids context menu.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the SQL menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowSQLFunctionality() As Boolean
        Get
            Return _AllowSQLScriptFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowSQLScriptFunctionality = Value
        End Set
    End Property

    '	AllowMathFunctionality
    ''' <summary>
    ''' Allows or disallows the Math submenu off of the grids context menu.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Math menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowMathFunctionality() As Boolean
        Get
            Return _AllowMathFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowMathFunctionality = Value
        End Set
    End Property

    '	AllowFormatFunctionality
    ''' <summary>
    ''' Allows or disallows the Format submenu off of the grids context menu. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Format menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowFormatFunctionality() As Boolean
        Get
            Return _AllowFormatFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowFormatFunctionality = Value
        End Set
    End Property

    '	AllowSettingsFunctionality
    ''' <summary>
    ''' Allows or disallows the settings submenu off of the grids context menu.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Settings menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowSettingsFunctionality() As Boolean
        Get
            Return _AllowSettingsFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowSettingsFunctionality = Value
        End Set
    End Property

    '	AllowSortFunctionality
    ''' <summary>
    ''' Allows or disallows the Sort submenu off of the grids context menu.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the Sort menu of the builtin context menu"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowSortFunctionality() As Boolean
        Get
            Return _AllowSortFunctionality
        End Get
        Set(ByVal Value As Boolean)
            _AllowSortFunctionality = Value
        End Set
    End Property

    '	AllowColumnSelection
    ''' <summary>
    ''' Allows or disallows the selection of a column by clicking the header of a column with the mouse
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow selection of a column visually by single clicking on the header of that column"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowColumnSelection() As Boolean
        Get
            Return _AllowColumnSelection
        End Get
        Set(ByVal Value As Boolean)
            _AllowColumnSelection = Value

            Me.Invalidate()
        End Set
    End Property

    '	AllowControlKeyMenuPopup
    ''' <summary>
    ''' Allows or disallows the pulling up of the grids context menu my pressintg the ctrl key while right mousebuttoning
    ''' over the grid itself. This allows programs that are hosting the grids to create their own context menus but to
    ''' still have the grids context menus available via the ctrl/right mousebutton combination.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the menu from poping up on a CTRL menubutton."), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowControlKeyMenuPopup() As Boolean
        Get
            Return _AllowControlKeyMenuPopup
        End Get
        Set(ByVal Value As Boolean)
            _AllowControlKeyMenuPopup = Value

            Me.Invalidate()
        End Set
    End Property

    '	AllowInGridEdits
    ''' <summary>
    ''' Allows or disallows the editing of grids contents. This is not an all or nothing process. The developer hase to turn this
    ''' on and explicitly set the columns where they want to allow editing in order for in grid edits to function.
    ''' Alternately they might elect to restrict editing of a cells contents to s list of available selections
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the editing of the grid contents at the column level"), _
    DefaultValue(GetType(Boolean), "False")> _
    Public Property AllowInGridEdits() As Boolean
        Get
            Return _AllowInGridEdits
        End Get
        Set(ByVal Value As Boolean)
            _AllowInGridEdits = Value
        End Set
    End Property

    '	AllowMultipleRowSelections
    ''' <summary>
    ''' Allows or disallows the ability of the user to select more than a single row in the grid at one time
    ''' via the standard CTRL/SHIFT key click mechanism used in the Windows OS. The rows selected will then
    ''' be exposed via the <c>SelectedRows</c> collection
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the selection of Multiple rows in the grid with the CTRL key "), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowMultipleRowSelections() As Boolean
        Get
            Return _AllowMultipleRowSelections
        End Get
        Set(ByVal Value As Boolean)
            _AllowMultipleRowSelections = Value
        End Set
    End Property

    '	AllowPopupMenu
    ''' <summary>
    ''' Allow or disallow the grids own context menu to appear via the right mousebutton
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the builtin popup menu for font selection and sizing to occur"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowPopupMenu() As Boolean
        Get
            Return _AllowPopupMenu
        End Get
        Set(ByVal Value As Boolean)
            _AllowPopupMenu = Value
        End Set
    End Property

    '	AllowRowSelection
    ''' <summary>
    ''' Allow or disallow the ability to select a single or multiple rows in the grid with the mouse,
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow selection of a Row or Multiple Rows visually by single clicking in the Row itself"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowRowSelection() As Boolean
        Get
            Return _AllowRowSelection
        End Get
        Set(ByVal Value As Boolean)
            _AllowRowSelection = Value

            Me.Invalidate()
        End Set
    End Property

    '   AllowWhiteSpaceInCells
    ''' <summary>
    ''' Will allow/disallow Whitespace in cells (newlines and what not)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow Whitespace in cells (newlines and what not)"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property AllowWhiteSpaceInCells() As Boolean
        Get
            Return _AllowWhiteSpaceInCells
        End Get
        Set(ByVal Value As Boolean)
            _AllowWhiteSpaceInCells = Value

            Me.Invalidate()
        End Set
    End Property


    '	AllowUserColumnResizing
    ''' <summary>
    ''' Allow or disallow user column resizing with the mouse
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the user to resize a column"), _
    DefaultValue(GetType(Boolean), "True")> _
        Public Property AllowUserColumnResizing() As Boolean
        Get
            Return _AllowUserColumnResizing
        End Get
        Set(ByVal Value As Boolean)
            _AllowUserColumnResizing = Value
        End Set
    End Property

    '	AlternateColoration
    ''' <summary>
    ''' Turns on or off the Alternate coloration mode of the grids display where it will alternate the background
    ''' color of the rows inserted between <c>AlternateColorationAltColor</c> and <c>AlternateColorationBaseColor</c>
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will trun on/off the alternate coloration mode of the grid. Rows will alternate between defined backcolor and the defined alternatecolor")> _
        Public Property AlternateColoration() As Boolean
        Get
            Return _alternateColorationMode
        End Get
        Set(ByVal Value As Boolean)
            _alternateColorationMode = Value

            'Me.TAIGCanvas_Paint(Me, _
            '              New System.Windows.Forms.PaintEventArgs(Me.CreateGraphics, _
            '              New System.Drawing.Rectangle(0, 0, Me.Width, Me.Height)))

            'Me.TAIGPanel.Invalidate()
            'Me.Refresh()

        End Set
    End Property

    '	AlternateColorationAltColor
    ''' <summary>
    ''' One of the colors used when the grid is rendering in AlternateColoration mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the alternate color for the alternate Coloration mode of operation")> _
        Public Property AlternateColorationAltColor() As Color
        Get
            Return _alternateColorationALTColor
        End Get
        Set(ByVal Value As Color)
            _alternateColorationALTColor = Value
        End Set
    End Property

    '	AlternateColorationBaseColor
    ''' <summary>
    '''  One of the colors used when the grid is rendering in AlternateColoration mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the base color for the alternate Coloration mode of operation")> _
        Public Property AlternateColorationBaseColor() As Color
        Get
            Return _alternateColorationBaseColor
        End Get
        Set(ByVal Value As Color)
            _alternateColorationBaseColor = Value
        End Set
    End Property

    '	Antialias
    ''' <summary>
    ''' Turns on or off the smoothing mode of the grids textual rendering engine
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Turns on/off the antialias mode of the grids rendering engine"), _
    DefaultValue(GetType(Boolean), "False")> _
        Public Property Antialias() As Boolean
        Get
            Return _antialias
        End Get
        Set(ByVal Value As Boolean)
            _antialias = Value

            Me.Invalidate()

        End Set
    End Property

    '	AutoSizeCellsToContents
    ''' <summary>
    ''' Allows or disallows the grids rendering engine to automagically resize the grids row and column metrics
    ''' to accomodate the contents being inserted into the grid manually or via one of the PopulateFromDatabase calls.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the grid to automatically adjust the grid cells heigh and width to match the textual contents of the cells"), _
    DefaultValue(GetType(Boolean), "False")> _
    Public Property AutoSizeCellsToContents() As Boolean
        Get
            Return _AutosizeCellsToContents
        End Get
        Set(ByVal Value As Boolean)
            _AutosizeCellsToContents = Value
            If Value Then
                _AutoSizeAlreadyCalculated = False
                _AutoSizeSemaphore = True
                DoAutoSizeCheck(Me.CreateGraphics)
            End If
            Me.Invalidate()
        End Set
    End Property

    '	AutoFocus
    ''' <summary>
    ''' Allows or disallows the grids ability to automagically gain focus as the user mouses over the grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow/disallow the grid to automatically gain focus on mouseover"), _
    DefaultValue(GetType(Boolean), "False")> _
    Public Property AutoFocus() As Boolean
        Get
            Return _AutoFocus
        End Get
        Set(ByVal Value As Boolean)
            _AutoFocus = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or Sets the background color for the control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the background color for the control"), _
    DefaultValue(GetType(System.Drawing.Color), "GradientActiveCaption")> _
    Public Overrides Property BackColor() As System.Drawing.Color
        Get
            Return MyBase.BackColor
        End Get
        Set(ByVal value As System.Drawing.Color)
            MyBase.BackColor = value
        End Set
    End Property

    '	BorderColor
    ''' <summary>
    ''' The color used to render the border of the grid itself when <c>BorderStyle</c> = something
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the border color for the drawn border on Borderstyle = something")> _
        Public Property BorderColor() As Color
        Get
            Return _BorderColor
        End Get
        Set(ByVal Value As Color)
            _BorderColor = Value
        End Set
    End Property

    '	BorderStyle
    ''' <summary>
    ''' The border style use to draw the grid border itself
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the border style for the grids container object or frame")> _
    Public Overloads Property BorderStyle() As Windows.Forms.BorderStyle
        Get
            Return _BorderStyle
        End Get
        Set(ByVal Value As Windows.Forms.BorderStyle)
            _BorderStyle = Value
        End Set
    End Property

    '	ShowProgressBar
    ''' <summary>
    ''' Allows or disallows the display f the progress bar across the top of the grid itself when long database population
    ''' processes are underway.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow the grid to show a small progress bar along its top edges on long database populate methods"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property ShowProgressBar() As Boolean
        Get
            Return _ShowProgressBar
        End Get
        Set(ByVal Value As Boolean)
            _ShowProgressBar = Value
        End Set
    End Property

    '	ShowExcelExportMessage
    ''' <summary>
    ''' Allows or disallows the display of a topmost windows signaling to the end users that the grid
    ''' is sending it's content to excel. Because messaging excel is sometime a lengthy process the display of
    ''' the dialog might prove useful in those situations. Messaging excel though is an inherantly messy process
    ''' where a user interacting with a different instance of excel might confuse the system and make the
    ''' export process fail. In these cases the dialog might also prove useful in that the user can be instructed
    ''' 'Hands off' with the dialog is visable. It can however be confusing when this dialog is on top of everything.
    ''' As always its use might prove useful or it might not depending on environmental factors.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will allow the grid to show a small window topmost on exporting to excel with status information"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property ShowExcelExportMessage() As Boolean
        Get
            Return _ShowExcelExportMessage
        End Get
        Set(ByVal Value As Boolean)
            _ShowExcelExportMessage = Value
        End Set
    End Property

    '	Cols
    ''' <summary>
    ''' Get or Sets the number of columns in the grid itself
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("How many columns are in the current grid")> _
    Public Property Cols() As Integer
        Get
            Return _cols
        End Get
        Set(ByVal Value As Integer)
            SetCols(Value)
            Me.Invalidate()
        End Set
    End Property

    '	CellAlignment
    ''' <summary>
    ''' Get or sets the alignment of the textual element contained at Grid coordinates R (row) and C (col)
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the alignment of the textual element of a cell at Row R and Col C")> _
        Public Property CellAlignment(ByVal r As Integer, ByVal c As Integer) As StringFormat
        Get
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                Return _DefaultStringFormat
            Else
                Return _gridCellAlignmentList(_gridCellAlignment(r, c))
            End If
        End Get
        Set(ByVal Value As StringFormat)
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                ' trouble here
            Else
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(Value)
                Me.Invalidate()
            End If

        End Set
    End Property

    '	CellBackColor
    ''' <summary>
    ''' Gets or sets the background color of the specificied cell at R (row) and C (col)
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the background color for the cell at Row R and Col C")> _
        Public Property CellBackColor(ByVal r As Integer, ByVal c As Integer) As Brush
        Get
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                Return New SolidBrush(Color.AntiqueWhite)
            Else
                Return _gridBackColorList(_gridBackColor(r, c))
            End If
        End Get
        Set(ByVal Value As Brush)
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                ' trouble here
            Else
                _gridBackColor(r, c) = GetGridBackColorListEntry(Value)
                Me.Invalidate()
            End If
        End Set
    End Property

    '	ColBackColorEdit
    ''' <summary>
    ''' Gets or sets the background color used to render a cell when that cell is in edit mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the background color for the cell when its being edited")> _
        Public Property ColBackColorEdit() As Color
        Get
            Return _colEditableTextBackColor
        End Get
        Set(ByVal Value As Color)
            _colEditableTextBackColor = Value
            Me.txtInput.BackColor = Value
        End Set
    End Property

    '	CellFont
    ''' <summary>
    ''' Gets or sets the font used to render a the cells contents. The cell is designated by 
    ''' its cordinates R (row) and C (col)
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the font use for the textual elements of the cell at Row R and Col C")> _
        Public Property CellFont(ByVal r As Integer, ByVal c As Integer) As Font
        Get
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                Return _gridCellFontsList(0)
            Else
                Return _gridCellFontsList(_gridCellFonts(r, c))
            End If
        End Get
        Set(ByVal Value As Font)
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                ' trouble here
            Else
                _gridCellFonts(r, c) = Me.GetGridCellFontListEntry(Value)
                Me.Invalidate()
            End If
        End Set
        'Get
        '    If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
        '        Return _DefaultCellFont
        '    Else
        '        Return _gridCellFonts(r, c)
        '    End If
        'End Get
        'Set(ByVal Value As Font)
        '    If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
        '        ' trouble here
        '    Else
        '        _gridCellFonts(r, c) = Value
        '        Me.Invalidate()
        '    End If
        'End Set
    End Property

    '	CellForeColor
    ''' <summary>
    ''' Gets or sets the foreground color used to render a cell at coordinated R (row) and C (col)
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the forground color for the text rendered for the cell at Row R and Col C")> _
        Public Property CellForeColor(ByVal r As Integer, ByVal c As Integer) As Pen
        Get
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                Return _gridForeColorList(0) 'New Pen(_DefaultForeColor.Blue)
            Else
                Return _gridForeColorList(_gridForeColor(r, c))
            End If
        End Get
        Set(ByVal Value As Pen)
            If r > _rows - 1 Or c > _cols - 1 Or r < 0 Or c < 0 Then
                ' trouble here
            Else
                _gridForeColor(r, c) = GetGridForeColorListEntry(Value)
                Me.Invalidate()
            End If
        End Set
    End Property

    '	CellOutlines
    ''' <summary>
    ''' Allows or disallows the rendering engins outlining of cells as it draws their contents
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Turns Cell outlining on or off")> _
        Public Property CellOutlines() As Boolean
        Get
            Return _CellOutlines
        End Get
        Set(ByVal Value As Boolean)
            _CellOutlines = Value
            Me.Invalidate()
        End Set
    End Property

    '	ColCheckBox
    ''' <summary>
    ''' Gets or sets the columns status as a boolean value where it will interpret the contents of a column
    ''' as boolean values. 1,True,Y,y,Yes,yes and other variations will be rendered as a check checkbok
    ''' 0,False,n,N,No,no and other variatons will be rendered as unchecked checkboxes. ALl other values will
    ''' be rendered as disabled checkboxes that are unchecked. If the column is editable then the grid will
    ''' manage checkbox state for you toggling the contents as the user interacts with the cells contents via
    ''' the mouse.
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the Boolean (Draw as a checkbox) status of a given column in the grid")> _
        Public Property ColCheckBox(ByVal idx As Integer) As Boolean
        Get
            If idx < 0 Or idx > _cols Then
                Return False
            Else
                Return _colboolean(idx)
            End If
        End Get
        Set(ByVal Value As Boolean)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _colboolean(idx) = Value
            End If
        End Set
    End Property

    '	ColEditable
    ''' <summary>
    ''' Gets or sets the editable status of a column at index idx.
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the editable status of a given column in the grid")> _
        Public Property ColEditable(ByVal idx As Integer) As Boolean
        Get
            If idx < 0 Or idx > _cols Then
                Return False
            Else
                Return _colEditable(idx)
            End If
        End Get
        Set(ByVal Value As Boolean)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _colEditable(idx) = Value
            End If
        End Set
    End Property

    '	RowEditable
    ''' <summary>
    ''' Gets or sets the editable status of a Row at index idx.
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the editable status of a given row in the grid")> _
    Public Property RowEditable(ByVal idx As Integer) As Boolean
        Get
            If idx < 0 Or idx > _cols Then
                Return False
            Else
                Return _rowEditable(idx)
            End If
        End Get
        Set(ByVal Value As Boolean)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _rowEditable(idx) = Value
            End If
        End Set
    End Property

    '	ColWidth
    ''' <summary>
    ''' Gets or sets the width of a column at index idx in pixels
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Returns or Sets the column width in pixels for the column at index IDX")> _
        Public Property ColWidth(ByVal idx As Integer) As Integer
        Get
            If idx < 0 Or idx > _cols Then
                Return 0
            Else
                Return _colwidths(idx)
            End If
        End Get
        Set(ByVal Value As Integer)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _AutosizeCellsToContents = False
                _colwidths(idx) = Value
                Me.Invalidate()
            End If

        End Set
    End Property

    '	ColPassword
    ''' <summary>
    ''' Gets or sets the rendering text to be used for a column at index idx to be set as a password column.
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the rendered text to be used to set a column as a password column. Col at index IDX will be rendered as a password column if this is set")> _
        Public Property ColPassword(ByVal idx As Integer) As String
        Get
            If idx < 0 Or idx > _cols Then
                Return ""
            Else
                Return _colPasswords(idx)
            End If
        End Get
        Set(ByVal Value As String)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _colPasswords(idx) = Value
                Me.Refresh()
            End If
        End Set
    End Property

    '	ColMaxCharacters
    ''' <summary>
    ''' Gets or sets the number of characters that a clumn at index idx will display before the rendering engine will
    ''' display the elipsis ... characters at the end.
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the Maximum characters a colum will display before elipsis... 0 = show entire column")> _
        Public Property ColMaxCharacters(ByVal idx As Integer) As Integer
        Get
            If idx < 0 Or idx > _cols Then
                Return 0
            Else
                Return _colMaxCharacters(idx)
            End If
        End Get
        Set(ByVal Value As Integer)
            If idx < 0 Or idx > _cols Then
                ' trouble in paradise
            Else
                _colMaxCharacters(idx) = Value

                If _AutosizeCellsToContents Then
                    _AutoSizeAlreadyCalculated = False
                    _AutoSizeSemaphore = True
                    DoAutoSizeCheck(Me.CreateGraphics)
                End If

                Me.Refresh()
            End If
        End Set
    End Property

    '	DataBaseTimeOut
    ''' <summary>
    ''' Gets or sets the time in seconds for a global database timeout value for all the Populate with database calls
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the database timeout value associated with the various populate from database methods of the grid"), _
    DefaultValue(GetType(Integer), "500")> _
        Public Property DataBaseTimeOut() As Integer
        Get
            Return _dataBaseTimeOut
        End Get
        Set(ByVal Value As Integer)
            _dataBaseTimeOut = Value
        End Set
    End Property

    '	DefaultCellFont
    ''' <summary>
    ''' Gets or sets the default font used to render cells.  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the font used for cell additions by default")> _
        Public Property DefaultCellFont() As Font
        Get
            Return _DefaultCellFont
        End Get
        Set(ByVal Value As Font)
            _DefaultCellFont = Value
        End Set
    End Property

    '	DefaultBackColor
    ''' <summary>
    ''' Gets or sets the default background color used to render cells
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the Background color use by default Cells in the grid")> _
        Public Property DefaultBackgroundColor() As System.Drawing.Color
        Get
            Return _DefaultBackColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            _DefaultBackColor = Value
            _gridBackColorList(0) = New SolidBrush(Value)
        End Set
    End Property

    '	DefaultForeColor
    ''' <summary>
    ''' Gets or sets the default foreground color user to render cells
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the color use by default for text added to grid")> _
        Public Property DefaultForegroundColor() As System.Drawing.Color
        Get
            Return _DefaultForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            _DefaultForeColor = Value
            _gridForeColorList(0) = New Pen(Value)
        End Set
    End Property

    '   Delimiter
    ''' <summary>
    ''' Gets or sets the default field delimiter used for the export to text methods
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field delimiter for export to text methods")> _
    Public Property Delimiter() As String
        Get
            Return Me._delimiter
        End Get
        Set(ByVal Value As String)
            Me._delimiter = Value
        End Set
    End Property

    ' GridEditMode
    ''' <summary>
    ''' Gets or sets the field forcing a return key on an edited cell to edit its contents or just losing focus will fire
    ''' a cell edited event.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field forcing a return key to edit or kist losing focus to edit")> _
    Public Property GridEditMode() As GridEditModes
        Get
            Return _GridEditMode
        End Get
        Set(ByVal Value As GridEditModes)
            _GridEditMode = Value
        End Set
    End Property

    ''' <summary>
    ''' When the grid is in editmode this is the column that is currently being edited
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EditModeCol() As Integer
        Get
            If _EditMode Then
                Return _EditModeCol
            Else
                Return -1
            End If
        End Get
    End Property

    ''' <summary>
    ''' When the grid is in editmode this is the row currently being edited
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EditModeRow() As Integer
        Get
            If _EditMode Then
                Return _EditModeRow
            Else
                Return -1
            End If
        End Get
    End Property

    ''' <summary>
    ''' When the grid is maintaining i set of tearaway columns this will return True false otherwise
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GridDoingTearAwayWork() As Boolean
        Get
            Return _TearAwayWork
        End Get
    End Property

    ''' <summary>
    ''' When the user brings up the context menu this will retiurn the column they were over when the menu was called up
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ColOverOnMenuButton() As Integer
        Get
            Return _ColOverOnMenuButton
        End Get
    End Property

    ''' <summary>
    ''' When the user brings up the context menu this will return the row they were over when the menu was called up
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RowOverOnMenuButton() As Integer
        Get
            Return _RowOverOnMenuButton
        End Get
    End Property

    '	GridHeaderFont
    ''' <summary>
    ''' Gets or sets the font used to render the column header
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the font used for the Grid Header by default")> _
        Public Property GridHeaderFont() As Font
        Get
            Return _GridHeaderFont
        End Get
        Set(ByVal Value As Font)
            _GridHeaderFont = Value
        End Set
    End Property

    '	GridHeaderStringFormat
    ''' <summary>
    ''' Gets or sets the formatting characteristics of the grid header line. (left,right,centered etc.)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the stringformat object used for the Grid Header by default")> _
    Public Property GridHeaderStringFormat() As StringFormat
        Get
            Return _GridHeaderStringFormat
        End Get
        Set(ByVal Value As StringFormat)
            _GridHeaderStringFormat = Value
            Me.Invalidate()
        End Set
    End Property

    '	GridHeaderVisible
    ''' <summary>
    ''' Allows or disallows the display of the grid header line
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("is the Gridheader visible or not")> _
    Public Property GridheaderVisible() As Boolean
        Get
            Return _GridHeaderVisible
        End Get
        Set(ByVal Value As Boolean)
            _GridHeaderVisible = Value
            Me.Invalidate()
        End Set
    End Property

    '	GridHeaderHeight
    ''' <summary>
    ''' Gets or sets the height of the grids header in pixels
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets how hight the grid header is drawn in pixels")> _
    Public Property GridHeaderHeight() As Integer
        Get
            Return _GridHeaderHeight
        End Get
        Set(ByVal Value As Integer)
            _GridHeaderHeight = Value
            Me.Invalidate()
        End Set
    End Property

    '	GridHeaderBackColor
    ''' <summary>
    ''' Gets or sets the background color used to render the grids header
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the color used for the Grid Header background")> _
    Public Property GridHeaderBackColor() As Color
        Get
            Return _GridHeaderBackcolor
        End Get
        Set(ByVal Value As Color)
            _GridHeaderBackcolor = Value
            Me.Invalidate()
        End Set
    End Property

    '	GridHeaderForeColor
    ''' <summary>
    ''' Gets or sets the foreground color useed to render the grids header
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the foreground color used to draw the grid header")> _
    Public Property GridHeaderForeColor() As Color
        Get
            Return _GridHeaderForecolor
        End Get
        Set(ByVal Value As Color)
            _GridHeaderForecolor = Value
            Me.Invalidate()
        End Set
    End Property

    '   GridReportOrientLandscape
    ''' <summary>
    ''' Gets or sets the grids output for reporting to be landscape mode or not.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report output to landscape mode"), _
    DefaultValue(GetType(Boolean), "False")> _
    Public Property GridReportOrientLandscape() As Boolean
        Get
            Return _gridReportOrientLandscape
        End Get
        Set(ByVal Value As Boolean)
            _gridReportOrientLandscape = Value
        End Set
    End Property

    '   GridReportOutlineCells
    ''' <summary>
    ''' Gets or sets the grids reporting engine to outline cells or not
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report to outline printed cells"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property GridReportOutlineCells() As Boolean
        Get
            Return _gridReportOutlineCells
        End Get
        Set(ByVal Value As Boolean)
            _gridReportOutlineCells = Value
        End Set
    End Property

    '   GridReportNumberPages
    ''' <summary>
    ''' Gets or sets the grids reporting engine to number the output pages or not
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report generation to number pages"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property GridReportNumberPages() As Boolean
        Get
            Return _gridReportNumberPages
        End Get
        Set(ByVal Value As Boolean)
            _gridReportNumberPages = Value
        End Set
    End Property

    '   GridReportMatchColors
    ''' <summary>
    ''' Gets or sets the grids reporting engine to attempt to match reported output coloration with the onscreen
    ''' display engines coloration scheme. Some on-screen colors dont look all that well when printed to paper,
    ''' this is especially true wjen the printer is a black and white printer and the screen representation is
    ''' full of various colors. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report generation to match grid colors"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property GridReportMatchColors() As Boolean
        Get
            Return _gridReportMatchColors
        End Get
        Set(ByVal Value As Boolean)
            _gridReportMatchColors = Value
        End Set
    End Property

    '   GridReportTitle
    ''' <summary>
    ''' Gets or sets the textual tile to apply to reported output
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report generation Title"), _
    DefaultValue(GetType(String), "")> _
    Public Property GridReportTitle() As String
        Get
            Return _gridReportTitle
        End Get
        Set(ByVal Value As String)
            _gridReportTitle = Value
        End Set
    End Property

    '   GridReportPreviewFirst
    ''' <summary>
    ''' Allows or disallows the preview display when printing output from the grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will set grid auto report generation to preview output first"), _
    DefaultValue(GetType(Boolean), "True")> _
    Public Property GridReportPreviewFirst() As Boolean
        Get
            Return _gridReportPreviewFirst
        End Get
        Set(ByVal Value As Boolean)
            _gridReportPreviewFirst = Value
        End Set
    End Property

#Region " Excel Properties "

    '   AutoFitColumn
    ''' <summary>
    ''' Allows or disallows the export to excel engines resizing excel columns to fit the contents
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the columns are autosized")> _
    Public Property ExcelAutoFitColumn() As Boolean
        Get
            Return Me._excelAutoFitColumn
        End Get
        Set(ByVal Value As Boolean)
            Me._excelAutoFitColumn = Value
        End Set
    End Property

    '   AutoFitRow
    ''' <summary>
    ''' Allows or disallows the export to excel engins resize rows to fit the contents
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the rows are autosized")> _
    Public Property ExcelAutoFitRow() As Boolean
        Get
            Return Me._excelAutoFitRow
        End Get
        Set(ByVal Value As Boolean)
            Me._excelAutoFitRow = Value
        End Set
    End Property

    '	ExcelAlternateColoration
    ''' <summary>
    ''' Gets or sets the color used to decorate alternate rows on the excel output when matching grid color
    ''' scheme is turned off.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the Color used by the to populoate each alternate row on exports to excel ")> _
        Public Property ExcelAlternateColoration() As Color
        Get
            Return _excelAlternateRowColor
        End Get
        Set(ByVal Value As Color)

            _excelAlternateRowColor = Value

        End Set
    End Property

    '	ExcelMatchGridColorScheme
    ''' <summary>
    ''' Allows or disallows the export to excel engine ability to attemt to color the excel output to match the
    ''' colors used on the onscreen grid. Not all screen colors convert to excel cleanly, and different versions
    ''' of excel interpret colors differently. The export engine attempts to match but those matches are not always
    ''' perfect.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Forces the Export to Excel function to attempt to match the grids color scheme on export ")> _
        Public Property ExcelMatchGridColorScheme() As Boolean
        Get
            Return _excelMatchGridColorScheme
        End Get
        Set(ByVal Value As Boolean)

            _excelMatchGridColorScheme = Value

        End Set
    End Property

    '   Filename
    ''' <summary>
    ''' Gets or sets the name of the excel spreadsheet that is generated when the export operation is complete
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the name of the excel spreadsheet")> _
    Public Property ExcelFilename() As String
        Get
            Return Me._excelFilename
        End Get
        Set(ByVal Value As String)
            Me._excelFilename = Value
        End Set
    End Property

    '   IncludeColumnHeaders
    ''' <summary>
    ''' Allows or disallows the insertion of the header column of the grid into the excel output.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the grid column header row is included in the spreadsheet")> _
    Public Property ExcelIncludeColumnHeaders() As Boolean
        Get
            Return Me._excelIncludeColumnHeaders
        End Get
        Set(ByVal Value As Boolean)
            Me._excelIncludeColumnHeaders = Value
        End Set
    End Property

    '   Keep Alive
    ''' <summary>
    ''' Gets or sets the setting to keep excell alive and kicking after the grid has been sent to excel.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the spreadsheet should remain open after it is filled")> _
    Public Property ExcelKeepAlive() As Boolean
        Get
            Return Me._excelKeepAlive
        End Get
        Set(ByVal Value As Boolean)
            Me._excelKeepAlive = Value
        End Set
    End Property

    '   Maximized
    ''' <summary>
    ''' Gets or sets the making excel opening up maximized or not when its instances and 
    ''' messaged during the export process
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the spreadsheet should be maximized")> _
    Public Property ExcelMaximized() As Boolean
        Get
            Return Me._excelMaximized
        End Get
        Set(ByVal Value As Boolean)
            Me._excelMaximized = Value
        End Set
    End Property

    '   Page Orientation
    ''' <summary>
    ''' Gets or sets the orientation of the excel output for printing
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls the orientation of the spreadsheet output")> _
    Public Property ExcelPageOrientation() As Integer
        Get
            Return Me._excelPageOrientation
        End Get
        Set(ByVal Value As Integer)
            If Value = TAIGridControl.xlPortrait Or Value = TAIGridControl.xlLandscape Then
                Me._excelPageOrientation = Value
            End If
        End Set
    End Property

    '   OutlineCells
    ''' <summary>
    ''' Turns on or off the outlining of cells that are populated during the export process
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Turns on or off the outlineing of all exported grid cells in a solid line")> _
    Public Property ExcelOutlineCells() As Boolean
        Get
            Return Me._excelOutlineCells
        End Get
        Set(ByVal Value As Boolean)
            Me._excelOutlineCells = Value
        End Set
    End Property

    '   Show Borders
    ''' <summary>
    ''' gets or sets the showborders cetting of cells that are populated during the export to excel process
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls whether the cells of the spreadsheet should be outlined")> _
    Public Property ExcelShowBorders() As Boolean
        Get
            Return Me._excelShowBorders
        End Get
        Set(ByVal Value As Boolean)
            Me._excelShowBorders = Value
        End Set
    End Property

    '   UseAlternateRowColor
    ''' <summary>
    ''' Gets or sets the using of the alternat coloring scheme for excel output
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the property which determines if the spreadsheet uses alternating color scheme")> _
    Public Property ExcelUseAlternateRowColor() As Boolean
        Get
            Return Me._excelUseAlternateRowColor
        End Get
        Set(ByVal Value As Boolean)
            Me._excelUseAlternateRowColor = Value
        End Set
    End Property

    '   Workbook Name
    ''' <summary>
    ''' Gets or sets the name of the worksheet name to be used when the export to excel process is underway
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the field that controls the name of the sheet")> _
    Public Property ExcelWorksheetName() As String
        Get
            Return Me._excelWorkSheetName
        End Get
        Set(ByVal Value As String)
            If Value.Length > 31 Then
                Me._excelWorkSheetName = Value.Substring(0, 31)
            Else
                Me._excelWorkSheetName = Value
            End If
        End Set
    End Property

    '   MaxrowsperSheet 
    ''' <summary>
    ''' Gets or sets the maximum rows to be sent to excel before anoher worksheet will be created to carry the remainder
    ''' of data during the export to excel process. This Add a new worksheet and continue will take place until all the
    ''' data in the grid is sent to excel. Excel has a limit of 65535 rows per worksheet.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the maximum number of rows that will be put on a single worksheet during excel export")> _
    Public Property ExcelMaxRowsPerSheet() As Integer
        Get
            Return Me._excelMaxRowsPerSheet
        End Get
        Set(ByVal Value As Integer)
            If Value > 65535 Then
                Me._excelMaxRowsPerSheet = 65535
            Else
                If Value < 100 Then
                    Me._excelMaxRowsPerSheet = 100
                Else
                    Me._excelMaxRowsPerSheet = Value
                End If
            End If
        End Set
    End Property

#End Region

    '   HealerLabel
    ''' <summary>
    ''' Gets or sets the column header label use for the column at index idx
    ''' </summary>
    ''' <param name="columnID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the column header text. Return an empty string on an illegal column ordinal")> _
    Public Property HeaderLabel(ByVal columnID As Integer) As String
        Get
            If columnID < 0 Or columnID > _cols - 1 Then
                Return ""
            Else
                If _GridHeader(columnID) Is Nothing Then
                    Return ""
                Else
                    Return _GridHeader(columnID)
                End If

            End If
        End Get
        Set(ByVal Value As String)
            If columnID < 0 Or columnID > _cols - 1 Then
                ' do nothing
            Else
                _GridHeader(columnID) = Value
                Me.Invalidate()
            End If
        End Set
    End Property

    '   IncludeFieldNames
    ''' <summary>
    ''' Allows or disallows the inclusion of the header lable on grids outoput when exporting to text
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets whether the first row of the grid with the column names should be " & _
                                       "included in the export to text file"), _
    DefaultValue(GetType(Boolean), "True")> _
   Public Property IncludeFieldNames() As Boolean
        Get
            Return Me._includeFieldNames
        End Get
        Set(ByVal Value As Boolean)
            Me._includeFieldNames = Value
        End Set
    End Property

    '   IncludeLineTerminator
    ''' <summary>
    ''' Allows or disallows the inclusion of line termination characters when exporting the grids contents to text
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets whether the export should add a line terminator at the end of each row"), _
    DefaultValue(GetType(Boolean), "True")> _
   Public Property IncludeLineTerminator() As Boolean
        Get
            Return Me._includeLineTerminator
        End Get
        Set(ByVal Value As Boolean)
            Me._includeLineTerminator = Value
        End Set
    End Property

    '	Item
    ''' <summary>
    ''' Gets or sets the contents of the grid cell at row R and col C
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="c"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the text displayed at Row R and Col C of the grid control")> _
    Public Property item(ByVal r As Integer, ByVal c As Integer) As String
        Get
            If r < 0 Or c < 0 Or r > _rows - 1 Or c > _cols - 1 Then
                Return ""
            Else
                If _grid(r, c) Is Nothing Then
                    Return ""
                Else
                    Return _grid(r, c)
                End If
                ' Return _grid(r, c)
            End If
        End Get
        Set(ByVal Value As String)
            If r < 0 Or c < 0 Or r > _rows - 1 Or c > _cols - 1 Then
                'trouble in paradise
            Else
                _grid(r, c) = Value
                Me.Invalidate()
            End If
        End Set
    End Property

    '	Item
    ''' <summary>
    ''' Gets or sets the contents of the grid cell at row R and column name colname
    ''' </summary>
    ''' <param name="r"></param>
    ''' <param name="colname"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the text displayed at Row R and Col named colname of the grid control")> _
    Public Property item(ByVal r As Integer, ByVal colname As String) As String
        Get

            Dim c As Integer = -1

            c = Me.GetColumnIDByName(colname)

            If r < 0 Or c < 0 Or r > _rows - 1 Or c > _cols - 1 Then
                Return ""
            Else
                If _grid(r, c) Is Nothing Then
                    Return ""
                Else
                    Return _grid(r, c)
                End If
                ' Return _grid(r, c)
            End If
        End Get
        Set(ByVal Value As String)

            Dim c As Integer = -1

            c = Me.GetColumnIDByName(colname)

            If r < 0 Or c < 0 Or r > _rows - 1 Or c > _cols - 1 Then
                'trouble in paradise
            Else
                _grid(r, c) = Value
                Me.Invalidate()
            End If
        End Set
    End Property

    '	MaxRowsSelected
    ''' <summary>
    ''' Gets or sets the maximum number of rows to populate the grid with when using the various database populate
    ''' calls. Set to 0 to have the parameter unbounded. With th rewrite in 2005 the grid can accomodate millions of
    ''' rows of data so this setting is largely unnecessary now.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the limit of the number or rows filled in a grid by the various populate from database methods. Will raise an event if this is set and the database fill operation exceeds this threshold. If set to 0 the ALL rows will be selected and no events will fire"), _
    DefaultValue(GetType(Integer), "0")> _
        Public Property MaxRowsSelected() As Integer
        Get
            Return _MaxRowsSelected
        End Get
        Set(ByVal Value As Integer)
            _MaxRowsSelected = Value
        End Set
    End Property

    '	OmitNulls
    ''' <summary>
    ''' Allows or disallows the rendering of the work (NULL) on reading nulls from the varous database population methods
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will turn on or off the rendering of the work (NULL) on reading nulls from the database with the various populate from database methods of the grid"), _
    DefaultValue(GetType(Boolean), "False")> _
        Public Property OmitNulls() As Boolean
        Get
            Return _omitNulls
        End Get
        Set(ByVal Value As Boolean)
            _omitNulls = Value
        End Set
    End Property

    '	PaginationSize
    ''' <summary>
    ''' Gets or sets the number of rows to scroll up or down on the pageup and pagedown keys
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("How many rows to scroll on a Page up or Down")> _
        Public Property PaginationSize() As Integer
        Get
            Return _PaginationSize
        End Get
        Set(ByVal Value As Integer)
            _PaginationSize = Value
        End Set
    End Property

    '	PageSettings
    ''' <summary>
    ''' Gets or sets the PageSettings object used print the grids contents to windows printer devices
    ''' allows for the developer to hand into the grid preconfigured print environments to support their
    ''' special needs.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("The System.Drawing.Printing.PageSettings object used to print the grid to windows printers")> _
        Public Property PageSettings() As System.Drawing.Printing.PageSettings
        Get
            Return _psets
        End Get
        Set(ByVal Value As System.Drawing.Printing.PageSettings)
            _psets = Value
        End Set
    End Property

    '	Rows
    ''' <summary>
    ''' Get to sets the number or rows of data contained in the current grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("How many rows in the current grid control")> _
    Public Property Rows() As Integer
        Get
            Return _rows
        End Get
        Set(ByVal Value As Integer)
            SetRows(Value)
            Me.Invalidate()
        End Set
    End Property

    '	RowHeight
    ''' <summary>
    ''' Gets or sets the height of the row at index idx in pixels
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the rowheight of the row at index IDX in pixels")> _
        Public Property RowHeight(ByVal idx As Integer) As Integer
        Get
            If idx < 0 Or idx > _rows Then
                Return 0
            Else
                Return _rowheights(idx)
            End If
        End Get
        Set(ByVal Value As Integer)
            If idx < 0 Or idx > _rows Then
                ' trouble in paradise
            Else
                _AutosizeCellsToContents = False
                _rowheights(idx) = Value
                Me.Invalidate()
            End If

        End Set
    End Property

    '	SCrollBarWeight
    ''' <summary>
    ''' Gets or sets the height or width of the horizontal and verticle scroll bars on the surface of the grid itself.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets Height or the Horizontal Scroll Bar or Width of the Vertical Scroll bar in Pixels"), _
    DefaultValue(GetType(Integer), "14")> _
    Public Property ScrollBarWeight() As Integer
        Get
            Return _ScrollBarWeight
        End Get
        Set(ByVal Value As Integer)
            _ScrollBarWeight = Value
            Me.Invalidate()
        End Set
    End Property

    '	ScrollInterval
    ''' <summary>
    ''' Gets or sets the amount of screen scroll that the scroll bars will move in pixels
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Sets the amount of screen that a scroll operation will make in Pixels")> _
       Public Property ScrollInterval() As Integer
        Get
            Return _scrollinterval
        End Get
        Set(ByVal Value As Integer)
            _scrollinterval = Value
            Me.Invalidate()
        End Set
    End Property

    '	SelectedColumn
    ''' <summary>
    ''' Gets or sets the currently selected column ID. If more than one column is selected then the 
    ''' <c>SelectedColumns</c> arraylist will contain the set of IDs representative of the selected columns
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the currently selected Column on the grid control")> _
        Public Property SelectedColumn() As Integer
        Get
            Return _SelectedColumn
        End Get
        Set(ByVal Value As Integer)

            If Value > _cols - 1 Or Value < 0 Then
                Exit Property
            Else
                _SelectedColumn = Value
                Me.Invalidate()
            End If

        End Set
    End Property

    '	SelectedRow
    ''' <summary>
    ''' Gets or sets the currently selected row ID. If more than one row is selected then the <c>SelectedRows</c>
    ''' Arraylist will contain the set of IDs representative of the selected rows
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the currently selected row on the grid control")> _
        Public Property SelectedRow() As Integer
        Get
            Return _SelectedRow
        End Get
        Set(ByVal Value As Integer)

            If Value > _rows - 1 Or Value < -1 Then
                Exit Property
            Else
                _SelectedRow = Value

                _SelectedRows.Clear()
                _SelectedRows.Add(_SelectedRow)

                If _SelectedRow <> -1 And vs.Visible Then
                    vs.Value = _SelectedRow
                End If

                Me.Invalidate()
            End If

        End Set
    End Property

    '	SelectedRows
    ''' <summary>
    ''' Gets or sets the currently selected row list in the current grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the currently selected row list on the grid control")> _
        Public Property SelectedRows() As ArrayList
        Get
            Return _SelectedRows
        End Get
        Set(ByVal Value As ArrayList)

            _SelectedRows = Value

            Me.Invalidate()

        End Set
    End Property

    '   SelectedColBackColor
    ''' <summary>
    ''' Gets or sets the background color for the currently selected column in the grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DescriptionAttribute("Sets the background color for the highlighted column in the grid")> _
    Public Property SelectedColBackColor() As Color
        Get
            Return _ColHighliteBackColor
        End Get
        Set(ByVal Value As Color)
            _ColHighliteBackColor = Value
        End Set
    End Property

    '   SelectedColForeColor
    ''' <summary>
    ''' Gets or sets the currently selected column foreground color in the grid
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DescriptionAttribute("Sets the foreground color for the selected column in the grid")> _
    Public Property SelectedColForeColor() As Color
        Get
            Return _ColHighliteForeColor
        End Get
        Set(ByVal Value As Color)
            _ColHighliteForeColor = Value
        End Set
    End Property

    '   SelectedRowBackColor
    ''' <summary>
    ''' Gets or sets the currently selected row background color
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DescriptionAttribute("Sets the background color for the highlighted row in the grid")> _
    Public Property SelectedRowBackColor() As Color
        Get
            Return _RowHighLiteBackColor
        End Get
        Set(ByVal Value As Color)
            _RowHighLiteBackColor = Value
        End Set
    End Property

    '   SelectedRowForeColor
    ''' <summary>
    ''' Gets or sets the currently selected row foreground color
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DescriptionAttribute("Sets the foreground color for the selected row in the grid")> _
    Public Property SelectedRowForeColor() As Color
        Get
            Return _RowHighLiteForeColor
        End Get
        Set(ByVal Value As Color)
            _RowHighLiteForeColor = Value
        End Set
    End Property

    '	ShowDatesWithTime
    ''' <summary>
    ''' Allows or disallows the display of the time portion of datetime values read from the various database populators
    ''' if disallowed just the date portions of these datatypes will be displayed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Will Will Turn on or off the expansion of Dates to include time values if they are present"), _
    DefaultValue(GetType(Boolean), "False")> _
        Public Property ShowDatesWithTime() As Boolean
        Get
            Return _ShowDatesWithTime
        End Get
        Set(ByVal Value As Boolean)
            _ShowDatesWithTime = Value
            Me.Refresh()
        End Set
    End Property

    '   TitleBackColor
    ''' <summary>
    ''' Gets or sets the background color of the title bar in the displayed grid control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the background color for the title")> _
    Public Property TitleBackColor() As Color
        Get
            Return _GridTitleBackcolor
        End Get
        Set(ByVal Value As Color)
            Me._GridTitleBackcolor = Value
            Me.Invalidate()
        End Set
    End Property

    '   TitleFont
    ''' <summary>
    ''' Gets or sets the font used to display the title of the grid control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the font for the title")> _
    Public Property TitleFont() As Font
        Get
            Return Me._GridTitleFont
        End Get
        Set(ByVal Value As Font)
            Me._GridTitleFont = Value
            Me._GridTitleHeight = Me.CreateGraphics.MeasureString("Yy", Value).Height
            Me.Invalidate()
        End Set
    End Property

    '   TitleForeColor
    ''' <summary>
    ''' Gets or sets the foreground color used to render the title of the grid control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the color of the text for the title")> _
    Public Property TitleForeColor() As Color
        Get
            Return Me._GridTitleForeColor
        End Get
        Set(ByVal Value As Color)
            Me._GridTitleForeColor = Value
            Me.Invalidate()
        End Set
    End Property

    '   TitleText
    ''' <summary>
    ''' Gets or sets the actual title text displayed in the grid controls title bar
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the text for the title")> _
    Public Property TitleText() As String
        Get
            Return Me._GridTitle
        End Get
        Set(ByVal Value As String)
            Me._GridTitle = Value
            Me.Invalidate()
        End Set
    End Property

    '   TitleVisible
    ''' <summary>
    ''' Allows or disallows the display of the title bar on the grid control itself
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets whether the title should be displayed")> _
    Public Property TitleVisible() As Boolean
        Get
            Return _GridTitleVisible
        End Get
        Set(ByVal Value As Boolean)
            _GridTitleVisible = Value
            Me.Invalidate()
        End Set
    End Property

    '   UserColResizeMinimum
    ''' <summary>
    ''' Gets or sets the minimum size in pixels the user will be allowed to resize columns to if user column
    ''' resizeing is enabled. This settiing prevents users from resizeing columns to 0 pixels in width making them
    ''' difficult to make visable again. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the Minimum column width for a user to be able to resize a column to in Px."), _
    DefaultValue(GetType(Integer), "5")> _
    Public Property UserColResizeMinimum() As Integer
        Get
            Return _UserColResizeMinimum
        End Get
        Set(ByVal Value As Integer)
            _UserColResizeMinimum = Value
        End Set
    End Property

    '	VisibleHeight
    ''' <summary>
    ''' Gets the height of the grid portion of the visable grid in pixels. Minus the height of the visible scrollbars
    ''' if the scrollbars are visible
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets the height of the grid minus the horizontal scrollbars height if its visible")> _
    Public ReadOnly Property VisibleHeight() As Integer
        Get
            If hs.Visible Then
                Return Me.Height - hs.Height
            Else
                Return Me.Height
            End If
        End Get
    End Property

    '	VisibleWidth
    ''' <summary>
    ''' Gets the width in pixels of the grid area minus the width of the verticle scrollbar if its visible
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets the width of the grid minus the verticle scrollbars width if its visible")> _
    Public ReadOnly Property VisibleWidth() As Integer
        Get
            If vs.Visible Then
                Return Me.Width - vs.Width
            Else
                Return Me.Width
            End If
        End Get
    End Property


#Region " XML Properties "
    '   XMLDataSetName
    ''' <summary>
    ''' Gets or sets he name of the dataset used during the export to XML of the grids contents
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the dataset name used during the export of the grid to XML.")> _
    Public Property XMLDataSetName() As String
        Get
            Return _xmlDataSetName
        End Get
        Set(ByVal Value As String)
            _xmlDataSetName = Value
        End Set
    End Property

    '   XMLFileName
    ''' <summary>
    ''' Gets or sets the filename used to export the contents of the grid to or read from during an XML import operation
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the file name used during the export/import of the grid to XML.")> _
    Public Property XMLFileName() As String
        Get
            Return _xmlFilename
        End Get
        Set(ByVal Value As String)
            _xmlFilename = Value
        End Set
    End Property

    '   XMLIncludeSchema
    ''' <summary>
    ''' Allows or disallows the exporting of the sceme defination in the resulting xml output
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the flag which is used during the export of the grid to XML.")> _
    Public Property XMLIncludeSchema() As Boolean
        Get
            Return _xmlIncludeSchema
        End Get
        Set(ByVal Value As Boolean)
            _xmlIncludeSchema = Value
        End Set
    End Property

    '   XMLNameSpace
    ''' <summary>
    ''' Gets or sets the namespace used to embed the contents of the grid into when exporting to XML
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the name space used during the export of the grid to XML.")> _
    Public Property XMLNameSpace() As String
        Get
            Return _xmlNameSpace
        End Get
        Set(ByVal Value As String)
            _xmlNameSpace = Value
        End Set
    End Property

    '   XMLTableName
    ''' <summary>
    ''' Gets or sets the table named used to export the grids content into during xml export
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.ComponentModel.Description("Gets or Sets the table name used during the export of the grid to XML.")> _
    Public Property XMLTableName() As String
        Get
            Return _xmlTableName
        End Get
        Set(ByVal Value As String)
            _xmlTableName = Value
        End Set
    End Property
#End Region

#End Region

#Region " Private Methods and Functions"

    Private Function AllColWidths() As Integer
        Dim t As Integer
        Dim res As Integer = 0

        For t = 0 To _cols - 1
            res = res + _colwidths(t)
        Next

        Return res + 1
    End Function

    Private Function AllRowHeights() As Integer
        Dim t As Integer
        Dim res As Integer = 0

        For t = 0 To _rows - 1
            res = res + _rowheights(t)
        Next

        If Me.TitleVisible Then
            res = res + Me.TitleFont.Height
        End If

        Return res + 1
    End Function

    Private Function CalculatePageRange() As Integer

        'Dim psets As New System.Drawing.printing.PageSettings

        If _psets.PrinterSettings.PrintRange = Printing.PrintRange.SomePages Then
            _gridPrintingAllPages = False
            _gridStartPage = _psets.PrinterSettings.FromPage
            _gridEndPage = _psets.PrinterSettings.ToPage

            Dim ppea As New System.Drawing.Printing.PrintPageEventArgs(Me.CreateGraphics(), _
                   New Rectangle(_psets.Margins.Left, _psets.Margins.Top, _psets.Margins.Right - _psets.Margins.Left, _
                                 _psets.Margins.Bottom - _psets.Margins.Top), _
                    _psets.Bounds, _psets)

            _gridReportPageNumbers = 1
            _gridReportCurrentrow = 0
            _gridReportCurrentColumn = 0

            Fake_PrintPage(Me, ppea)

            While ppea.HasMorePages
                Fake_PrintPage(Me, ppea)
            End While

            Dim maxpage As Integer = _gridReportPageNumbers

            _gridReportPageNumbers = 1
            _gridReportCurrentrow = 0
            _gridReportCurrentColumn = 0
            Return maxpage


        Else

            _gridPrintingAllPages = True
            _gridStartPage = 1
            _gridStartPageRow = -1

            Dim ppea As New System.Drawing.Printing.PrintPageEventArgs(Me.CreateGraphics(), _
                   New Rectangle(_psets.Margins.Left, _psets.Margins.Top, _psets.Margins.Right - _psets.Margins.Left, _
                                 _psets.Margins.Bottom - _psets.Margins.Top), _
                    _psets.Bounds, _psets)

            _gridReportPageNumbers = 1
            _gridReportCurrentrow = 0
            _gridReportCurrentColumn = 0

            Fake_PrintPage(Me, ppea)

            While ppea.HasMorePages
                Fake_PrintPage(Me, ppea)
            End While

            Dim maxpage As Integer = _gridReportPageNumbers

            _gridReportPageNumbers = 1
            _gridReportCurrentrow = 0
            _gridReportCurrentColumn = 0

            _gridEndPage = maxpage

            Return maxpage

        End If

    End Function

    Private Sub CheckGridTearAways(ByVal colid As Integer)
        '
        ' to be use in methods that affect a specific grid column
        ' like RemoveColFromGrid
        '

        ' dont bother unless we actualy have some to act on 
        If TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)

            If ta.ColID = colid Then
                ' we are showing the column that needs to get the boot
                ta.KillTearAway()
                TearAways.RemoveAt(t)
            End If
        Next

        If TearAways.Count > 0 Then
            ' we still have some so lets look to see if any colids were greater than 
            ' the intended colid for deletion if so we need to decrement them by one
            For t = 0 To TearAways.Count - 1
                If DirectCast(TearAways.Item(t), TearAwayWindowEntry).ColID > colid Then
                    DirectCast(TearAways.Item(t), TearAwayWindowEntry).ColID -= 1
                End If
            Next
        End If

    End Sub

    Private Function CleanMoneyString(ByVal s As String) As String
        Return s.Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "")
    End Function

    Private Sub ClearToBackgroundColor()
        Dim gr As Graphics = Me.CreateGraphics
        gr.FillRectangle(New SolidBrush(Me.BackColor), gr.ClipBounds)
    End Sub

    Private Sub ClearToBackgroundColor(ByVal gr As Graphics)

        gr.FillRectangle(New SolidBrush(Me.BackColor), gr.ClipBounds)

    End Sub

    Private Sub DoAutoSizeCheck(ByVal gr As Graphics)

        Dim r As Integer
        Dim c As Integer
        Dim rr As Integer = 0
        Dim cc As Integer = 0
        Dim t As String

        Dim rrr As Integer = 0
        Dim sz As New SizeF(0, 0)

        If Not _AutoSizeSemaphore Or _AutoSizeAlreadyCalculated Then
            Exit Sub
        End If

        If _AutosizeCellsToContents Then

            _AutoSizeSemaphore = False

            For r = 0 To _rows - 1
                _rowheights(r) = 0
            Next

            For c = 0 To _cols - 1

                t = " " & _GridHeader(c) & " "
                cc = gr.MeasureString(t, _GridHeaderFont).Width
                If cc > rr Then
                    rr = cc
                End If

                For r = 0 To _rows - 1

                    If _colPasswords(c) <> "" Then
                        t = " " & _colPasswords(c) & " "
                    Else
                        If _grid(r, c) Is Nothing Then
                            t = "  "
                        Else
                            t = " " & _grid(r, c) & " "
                        End If
                    End If

                    If _colMaxCharacters(c) <> 0 Then
                        If t.Length > _colMaxCharacters(c) Then
                            t = t.Substring(0, _colMaxCharacters(c))
                        End If
                    End If

                    If (Not _AllowWhiteSpaceInCells) Then
                        t = t.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ").Replace(vbFormFeed, " ")

                        While (t.Replace("  ", " ") <> t)
                            t = t.Replace("  ", " ")
                        End While

                    End If

                    sz = gr.MeasureString(t, _gridCellFontsList(_gridCellFonts(r, c)))

                    cc = sz.Width

                    If cc > rr Then
                        rr = cc
                    End If

                    If _rowheights(r) < sz.Height Then
                        _rowheights(r) = sz.Height
                    End If

                Next

                _colwidths(c) = rr
                rr = 0

            Next

            cc = gr.MeasureString("Yy", _GridHeaderFont).Height

            _GridHeaderHeight = cc

            cc = gr.MeasureString("Yy", _GridTitleFont).Height

            _GridTitleHeight = cc

            'For r = 0 To _rows - 1
            '    For c = 0 To _cols - 1
            '        t = " " & _grid(r, c) & " "
            '        cc = gr.MeasureString(t, _gridCellFontsList(_gridCellFonts(r, c))).Height
            '        If cc > rr Then
            '            rr = cc
            '        End If
            '    Next

            '    _rowheights(r) = rr
            '    rr = 0

            'Next

            _AutoSizeSemaphore = True

            _AutoSizeAlreadyCalculated = True

        Else

        End If
    End Sub

    Private Sub Fake_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)

        Dim x, y, xx, r, c As Integer
        Dim fnt As New System.Drawing.Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Regular, GraphicsUnit.Pixel)
        Dim fnt2 As New System.Drawing.Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Bold, GraphicsUnit.Pixel)

        Dim m As Single

        Dim ft As Font

        Dim greypen As New Pen(Color.Gray)

        Dim pagewidth As Integer = e.PageSettings.Bounds.Size.Width
        Dim pageheight As Integer = e.PageSettings.Bounds.Size.Height

        Dim lrmargin As Integer = 40
        Dim tbmargin As Integer = 70

        Dim colprintedonpage As Boolean = False

        If (AllColWidths() * _gridReportScaleFactor) < pagewidth - (2 * lrmargin) Then
            xx = ((pagewidth - (2 * lrmargin)) - (AllColWidths() * _gridReportScaleFactor)) / 2
        Else
            xx = 0
        End If


        Dim rect As New RectangleF(0, 0, 1, 1)

        x = lrmargin
        y = tbmargin

        Dim coloffset As Integer = 0
        Dim morecols As Boolean = True
        Dim currow As Integer = _gridReportCurrentrow

        If e.PageSettings.PrinterSettings.PrintRange = Printing.PrintRange.SomePages Then
            ' we may be printing just a range so lets see if we can calculate the row to start printing on
            If _gridStartPageRow <= 0 Then
                ' we have not set this up yet so lets check the bounds
                If _gridReportPageNumbers >= _gridStartPage Then
                    _gridStartPageRow = currow
                End If
            End If
        End If

        ft = _GridHeaderFont

        ft = New Font(_GridHeaderFont.FontFamily, _
                        (_GridHeaderFont.SizeInPoints - 1) * _gridReportScaleFactor, _
                        _GridHeaderFont.Style, _GridHeaderFont.Unit)

        ' calculate size and place te printed on date on the page

        m = e.Graphics.MeasureString(_gridReportPrintedOn.ToLongDateString + vbCrLf _
                                    + _gridReportPrintedOn.ToLongTimeString, fnt).Width

        If _gridReportNumberPages Then
            ' we want to number the pages here

            m = e.Graphics.MeasureString("Page " + _gridReportPageNumbers.ToString(), fnt).Height

        End If

        ' print the grid header

        For c = _gridReportCurrentColumn To Me.Cols - 1

            If x + _colwidths(c) + xx > pagewidth - lrmargin And colprintedonpage Then
                Exit For
            End If

            colprintedonpage = True

            rect.X = Convert.ToSingle(x + xx)
            rect.Y = Convert.ToSingle(y)
            rect.Width = Convert.ToSingle(_colwidths(c))
            rect.Height = Convert.ToSingle(_GridHeaderHeight)

            x = x + _colwidths(c)
        Next


        y += _GridHeaderHeight
        x = lrmargin

        For r = _gridReportCurrentrow To Me.Rows - 1
            For c = _gridReportCurrentColumn To Me.Cols - 1
                If x + _colwidths(c) + xx > pagewidth - lrmargin And colprintedonpage Then
                    coloffset = c
                    morecols = True
                    Exit For
                Else
                    morecols = False
                End If

                colprintedonpage = True

                rect.X = Convert.ToSingle(x + xx)
                rect.Y = Convert.ToSingle(y)
                rect.Width = Convert.ToSingle(_colwidths(c))
                rect.Height = Convert.ToSingle(_rowheights(r))

                'ft = New Font(_gridCellFontsList(_gridCellFonts(r, c)).FontFamily, _
                '              _gridCellFontsList(_gridCellFonts(r, c)).SizeInPoints - 1, _
                '              _gridCellFontsList(_gridCellFonts(r, c)).Style, _
                '              _gridCellFontsList(_gridCellFonts(r, c)).Unit)

                'e.Graphics.DrawString(_grid(r, c), ft, _
                '                      Brushes.Black, rect, _gridCellAlignmentList(_gridCellAlignment(r, c)))

                x = x + _colwidths(c)

            Next
            x = lrmargin
            y += _rowheights(r)
            _gridReportCurrentrow += 1

            ' do we need to skip to next page here
            If y >= pageheight - tbmargin Then
                Exit For
            Else
                ' nope
            End If

            Application.DoEvents()
        Next

        If _gridReportCurrentrow >= Me.Rows - 1 And Not morecols Then
            e.HasMorePages = False
            '_gridReportPageNumbers = 1
            _gridReportCurrentrow = 0
            _gridReportCurrentColumn = 0
        Else
            If morecols Then
                _gridReportCurrentColumn = coloffset
                _gridReportCurrentrow = currow
            Else
                _gridReportCurrentColumn = 0
            End If
            e.HasMorePages = True
            _gridReportPageNumbers += 1
        End If

    End Sub

    Private Function GetGridBackColorListEntry(ByVal bcol As Brush) As Integer
        Dim t As Integer
        Dim flag As Integer = -1

        Dim bbcol As System.Drawing.SolidBrush
        Dim aacol As System.Drawing.SolidBrush

        For t = 0 To _gridBackColorList.GetUpperBound(0)
            If _gridBackColorList(t) Is Nothing Then
                ' we got nothing in the color list

            Else
                bbcol = DirectCast(_gridBackColorList(t), System.Drawing.SolidBrush)
                aacol = DirectCast(bcol, System.Drawing.SolidBrush)

                If aacol.Color.A = bbcol.Color.A Then
                    If aacol.Color.R = bbcol.Color.R Then
                        If aacol.Color.G = bbcol.Color.G Then
                            If aacol.Color.B = bbcol.Color.B Then
                                flag = t
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If flag = -1 Then
            ' we dont have that fnt2find in the list so we need to add it
            t = _gridBackColorList.GetUpperBound(0)
            t += 1
            ReDim Preserve _gridBackColorList(t + 1)
            _gridBackColorList(t) = bcol
            flag = t
        End If

        Return flag

    End Function

    Private Function GetGridCellAlignmentListEntry(ByVal sfmt As StringFormat) As Integer
        Dim t As Integer
        Dim flag As Integer = -1

        For t = 0 To _gridCellAlignmentList.GetUpperBound(0)
            If sfmt.Equals(_gridCellAlignmentList(t)) Then
                flag = t
                Exit For
            End If
        Next

        If flag = -1 Then
            ' we dont have that fnt2find in the list so we need to add it
            t = _gridCellAlignmentList.GetUpperBound(0)
            t += 1
            ReDim Preserve _gridCellAlignmentList(t + 1)
            _gridCellAlignmentList(t) = sfmt
            flag = t
        End If

        Return flag

    End Function

    Private Function GetGridCellFontListEntry(ByVal fnt2find As Font) As Integer
        Dim t As Integer
        Dim flag As Integer = -1

        For t = 0 To _gridCellFontsList.GetUpperBound(0)
            If fnt2find.Equals(_gridCellFontsList(t)) Then
                flag = t
                Exit For
            End If
        Next

        If flag = -1 Then
            ' we dont have that fnt2find in the list so we need to add it
            t = _gridCellFontsList.GetUpperBound(0)
            t += 1
            ReDim Preserve _gridCellFontsList(t + 1)
            _gridCellFontsList(t) = fnt2find
            flag = t
        End If

        Return flag

    End Function

    Private Function GetGridForeColorListEntry(ByVal fcol As Pen) As Integer
        Dim t As Integer
        Dim flag As Integer = -1

        For t = 0 To _gridForeColorList.GetUpperBound(0)

            If _gridForeColorList(t) Is Nothing Then
                ' we got nothing in the color list
            Else
                If fcol.Color.A = _gridForeColorList(t).Color.A Then
                    If fcol.Color.R = _gridForeColorList(t).Color.R Then
                        If fcol.Color.G = _gridForeColorList(t).Color.G Then
                            If fcol.Color.B = _gridForeColorList(t).Color.B Then
                                flag = t
                                Exit For
                            End If
                        End If

                    End If
                End If
            End If

            'If fcol Is _gridForeColorList(t) Then
            '    flag = t
            '    Exit For
            'End If

            'If fcol.Equals(_gridForeColorList(t)) Then
            '    flag = t
            '    Exit For
            'End If
        Next

        If flag = -1 Then
            ' we dont have that fnt2find in the list so we need to add it
            t = _gridForeColorList.GetUpperBound(0)
            t += 1
            ReDim Preserve _gridForeColorList(t + 1)
            _gridForeColorList(t) = fcol
            flag = t
        End If

        Return flag

    End Function

    Private Function GetLetter(ByVal iNumber As Integer) As String

        Try

            Select Case iNumber
                Case 0
                    Return ""
                Case 1
                    Return "A"
                Case 2
                    Return "B"
                Case 3
                    Return "C"
                Case 4
                    Return "D"
                Case 5
                    Return "E"
                Case 6
                    Return "F"
                Case 7
                    Return "G"
                Case 8
                    Return "H"
                Case 9
                    Return "I"
                Case 10
                    Return "J"
                Case 11
                    Return "K"
                Case 12
                    Return "L"
                Case 13
                    Return "M"
                Case 14
                    Return "N"
                Case 15
                    Return "O"
                Case 16
                    Return "P"
                Case 17
                    Return "Q"
                Case 18
                    Return "R"
                Case 19
                    Return "S"
                Case 20
                    Return "T"
                Case 21
                    Return "U"
                Case 22
                    Return "V"
                Case 23
                    Return "W"
                Case 24
                    Return "X"
                Case 25
                    Return "Y"
                Case 26
                    Return "Z"
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.GetUpperColumn Error...")
            Return ""
        End Try

    End Function

    Private Function GetUpperColumn(ByVal iCols As Integer) As String

        Try
            Dim iMajor As Integer = iCols / 26
            Dim iMinor As Integer = iCols Mod 26

            Return Me.GetLetter(iMajor) & Me.GetLetter(iMinor)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.GetLetter Error...")
            Return Me.GetLetter(1) & Me.GetLetter(1)
        End Try

    End Function

    Private Function GimmeGridSize() As Point
        Dim pnt As Point
        Dim x, y As Integer
        Dim siz As Integer = 0

        For x = 0 To _cols - 1
            siz = siz + _colwidths(x)
        Next

        pnt.X = siz
        siz = 0

        For y = 0 To _rows - 1
            siz = siz + _rowheights(y)
        Next

        If _GridTitleVisible Then
            siz = siz + _GridTitleHeight
        End If

        If _GridHeaderVisible Then
            siz = siz + _GridHeaderHeight
        End If

        pnt.Y = siz

        Return pnt

    End Function

    Private Function GimmeXOffset(ByVal col As Integer) As Integer
        Dim t As Integer
        Dim ret As Integer = 0

        If col = 0 Then
            ret = 0
        Else
            For t = 0 To col - 1
                ret = ret + _colwidths(t)
            Next
        End If
        Return ret

    End Function

    Private Function GimmeYOffset(ByVal row As Integer) As Integer
        Dim t As Integer
        Dim ret As Integer = 0

        If row = 0 Then
            ret = 0
        Else
            For t = 0 To row - 1
                ret = ret + _rowheights(t)
            Next
        End If
        Return ret

    End Function

    Private Sub InitializeTheGrid()

        Dim r, c As Integer

        ReDim _grid(2, 2)
        ReDim _gridBackColor(2, 2)
        ReDim _gridForeColor(2, 2)
        ReDim _gridCellFonts(2, 2)
        ReDim _gridCellFontsList(1)
        ReDim _gridForeColorList(1)
        ReDim _gridBackColorList(1)
        ReDim _gridCellAlignment(2, 2)
        ReDim _gridCellAlignmentList(1)
        ReDim _colwidths(2)
        ReDim _colhidden(2)
        ReDim _rowhidden(2)
        ReDim _colEditable(2)
        ReDim _rowEditable(2)
        ReDim _colboolean(2)
        ReDim _colPasswords(2)
        ReDim _colMaxCharacters(2)
        ReDim _rowheights(2)
        ReDim _GridHeader(2)
        _rows = 2
        _cols = 2

        _SelectedRow = -1
        _SelectedColumn = -1

        _gridCellFontsList(0) = _DefaultCellFont
        _gridForeColorList(0) = New Pen(_DefaultForeColor)
        _gridCellAlignmentList(0) = _DefaultStringFormat
        _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)


        hs.Visible = False
        vs.Visible = False
        hs.Value = 0
        vs.Value = 0

        For r = 0 To _rows - 1
            For c = 0 To _cols - 1

                _gridBackColor(r, c) = 0 ' New SolidBrush(_DefaultBackColor.AntiqueWhite)
                _gridForeColor(r, c) = 0 ' New Pen(_DefaultForeColor.Blue)
                _gridCellFonts(r, c) = 0 ' _DefaultCellFont
                _gridCellAlignment(r, c) = 0 ' _DefaultStringFormat

            Next
        Next

    End Sub

    Private Sub InitializeTheGrid(ByVal row As Integer, ByVal col As Integer)

        Dim r, c As Integer

        ReDim _grid(row, col)
        ReDim _gridBackColor(row, col)
        ReDim _gridForeColor(row, col)
        ReDim _gridCellFonts(row, col)
        ReDim _gridCellFontsList(1)
        ReDim _gridForeColorList(1)
        ReDim _gridBackColorList(1)
        ReDim _gridCellAlignment(row, col)
        ReDim _gridCellAlignmentList(1)
        ReDim _colwidths(col)
        ReDim _colEditable(col)
        ReDim _rowEditable(row)
        ReDim _colhidden(col)
        ReDim _colboolean(col)
        ReDim _rowhidden(row)
        ReDim _colPasswords(col)
        ReDim _colMaxCharacters(col)
        ReDim _rowheights(row)
        ReDim _GridHeader(col)

        _SelectedRows = New ArrayList()
        ''_SelectedRows.Clear()

        _rows = row
        _cols = col

        _SelectedRow = -1
        _SelectedColumn = -1

        _gridCellFontsList(0) = _DefaultCellFont
        _gridForeColorList(0) = New Pen(_DefaultForeColor)
        _gridCellAlignmentList(0) = _DefaultStringFormat
        _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)


        hs.Visible = False
        vs.Visible = False
        hs.Value = 0
        vs.Value = 0

        For r = 0 To _rows - 1
            _rowEditable(r) = True
            For c = 0 To _cols - 1

                _gridBackColor(r, c) = 0 ' New SolidBrush(_DefaultBackColor)
                _gridForeColor(r, c) = 0 ' New Pen(_DefaultForeColor)
                _gridCellFonts(r, c) = 0 '_DefaultCellFont
                _gridCellAlignment(r, c) = 0 ' _DefaultStringFormat

            Next
        Next

        For c = 0 To _cols - 1
            _colPasswords(c) = ""
            _colEditable(c) = False
            _colhidden(c) = False
            _colboolean(c) = False
        Next

    End Sub

    Private Sub NormalizeTearaways()
        If TearAways.Count = 0 Then
            ' we dont have any tearaways lets blow this pop stand
            Exit Sub
        End If

        ' we have some so lets fix things here

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)
            If ta.ColID >= _cols Then
                ' we have a tearaway open on a column that no longer exists so lets close it
                ta.Winform.KillMe(ta.ColID)
            Else
                ' we the column is still there so lets change its title and its contents
                ta.Winform.Text = Me.HeaderLabel(ta.ColID)
                ta.Winform.ListItems = Me.GetColAsArrayList(ta.ColID)
                ta.SetTearAwayScrollParameters(vs.Minimum, vs.Maximum, vs.Visible)
                ta.SetTearAwayScrollIndex(vs.Value)
            End If
        Next
    End Sub

    Private Sub OleRenderGrid(ByVal gr As Graphics)
        Dim w As Integer = Me.AllColWidths()
        Dim h As Integer = Me.AllRowHeights()
        Dim orig As Point
        Dim t As Integer
        Dim xof As Integer
        Dim xxof, yyof As Integer
        Dim r, c As Integer
        Dim rh, rhy, rhx As Integer ' use for checkbox renderings
        Dim rowstart As Integer = -1
        Dim rowend As Integer = -1
        Dim colstart As Integer = -1
        Dim colend As Integer = -1
        Dim gyofset As Integer
        Dim renderstring As String = ""

        If _gridForeColorList(0) Is Nothing Then
            _gridForeColorList(0) = New Pen(_DefaultForeColor)
        End If

        If _gridBackColorList(0) Is Nothing Then
            _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)
        End If

        If _GridHeaderVisible Then
            h += _GridHeaderHeight
        End If

        If _antialias Then
            gr.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        Else
            gr.SmoothingMode = Drawing2D.SmoothingMode.Default
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault
        End If

        ClearToBackgroundColor(gr)

        ' If we are disallowing selection of columns then make sure the Selected column variable is out of bounds
        If Not _AllowColumnSelection Then
            _SelectedColumn = -1
        End If

        If _GridTitleVisible Then
            ' we need to draw the title
            gr.FillRectangle(New SolidBrush(_GridTitleBackcolor), 0, 0, w, _GridTitleHeight)
            gr.DrawString(_GridTitle, _GridTitleFont, New SolidBrush(_GridTitleForeColor), 0, 0)
            orig.X = 0
            orig.Y = _GridTitleHeight
        Else
            orig.X = 0
            orig.Y = 0
        End If

        If _cols <> 0 And _GridHeaderVisible Then
            orig.Y = orig.Y + _GridHeaderHeight
        End If

        yyof = 0
        xxof = 0

        If _rows = 0 And _cols = 0 Then
            ' We have nothing else to draw so lets bail
            '_Painting = False
        Else

            rowstart = 0
            rowend = _rows - 1

            colstart = 0
            colend = _cols - 1

            ' time to render the grid here
            For r = rowstart To rowend
                gyofset = GimmeYOffset(r)
                For c = colstart To colend
                    xof = GimmeXOffset(c)
                    If _colwidths(c) > 0 Then

                        If _colPasswords(c) Is Nothing Then
                            renderstring = _grid(r, c)
                        Else
                            If _colPasswords(c) = "" Then
                                renderstring = _grid(r, c)
                            Else
                                renderstring = _colPasswords(c)
                            End If
                        End If

                        ' handle the Max characters display here

                        If _colMaxCharacters(c) <> 0 Then
                            If renderstring.Length > _colMaxCharacters(c) Then
                                renderstring = renderstring.Substring(0, _colMaxCharacters(c)) + "..."
                            End If
                        End If

                        If r = _SelectedRow Or c = _SelectedColumn Or _SelectedRows.Contains(r) Then
                            If r = _SelectedRow Or _SelectedRows.Contains(r) Then
                                ' we have a selected row override of selected column

                                gr.FillRectangle(New SolidBrush(_RowHighLiteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    ' we have to render the the checkbox

                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If

                                Else
                                    gr.DrawString(renderstring, _
                                                  _gridCellFontsList(_gridCellFonts(r, c)), _
                                                  New SolidBrush(_RowHighLiteForeColor), _
                                                  New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                                  _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If


                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If
                            Else
                                ' we have a selected Col

                                gr.FillRectangle(New SolidBrush(_ColHighliteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    ' we have to render the the checkbox
                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If
                                Else
                                    gr.DrawString(renderstring, _
                                              _gridCellFontsList(_gridCellFonts(r, c)), _
                                             New SolidBrush(_ColHighliteForeColor), _
                                             New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                             _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If

                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If

                            End If
                        Else
                            gr.FillRectangle(_gridBackColorList(_gridBackColor(r, c)), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                            If _colboolean(c) Then
                                ' we have to render the the checkbox
                                rh = _rowheights(r) - 2

                                If rh > 14 Then

                                    rh = 14

                                End If

                                If rh < 6 Then

                                    rh = 6

                                End If

                                rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                If rhx < 0 Then
                                    rhx = 0
                                End If

                                rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                If rhy < 0 Then
                                    rhy = 0
                                End If

                                If UCase(renderstring) = "TRUE" Or _
                                   UCase(renderstring) = "YES" Or _
                                   UCase(renderstring) = "Y" Or _
                                   UCase(renderstring) = "1" Then
                                    ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                Else
                                    If UCase(renderstring) = "" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                    Else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                    End If
                                End If
                            Else
                                gr.DrawString(renderstring, _
                                        _gridCellFontsList(_gridCellFonts(r, c)), _
                                       New SolidBrush(_gridForeColorList(_gridForeColor(r, c)).Color), _
                                       New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                       _gridCellAlignmentList(_gridCellAlignment(r, c)))

                            End If
                            If _CellOutlines Then
                                gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                            _colwidths(c), _rowheights(r)))
                            End If
                        End If
                    End If
                Next
            Next

            ' recalc the top area so we can draw the header if its vivible
            If _GridTitleVisible Then
                orig.X = 0
                orig.Y = _GridTitleHeight
            Else
                orig.X = 0
                orig.Y = 0
            End If

            gr.SetClip(New RectangleF(0, 0, w, h))

            If _cols <> 0 And _GridHeaderVisible Then
                ' we need to render the Header

                For t = 0 To _cols - 1
                    xof = GimmeXOffset(t)
                    If _colwidths(t) > 0 Then
                        gr.FillRectangle(New SolidBrush(_GridHeaderBackcolor), xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight)
                        gr.DrawString(_GridHeader(t), _GridHeaderFont, New SolidBrush(_GridHeaderForecolor), _
                                      New RectangleF(xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight), _GridHeaderStringFormat)
                        If _CellOutlines Then
                            gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight))
                        End If
                    End If
                Next
                orig.Y = orig.Y + _GridHeaderHeight
            End If

            ' do we need to display the scrollbars

            'RecalcScrollBars()

            If _BorderStyle = BorderStyle.Fixed3D Or _BorderStyle = BorderStyle.FixedSingle Then
                gr.DrawRectangle(New Pen(_BorderColor, 1), 0, 0, w - 1, h - 1)
            End If

#If EVALUATION Then
            Dim evalfont As Font = New Font("Courier New", 50, FontStyle.Bold)
            Dim evalstring As String = "EVALUATION VERSION"
            Dim evalmetrics As SizeF = gr.MeasureString(evalstring, evalfont)

            gr.DrawString(evalstring, evalfont, New SolidBrush(Color.IndianRed), (w / 2) - (evalmetrics.Width / 2), (h / 2) - (evalmetrics.Height / 2))
#End If

            _Painting = False

        End If

        'grview.DrawImageUnscaled(bmp, 0, 0)
        'gr.Dispose()
        'bmp.Dispose()
        'bmp = Nothing

    End Sub

    Private Sub PrivatePopulateGridFromArray(ByVal arr(,) As String, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean)

        Dim x, y As Integer
        Dim r, c As Integer

        r = arr.GetUpperBound(0) + 1
        c = arr.GetUpperBound(1) + 1

        If FirstRowHeader Then
            InitializeTheGrid(r - 1, c)
            For y = 0 To c - 1
                _GridHeader(y) = arr(0, y)
            Next
            For x = 1 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y)
                Next
            Next
        Else
            'InitializeTheGrid(r, c)
            'For y = 0 To c - 1
            '    _GridHeader(y) = "Column - " & y.ToString
            'Next
            For x = 0 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y)
                Next
            Next
        End If

        AllCellsUseThisFont(gridfont)
        AllCellsUseThisForeColor(col)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

    End Sub

    Private Sub RecalcScrollBars()
        ' ##Recalculates the positions and the visibility for the scroll bars

        Dim ClientHeight As Integer

        ClientHeight = Me.Height


        _GridSize = GimmeGridSize()

        If _GridHeaderVisible Then
            ClientHeight -= _GridHeaderHeight
        End If

        If _GridTitleVisible Then
            ClientHeight -= _GridTitleHeight
        End If

        If _GridSize.X > Me.Width Then
            hs.Visible = True
            hs.Height = _ScrollBarWeight

            hs.Maximum = _cols + 2
            hs.LargeChange = 4
            hs.SmallChange = 1
            ClientHeight -= _ScrollBarWeight
        Else
            hs.Visible = False
            hs.Maximum = 1
            hs.Minimum = 0
            hs.Value = 0
        End If

        If _GridSize.Y > ClientHeight Then
            vs.Visible = True
            vs.Width = _ScrollBarWeight

            vs.Maximum = _rows + 10
            vs.LargeChange = 10
            vs.SmallChange = 1
        Else
            vs.Visible = False
            vs.Maximum = 1
            vs.Minimum = 0
            vs.Value = 0
        End If

    End Sub

    Private Sub RedimTable()

        Dim oldrowhidden(_rowhidden.GetUpperBound(0)) As Boolean
        Dim oldcolhidden(_colhidden.GetUpperBound(0)) As Boolean
        Dim oldcolboolean(_colboolean.GetUpperBound(0)) As Boolean
        Dim oldcoleditable(_colEditable.GetUpperBound(0)) As Boolean
        Dim oldroweditable(_rowEditable.GetUpperBound(0)) As Boolean
        Dim oldcolwidths(_colwidths.GetUpperBound(0)) As Integer
        Dim oldrowheights(_rowheights.GetUpperBound(0)) As Integer
        Dim oldgridheader(_GridHeader.GetUpperBound(0)) As String
        Dim oldgrid(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As String
        Dim oldgridbcolor(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridfcolor(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridfonts(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridcolpasswords(_colPasswords.GetUpperBound(0)) As String
        Dim oldcolmaxcharacters(_colMaxCharacters.GetUpperBound(0)) As Integer
        Dim oldgridcellalignment(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim r, c As Integer
        Dim x, y As Integer
        
        x = oldgrid.GetUpperBound(0)
        y = oldgrid.GetUpperBound(1)

        For r = 0 To x
            For c = 0 To y
                oldgrid(r, c) = _grid(r, c)
                oldgridbcolor(r, c) = _gridBackColor(r, c)
                oldgridfcolor(r, c) = _gridForeColor(r, c)
                oldgridfonts(r, c) = _gridCellFonts(r, c)
                oldgridcellalignment(r, c) = _gridCellAlignment(r, c)
            Next
        Next

        For c = 0 To Math.Min(_GridHeader.GetUpperBound(0), _colwidths.GetUpperBound(0))
            oldgridheader(c) = _GridHeader(c)
            oldcolwidths(c) = _colwidths(c)
            oldgridcolpasswords(c) = _colPasswords(c)
            oldcolhidden(c) = _colhidden(c)
            oldcolboolean(c) = _colboolean(c)
            oldcoleditable(c) = _colEditable(c)
            oldcolmaxcharacters(c) = _colMaxCharacters(c)
        Next

        For r = 0 To _rowheights.GetUpperBound(0)
            oldrowheights(r) = _rowheights(r)
        Next

        For r = 0 To _rowhidden.GetUpperBound(0)
            oldrowhidden(r) = _rowhidden(r)
        Next

        For r = 0 To _rowEditable.GetUpperBound(0)
            oldroweditable(r) = _rowEditable(r)
        Next

        ReDim _rowhidden(_rows)
        ReDim _colhidden(_cols)
        ReDim _colboolean(_cols)
        ReDim _colEditable(_cols)
        ReDim _rowEditable(_rows)
        ReDim _rowheights(_rows)
        ReDim _colwidths(_cols)
        ReDim _GridHeader(_cols)
        ReDim _grid(_rows, _cols)
        ReDim _gridBackColor(_rows, _cols)
        ReDim _gridForeColor(_rows, _cols)
        ReDim _gridCellFonts(_rows, _cols)
        ReDim _gridCellAlignment(_rows, _cols)
        ReDim _colPasswords(_cols)
        ReDim _colMaxCharacters(_cols)

        If _rows < x Then
            x = _rows
        End If

        If _cols < y Then
            y = _cols
        End If

        For c = 0 To y
            _colPasswords(c) = oldgridcolpasswords(c)
            _GridHeader(c) = oldgridheader(c)
            _colwidths(c) = oldcolwidths(c)
            _colhidden(c) = oldcolhidden(c)
            _colboolean(c) = oldcolboolean(c)
            _colEditable(c) = oldcoleditable(c)
            _colMaxCharacters(c) = oldcolmaxcharacters(c)
        Next

        For r = 0 To x
            _rowheights(r) = oldrowheights(r)
            _rowhidden(r) = oldrowhidden(r)
            _rowEditable(r) = oldroweditable(r)
        Next

        If x = 0 Then
            r = x
            For c = 0 To y
                _grid(r, c) = oldgrid(r, c)
                _gridBackColor(r, c) = GetGridBackColorListEntry(New SolidBrush(_DefaultBackColor))
                _gridForeColor(r, c) = GetGridForeColorListEntry(New Pen(_DefaultForeColor))
                _gridCellFonts(r, c) = GetGridCellFontListEntry(_DefaultCellFont)
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(_DefaultStringFormat)
            Next
        Else
            For r = 0 To x
                For c = 0 To y
                    _grid(r, c) = oldgrid(r, c)
                    _gridBackColor(r, c) = oldgridbcolor(r, c)
                    _gridForeColor(r, c) = oldgridfcolor(r, c)
                    _gridCellFonts(r, c) = oldgridfonts(r, c)
                    _gridCellAlignment(r, c) = oldgridcellalignment(r, c)
                Next
            Next
        End If

        If oldcolwidths.GetUpperBound(0) < _colwidths.GetUpperBound(0) Then
            For c = oldcolwidths.GetUpperBound(0) + 1 To _colwidths.GetUpperBound(0)
                _colwidths(c) = _DefaultColWidth
                _colEditable(c) = False ' default all new columns to not editable
                _colhidden(c) = False ' cols default to not hidden
                _colboolean(c) = False ' cols default to not boolean
            Next
        End If

        If oldrowheights.GetUpperBound(0) < _rowheights.GetUpperBound(0) Then
            For r = oldrowheights.GetUpperBound(0) + 1 To _rowheights.GetUpperBound(0)
                _rowheights(r) = _DefaultRowHeight
            Next
        End If

        For c = 0 To _cols - 1
            If _colwidths(c) = 0 And Not _colhidden(c) Then
                _colwidths(c) = _DefaultColWidth
            End If
        Next

        For r = 0 To _rows - 1
            If _rowheights(r) = 0 And Not _rowhidden(r) Then
                _rowheights(r) = _DefaultRowHeight
            End If
        Next

    End Sub

    Private Sub RenderGrid(ByVal grview As Graphics)
        Dim w As Integer = grview.VisibleClipBounds.Width
        Dim h As Integer = grview.VisibleClipBounds.Height
        Dim orig As Point
        Dim t As Integer
        Dim xof As Integer
        Dim xxof, yyof As Integer
        Dim r, c As Integer
        Dim rh, rhy, rhx As Integer ' use for checkbox renderings
        Dim rowstart As Integer = -1
        Dim rowend As Integer = -1
        Dim colstart As Integer = -1
        Dim colend As Integer = -1
        Dim gyofset As Integer
        Dim renderstring As String = ""

        '
        ' Here we want to just bail if the size is less than some small size
        ' 

        If w < 10 Or h < 10 Then
            Exit Sub
        End If

        If _Painting Then
            Exit Sub
        Else
            _Painting = True
        End If

        If _gridForeColorList(0) Is Nothing Then
            _gridForeColorList(0) = New Pen(_DefaultForeColor)
        End If

        If _gridBackColorList(0) Is Nothing Then
            _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)
        End If

        Dim gr As Graphics
        Dim bmp As Bitmap

        bmp = New Bitmap(w, h, grview)
        gr = Graphics.FromImage(bmp)

        If _antialias Then
            gr.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        Else
            gr.SmoothingMode = Drawing2D.SmoothingMode.Default
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault
        End If

        DoAutoSizeCheck(gr)

        ClearToBackgroundColor(gr)

        RecalcScrollBars()

        ' If we are disallowing selection of columns then make sure the Selectedcolumn variable is out of bounds
        If Not _AllowColumnSelection Then
            _SelectedColumn = -1
        End If

        If _GridTitleVisible Then
            ' we need to draw the title
            gr.FillRectangle(New SolidBrush(_GridTitleBackcolor), 0, 0, w, _GridTitleHeight)
            gr.DrawString(_GridTitle, _GridTitleFont, New SolidBrush(_GridTitleForeColor), 0, 0)
            orig.X = 0
            orig.Y = _GridTitleHeight
        Else
            orig.X = 0
            orig.Y = 0
        End If

        If _cols <> 0 And _GridHeaderVisible Then
            orig.Y = orig.Y + _GridHeaderHeight
        End If

        If vs.Visible Then
            yyof = GimmeYOffset(vs.Value)
        Else
            yyof = 0
        End If

        If hs.Visible Then
            xxof = GimmeXOffset(hs.Value)
        Else
            xxof = 0
        End If

        If _rows = 0 And _cols = 0 Then
            ' We have nothing else to draw so lets bail
            _Painting = False
        Else
            ' if we are possible needing to draw the background we had better do it here
            If Not (hs.Visible And vs.Visible) Then
                If _GridTitleVisible Then
                    gr.FillRectangle(New SolidBrush(Me.BackColor), New RectangleF(0, _GridTitleHeight, w, h - _GridTitleHeight))
                Else
                    gr.FillRectangle(New SolidBrush(Me.BackColor), gr.VisibleClipBounds)
                End If
            End If

            ' here we want to validate the starting and ending rows for the render process
            If vs.Visible Then

                'If _SelectedRow <> -1 Then
                '    vs.Value = _SelectedRow
                'End If

                rowstart = vs.Value
                rowend = _rows - 1

                For r = rowstart To _rows - 1
                    If (GimmeYOffset(r) - yyof) >= h Then
                        rowend = r
                        Exit For
                    End If
                Next

            Else
                rowstart = 0
                rowend = _rows - 1
            End If

            If hs.Visible Then
                colstart = hs.Value
                colend = _cols - 1

                For c = colstart To _cols - 1
                    If (GimmeXOffset(c) - xxof) >= w Then
                        colend = c
                        Exit For
                    End If
                Next
            Else
                colstart = 0
                colend = _cols - 1
            End If

            'If _SelectedRow <> -1 And vs.Visible Then
            '    If _SelectedRow < rowstart Then
            '        vs.Value = vs.Value - (rowstart - _SelectedRow)
            '    End If
            '    If _SelectedRow > rowend Then
            '        vs.Value = vs.Value + (_SelectedRow - rowend)
            '    End If
            'End If

            ' from now on all drawing ops occur below the grid title if its visible and the header if its visible

            gr.SetClip(New RectangleF(0, orig.Y, w, h - orig.Y))

            ' Console.WriteLine(rowstart.ToString & " - " & rowend.ToString & " ------- " & colstart.ToString & " - " & colend)

            ' time to render the grid here
            For r = rowstart To rowend
                gyofset = GimmeYOffset(r)
                For c = colstart To colend
                    xof = GimmeXOffset(c)
                    If _colwidths(c) > 0 Then

                        If _colPasswords(c) Is Nothing Then
                            renderstring = _grid(r, c)
                        Else
                            If _colPasswords(c) = "" Then
                                renderstring = _grid(r, c)
                            Else
                                renderstring = _colPasswords(c)
                            End If
                        End If

                        ' handle the Max characters display here

                        If _colMaxCharacters(c) <> 0 Then
                            If renderstring.Length > _colMaxCharacters(c) Then
                                renderstring = renderstring.Substring(0, _colMaxCharacters(c)) + "..."
                            End If
                        End If

                        If r = _SelectedRow Or c = _SelectedColumn Or _SelectedRows.Contains(r) Then
                            If r = _SelectedRow Or _SelectedRows.Contains(r) Then
                                ' we have a selected row override of selected column

                                gr.FillRectangle(New SolidBrush(_RowHighLiteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    ' we have to render the the checkbox

                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If

                                Else
                                    gr.DrawString(renderstring, _
                                                  _gridCellFontsList(_gridCellFonts(r, c)), _
                                                  New SolidBrush(_RowHighLiteForeColor), _
                                                  New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                                  _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If


                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If
                            Else
                                ' we have a selected Col

                                gr.FillRectangle(New SolidBrush(_ColHighliteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    ' we have to render the the checkbox
                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If
                                Else
                                    gr.DrawString(renderstring, _
                                              _gridCellFontsList(_gridCellFonts(r, c)), _
                                             New SolidBrush(_ColHighliteForeColor), _
                                             New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                             _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If

                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If

                            End If
                        Else
                            gr.FillRectangle(_gridBackColorList(_gridBackColor(r, c)), xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r))

                            If _colboolean(c) Then
                                ' we have to render the the checkbox
                                rh = _rowheights(r) - 2

                                If rh > 14 Then

                                    rh = 14

                                End If

                                If rh < 6 Then

                                    rh = 6

                                End If

                                rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                If rhx < 0 Then
                                    rhx = 0
                                End If

                                rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                If rhy < 0 Then
                                    rhy = 0
                                End If

                                If UCase(renderstring) = "TRUE" Or _
                                   UCase(renderstring) = "YES" Or _
                                   UCase(renderstring) = "Y" Or _
                                   UCase(renderstring) = "1" Then
                                    ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                Else
                                    If UCase(renderstring) = "" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                    Else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                    End If
                                End If
                            Else
                                gr.DrawString(renderstring, _
                                        _gridCellFontsList(_gridCellFonts(r, c)), _
                                       New SolidBrush(_gridForeColorList(_gridForeColor(r, c)).Color), _
                                       New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                       _gridCellAlignmentList(_gridCellAlignment(r, c)))

                            End If
                            If _CellOutlines Then
                                gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y + gyofset - yyof, _
                                                                                            _colwidths(c), _rowheights(r)))
                            End If
                        End If
                    End If
                Next
            Next

            ' recalc the top area so we can draw the header if its vivible
            If _GridTitleVisible Then
                orig.X = 0
                orig.Y = _GridTitleHeight
            Else
                orig.X = 0
                orig.Y = 0
            End If

            gr.SetClip(New RectangleF(0, 0, w, h))

            If _cols <> 0 And _GridHeaderVisible Then
                ' we need to render the Header

                For t = 0 To _cols - 1
                    xof = GimmeXOffset(t)
                    If _colwidths(t) > 0 Then
                        gr.FillRectangle(New SolidBrush(_GridHeaderBackcolor), xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight)
                        gr.DrawString(_GridHeader(t), _GridHeaderFont, New SolidBrush(_GridHeaderForecolor), _
                                      New RectangleF(xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight), _GridHeaderStringFormat)
                        If _CellOutlines Then
                            gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof, orig.Y, _colwidths(t), _GridHeaderHeight))
                        End If
                    End If
                Next
                orig.Y = orig.Y + _GridHeaderHeight
            End If

            ' do we need to display the scrollbars

            RecalcScrollBars()

            If _BorderStyle = BorderStyle.Fixed3D Or _BorderStyle = BorderStyle.FixedSingle Then
                gr.DrawRectangle(New Pen(_BorderColor, 1), 0, 0, w - 1, h - 1)
            End If

#If EVALUATION Then
            Dim evalfont As Font = New Font("Courier New", 50, FontStyle.Bold)
            Dim evalstring As String = "EVALUATION VERSION"
            Dim evalmetrics As SizeF = gr.MeasureString(evalstring, evalfont)

            gr.DrawString(evalstring, evalfont, New SolidBrush(Color.IndianRed), (w / 2) - (evalmetrics.Width / 2), (h / 2) - (evalmetrics.Height / 2))
#End If

            _Painting = False

        End If

        grview.DrawImageUnscaled(bmp, 0, 0)
        gr.Dispose()
        bmp.Dispose()
        bmp = Nothing

    End Sub

    Private Sub RenderGridToGraphicsContext(ByVal gr As Graphics, ByVal Cliprect As Rectangle)
        Dim w As Integer = Me.AllColWidths()
        Dim h As Integer = Me.AllRowHeights()
        Dim orig As Point
        Dim t As Integer
        Dim xof As Integer
        Dim xxof, yyof, ofx, ofy As Integer
        Dim r, c As Integer
        Dim rh, rhy, rhx As Integer ' use for checkbox renderings
        Dim rowstart As Integer = -1
        Dim rowend As Integer = -1
        Dim colstart As Integer = -1
        Dim colend As Integer = -1
        Dim gyofset As Integer
        Dim renderstring As String = ""

        gr.SetClip(Cliprect)
        ofx = Cliprect.X
        ofy = Cliprect.Y

        If _gridForeColorList(0) Is Nothing Then
            _gridForeColorList(0) = New Pen(_DefaultForeColor)
        End If

        If _gridBackColorList(0) Is Nothing Then
            _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)
        End If

        If _GridHeaderVisible Then
            h += _GridHeaderHeight
        End If

        If _antialias Then
            gr.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        Else
            gr.SmoothingMode = Drawing2D.SmoothingMode.Default
            gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault
        End If

        ClearToBackgroundColor(gr)

        ' If we are disallowing selection of columns then make sure the Selected column variable is out of bounds
        If Not _AllowColumnSelection Then
            _SelectedColumn = -1
        End If

        If _GridTitleVisible Then
            ' we need to draw the title
            gr.FillRectangle(New SolidBrush(_GridTitleBackcolor), 0 + ofx, 0 + ofy, w, _GridTitleHeight)
            gr.DrawString(_GridTitle, _GridTitleFont, New SolidBrush(_GridTitleForeColor), 0 + ofx, 0 + ofy)
            orig.X = 0
            orig.Y = _GridTitleHeight
        Else
            orig.X = 0
            orig.Y = 0
        End If

        If _cols <> 0 And _GridHeaderVisible Then
            orig.Y = orig.Y + _GridHeaderHeight
        End If

        yyof = 0
        xxof = 0

        If _rows = 0 And _cols = 0 Then
            ' We have nothing else to draw so lets bail
            '_Painting = False
        Else

            rowstart = 0
            rowend = _rows - 1

            colstart = 0
            colend = _cols - 1

            ' time to render the grid here
            For r = rowstart To rowend
                gyofset = GimmeYOffset(r)
                For c = colstart To colend
                    xof = GimmeXOffset(c)
                    If _colwidths(c) > 0 Then

                        If _colPasswords(c) Is Nothing Then
                            renderstring = _grid(r, c)
                        Else
                            If _colPasswords(c) = "" Then
                                renderstring = _grid(r, c)
                            Else
                                renderstring = _colPasswords(c)
                            End If
                        End If

                        ' handle the Max characters display here

                        If _colMaxCharacters(c) <> 0 Then
                            If renderstring.Length > _colMaxCharacters(c) Then
                                renderstring = renderstring.Substring(0, _colMaxCharacters(c)) + "..."
                            End If
                        End If

                        If r = _SelectedRow Or c = _SelectedColumn Or _SelectedRows.Contains(r) Then
                            If r = _SelectedRow Or _SelectedRows.Contains(r) Then
                                ' we have a selected row override of selected column

                                gr.FillRectangle(New SolidBrush(_RowHighLiteBackColor), xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If

                                Else
                                    gr.DrawString(renderstring, _
                                                  _gridCellFontsList(_gridCellFonts(r, c)), _
                                                  New SolidBrush(_RowHighLiteForeColor), _
                                                  New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                                  _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If

                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If
                            Else
                                ' we have a selected Col

                                gr.FillRectangle(New SolidBrush(_ColHighliteBackColor), xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths(c), _rowheights(r))

                                If _colboolean(c) Then
                                    ' we have to render the the checkbox
                                    rh = _rowheights(r) - 2

                                    If rh > 14 Then

                                        rh = 14

                                    End If

                                    If rh < 6 Then

                                        rh = 6

                                    End If

                                    rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                    If rhx < 0 Then
                                        rhx = 0
                                    End If

                                    rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                    If rhy < 0 Then
                                        rhy = 0
                                    End If

                                    If UCase(renderstring) = "TRUE" Or _
                                       UCase(renderstring) = "YES" Or _
                                       UCase(renderstring) = "Y" Or _
                                       UCase(renderstring) = "1" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                    Else
                                        If UCase(renderstring) = "" Then
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                        Else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                        End If
                                    End If

                                Else
                                    gr.DrawString(renderstring, _
                                                  _gridCellFontsList(_gridCellFonts(r, c)), _
                                                  New SolidBrush(_RowHighLiteForeColor), _
                                                  New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                                  _gridCellAlignmentList(_gridCellAlignment(r, c)))

                                End If

                                If _CellOutlines Then
                                    gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _
                                                                                                _colwidths(c), _rowheights(r)))
                                End If

                            End If
                        Else
                            gr.FillRectangle(_gridBackColorList(_gridBackColor(r, c)), xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths(c), _rowheights(r))

                            If _colboolean(c) Then
                                ' we have to render the the checkbox
                                rh = _rowheights(r) - 2

                                If rh > 14 Then

                                    rh = 14

                                End If

                                If rh < 6 Then

                                    rh = 6

                                End If

                                rhx = (_colwidths(c) \ 2) - (rh \ 2)

                                If rhx < 0 Then
                                    rhx = 0
                                End If

                                rhy = (_rowheights(r) \ 2) - (rh \ 2)

                                If rhy < 0 Then
                                    rhy = 0
                                End If

                                If UCase(renderstring) = "TRUE" Or _
                                   UCase(renderstring) = "YES" Or _
                                   UCase(renderstring) = "Y" Or _
                                   UCase(renderstring) = "1" Then
                                    ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked)
                                Else
                                    If UCase(renderstring) = "" Then
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive)
                                    Else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal)
                                    End If
                                End If

                            Else
                                gr.DrawString(renderstring, _
                                              _gridCellFontsList(_gridCellFonts(r, c)), _
                                              New SolidBrush(_RowHighLiteForeColor), _
                                              New RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths(c), _rowheights(r)), _
                                              _gridCellAlignmentList(_gridCellAlignment(r, c)))

                            End If
                            If _CellOutlines Then
                                gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _
                                                                                            _colwidths(c), _rowheights(r)))
                            End If
                        End If
                    End If
                Next
            Next

            ' recalc the top area so we can draw the header if its vivible
            If _GridTitleVisible Then
                orig.X = 0
                orig.Y = _GridTitleHeight
            Else
                orig.X = 0
                orig.Y = 0
            End If

            gr.SetClip(New RectangleF(0 + ofx, 0 + ofy, Cliprect.Width, Cliprect.Height))

            If _cols <> 0 And _GridHeaderVisible Then
                ' we need to render the Header

                For t = 0 To _cols - 1
                    xof = GimmeXOffset(t)
                    If _colwidths(t) > 0 Then
                        gr.FillRectangle(New SolidBrush(_GridHeaderBackcolor), xof - xxof + ofx, orig.Y + ofy, _colwidths(t), _GridHeaderHeight)
                        gr.DrawString(_GridHeader(t), _GridHeaderFont, New SolidBrush(_GridHeaderForecolor), _
                                      New RectangleF(xof - xxof + ofx, orig.Y + ofy, _colwidths(t), _GridHeaderHeight), _GridHeaderStringFormat)
                        If _CellOutlines Then
                            gr.DrawRectangle(New Pen(_CellOutlineColor), New Rectangle(xof - xxof + ofx, orig.Y + ofy, _colwidths(t), _GridHeaderHeight))
                        End If
                    End If
                Next
                orig.Y = orig.Y + _GridHeaderHeight
            End If

            ' do we need to display the scrollbars

            'RecalcScrollBars()

            If _BorderStyle = BorderStyle.Fixed3D Or _BorderStyle = BorderStyle.FixedSingle Then
                gr.DrawRectangle(New Pen(_BorderColor, 1), 0 + ofx, 0 + ofy, Cliprect.Width - 1, Cliprect.Height - 1)
            End If

#If EVALUATION Then
            Dim evalfont As Font = New Font("Courier New", 50, FontStyle.Bold)
            Dim evalstring As String = "EVALUATION VERSION"
            Dim evalmetrics As SizeF = gr.MeasureString(evalstring, evalfont)

            gr.DrawString(evalstring, evalfont, New SolidBrush(Color.IndianRed), (w / 2) - (evalmetrics.Width / 2), (h / 2) - (evalmetrics.Height / 2))
#End If

            _Painting = False

        End If

        'grview.DrawImageUnscaled(bmp, 0, 0)
        'gr.Dispose()
        'bmp.Dispose()
        'bmp = Nothing

    End Sub

    Private Function ReturnExcelColumn(ByVal intColumn As Integer) As String
        Try
            Dim arrAlphabet As New ArrayList
            arrAlphabet.Add("A")
            arrAlphabet.Add("B")
            arrAlphabet.Add("C")
            arrAlphabet.Add("D")
            arrAlphabet.Add("E")
            arrAlphabet.Add("F")
            arrAlphabet.Add("G")
            arrAlphabet.Add("H")
            arrAlphabet.Add("I")
            arrAlphabet.Add("J")
            arrAlphabet.Add("K")
            arrAlphabet.Add("L")
            arrAlphabet.Add("M")
            arrAlphabet.Add("N")
            arrAlphabet.Add("O")
            arrAlphabet.Add("P")
            arrAlphabet.Add("Q")
            arrAlphabet.Add("R")
            arrAlphabet.Add("S")
            arrAlphabet.Add("T")
            arrAlphabet.Add("U")
            arrAlphabet.Add("V")
            arrAlphabet.Add("W")
            arrAlphabet.Add("X")
            arrAlphabet.Add("Y")
            arrAlphabet.Add("Z")

            If intColumn <= 25 Then
                Return arrAlphabet.Item(intColumn)
            Else
                Dim idx As Integer = (intColumn \ 26)
                If idx = 0 Then
                    idx += 1
                End If
                If idx >= 1 Then
                    'If (intColumn - 1) - (idx * 26) < 0 Then
                    '    Return arrAlphabet.Item(idx - 1) + arrAlphabet.Item((intColumn) - (idx * 26))
                    'Else
                    Return arrAlphabet.Item(idx - 1) + arrAlphabet.Item(intColumn - (idx * 26))
                    'End If
                Else
                    Return "A" + arrAlphabet.Item((intColumn - (idx * 26)))
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToExcel.ReturnExcelColumn Error...")
            Return ""
        End Try

    End Function

    Private Function ReturnByteArrayAsHexString(ByVal Bytes() As Byte) As String
        Dim t As Integer
        Dim result As String = ""
        Dim a As String

        Try
            For t = 0 To Bytes.GetLength(0) - 1
                a = Microsoft.VisualBasic.Right("00" & Hex(Bytes(t)), 2)
                result = result & a
            Next
        Catch ex As Exception
            ' we should do something here besides bail

        End Try

        Return result

    End Function

    Private Function ReturnHTMLColor(ByVal col As Color) As String

        Dim result As String = Chr(34) + "#"

        result += Microsoft.VisualBasic.Right("0" + Hex(col.R), 2)
        result += Microsoft.VisualBasic.Right("0" + Hex(col.G), 2)
        result += Microsoft.VisualBasic.Right("0" + Hex(col.B), 2) + Chr(34)

        Return result

    End Function

    Private Sub SetCols(ByVal newColVal As Integer)
        Dim t As Integer

        hs.Value = 0

        If _cols = 0 Then
            ' we have no columns now so lets just set them
            ReDim _colwidths(newColVal)
            For t = 0 To newColVal - 1
                _colwidths(t) = _DefaultColWidth
            Next
            _cols = newColVal
        Else
            _cols = newColVal
        End If

        RedimTable()

    End Sub

    Private Sub SetRows(ByVal newRowVal As Integer)
        Dim t As Integer

        vs.Value = 0

        If _rows = 0 Then
            ' we have no rows now so lets start things off
            ReDim _rowheights(newRowVal)
            For t = 0 To newRowVal - 1
                _rowheights(t) = _DefaultRowHeight
            Next
            _rows = newRowVal
        Else
            _rows = newRowVal

        End If

        RedimTable()

    End Sub

    Private Function SplitLongString(ByVal input As String, ByVal breaklen As Integer) As String

        Dim splitstringarray As String() = input.Split(" ".ToCharArray())

        Dim ret As String = ""
        Dim subret As String = ""

        Dim t As Integer = 0

        For t = 0 To splitstringarray.GetUpperBound(0)

            If splitstringarray(t).Trim = System.Environment.NewLine Then
                splitstringarray(t) = ""
            End If

            subret += " " + splitstringarray(t).Trim()

            If subret.Length >= breaklen Then
                ret += subret + System.Environment.NewLine
                subret = ""
            End If
        Next


        ret += subret

        ret = ret.Trim()

        If ret.EndsWith(System.Environment.NewLine) Then
            ret = ret.Substring(1, ret.Length - System.Environment.NewLine.Length)
        End If


        Return ret

    End Function

    Private Sub TearAwayColumID(ByVal id As Integer)

        If TearAways.Count > 0 Then
            Dim t As Integer
            For t = 0 To TearAways.Count - 1
                Dim ta As TearAwayWindowEntry = TearAways.Item(t)
                If ta.ColID = id Then
                    ' we already got one of these
                    ta.Winform.BringToFront()
                    ta.Winform.Focus()
                    Exit Sub
                End If
            Next
        End If

        Dim tear As New TearAwayWindowEntry
        Dim TearItem As New frmColumnTearAway(HeaderLabel(id))
        TearItem.Show()

        TearItem.ListItems = Me.GetColAsArrayList(id)
        TearItem.GridParent = Me
        TearItem.Colid = id
        TearItem.DefaultSelectionColor = _RowHighLiteBackColor
        TearItem.GridDefaultBackColor = _DefaultBackColor
        TearItem.GridDefaultForeColor = _DefaultForeColor
        TearItem.SelectedRow = _SelectedRow

        tear.Winform = TearItem

        tear.ColID = id
        tear.SetTearAwayScrollParameters(vs.Minimum, vs.Maximum, vs.Visible)

        'tear.ShowTearAway()
        TearAways.Add(tear)

    End Sub

#End Region

#Region " Public Methods "
    ''' <summary>
    ''' Takes the DelimitedSTringArray string and splits it up on the Delimiter. Then adds a row to the grids contents
    ''' filling the newly added row with the split fields from the DelimitedStringArray.
    ''' </summary>
    ''' <param name="DelimitedStringArray"></param>
    ''' <param name="Delimiter"></param>
    ''' <remarks></remarks>
    Public Sub AddRowToGrid(ByVal DelimitedStringArray As String, ByVal Delimiter As String)

        Dim _oldPainting As Boolean

        _oldPainting = _Painting


        _Painting = True


        ' added Aug 2, 2004
        ' Larry found an oddity in this routine when he was adding to the grid that had no columns
        ' we decided to make it generate columns by default if none where already in the grid
        '
        If _cols = 0 Then
            ' we have no columns in the grid so we need to add some now
            Dim b() As String = DelimitedStringArray.Split(Delimiter)
            Dim xx As Integer = 0

            Me.Cols = b.GetUpperBound(0) + 1

            For xx = 0 To _cols - 1
                Me.HeaderLabel(xx) = "COLUMN " & xx.ToString
            Next

            Me.AutoSizeCellsToContents = True

            Me.Refresh()

        End If

        ' end of Aug 2, 2004

        Dim a() As String = DelimitedStringArray.Split(Delimiter.ToCharArray(), _cols)
        Dim x As Integer = 0

        Me.Rows = _rows + 1

        For x = 0 To _cols
            _grid(_rows - 1, x) = ""
        Next

        For x = 0 To a.GetUpperBound(0)
            _grid(_rows - 1, x) = a(x)
        Next

        _AutoSizeAlreadyCalculated = False

        _Painting = _oldPainting

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Takes the DelimitedSTringArray string and splits it up on the default delimiter of '|'.
    ''' Then adds a row to the grids contents filling the newly added row with the split fields 
    ''' from the DelimitedStringArray.
    ''' </summary>
    ''' <param name="DelimitedStringArray"></param>
    ''' <remarks></remarks>
    Public Sub AddRowToGrid(ByVal DelimitedStringArray As String)

        AddRowToGrid(DelimitedStringArray, "|")

    End Sub

    ''' <summary>
    ''' Sets all cells in the grid to be rendered using the Font style specified by fnt
    ''' </summary>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub AllCellsUseThisFont(ByVal fnt As Font)
        Dim r, c As Integer
        Dim selrow As Integer = -1

        'If _MouseRow >= 0 And _MouseRow <= _rows - 1 Then
        '    selrow = _MouseRow
        'End If

        For r = 0 To _rows - 1
            For c = 0 To _cols - 1
                _gridCellFonts(r, c) = 0 ' fnt
            Next
        Next

        'Me.txtKeyHandler.Top = Me.TAIGCanvas.Top
        'Me.txtKeyHandler.Left = Me.TAIGCanvas.Left
        'TAIGPanel.ScrollControlIntoView(txtKeyHandler)

        _DefaultCellFont = fnt
        _gridCellFontsList(0) = fnt
        _AutoSizeAlreadyCalculated = False

        Me.Invalidate()

        'If selrow <> -1 Then
        '    Me.DoSelectedRowHighlight(selrow)
        'End If

    End Sub

    ''' <summary>
    ''' Sets all cells in the grid to use the forgroundcolor specified by fcol
    ''' </summary>
    ''' <param name="fcol"></param>
    ''' <remarks></remarks>
    Public Sub AllCellsUseThisForeColor(ByVal fcol As Color)
        Dim r, c As Integer

        For r = 0 To _rows - 1
            For c = 0 To _cols - 1
                _gridForeColor(r, c) = 0
            Next
        Next

        _DefaultForeColor = fcol
        _gridForeColorList(0) = New Pen(fcol)

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Will decrease the size of all displayed fonts in the grid by a single point
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AllFontsSmaller()

        miFontsSmaller_Click(Me, New System.EventArgs)
        miHeaderFontSmaller_Click(Me, New System.EventArgs)
        miTitleFontSmaller_Click(Me, New System.EventArgs)


    End Sub

    ''' <summary>
    ''' Will increase the size of all displayed fonts in the grid by a single point
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AllFontsLarger()

        miFontsLarger_Click(Me, New System.EventArgs)
        miHeaderFontLarger_Click(Me, New System.EventArgs)
        miTitleFontLarger_Click(Me, New System.EventArgs)

    End Sub

    ''' <summary>
    ''' Will set all the rows in the grid to use the background color specified by startcolor
    ''' </summary>
    ''' <param name="startcolor"></param>
    ''' <remarks></remarks>
    Public Sub AllRowsThisColor(ByVal startcolor As Color)
        Dim r, c As Integer

        _Painting = True

        For r = 0 To _rows - 1
            For c = 0 To Cols - 1
                _gridBackColor(r, c) = GetGridBackColorListEntry(New SolidBrush(startcolor))
            Next
        Next

        _Painting = False

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Will instruct all data represented in the grid to be colored using the specified color Startcolor
    ''' </summary>
    ''' <param name="startcolor"></param>
    ''' <remarks></remarks>
    Public Sub AllTextThisColor(ByVal startcolor As Color)
        Dim r, c As Integer

        _Painting = True

        For r = 0 To _rows - 1
            For c = 0 To Cols - 1
                _gridForeColor(r, c) = GetGridForeColorListEntry(New Pen(startcolor))
            Next
        Next

        _Painting = False

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Will take the parmeters startcolor and alternatecolor and color every other row in the grid using these two
    ''' colors 
    ''' </summary>
    ''' <param name="startcolor"></param>
    ''' <param name="alternatecolor"></param>
    ''' <remarks></remarks>
    Public Sub AlternateRowColoration(ByVal startcolor As Color, ByVal alternatecolor As Color)
        Dim r, c As Integer
        Dim flag As Boolean = False

        If _rows < 2 Then
            ' we ain't got enough rows to alternate colorize
            Exit Sub
        End If

        _Painting = True

        _alternateColorationMode = True
        _alternateColorationBaseColor = startcolor
        _alternateColorationALTColor = alternatecolor

        For r = 1 To _rows - 1
            For c = 0 To Cols - 1
                If flag Then
                    _gridBackColor(r, c) = GetGridBackColorListEntry(New SolidBrush(alternatecolor))
                Else
                    _gridBackColor(r, c) = GetGridBackColorListEntry(New SolidBrush(startcolor))
                End If
            Next
            flag = Not flag
        Next

        _Painting = False

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' will take the property defined basecolor and altcolor and apply the alternaterowcoloration 
    ''' process to the contents of the grid
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AlternateRowColoration()
        AlternateRowColoration(_alternateColorationBaseColor, _alternateColorationALTColor)
    End Sub

    ''' <summary>
    ''' Will iterate through the maintained list of tearaway windows attemptiong to place 
    ''' them on the screen so that they dont overlap each other. Simillar to the old windows arrange windows
    ''' menu item from the wfw 1.1 and windows 95/98 days
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ArrangeTearAwayWindows()

        Dim maxy As Integer = 0

        Dim t As Integer
        Dim rect As Rectangle = System.Windows.Forms.SystemInformation.WorkingArea
        Dim x, y As Integer
        Dim tear As TearAwayWindowEntry


        If TearAways.Count = 0 Then
            ' we ain't got any tearaways so lets bail
            Exit Sub
        End If

        If _TearAwayWork Then
            Exit Sub
        End If

        _TearAwayWork = True

        ' lets see if we can minimize the windows size first here

        For t = 0 To TearAways.Count - 1
            tear = TearAways(t)
            y = tear.Winform.MaxRenderHeight()

            If maxy < y Then
                maxy = y
            End If

        Next

        ' now maxy is the largest windows maximum render height so lets compare

        If maxy < 350 Then
            For t = 0 To TearAways.Count - 1
                tear = TearAways(t)
                tear.Winform.Height = maxy + 10
            Next
        End If

        maxy = 0

        ' now to so the moving about


        ' first we need to get the height of the largest window
        For t = 0 To TearAways.Count - 1
            tear = TearAways(t)
            If maxy < tear.Winform.Height Then
                maxy = tear.Winform.Height
            End If
        Next

        x = 0
        y = 0

        ' ok now we have the height so  lets start arranging them
        For t = 0 To TearAways.Count - 1
            tear = TearAways(t)

            If x + tear.Winform.Width > rect.Width Then
                ' that window is off screen so lets organize it down a bit
                x = 0
                If y + (maxy * 2) > rect.Height Then
                    y = 0
                Else
                    y += maxy
                End If
            End If

            Dim loc As New Point(x, y)

            tear.Winform.Location = loc

            x += tear.Winform.Width

        Next

        _TearAwayWork = False

        'System.Threading.Thread.Sleep(50)

        'For t = 0 To TearAways.Count - 1
        '    Dim tear As TearAwayWindowEntry = TearAways(t)
        '    tear.Winform.Invalidate()
        'Next

    End Sub

    ''' <summary>
    ''' Erase's the contents of the grid and sets it up to contain 1 row and 1 column
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BlankTheGrid()

        Me.Cols = 1
        Me.Rows = 1

        Me.ClearAllText()

        Me.ColWidth(0) = Me.Width
        Me.RowHeight(0) = Me.Height

        _AutoSizeAlreadyCalculated = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' resets all the columns of the grid to not be displaying boolean datatypes ( CheckBoxes )
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearAllGridCheckboxStates()
        Dim c As Integer
        For c = 0 To _cols - 1

            _colboolean(c) = False

        Next
        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Clears the text in the grid but leaves the columns and rows in place
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearAllText()
        ReDim _grid(_rows, _cols)
        _AutoSizeAlreadyCalculated = False
        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Clears the internal column restriction list allows all editable columns to contain any arbritrary
    ''' textual data.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearAllColumnEditRestrictionLists()
        _colEditRestrictions.Clear()
    End Sub

    ''' <summary>
    ''' Removes the column edit restrinctions from the columnid designated by colid
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <remarks></remarks>
    Public Sub ClearSpecificColumnEditRestrictionList(ByVal colid As Integer)

        For Each it As EditColumnRestrictor In _colEditRestrictions
            If it.ColumnID = colid Then
                _colEditRestrictions.Remove(it)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Takes the contents of the grid and copys it to the clipboard as a Tab delimited array of text elements
    ''' suitable for pasting into excel or word.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CopyGridToClipboard()
        Dim x, y As Integer
        Dim s As String = ""

        For y = 0 To _rows - 1
            For x = 0 To _cols - 1
                s = s & _grid(y, x)
                If x = _cols - 1 Then
                    s = s & vbCrLf
                Else
                    s = s & vbTab
                End If
            Next
        Next

        System.Windows.Forms.Clipboard.SetDataObject(s)
    End Sub

    ''' <summary>
    ''' Applys the designated stringformat object sf to the contents of the colum designated as c
    ''' </summary>
    ''' <param name="c"></param>
    ''' <param name="sf"></param>
    ''' <remarks></remarks>
    Public Sub ColumnFormat(ByVal c As Integer, ByVal sf As StringFormat)
        Dim r As Integer

        If c >= _cols Or _rows < 1 Or c < 0 Then
            Exit Sub
        End If

        For r = 0 To _rows - 1
            _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)
        Next

        _AutoSizeAlreadyCalculated = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' applys the standard format specification of currency to the designated column at colume id C
    ''' </summary>
    ''' <param name="c"></param>
    ''' <remarks></remarks>
    Public Sub ColumnFormatasMoney(ByVal c As Integer)
        Dim r As Integer
        Dim sf As New StringFormat

        If c >= _cols Or _rows < 1 Or c < 0 Then
            Exit Sub
        End If

        'sf.LineAlignment = StringAlignment.Far
        sf.LineAlignment = StringAlignment.Near
        sf.Alignment = StringAlignment.Far

        For r = 0 To _rows - 1
            If IsNumeric(_grid(r, c)) Then
                _grid(r, c) = _
                    Format(Val(_grid(r, c).Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")), "C")
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)
            End If
        Next

        _AutoSizeAlreadyCalculated = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' applys the standard format specification of numbers to the designated column at colume id C
    ''' </summary>
    ''' <param name="c"></param>
    ''' <param name="sFormat"></param>
    ''' <remarks></remarks>
    Public Sub ColumnFormatasNumber(ByVal c As Integer, ByVal sFormat As String)
        Dim r As Integer
        Dim sf As New StringFormat

        If c >= _cols Or _rows < 1 Or c < 0 Then
            Exit Sub
        End If

        'sf.LineAlignment = StringAlignment.Far
        sf.LineAlignment = StringAlignment.Near
        sf.Alignment = StringAlignment.Far

        For r = 0 To _rows - 1
            If IsNumeric(_grid(r, c)) Then
                _grid(r, c) = _
                    Format(Val(_grid(r, c).Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")), sFormat)
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)
            End If
        Next

        _AutoSizeAlreadyCalculated = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' applys the standard format specification of short date to the designated column at colume id C
    ''' </summary>
    ''' <param name="c"></param>
    ''' <remarks></remarks>
    Public Sub ColumnFormatasShortDate(ByVal c As Integer)
        Dim r As Integer

        If c >= _cols Or _rows < 1 Or c < 0 Then
            Exit Sub
        End If

        For r = 0 To _rows - 1
            If IsDate(_grid(r, c)) Then
                _grid(r, c) = Format(_grid(r, c), "Short Date")
            End If
        Next

        _AutoSizeAlreadyCalculated = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Will return a string that is representative of an SQL script that will write the gridshape and its contents
    ''' to a table in a database that the resulting script is handed to.
    ''' </summary>
    ''' <param name="tname"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreatePersistanceScript(ByVal tname As String) As String

        Dim result As String = ""
        Dim fname As String = ""
        Dim maxl(_cols) As Integer

        Dim sb As New StringBuilder

        Dim m, r, c As Integer

        ' figure out how large each column needs to be to persist the grids contents

        For c = 0 To _cols - 1
            m = 0
            For r = 0 To _rows - 1
                If _grid(r, c) Is Nothing Then
                    ' do nothing here then
                Else
                    If _grid(r, c).Length > m Then
                        m = _grid(r, c).Length
                    End If
                End If

            Next
            ' range check the size of the result to the maximum varchar size
            If m > 8000 Then
                m = 8000
            End If

            ' if we got no data then set the filed to hold something
            If m = 0 Then
                m = 10
            End If

            ' set the size in the array
            maxl(c) = m
        Next

        result = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" + tname + "]') " + _
                 "and OBJECTPROPERTY(id, N'IsUserTable') = 1)" + vbCrLf + _
                 "drop table [dbo].[" + tname + "] " + vbCrLf + _
                 "GO" + vbCrLf + vbCrLf



        result += "CREATE TABLE [dbo].[" + tname + "] (" + vbCrLf

        result += vbTab + "[ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," + vbCrLf

        For c = 0 To _cols - 1
            fname = HeaderLabel(c).ToUpper()

            If fname = "ID" Then
                fname = "ID_DATA"
            End If

            result += vbTab + "[" + fname + "] [VARCHAR] (" + maxl(c).ToString() + ") NULL ," + vbCrLf

        Next

        result += ") ON [PRIMARY]" + vbCrLf + "GO " + vbCrLf + vbCrLf

        sb.Append(result)

        result = ""

        For r = 0 To _rows - 1
            result = "INSERT INTO [" + tname + "] ("
            For c = 0 To _cols - 1
                fname = HeaderLabel(c).ToUpper()

                If fname = "ID" Then
                    fname = "ID_DATA"
                End If
                result += "[" + fname + "],"
            Next

            result = result.Substring(0, result.Length - 1) + ") VALUES ("

            For c = 0 To _cols - 1


                If _grid(r, c) Is Nothing Then
                    fname = "{null}"
                Else
                    fname = _grid(r, c)
                End If

                If fname.Length > maxl(c) Then
                    fname = fname.Substring(0, maxl(c)).Replace("'", "''")
                Else
                    fname = fname.Replace("'", "''")
                End If

                If fname = "{null}" Then
                    result += "NULL"
                Else
                    result += "'" + fname + "'"
                End If

                If c < Cols - 1 Then
                    result += ","
                Else
                    result += ")" + vbCrLf + "GO" + vbCrLf + vbCrLf
                End If

            Next

            sb.Append(result)

        Next

        Return sb.ToString()

    End Function

    ''' <summary>
    ''' Will return a string containing an HTML table representation of the grids contents
    ''' Borderval is the size parameter of the tables borders
    ''' Matchcolors will turn on or off the attempt to set the colors of the table to match thos of the grid itself
    ''' OmitNulls will have the rendering of empty cells in the grid or not. (creating holes in the resuting
    ''' html output)
    ''' </summary>
    ''' <param name="BorderVal"></param>
    ''' <param name="MatchColors"></param>
    ''' <param name="OmitNulls"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateHTMLTableScript(ByVal BorderVal As Integer, _
                                          ByVal MatchColors As Boolean, _
                                          ByVal OmitNulls As Boolean) As String

        Dim rb As New StringBuilder
        Dim result As String = ""
        Dim rr As String = ""

        Dim r, c As Integer

        If MatchColors Then

            If BorderVal > 0 Then
                rb.Append("<TABLE BGCOLOR = " + ReturnHTMLColor(Me.BackColor) + " " + _
                "BORDER = " + Chr(34) + BorderVal.ToString() + Chr(34) + ">" + vbCrLf)
            Else
                rb.Append("<TABLE BGCOLOR = " + ReturnHTMLColor(Me.BackColor) + " " + ">" + vbCrLf)
            End If

            rb.Append("<TR><TD BGCOLOR = " + ReturnHTMLColor(Me.TitleBackColor) + " COLSPAN =" + Chr(34) + _cols.ToString() + Chr(34) + ">")
            rb.Append(Me.TitleText + "</TD></TR>" + vbCrLf)

            rb.Append("<TR>")

            For c = 0 To _cols - 1
                rb.Append("<TH BGCOLOR=" + ReturnHTMLColor(Me.GridHeaderBackColor) + ">" + HeaderLabel(c) + "</TH>")
            Next

            rb.Append("</TR>" + vbCrLf)

            For r = 0 To _rows - 1
                rb.Append("<TR>")
                For c = 0 To _cols - 1

                    Dim sb As SolidBrush = _gridBackColorList(_gridBackColor(r, c))
                    Dim txt As String = _grid(r, c) + ""

                    If OmitNulls And txt = "{null}" Then
                        txt = ""
                    End If

                    rb.Append("<TD BGCOLOR =" + ReturnHTMLColor(sb.Color) + ">" + txt + "</TD>")
                Next
                rb.Append("</TR>" + vbCrLf)
            Next

            rb.Append("</TABLE>" + vbCrLf)

        Else

            If BorderVal > 0 Then
                rb.Append("<TABLE BORDER = " + Chr(34) + BorderVal.ToString() + Chr(34) + ">" + vbCrLf)
            Else
                rb.Append("<TABLE>" + vbCrLf)
            End If

            rb.Append("<TR><TD COLSPAN =" + Chr(34) + _cols.ToString() + Chr(34) + ">")
            rb.Append(Me.TitleText + "</TD></TR>" + vbCrLf)

            result += "<TR>"

            For c = 0 To _cols - 1
                rb.Append("<TH>" + HeaderLabel(c) + "</TH>")
            Next

            rb.Append("</TR>" + vbCrLf)

            For r = 0 To _rows - 1
                rb.Append("<TR>")
                For c = 0 To _cols - 1
                    Dim txt As String = _grid(r, c) + ""

                    If OmitNulls And txt = "{null}" Then
                        txt = ""
                    End If

                    rb.Append("<TD>" + txt + "</TD>")
                Next
                rb.Append("</TR>" + vbCrLf)
            Next

            rb.Append("</TABLE>" + vbCrLf)

        End If

        Return rb.ToString()

    End Function

    ''' <summary>
    ''' Overload that will set the border to 1 pixel, Matchcolors, and omitnulls
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateHTMLTableScript() As String
        ' Will use the default border of 1
        ' Will Matchcolors
        ' Will OmitNulls

        Return CreateHTMLTableScript(1, True, True)

    End Function

    ''' <summary>
    ''' Will deslect all rows in the grid if any are selected
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DeSelectAllRows()
        _SelectedRows.Clear()
        _SelectedRow = -1
        Me.Refresh()
    End Sub

    ''' <summary>
    ''' Will set the tooltip on the mouse to match the cells contents that its hovering over
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ttText"></param>
    ''' <remarks></remarks>
    Public Sub DisplayGridToolTip(ByVal sender As Object, ByVal ttText As String)

        If _TearAwayWork Then
            Exit Sub
        End If


        If TypeOf sender Is TAIGridControl Then
            _TTip.SetToolTip(sender, ttText)
            _TTip.ShowAlways = True
            _TTip.Active = True

        Else
            If TypeOf sender Is frmColumnTearAway Then
                ' we are a hovering on a tearaway form so lets pass it to that form

                Dim t As Integer
                Dim f As frmColumnTearAway = DirectCast(sender, frmColumnTearAway)

                For t = 0 To TearAways.Count - 1
                    Dim ti As TearAwayWindowEntry = DirectCast(TearAways.Item(t), TearAwayWindowEntry)

                    If f.Colid = ti.ColID Then
                        ti.Winform.ShowToolTipOnForm(ttText)
                        Exit For
                    End If
                Next

            Else
                _TTip.SetToolTip(sender, ttText)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Will hide the tooltip on the mouse pointer if its visible
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub HideGridToolTip()

        If _TearAwayWork Then
            Exit Sub
        End If

        _TTip.Active = False

        If TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = 0 To TearAways.Count - 1
            Dim ti As TearAwayWindowEntry

            ti = TearAways.Item(t)

            ti.Winform.HideToolTipOnForm()
        Next

    End Sub

    ''' <summary>
    ''' The <c>DoControlBreakProcessing</c> method will as its name indicates conduct a good oldfasioned sub-total
    ''' and grand-total parse on the contents of the grid using the old style Cobol rules for control break processing.
    ''' <list type="bullet">
    ''' <item>
    ''' <c>BreakColIntArrayList</c> list needs to contain the column IDs that will be looked at to determine 
    ''' where to break and subtotal.
    ''' </item>
    ''' <item>
    ''' <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
    ''' be calculated. 
    ''' </item> 
    ''' <item>
    ''' <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
    ''' before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
    ''' </item>
    ''' <item>
    ''' <c>ColumnToPlaceSubTotalTextIn</c> indicates the column to reiterate the criteria for the break into.
    ''' Think of this as the label to apply to the subtotal rows. 
    ''' </item>
    ''' <item>
    ''' <c>SubtotalText</c> is an arbritrary string of text to be appended to the lable defined above. 
    ''' </item>
    ''' <item>
    ''' <c>RightAlignSubTotalText</c> will allow or disallow the right aligning of the resulting subtotal lables. 
    ''' </item>
    ''' <item>
    ''' <c>ColorForSubTotalRows</c> defines the backgroundcolor to use when inserting a subtotal row into the 
    ''' resulting output. 
    ''' </item>
    ''' <item>
    ''' <c>BlankSeperateBreaks</c> will allow or disallow the insertion of an additional blank row after a 
    ''' subtotal operation.
    ''' </item>
    ''' <item> 
    ''' <c>EchoBreakFieldsOnSubTotalLines</c> will allow or disallow the echoing of the criteria for the
    ''' above subtotal operation on the line with the subtotal figures itself. 
    ''' </item>
    ''' <item>
    ''' <c>TreatBlanksAsSame</c> will force the parser to treat a blank field in a row to be treated as the 
    ''' most recent previous non blank field for the purposes of determining that the break is necessary.
    ''' </item>
    ''' </list>
    ''' Note:
    '''      Because the control break process works from top to bottom on the current contents of the grid
    ''' those contents should be sorted as the results will not have any real meaning if the grids contents 
    ''' are not sorted before the call to this method.
    ''' </summary>
    ''' <param name="BreakColIntArrayList"></param>
    ''' <param name="SumColumnIntegerArraylist"></param>
    ''' <param name="IgnoreCase"></param>
    ''' <param name="ColumnToPlaceSubtotalTextIn"></param>
    ''' <param name="SubtotalText"></param>
    ''' <param name="RightAlignSubTotalText"></param>
    ''' <param name="ColorForSubTotalRows"></param>
    ''' <param name="BlankSeperateBreaks"></param>
    ''' <param name="EchoBreakFieldsOnSubtotalLines"></param>
    ''' <param name="TreatBlanksAsSame"></param>
    ''' <remarks></remarks>
    Public Sub DoControlBreakProcessing(ByVal BreakColIntArrayList As ArrayList, _
                                                    ByVal SumColumnIntegerArraylist As ArrayList, _
                                                    ByVal IgnoreCase As Boolean, _
                                                    ByVal ColumnToPlaceSubtotalTextIn As Integer, _
                                                    ByVal SubtotalText As String, _
                                                    ByVal RightAlignSubTotalText As Boolean, _
                                                    ByVal ColorForSubTotalRows As System.Drawing.Color, _
                                                    ByVal BlankSeperateBreaks As Boolean, _
                                                    ByVal EchoBreakFieldsOnSubtotalLines As Boolean, _
                                                    ByVal TreatBlanksAsSame As Boolean)

        Dim breakstring As String = ""
        Dim oldbreak As String = ""
        Dim gvalue As String = ""
        Dim scols(SumColumnIntegerArraylist.Count - 1) As Double
        Dim tscols(SumColumnIntegerArraylist.Count - 1) As Double

        Dim sfmt As New StringFormat

        sfmt.Alignment = StringAlignment.Far
        sfmt.LineAlignment = StringAlignment.Center


        Dim hdr() As String = _GridHeader

        Dim x, y, xx, ngridcurrow As Integer

        For x = 0 To scols.GetUpperBound(0)
            ' init the sums colums here to 0
            scols(x) = 0
            tscols(x) = 0
        Next

        For y = 0 To _rows - 1

            ' lets calculate how many breaks we are going to have

            For x = 0 To BreakColIntArrayList.Count - 1
                ' start off the Break

                If _grid(y, CInt(BreakColIntArrayList.Item(x))) = "" And TreatBlanksAsSame Then
                    Dim sss() As String = oldbreak.Split("|".ToCharArray())
                    If sss.GetUpperBound(0) >= x Then
                        breakstring += sss(x) + "|"
                    Else
                        breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                    End If
                Else
                    breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                End If


            Next

            If IgnoreCase Then
                breakstring = breakstring.ToUpper()
            End If

            If breakstring <> oldbreak Then
                xx += 1
                If BlankSeperateBreaks Then
                    xx += 1
                End If
                oldbreak = breakstring
            End If

            breakstring = ""

        Next

        ' dimension our new grid to get it ready to hold the manipulated data

        Dim ngrid(_rows + xx, _cols - 1) As String

        ' calculate the first breakstring

        breakstring = ""

        For x = 0 To BreakColIntArrayList.Count - 1
            ' start off the break
            If _grid(y, CInt(BreakColIntArrayList.Item(x))) = "" And TreatBlanksAsSame Then
                Dim sss() As String = oldbreak.Split("|".ToCharArray())
                If sss.GetUpperBound(0) >= x Then
                    breakstring += sss(x) + "|"
                Else
                    breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                End If
            Else
                breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
            End If
        Next

        If IgnoreCase Then
            breakstring = breakstring.ToUpper()
        End If

        oldbreak = ""

        For y = 0 To _rows - 1
            ' our main loop here

            breakstring = ""

            For x = 0 To BreakColIntArrayList.Count - 1
                ' start off the break
                If _grid(y, CInt(BreakColIntArrayList.Item(x))) = "" And TreatBlanksAsSame Then
                    Dim sss() As String = oldbreak.Split("|".ToCharArray())
                    If sss.GetUpperBound(0) >= x Then
                        breakstring += sss(x) + "|"
                    Else
                        breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                    End If
                Else
                    breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                End If
            Next

            If IgnoreCase Then
                breakstring = breakstring.ToUpper()
            End If

            If breakstring <> oldbreak Or oldbreak = "" Then
                ' we have a break
                ' are we on the first break
                If oldbreak = "" Then
                    ' yes we is
                    For x = 0 To BreakColIntArrayList.Count - 1
                        'breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                        ngrid(ngridcurrow, CInt(BreakColIntArrayList.Item(x))) = _grid(y, CInt(BreakColIntArrayList.Item(x)))
                    Next

                    oldbreak = breakstring
                Else

                    ' we have a real break here so we need to display the subtotal lines

                    For x = 0 To SumColumnIntegerArraylist.Count - 1
                        ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
                    Next

                    If EchoBreakFieldsOnSubtotalLines Then

                        If oldbreak.EndsWith("|") Then
                            oldbreak = oldbreak.Substring(0, oldbreak.Length - 1)
                        End If

                        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = oldbreak + " " + SubtotalText
                    Else
                        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = SubtotalText
                    End If


                    ' clear the sum columns now
                    For x = 0 To scols.GetUpperBound(0)
                        ' init the sums colums here to 0
                        scols(x) = 0
                    Next

                    ngridcurrow += 1

                    If BlankSeperateBreaks Then
                        ngridcurrow += 1
                    End If

                    For x = 0 To BreakColIntArrayList.Count - 1
                        'breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                        ngrid(ngridcurrow, CInt(BreakColIntArrayList.Item(x))) = _grid(y, CInt(BreakColIntArrayList.Item(x)))
                    Next

                    oldbreak = breakstring

                End If
            End If

            ' here we have to sum up the selected columns

            For x = 0 To SumColumnIntegerArraylist.Count - 1

                ' lets sanitize the grid value to make it convert numerically
                gvalue = _grid(y, CInt(SumColumnIntegerArraylist.Item(x))).Replace("$", "").Replace("(", "-").Replace(")", "")

                If IsNumeric(gvalue) Then
                    If gvalue.Split(".".ToCharArray())(0) = gvalue Then
                        ' singe a split returns the same value in position 0 then we dont have a decimal
                        ' treat it as an integer

                        scols(x) += CInt(gvalue)
                        tscols(x) += CInt(gvalue)

                    Else
                        ' its got a decimal in it so treat it as a double

                        scols(x) += CDbl(gvalue)
                        tscols(x) += CDbl(gvalue)
                    End If
                End If
                'ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
            Next

            For x = 0 To _cols - 1
                ' here we have to echo the Non Break columns and their values into the results grid
                If BreakColIntArrayList.Contains(x) Then
                    ' we are breaking on this column
                    ' so don't echo it
                Else
                    ngrid(ngridcurrow, x) = _grid(y, x)
                End If
            Next

            ngridcurrow += 1

        Next

        ' now lets do the final break here

        For x = 0 To SumColumnIntegerArraylist.Count - 1
            ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
        Next

        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = SubtotalText

        ngridcurrow += 1

        If BlankSeperateBreaks Then
            ngridcurrow += 1
        End If

        For x = 0 To SumColumnIntegerArraylist.Count - 1
            ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = tscols(x).ToString()
        Next

        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = "Grand Total"


        ' now lets push the new grid into the old grids contents

        PopulateGridFromArray(ngrid, _DefaultCellFont, _DefaultForeColor, False, False, hdr)

        For y = 0 To _rows - 1
            If _grid(y, ColumnToPlaceSubtotalTextIn) Is Nothing Then
                ' when there be nothing there we do nothing
            Else
                If RightAlignSubTotalText And _grid(y, ColumnToPlaceSubtotalTextIn).EndsWith(SubtotalText) Then
                    CellAlignment(y, ColumnToPlaceSubtotalTextIn) = sfmt
                    For xx = 0 To _cols - 1
                        CellBackColor(y, xx) = New SolidBrush(ColorForSubTotalRows)
                    Next
                Else
                    If _grid(y, ColumnToPlaceSubtotalTextIn).EndsWith(SubtotalText) Then
                        For xx = 0 To _cols - 1
                            CellBackColor(y, xx) = New SolidBrush(ColorForSubTotalRows)
                        Next
                    End If
                End If
            End If

        Next

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' The <c>DoControlBreakProcessing</c> method will as its name indicates conduct a good oldfasioned sub-total
    ''' and grand-total parse on the contents of the grid using the old style Cobol rules for control break processing.
    ''' <list type="bullet">
    ''' <item>
    ''' <c>BreakColIntArrayList</c> list needs to contain the column IDs that will be looked at to determine 
    ''' where to break and subtotal.
    ''' </item>
    ''' <item>
    ''' <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
    ''' be calculated. 
    ''' </item> 
    ''' <item>
    ''' <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
    ''' before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
    ''' </item>
    ''' <item>
    ''' <c>ColumnToPlaceSubTotalTextIn</c> indicates the column to reiterate the criteria for the break into.
    ''' Think of this as the label to apply to the subtotal rows. 
    ''' </item>
    ''' <item>
    ''' <c>SubtotalText</c> is an arbritrary string of text to be appended to the lable defined above. 
    ''' </item>
    ''' <item>
    ''' <c>RightAlignSubTotalText</c> will allow or disallow the right aligning of the resulting subtotal lables. 
    ''' </item>
    ''' <item>
    ''' <c>ColorForSubTotalRows</c> defines the backgroundcolor to use when inserting a subtotal row into the 
    ''' resulting output. 
    ''' </item>
    ''' <item>
    ''' <c>BlankSeperateBreaks</c> will allow or disallow the insertion of an additional blank row after a 
    ''' subtotal operation.
    ''' </item>
    ''' <item> 
    ''' <c>EchoBreakFieldsOnSubTotalLines</c> will allow or disallow the echoing of the criteria for the
    ''' above subtotal operation on the line with the subtotal figures itself. 
    ''' </item>
    ''' 
    ''' </list>
    ''' Note:
    '''      Because the control break process works from top to bottom on the current contents of the grid
    ''' those contents should be sorted as the results will not have any real meaning if the grids contents 
    ''' are not sorted before the call to this method.
    ''' </summary>
    ''' <param name="BreakColIntArrayList"></param>
    ''' <param name="SumColumnIntegerArraylist"></param>
    ''' <param name="IgnoreCase"></param>
    ''' <param name="ColumnToPlaceSubtotalTextIn"></param>
    ''' <param name="SubtotalText"></param>
    ''' <param name="RightAlignSubTotalText"></param>
    ''' <param name="ColorForSubTotalRows"></param>
    ''' <param name="BlankSeperateBreaks"></param>
    ''' <param name="EchoBreakFieldsOnSubtotalLines"></param>
    ''' <remarks></remarks>
    Public Sub DoControlBreakProcessing(ByVal BreakColIntArrayList As ArrayList, _
                                        ByVal SumColumnIntegerArraylist As ArrayList, _
                                        ByVal IgnoreCase As Boolean, _
                                        ByVal ColumnToPlaceSubtotalTextIn As Integer, _
                                        ByVal SubtotalText As String, _
                                        ByVal RightAlignSubTotalText As Boolean, _
                                        ByVal ColorForSubTotalRows As System.Drawing.Color, _
                                        ByVal BlankSeperateBreaks As Boolean, _
                                        ByVal EchoBreakFieldsOnSubtotalLines As Boolean)

        Dim breakstring As String = ""
        Dim oldbreak As String = ""
        Dim gvalue As String = ""
        Dim scols(SumColumnIntegerArraylist.Count - 1) As Double
        Dim tscols(SumColumnIntegerArraylist.Count - 1) As Double

        Dim sfmt As New StringFormat

        sfmt.Alignment = StringAlignment.Far
        sfmt.LineAlignment = StringAlignment.Center


        Dim hdr() As String = _GridHeader

        Dim x, y, xx, ngridcurrow As Integer

        For x = 0 To scols.GetUpperBound(0)
            ' init the sums colums here to 0
            scols(x) = 0
            tscols(x) = 0
        Next

        For y = 0 To _rows - 1

            ' lets calculate how many breaks we are going to have

            For x = 0 To BreakColIntArrayList.Count - 1
                ' start off the Break
                breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
            Next

            If IgnoreCase Then
                breakstring = breakstring.ToUpper()
            End If

            If breakstring <> oldbreak Then
                xx += 1
                If BlankSeperateBreaks Then
                    xx += 1
                End If
                oldbreak = breakstring
            End If

            breakstring = ""

        Next

        ' dimension our new grid to get it ready to hold the manipulated data

        Dim ngrid(_rows + xx, _cols - 1) As String

        ' calculate the first breakstring

        breakstring = ""

        For x = 0 To BreakColIntArrayList.Count - 1
            ' start off the Break
            breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
        Next

        If IgnoreCase Then
            breakstring = breakstring.ToUpper()
        End If

        oldbreak = ""

        For y = 0 To _rows - 1
            ' our main loop here

            breakstring = ""

            For x = 0 To BreakColIntArrayList.Count - 1
                ' start off the Break
                breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
            Next

            If IgnoreCase Then
                breakstring = breakstring.ToUpper()
            End If

            If breakstring <> oldbreak Or oldbreak = "" Then
                ' we have a break
                ' are we on the first break
                If oldbreak = "" Then
                    ' yes we is
                    For x = 0 To BreakColIntArrayList.Count - 1
                        'breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                        ngrid(ngridcurrow, CInt(BreakColIntArrayList.Item(x))) = _grid(y, CInt(BreakColIntArrayList.Item(x)))
                    Next

                    oldbreak = breakstring
                Else

                    ' we have a real break here so we need to display the subtotal lines

                    For x = 0 To SumColumnIntegerArraylist.Count - 1
                        ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
                    Next

                    If EchoBreakFieldsOnSubtotalLines Then

                        If oldbreak.EndsWith("|") Then
                            oldbreak = oldbreak.Substring(0, oldbreak.Length - 1)
                        End If

                        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = oldbreak + " " + SubtotalText
                    Else
                        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = SubtotalText
                    End If


                    ' clear the sum columns now
                    For x = 0 To scols.GetUpperBound(0)
                        ' init the sums colums here to 0
                        scols(x) = 0
                    Next

                    ngridcurrow += 1

                    If BlankSeperateBreaks Then
                        ngridcurrow += 1
                    End If

                    For x = 0 To BreakColIntArrayList.Count - 1
                        'breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                        ngrid(ngridcurrow, CInt(BreakColIntArrayList.Item(x))) = _grid(y, CInt(BreakColIntArrayList.Item(x)))
                    Next

                    oldbreak = breakstring

                End If
            End If

            ' here we have to sum up the selected columns

            For x = 0 To SumColumnIntegerArraylist.Count - 1

                ' lets sanitize the grid value to make it convert numerically
                gvalue = _grid(y, CInt(SumColumnIntegerArraylist.Item(x))).Replace("$", "").Replace("(", "-").Replace(")", "")

                If IsNumeric(gvalue) Then
                    If gvalue.Split(".".ToCharArray())(0) = gvalue Then
                        ' singe a split returns the same value in position 0 then we dont have a decimal
                        ' treat it as an integer

                        scols(x) += CInt(gvalue)
                        tscols(x) += CInt(gvalue)

                    Else
                        ' its got a decimal in it so treat it as a double

                        scols(x) += CDbl(gvalue)
                        tscols(x) += CDbl(gvalue)
                    End If
                End If
                'ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
            Next

            For x = 0 To _cols - 1
                ' here we have to echo the Non Break columns and their values into the results grid
                If BreakColIntArrayList.Contains(x) Then
                    ' we are breaking on this column
                    ' so don't echo it
                Else
                    ngrid(ngridcurrow, x) = _grid(y, x)
                End If
            Next

            ngridcurrow += 1

        Next

        ' now lets do the final break here

        For x = 0 To SumColumnIntegerArraylist.Count - 1
            ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = scols(x).ToString()
        Next

        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = SubtotalText

        ngridcurrow += 1

        If BlankSeperateBreaks Then
            ngridcurrow += 1
        End If

        For x = 0 To SumColumnIntegerArraylist.Count - 1
            ngrid(ngridcurrow, CInt(SumColumnIntegerArraylist.Item(x))) = tscols(x).ToString()
        Next

        ngrid(ngridcurrow, ColumnToPlaceSubtotalTextIn) = "Grand Total"


        ' now lets push the new grid into the old grids contents

        PopulateGridFromArray(ngrid, _DefaultCellFont, _DefaultForeColor, False, False, hdr)

        For y = 0 To _rows - 1
            If _grid(y, ColumnToPlaceSubtotalTextIn) Is Nothing Then
                ' when there be nothing there we do nothing
            Else
                If RightAlignSubTotalText And (_grid(y, ColumnToPlaceSubtotalTextIn).EndsWith(SubtotalText) Or _
                                                _grid(y, ColumnToPlaceSubtotalTextIn).EndsWith("Grand Total")) Then
                    CellAlignment(y, ColumnToPlaceSubtotalTextIn) = sfmt
                    For xx = 0 To _cols - 1
                        CellBackColor(y, xx) = New SolidBrush(ColorForSubTotalRows)
                    Next
                Else
                    If _grid(y, ColumnToPlaceSubtotalTextIn).EndsWith(SubtotalText) Or _
                        _grid(y, ColumnToPlaceSubtotalTextIn).EndsWith("Grand Total") Then
                        For xx = 0 To _cols - 1
                            CellBackColor(y, xx) = New SolidBrush(ColorForSubTotalRows)
                        Next
                    End If
                End If
            End If

        Next

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' This will take the supplied <c>BreakColIntArrayList</c> and do a cell colorization operation
    ''' on the current grids contents alternating between <c>StartColor</c> and <c>AltColor</c> on a change
    ''' in the cols id'd in the supplied arraylist.
    ''' </summary>
    ''' <param name="BreakColIntArrayList"></param>
    ''' <param name="StartColor"></param>
    ''' <param name="AltColor"></param>
    ''' <remarks></remarks>
    Public Sub DoControlBreakColorization(ByVal BreakColIntArrayList As ArrayList, _
                                          ByVal StartColor As Color, _
                                          ByVal AltColor As Color)


        Dim start As Boolean = False
        Dim a As String = "--------------------------------------------------------"

        Dim t As Integer = 0

        For t = 0 To Me.Rows - 1

            Dim aa As String = ""

            Dim tt As Integer = 0

            For tt = 0 To BreakColIntArrayList.Count - 1
                aa += _grid(t, tt)
            Next

            If aa <> a Then
                start = Not start
                a = aa
            End If

            For tt = 0 To Me.Cols - 1
                If start Then
                    Me.CellBackColor(t, tt) = New SolidBrush(StartColor)

                Else
                    Me.CellBackColor(t, tt) = New SolidBrush(AltColor)
                End If
            Next
        Next

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Another control break process, 
    ''' <list type="bullet">
    ''' <item>
    ''' <c>BreakColIntArrayValues</c> list needs to contain the column IDs that will be looked at to determine 
    ''' where to break and subtotal.
    ''' </item>
    ''' <item>
    ''' <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
    ''' before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
    ''' </item>
    ''' <item>
    ''' <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
    ''' be calculated. 
    ''' </item> 
    ''' <item>
    ''' <c>ColorForBreakSubtotals</c> defines the backgroundcolor to use when inserting a subtotal row into the 
    ''' resulting output. 
    ''' </item>
    ''' <item> 
    ''' <c>CutoffRow</c> The maximum row to search for in the grid for processing.
    ''' will stop processing at <c>CutoffRow</c>
    ''' </item>
    ''' 
    ''' </list>
    ''' Note:
    '''      Because the control break process works from top to bottom on the current contents of the grid
    ''' those contents should be sorted as the results will not have any real meaning if the grids contents 
    ''' are not sorted before the call to this method.
    ''' </summary>
    ''' <param name="BreakColArrayValues"></param>
    ''' <param name="ColToFindValues"></param>
    ''' <param name="IgnoreCase"></param>
    ''' <param name="SumColumnIntegerArrayList"></param>
    ''' <param name="ColorForBreakSubtotals"></param>
    ''' <param name="CutoffRow"></param>
    ''' <remarks></remarks>
    Public Sub DoControlBreakSubTotals(ByVal BreakColArrayValues As ArrayList, _
                                       ByVal ColToFindValues As Integer, _
                                       ByVal IgnoreCase As Boolean, _
                                       ByVal SumColumnIntegerArrayList As ArrayList, _
                                       ByVal ColorForBreakSubtotals As System.Drawing.Color, _
                                       ByVal CutoffRow As Integer)


        Dim r As Integer = _rows + 1
        Dim x, y, xx As Integer
        Dim s, ss As String

        Dim scols(SumColumnIntegerArrayList.Count - 1) As Double

        Me.Rows += BreakColArrayValues.Count + 1

        For x = 0 To BreakColArrayValues.Count - 1
            ' lets iterate through the BreakItems

            ' clear the sum columns first
            For xx = 0 To scols.GetUpperBound(0)
                scols(xx) = 0
            Next

            s = BreakColArrayValues.Item(x)

            For xx = 0 To CutoffRow
                If IgnoreCase Then
                    If UCase(s) = UCase(_grid(xx, ColToFindValues)) Then
                        For y = 0 To SumColumnIntegerArrayList.Count - 1

                            ss = _grid(xx, SumColumnIntegerArrayList.Item(y)) + ""
                            ss = ss.Replace("(", "-").Replace("$", "").Replace(")", "")

                            If IsNumeric(ss) Then
                                scols(y) += Convert.ToDouble(ss)
                            End If


                        Next
                    End If
                Else
                    If s = _grid(xx, ColToFindValues) Then
                        For y = 0 To SumColumnIntegerArrayList.Count - 1

                            ss = _grid(xx, SumColumnIntegerArrayList.Item(y)) + ""
                            ss = ss.Replace("(", "-").Replace("$", "").Replace(")", "")

                            If IsNumeric(ss) Then
                                scols(y) += Convert.ToDouble(ss)
                            End If


                        Next
                    End If
                End If
            Next

            _grid(r + x, ColToFindValues) = s

            For xx = 0 To SumColumnIntegerArrayList.Count - 1

                _grid(r + x, SumColumnIntegerArrayList.Item(xx)) = scols(xx).ToString

            Next

            Me.SetRowBackColor(r + x, ColorForBreakSubtotals)

        Next

        Me.Refresh()

    End Sub

    Public Sub DoSelectedRowHighlight()
        ' old stub routine

    End Sub

    ''' <summary>
    ''' Will select and highlight the indicated rownum as if the user had selected it with the mouse
    ''' </summary>
    ''' <param name="rownum"></param>
    ''' <remarks></remarks>
    Public Sub DoSelectedRowHighlight(ByVal rownum As Integer)

        If rownum < 0 Or rownum >= _rows Then
            Exit Sub
        End If

        If vs.Visible Then
            vs.Value = 0
        End If

        Me.SelectedRow = rownum
        '_SelectedRow = rownum

        Me.Invalidate()

    End Sub

#Region " Export to Excel Functions "

    ''' <summary>
    ''' Will instance Microsoft excel and place the contents of the grid on the first worksheet in the excel application
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ExportToExcel()

        Dim _excel As Object
        Dim _workbook As New Object

        Try

            _excel = CreateObject("Excel.Application")
            '_excel.Visible = True

            ' _excel.ScreenUpdating = False

            '_workbook = _excel.Workbooks.Add()

            ExportToExcel(_excel, _workbook, Me._excelWorkSheetName)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToExcel Error...")
            Exit Sub
        End Try

    End Sub

    ''' <summary>
    ''' Will take the supplied instance of Microsoft excel <c>_excel</c> and place the contents of the grid on the 
    ''' worksheet named <c>wsname</c> in the supplied workbook <c>_Workbook</c>
    ''' </summary>
    ''' <param name="_excel"></param>
    ''' <param name="_WorkBook"></param>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Public Sub ExportToExcel(ByVal _excel As Object, ByVal _WorkBook As Object, ByVal wsname As String)

        Dim frm As New frmExportingToExcelWorking
        Dim lastsheetname As String

        If _ShowExcelExportMessage Then
            frm.Show()
            frm.Refresh()
        End If

        Me.Refresh()
        Application.DoEvents()

        Dim TotalRows As Integer = -1
        Dim sh As Object
        Try
            Dim _br As System.Drawing.SolidBrush
            Dim rng As String
            Dim idx As Integer = 0
            Dim FirstColumn As String = "A"
            Dim LastColumn As String = ReturnExcelColumn(Me.Cols - 1)
            Dim CurrentSHidx As Integer = 1
            Dim ccidx As Integer = 0

            Dim rows As Integer = Me.Rows
            _WorkBook = _excel.Workbooks.Add

            If rows > _excelMaxRowsPerSheet Then
                idx = (rows / _excelMaxRowsPerSheet)

                ' get around the Int round up problem of VB

                If idx * _excelMaxRowsPerSheet > rows Then
                    idx -= 1
                End If

            End If

            '_WorkBook.Worksheets("Sheet2").Delete()
            '_WorkBook.Worksheets("Sheet3").Delete()

            Do Until idx = -1
                If idx > 0 Then
                    rows = _excelMaxRowsPerSheet
                Else
                    rows = Me.Rows - TotalRows
                End If

                If CurrentSHidx = 1 Then
                    lastsheetname = wsname + " " + CurrentSHidx.ToString
                    _WorkBook.Worksheets("Sheet1").Name = lastsheetname
                    sh = _WorkBook.ActiveSheet
                Else
                    _WorkBook.Worksheets.Add()
                    sh = _WorkBook.ActiveSheet

                    'Dim cmd As String = "After:=Sheets(" + Chr(34) + lastsheetname + Chr(34) + ")"

                    'sh.move(cmd)
                    lastsheetname = wsname + " " + CurrentSHidx.ToString
                    sh.Name = lastsheetname

                    'sh.Move(2)
                End If

                _excel.MaxChange = 0.001

                _excel.ActiveWorkbook.PrecisionAsDisplayed = False
                Dim arr(rows + 1, Me.Cols) As Object
                Dim r, c, rmod As Integer

                rmod = 0
                r = 0
                c = 0

                Do Until r = rows + 1
                    Do Until c = Me.Cols
                        If r > 0 Then
                            arr(r, c) = Me.item(TotalRows, c)
                        Else
                            arr(r, c) = Me.HeaderLabel(c)
                        End If
                        c += 1
                    Loop
                    r += 1
                    c = 0
                    TotalRows += 1
                Loop

                rng = FirstColumn + "1:" + LastColumn + (rows + 1).ToString

                sh.Range(rng).NumberFormat = "General"

                sh.Range(rng).Value = arr

                If _excelMatchGridColorScheme Then
                    ' header always on row 1 of the sheet
                    c = 1
                    rng = FirstColumn + "1:" + LastColumn + "1"
                    sh.Range(rng).interior.color = RGB(_GridHeaderBackcolor.R, _GridHeaderBackcolor.G, _GridHeaderBackcolor.B)

                    r = (CurrentSHidx - 1) * _excelMaxRowsPerSheet

                    Dim rmax As Integer = r + _excelMaxRowsPerSheet

                    If rmax > Me.Rows Then
                        rmax = Me.Rows
                    End If

                    c = 0

                    ' here we will blast the first range of standard color in a single shot

                    rng = FirstColumn + "2:" + LastColumn + (rows + 1).ToString
                    _br = Me._gridBackColorList(0) ' element 0 is the default/first backcolor

                    ' now lets blast this color into the background of the grid
                    sh.Range(rng).interior.color = RGB(_br.Color.R, _br.Color.G, _br.Color.B)

                    ccidx = 1

                    Do Until ccidx = _gridBackColorList.GetUpperBound(0)
                        If _gridBackColorList(ccidx) Is Nothing Then
                            'we are on a nothing here
                        Else
                            r = (CurrentSHidx - 1) * _excelMaxRowsPerSheet
                            rmod = (CurrentSHidx - 1) * _excelMaxRowsPerSheet
                            Do Until r = rmax
                                c = 0
                                Do Until c = Me.Cols
                                    If _gridBackColor(r, c) = ccidx Then
                                        ' we got a color thats different than the blasted backcolor


                                        rng = Me.ReturnExcelColumn(c) + (r - rmod + 2).ToString
                                        rng = rng + ":" + rng
                                        _br = Me._gridBackColorList(ccidx)
                                        sh.Range(rng).interior.color = RGB(_br.Color.R, _br.Color.G, _br.Color.B)
                                    End If
                                    c += 1
                                Loop
                                r += 1
                            Loop
                        End If
                        ccidx += 1
                    Loop

                    'Do Until r = rows
                    '    FirstColumn = "A"
                    '    FirstColumn = Me.ReturnExcelColumn(c)
                    '    Do Until c = Me.Cols
                    '        LastColumn = Me.ReturnExcelColumn(c)
                    '        _br = Me._gridBackColorList(_gridBackColor(r, c))

                    '        If _oldbr Is Nothing Then
                    '            _oldbr = _br.Clone
                    '        End If

                    '        If c = 0 Then
                    '            CurrentColor = _br
                    '        Else
                    '            If Not _br Is CurrentColor Then
                    '                LastColumn = Me.ReturnExcelColumn(c - 1)
                    '                rng = FirstColumn + (r + 2).ToString + ":" + LastColumn + (r + 2).ToString
                    '                sh.Range(rng).interior.color = RGB(CurrentColor.Color.R, CurrentColor.Color.G, CurrentColor.Color.B)
                    '                CurrentColor = _br
                    '                FirstColumn = Me.ReturnExcelColumn(c)
                    '                LastColumn = Me.ReturnExcelColumn(c)
                    '            End If
                    '        End If
                    '        c += 1
                    '    Loop
                    '    rng = FirstColumn + (r + 2).ToString + ":" + LastColumn + (r + 2).ToString
                    '    sh.Range(rng).interior.color = RGB(_br.Color.R, _br.Color.G, _br.Color.B)
                    '    r += 1
                    '    c = 0
                    'Loop
                Else
                    r = 2
                    If Me.ExcelUseAlternateRowColor Then
                        '38 seconds to load 5000 lines of claims data same as previous loop
                        Do Until r >= rows
                            rng = FirstColumn + r.ToString + ":" + LastColumn + r.ToString
                            'there are 56 possible colors, Lonnie uses number 35 in the grid.  Maybe 57, I didn't try index# 0...
                            sh.Range(rng).Interior.ColorIndex = 35
                            r += 2
                        Loop
                    End If
                End If

                If Me._excelIncludeColumnHeaders Then
                    sh.rows("1:1").Font.Bold = True
                    sh.rows("1:1").HorizontalAlignment = xlCenter
                End If

                If Me._excelShowBorders Or _excelOutlineCells Then

                    ' try to catch errors here just in case someone is putting massive amounts of text into
                    ' cells and the range selection diddys are breaking in excel versions less then 2007

                    Try
                        FirstColumn = "A"
                        LastColumn = ReturnExcelColumn(Me.Cols - 1)
                        rng = FirstColumn + "1:" + LastColumn + (rows + 1).ToString

                        sh.Range(rng).Borders(xlEdgeRight).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlEdgeRight).Weight = xlThin
                        sh.Range(rng).Borders(xlEdgeRight).ColorIndex = xlAutomatic

                        sh.Range(rng).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlEdgeLeft).Weight = xlThin
                        sh.Range(rng).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

                        sh.Range(rng).Borders(xlEdgeTop).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlEdgeTop).Weight = xlThin
                        sh.Range(rng).Borders(xlEdgeTop).ColorIndex = xlAutomatic

                        sh.Range(rng).Borders(xlEdgeBottom).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlEdgeBottom).Weight = xlThin
                        sh.Range(rng).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

                        sh.Range(rng).Borders(xlInsideVertical).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlInsideVertical).Weight = xlThin
                        sh.Range(rng).Borders(xlInsideVertical).ColorIndex = xlAutomatic

                        sh.Range(rng).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        sh.Range(rng).Borders(xlInsideHorizontal).Weight = xlThin
                        sh.Range(rng).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
                    Catch ex As Exception
                        ' yup it bombed but it should be ok to continue here

                    End Try

                End If

                If Me._excelAutoFitColumn Then
                    sh.Range(rng).EntireColumn.Autofit()
                End If

                If Me._excelAutoFitRow Then
                    sh.Range(rng).EntireRow.Autofit()
                End If

                'If Me._excelAutoFitColumn Then
                '    sh.Range(rng).EntireColumn.Autofit()
                'End If

                sh.PageSetup.Orientation = Me._excelPageOrientation

                _excel.ActiveWindow.WindowState = xlMaximized

                ' save the spreadsheet
                _excel.AlertBeforeOverwriting = False
                _excel.DisplayAlerts = False

                _excel.ScreenUpdating = True

                r = 0
                idx -= 1
                CurrentSHidx += 1
                TotalRows -= 1 'this subtracts 1 row incase there is another worksheet that is needed.  WIthout this the first row will be skipped
            Loop

            If _ShowExcelExportMessage Then
                frm.Hide()
                frm = Nothing
            End If

            _excel.Visible = True
            _WorkBook = Nothing
            sh = Nothing
            _excel = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()

        Catch ex As Exception

            If _ShowExcelExportMessage Then
                frm.Hide()
                frm = Nothing
            End If

            _WorkBook = Nothing
            sh = Nothing
            _excel = Nothing

            GC.Collect()
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRiDControl.ExportToExcel Error...")
        End Try

    End Sub

#End Region

#Region " Export to Text Functions "

    ''' <summary>
    ''' Will open the internal export filename dialog querying the user for the filename to export to
    ''' Will then export the contents of the grid to a textfile employing the properties setup by 
    ''' displayed dialog, Filename, Include column headers as fieldname, the field terminator, and the line 
    ''' terminator...
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ExportToText()

        Try

            Dim frmExport As New frmExportToText

            If frmExport.ShowDialog() = DialogResult.OK Then

                Dim sDelimiter As String = frmExport.Delimiter
                Dim sFilename As String = frmExport.Filename
                Dim bIncludeFieldNames As Boolean = frmExport.IncludeFieldNames
                Dim bIncludeLineTerminator As Boolean = frmExport.IncludeLineTerminator

                Me.ExportToText(sFilename, sDelimiter, bIncludeFieldNames, bIncludeLineTerminator)

            End If

            frmExport = Nothing


        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToText Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will export the contents of the grid to the supplied filename <c>sFilename</c> employing the internally set properties
    ''' to control the field delimiters, line terminators, and inclusion of column headers as field names...
    ''' </summary>
    ''' <param name="sFilename"></param>
    ''' <remarks></remarks>
    Public Sub ExportToText(ByVal sFilename As String)

        Try

            Me.ExportToText(sFilename, _delimiter, _includeFieldNames, _includeLineTerminator)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToText Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will export the contens of the grid to a file named <c>sFilename</c> with the field delimiters of <c>sDelimiter</c>
    ''' The internal properties will control the inclusion of column headers as field names and the characters used to terminate
    ''' the lines of output
    ''' </summary>
    ''' <param name="sFilename"></param>
    ''' <param name="sDelimiter"></param>
    ''' <remarks></remarks>
    Public Sub ExportToText(ByVal sFilename As String, ByVal sDelimiter As String)

        Try

            Me.ExportToText(sFilename, sDelimiter, _includeFieldNames, _includeLineTerminator)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToText Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will export the contents of the grid to a textfile name <c>sFilename</c> using <c>sDelimiter</c> for
    ''' field delimiters and using <c>bIncludeFieldNames</c> to control inclusion of the column headers as field names
    ''' the interna;l property for the end of the lines termination will be employed.
    ''' </summary>
    ''' <param name="sFilename"></param>
    ''' <param name="sDelimiter"></param>
    ''' <param name="bIncludeFieldNames"></param>
    ''' <remarks></remarks>
    Public Sub ExportToText(ByVal sFilename As String, ByVal sDelimiter As String, ByVal bIncludeFieldNames As Boolean)

        Try

            Me.ExportToText(sFilename, sDelimiter, bIncludeFieldNames, _includeLineTerminator)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToText Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will export the contents of the grid to <c>sFilename</c> employing <c>sDelimiter</c> for the field delimiters
    ''' <c>bIncludeFieldNames</c> to control inclusion of the column headers as field names, and the <c>bIncludeLineTerminator</c>
    ''' to control the CRLF at the end of the lines of output
    ''' </summary>
    ''' <param name="sFileName"></param>
    ''' <param name="sDelimiter"></param>
    ''' <param name="bIncludeFieldNames"></param>
    ''' <param name="bIncludeLineTerminator"></param>
    ''' <remarks></remarks>
    Public Sub ExportToText(ByVal sFileName As String, ByVal sDelimiter As String, ByVal bIncludeFieldNames As Boolean, _
                             ByVal bIncludeLineTerminator As Boolean)

        Try

            If sFileName.Trim = "" Then
                MsgBox("You need to specify the filename before I can continue!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, _
                       "TAIGRIDControl.ExportToText Error...")
                Exit Sub
            End If

            Dim sw As System.IO.StreamWriter = System.IO.File.CreateText(sFileName)

            Dim iCols As Integer = Me.Cols
            Dim iCol As Integer
            Dim iRows As Integer = Me.Rows
            Dim iRow As Integer
            Dim iRowStart As Integer = 0

            Dim sLine As String

            If bIncludeFieldNames Then

                If bIncludeLineTerminator Then

                    sLine = ""
                    For iCol = 0 To iCols - 1
                        sLine += Me.HeaderLabel(iCol) & sDelimiter
                    Next
                    sw.WriteLine(sLine)

                Else

                    sLine = ""
                    For iCol = 0 To iCols - 1
                        sLine += Me.HeaderLabel(iCol) & sDelimiter
                    Next
                    sw.Write(sLine)

                End If

            End If

            If bIncludeLineTerminator Then
                For iRow = iRowStart To iRows - 1
                    sLine = ""
                    For iCol = 0 To iCols - 1
                        sLine += Me.item(iRow, iCol) & sDelimiter
                    Next
                    sw.WriteLine(sLine)
                Next
            Else
                For iRow = iRowStart To iRows - 1
                    sLine = ""
                    For iCol = 0 To iCols - 1
                        sLine += Me.item(iRow, iCol) & sDelimiter
                    Next
                    sw.Write(sLine)
                Next
            End If

            sw.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ExportToText Error...")
        End Try

    End Sub


#End Region


    ''' <summary>
    ''' Will return a list(of string) of the unique values contained in ColId of the current grid contents
    ''' </summary>
    ''' <param name="ColId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DistinctInColumn(ByVal ColId As Integer) As List(Of String)
        Dim ret As List(Of String) = New List(Of String)

        Dim t As Integer = 0

        For t = 0 To _rows - 1

            If Not ret.Contains(item(t, ColId)) Then
                ret.Add(item(t, ColId))
            End If

        Next

        Return ret
    End Function

    ''' <summary>
    ''' will do a case insensitive rip through grid col colvalue searching for strvalue
    ''' on finding it will return id value of the row where the search was successful
    ''' -1 otherwise
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <param name="colvalue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindInColumn(ByVal strValue As String, ByVal colvalue As Integer) As Integer
        ' will do a case insensative rip through grid col colvalue searching for strvalue
        ' on finding it will return id value of the row where the search was successful
        ' -1 otherwise

        Dim t As Integer
        Dim ret As Integer = -1

        If colvalue <= _cols - 1 And colvalue > -1 Then

            For t = 0 To Rows - 1
                If UCase(Trim(strValue)) = UCase(Trim(_grid(t, colvalue))) Then
                    ret = t
                    Exit For
                End If
            Next

        End If

        Return ret

    End Function

    ''' <summary>
    ''' will do a case insensative or sensitive  (depending on the CaseSensitive parameter)
    ''' rip through grid col colvalue searching for strvalue
    ''' on finding it will return id value of the row where the search was successful
    ''' -1 otherwise
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <param name="colvalue"></param>
    ''' <param name="CaseSensitive"></param>
    ''' <returns>Will return the first row ID of the search or -1 if the search is unsuccessful</returns>
    ''' <remarks></remarks>
    Public Function FindInColumn(ByVal strValue As String, ByVal colvalue As Integer, ByVal CaseSensitive As Boolean) As Integer
        ' will do a eithyer a case insensative  or case sensative rip through grid col colvalue searching for strvalue
        ' on finding it will return id value of the row where the search was successful
        ' -1 otherwise

        Dim t As Integer
        Dim ret As Integer = -1

        If colvalue <= _cols - 1 And colvalue > -1 Then
            If CaseSensitive Then
                For t = 0 To Rows - 1
                    If Trim(strValue) = Trim(_grid(t, colvalue)) Then
                        ret = t
                        Exit For
                    End If
                Next
            Else
                ret = FindInColumn(strValue, colvalue)
            End If

        End If

        Return ret

    End Function

    ''' <summary>
    ''' will free the memory associated with the internal captured image of the grids contents
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub FreeGridContentImage()

        _image.Dispose()

    End Sub

    ''' <summary>
    ''' Will return the column ID of the first column matching name. CaseSensative will toggle the search matching
    ''' the case of the searched name.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="CaseSensitive"></param>
    ''' <returns>Column ID of the first match, -1 if not found</returns>
    ''' <remarks></remarks>
    Public Function GetColumnIDByName(ByVal name As String, ByVal CaseSensitive As Boolean) As Integer
        ' returns column ID of NAME string from current grid contents
        ' -1 if not found

        Dim ret As Integer = -1
        Dim t As Integer
        If CaseSensitive Then
            For t = 0 To _cols - 1
                If Trim(name) = Trim(_GridHeader(t)) Then
                    ' we have a match
                    ret = t
                    Exit For
                End If
            Next
        Else
            ret = GetColumnIDByName(name)
        End If

        Return ret

    End Function

    ''' <summary>
    ''' Will return the column ID of the first column matching name. The search is not case sensative.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns>Column ID of the first match, -1 if not found</returns>
    ''' <remarks></remarks>
    Public Function GetColumnIDByName(ByVal name As String) As Integer
        ' returns column ID of NAME string from current grid contents
        ' -1 if not found

        Dim ret As Integer = -1
        Dim t As Integer
        For t = 0 To _cols - 1
            If Trim(UCase(name)) = Trim(UCase(_GridHeader(t))) Then
                ' we have a match
                ret = t
                Exit For
            End If
        Next

        Return ret

    End Function

    ''' <summary>
    ''' Will return an ArrayList of distinct values contained in a given Column within the grids
    ''' current contents. Column searched indicated by colid. Search will ignore case differences.
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <returns>ArrayList of distinct values sorted</returns>
    ''' <remarks></remarks>
    Public Function GetDistinctColumnEntries(ByVal colid As Integer) As ArrayList

        Dim exl As New ArrayList

        Return GetDistinctColumnEntries(colid, exl, True)

    End Function

    ''' <summary>
    ''' Will return an ArrayList of distinct values contained in a given Column within the grids
    ''' current contents. Column searched indicated by colid. The ignorecase parameter will
    ''' allow or disallow the differences in case to be taken into account.
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <param name="ignorecase"></param>
    ''' <returns>ArrayList of distinct values sorted</returns>
    ''' <remarks></remarks>
    Public Function GetDistinctColumnEntries(ByVal colid As Integer, _
                                             ByVal ignorecase As Boolean) As ArrayList

        Dim exl As New ArrayList

        Return GetDistinctColumnEntries(colid, exl, ignorecase)

    End Function

    ''' <summary>
    '''  Will return an ArrayList of distinct values contained in a given Column within the grids
    ''' current contents. Column searched indicated by colid. The parameter for the Exclusionlist will
    ''' contain any values you wish to omit from the results. The ignorecase parameter will
    ''' allow or disallow the differences in case to be taken into account.
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <param name="exclusionlist"></param>
    ''' <param name="ignorecase"></param>
    ''' <returns>ArrayList of distinct values sorted</returns>
    ''' <remarks></remarks>
    Public Function GetDistinctColumnEntries(ByVal colid As Integer, _
                                             ByVal exclusionlist As ArrayList, _
                                             ByVal ignorecase As Boolean) As ArrayList

        Dim retval As New ArrayList
        Dim s As String

        Dim r As Integer

        If exclusionlist Is Nothing Then
            exclusionlist.Add("KJKJKJJHKJHKJHKJHKJH") ' some dummy value here
        End If

        For r = 0 To _rows - 1
            s = _grid(r, colid)
            If ignorecase Then
                If Not exclusionlist.Contains(UCase(s)) Then
                    If retval Is Nothing Then
                        retval.Add(UCase(s))
                    Else
                        If Not retval.Contains(UCase(s)) Then
                            retval.Add(UCase(s))
                        End If
                    End If
                End If
            Else
                If Not exclusionlist.Contains(s) Then
                    If retval Is Nothing Then
                        retval.Add(s)
                    Else
                        If Not retval.Contains(s) Then
                            retval.Add(s)
                        End If
                    End If
                End If
            End If
        Next

        retval.Sort()

        Return retval

    End Function

    ''' <summary>
    ''' Returns the grids contents as a two dimensional array of string values
    ''' </summary>
    ''' <returns>2 dimensional array of strings representative of the ccurrent contents of the grid</returns>
    ''' <remarks></remarks>
    Public Function GetGridAsArray() As String(,)

        Dim result(_rows, _cols - 1) As String
        Dim r, c As Integer

        For c = 0 To _cols - 1
            result(0, c) = _GridHeader(c)
        Next

        For r = 1 To _rows
            For c = 0 To _cols - 1
                result(r, c) = _grid(r - 1, c)
            Next
        Next

        Return result

    End Function

    ''' <summary>
    ''' Scans ColID for Values in ColVals and colors each row where ColID contains ColVal with corresponding ColorVal 
    ''' </summary>
    ''' <param name="Colid"></param>
    ''' <param name="Colvals"></param>
    ''' <param name="ColorVals"></param>
    ''' <remarks></remarks>
    Public Sub SetRowBackgroundsBasedOnValue(ByVal Colid As Integer, ByVal Colvals As List(Of String), ByVal ColorVals As List(Of Color))
        Dim t As Integer

        For t = 0 To _rows - 1
            Dim v As String = item(t, Colid)

            Dim i As Integer = 0
            For i = 0 To Colvals.Count - 1
                If Colvals(i) = v Then
                    SetRowBackColor(t, ColorVals(i))
                    i = Colvals.Count
                End If
            Next

        Next
    End Sub

    ''' <summary>
    ''' Returns the contents of the grid as a 24 bit bitmat System.Drawing.Bitmap
    ''' Useful for when the contents of the grid need to be image mapped onto another
    ''' surface like a 3D rotating cube (believe it or not) or perhaps a printer page context.
    ''' </summary>
    ''' <returns>24 Bit System.Drawing.Bitmap</returns>
    ''' <remarks></remarks>
    Public Function GetGridContentsAsImage() As Image

        ' here we will get a picture of the attached canvas 
        Dim h As Integer
        Dim w As Integer

        h = Me.AllRowHeights()
        w = Me.AllColWidths()

        'If _GridTitleVisible Then
        '    h = h + _GridTitleHeight
        'End If

        If _GridHeaderVisible Then
            h = h + _GridHeaderHeight
        End If

        If Not _image Is Nothing Then
            _image.Dispose()
            _image = Nothing    ' clear and release the last image gathered 
        End If

        _image = New Bitmap(w, h, Imaging.PixelFormat.Format24bppRgb)

        Dim g1 As Graphics = Graphics.FromImage(_image)

        OleRenderGrid(g1)

        Return _image

    End Function

    ''' <summary>
    ''' Will return a row in the grid as a string with '|' character delimitring the columns.
    ''' </summary>
    ''' <param name="rowid"></param>
    ''' <returns>String representation of the row in the grid at rowid with '|' characters between the fields</returns>
    ''' <remarks></remarks>
    Public Function GetRowAsString(ByVal rowid As Integer) As String

        Dim c As Integer = 0
        Dim result As String = ""

        If rowid <= _rows - 1 Then
            For c = 0 To _cols - 1
                result = result & _grid(rowid, c)
                If c < _cols - 1 Then
                    result = result & "|"
                End If
            Next
        End If

        Return result

    End Function

    ''' <summary>
    ''' Will return the indicated column at col as an Arraylist of values
    ''' </summary>
    ''' <param name="col"></param>
    ''' <returns>ArrayList indicating the contents of the column at col</returns>
    ''' <remarks></remarks>
    Public Function GetColAsArrayList(ByVal col As Integer) As ArrayList
        ' if col is illegal then will return an empty arraylist
        Dim ar As New ArrayList
        Dim r As Integer

        If col >= 0 And col < _cols Then
            For r = 0 To _rows - 1
                ar.Add(_grid(r, col))
            Next
        End If
        Return ar
    End Function

    ''' <summary>
    ''' Will return the indicated column at col as an Arraylist of values. The values will be cleaned
    ''' of specific formatting prior to being returned. Dollar values will have the $() and ,'s removed
    ''' numeric values will be converted represent the numeric string representation. This is useful to
    ''' subsequent insertion into an excel spreadsheet for example.
    ''' </summary>
    ''' <param name="col"></param>
    ''' <returns>ArrayList indicating the contents of the column at col</returns>
    ''' <remarks></remarks>
    Public Function GetColAsCleanedArrayList(ByVal col As Integer) As ArrayList
        ' if col is illegal then will return an empty arraylist
        Dim ar As New ArrayList
        Dim r As Integer
        Dim s As String

        If col >= 0 And col < _cols Then
            For r = 0 To _rows - 1
                If Not _grid(r, col) Is Nothing Then
                    If _grid(r, col).Trim() <> "" Then
                        s = _grid(r, col)
                        s = CleanMoneyString(s)
                        If IsNumeric(s) Then
                            Try
                                ar.Add(Convert.ToDouble(s))
                            Catch
                            End Try
                        End If
                    End If
                End If
            Next
        End If
        Return ar
    End Function

    ''' <summary>
    ''' Will return the indicated row as an arraylist of values
    ''' </summary>
    ''' <param name="row"></param>
    ''' <returns>Arraylist of values at row</returns>
    ''' <remarks></remarks>
    Public Function GetRowAsArrayList(ByVal row As Integer) As ArrayList
        ' if row is illegal then will return an empty arraylist
        Dim ar As New ArrayList
        Dim c As Integer

        If row >= 0 And row < _rows Then
            For c = 0 To _cols - 1
                ar.Add(_grid(row, c))
            Next
        End If
        Return ar
    End Function

    ''' <summary>
    ''' Will insert numrows of blank space into the grid at row atrow. 
    ''' </summary>
    ''' <param name="atrow"></param>
    ''' <param name="numrows"></param>
    ''' <remarks></remarks>
    Public Sub InsertRowsIntoGridAt(ByVal atrow As Integer, ByVal numrows As Integer)
        _Painting = True

        ' do the math here

        Dim oldrowhidden(_rowhidden.GetUpperBound(0)) As Boolean
        Dim oldcolhidden(_colhidden.GetUpperBound(0)) As Boolean
        Dim oldcoleditable(_colEditable.GetUpperBound(0)) As Boolean
        Dim oldroweditable(_rowEditable.GetUpperBound(0)) As Boolean
        Dim oldcolwidths(_colwidths.GetUpperBound(0)) As Integer
        Dim oldrowheights(_rowheights.GetUpperBound(0)) As Integer
        Dim oldgridheader(_GridHeader.GetUpperBound(0)) As String
        Dim oldgrid(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As String
        Dim oldgridbcolor(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridfcolor(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridfonts(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim oldgridcolpasswords(_colPasswords.GetUpperBound(0)) As String
        Dim oldgridcellalignment(_grid.GetUpperBound(0), _grid.GetUpperBound(1)) As Integer
        Dim r, c As Integer
        Dim x, y As Integer

        x = oldgrid.GetUpperBound(0)
        y = oldgrid.GetUpperBound(1)

        For r = 0 To x
            For c = 0 To y
                oldgrid(r, c) = _grid(r, c)
                oldgridbcolor(r, c) = _gridBackColor(r, c)
                oldgridfcolor(r, c) = _gridForeColor(r, c)
                oldgridfonts(r, c) = _gridCellFonts(r, c)
                oldgridcellalignment(r, c) = _gridCellAlignment(r, c)
            Next
        Next

        For c = 0 To Math.Min(_GridHeader.GetUpperBound(0), _colwidths.GetUpperBound(0))
            oldgridheader(c) = _GridHeader(c)
            oldcolwidths(c) = _colwidths(c)
            oldgridcolpasswords(c) = _colPasswords(c)
            oldcolhidden(c) = _colhidden(c)
            oldcoleditable(c) = _colEditable(c)
        Next

        For r = 0 To _rowheights.GetUpperBound(0)
            oldrowheights(r) = _rowheights(r)
        Next

        For r = 0 To _rowhidden.GetUpperBound(0)
            oldrowhidden(r) = _rowhidden(r)
        Next

        For r = 0 To _rowEditable.GetUpperBound(0)
            oldroweditable(r) = _rowEditable(r)
        Next

        ' we have the state

        _rows += numrows

        ReDim _rowhidden(_rows)
        ReDim _colhidden(_cols)
        ReDim _colEditable(_cols)
        ReDim _rowEditable(_rows)
        ReDim _rowheights(_rows)
        ReDim _colwidths(_cols)
        ReDim _GridHeader(_cols)
        ReDim _grid(_rows, _cols)
        ReDim _gridBackColor(_rows, _cols)
        ReDim _gridForeColor(_rows, _cols)
        ReDim _gridCellFonts(_rows, _cols)
        ReDim _gridCellAlignment(_rows, _cols)
        ReDim _colPasswords(_cols)

        ' columns aren't changing so we can just do the column only stuff here
        For c = 0 To y
            _colPasswords(c) = oldgridcolpasswords(c)
            _GridHeader(c) = oldgridheader(c)
            _colwidths(c) = oldcolwidths(c)
            _colhidden(c) = oldcolhidden(c)
            _colEditable(c) = oldcoleditable(c)
        Next

        If atrow = 0 Then
            ' we are just moving rows with an offset
            For r = 0 To x
                For c = 0 To y
                    _grid(r + numrows, c) = oldgrid(r, c)
                    _gridBackColor(r + numrows, c) = oldgridbcolor(r, c)
                    _gridForeColor(r + numrows, c) = oldgridfcolor(r, c)
                    _gridCellFonts(r + numrows, c) = oldgridfonts(r, c)
                    _gridCellAlignment(r + numrows, c) = oldgridcellalignment(r, c)
                Next
                _rowheights(r + numrows) = oldrowheights(r)
                _rowhidden(r + numrows) = oldrowhidden(r)
            Next

            For r = 0 To numrows - 1
                For c = 0 To y
                    _grid(r, c) = ""
                    _gridBackColor(r, c) = GetGridBackColorListEntry(New SolidBrush(_DefaultBackColor))
                    _gridForeColor(r, c) = GetGridForeColorListEntry(New Pen(_DefaultForeColor))
                    _gridCellFonts(r, c) = GetGridCellFontListEntry(_DefaultCellFont)
                    _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(_DefaultStringFormat)
                Next
                _rowheights(r) = _DefaultRowHeight
                _rowEditable(r) = True
                _rowhidden(r) = False
            Next
        Else
            For r = 0 To atrow - 1
                For c = 0 To y
                    _grid(r, c) = oldgrid(r, c)
                    _gridBackColor(r, c) = oldgridbcolor(r, c)
                    _gridForeColor(r, c) = oldgridfcolor(r, c)
                    _gridCellFonts(r, c) = oldgridfonts(r, c)
                    _gridCellAlignment(r, c) = oldgridcellalignment(r, c)
                Next
                _rowheights(r) = oldrowheights(r)
                _rowEditable(r) = oldroweditable(r)
                _rowhidden(r) = oldrowhidden(r)
            Next

            For r = atrow To x
                For c = 0 To y
                    _grid(r + numrows, c) = oldgrid(r, c)
                    _gridBackColor(r + numrows, c) = oldgridbcolor(r, c)
                    _gridForeColor(r + numrows, c) = oldgridfcolor(r, c)
                    _gridCellFonts(r + numrows, c) = oldgridfonts(r, c)
                    _gridCellAlignment(r + numrows, c) = oldgridcellalignment(r, c)
                Next
                _rowheights(r + numrows) = oldrowheights(r)
                _rowEditable(r + numrows) = True
                _rowhidden(r + numrows) = oldrowhidden(r)
            Next

            For r = 0 To numrows - 1
                For c = 0 To y
                    _grid(r + atrow, c) = ""
                    _gridBackColor(r + atrow, c) = GetGridBackColorListEntry(New SolidBrush(_DefaultBackColor))
                    _gridForeColor(r + atrow, c) = GetGridForeColorListEntry(New Pen(_DefaultForeColor))
                    _gridCellFonts(r + atrow, c) = GetGridCellFontListEntry(_DefaultCellFont)
                    _gridCellAlignment(r + atrow, c) = GetGridCellAlignmentListEntry(_DefaultStringFormat)
                Next
                _rowheights(r + atrow) = _DefaultRowHeight
                _rowEditable(r + atrow) = True
                _rowhidden(r + atrow) = False
            Next

        End If

        For c = 0 To _cols - 1
            If _colwidths(c) = 0 And Not _colhidden(c) Then
                _colwidths(c) = _DefaultColWidth
            End If
        Next

        For r = 0 To _rows - 1
            If _rowheights(r) = 0 And Not _rowhidden(r) Then
                _rowheights(r) = _DefaultRowHeight
            End If
        Next

        _Painting = False
        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Will close and release all open tearaway windows currentl being maintained by the grid
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub KillAllTearAwayColumnWindows()

        If TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)
            ta.Winform.KillMe(ta.ColID)
            'TearAways.RemoveAt(t)
        Next

    End Sub

    ''' <summary>
    ''' Will kill any tearaway windows being maintained by the grid for the indicated column it colid
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <remarks></remarks>
    Public Sub KillTearAwayColumnWindow(ByVal colid As Integer)

        If colid = -1 Or TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)
            If ta.ColID = colid Then
                'ta.KillTearAway()
                TearAways.RemoveAt(t)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Will fire the GridHoverleave event if the grid is not maintaining any tearaway windows at the moment
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LowerGridHoverEvents()

        If Not _TearAwayWork Then

            RaiseEvent GridHoverleave(Me)

        End If

    End Sub

#Region " Populate Database Calls "

#Region " Populate from WQL Statements "

    ' New stuff as of May 31 2005

    ''' <summary>
    ''' Will populate the grid with the results of a call to the System.Management classes with a properly
    ''' formatted wql query (Windows Query Language) call.
    ''' 
    ''' Example:
    ''' <c>Select * from Win32_Printer</c>
    ''' 
    ''' </summary>
    ''' <param name="wql"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromWQL(ByVal wql As String)

        RaiseEvent StartedDatabasePopulateOperation(Me)

        Try

            Dim moReturn As Management.ManagementObjectCollection

            Dim moSearch As Management.ManagementObjectSearcher

            Dim mo As Management.ManagementObject

            Dim prop As Management.PropertyData

            Dim HeaderDone As Boolean = False

            'moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Printer")

            moSearch = New Management.ManagementObjectSearcher(wql)

            moReturn = moSearch.Get

            If _ShowProgressBar Then
                pBar.Maximum = moReturn.Count
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If

            Dim x As Integer = 0

            For Each mo In moReturn

                Dim y As Integer = 0

                If Not HeaderDone Then

                    InitializeTheGrid(moReturn.Count, mo.Properties.Count)

                    For Each prop In mo.Properties

                        _GridHeader(y) = prop.Name
                        y += 1

                    Next

                    HeaderDone = True

                End If

                y = 0

                For Each prop In mo.Properties

                    _grid(x, y) = Convert.ToString(prop.Value)
                    y += 1

                Next

                If _ShowProgressBar Then
                    pBar.Increment(1)
                    pBar.Refresh()
                End If

                x += 1

            Next

            AllCellsUseThisFont(_DefaultCellFont)
            AllCellsUseThisForeColor(_DefaultForeColor)

            Me.AutoSizeCellsToContents = True

            Me.Refresh()

            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            NormalizeTearaways()

        Catch ex As Exception

            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try


    End Sub


#End Region

#Region " Populate from Array Calls "

    ' All the from Array Calls

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, employing the supplied parameters 
    ''' to control 
    ''' <list type="Bullet">
    ''' <item><c>gridfont</c> will employ this font for the cell contents</item>
    ''' <item><c>col</c> will be used as the color for the displays cell items</item>
    ''' <item><c>FirstRowHeader</c> if true will treat the first row in the array as the names for each column header</item>
    ''' <item><c>AutoHeader</c> if true will automatically name each column COLUMN - {ordinal} as it populates the grid</item>
    ''' <item><c>hdr</c> an array of strings that will be used as the column labels if the other column options are False</item>
    ''' </list>
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="FirstRowHeader"></param>
    ''' <param name="AutoHeader"></param>
    ''' <param name="hdr"></param>
    ''' <remarks></remarks>
    Private Sub PopulateGridFromArray(ByVal arr(,) As String, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean, ByVal AutoHeader As Boolean, ByVal hdr() As String)

        Dim x, y As Integer
        Dim r, c As Integer

        r = arr.GetUpperBound(0) + 1
        c = arr.GetUpperBound(1) + 1

        If FirstRowHeader Then
            InitializeTheGrid(r - 1, c)
            For y = 0 To c - 1
                _GridHeader(y) = arr(0, y)
            Next
            For x = 1 To r - 1
                For y = 0 To c - 1
                    _grid(x - 1, y) = arr(x, y)
                Next
            Next
        Else
            InitializeTheGrid(r, c)

            If AutoHeader Then
                For y = 0 To c - 1
                    _GridHeader(y) = "Column - " & y.ToString
                Next
            Else
                _GridHeader = hdr
            End If

            For x = 0 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y)
                Next
            Next
        End If

        AllCellsUseThisFont(gridfont)
        AllCellsUseThisForeColor(col)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()


    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, employing the supplied parameters 
    ''' to control 
    ''' <list type="Bullet">
    ''' <item><c>gridfont</c> will employ this font for the cell contents</item>
    ''' <item><c>col</c> will be used as the color for the displays cell items</item>
    ''' <item><c>FirstRowHeader</c> if true will treat the first row in the array as the names for each column header
    ''' if its false the columns will be automatically named</item>
    ''' </list>
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="FirstRowHeader"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As String, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean)

        PopulateGridFromArray(arr, gridfont, col, FirstRowHeader, True, _GridHeader)

    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
    ''' to named each column
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As String)
        PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
    ''' to name each column each cell will be displayed the the supplied <c>cellfont</c>
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="CellFont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As String, ByVal CellFont As Font)
        PopulateGridFromArray(arr, CellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
    ''' to name each column each cell will be displayed the the supplied <c>Forecolor</c>
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Forecolor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As String, ByVal Forecolor As Color)
        PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
    ''' to name each column each cell will be displayed the the supplied <c>cellfont</c> and <c>Forecolor</c>
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Cellfont"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As String, ByVal Cellfont As Font, ByVal ForeColor As Color)
        PopulateGridFromArray(arr, Cellfont, ForeColor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
    ''' to name each column each cell will be displayed the the supplied <c>gridfont</c> and <c>col</c> color and
    ''' if <c>FirstRowHeader</c> is true will use the first row to label each column, if not, then the first row will be auto
    ''' labled with Column - {ordinal}
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="FirstRowHeader"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Integer, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean)

        Dim x, y As Integer
        Dim r, c As Integer

        r = arr.GetUpperBound(0) + 1
        c = arr.GetUpperBound(1) + 1

        If FirstRowHeader Then
            InitializeTheGrid(r - 1, c)
            For y = 0 To c - 1
                _GridHeader(y) = arr(0, y).ToString
            Next
            For x = 1 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        Else
            InitializeTheGrid(r, c)
            For y = 0 To c - 1
                _GridHeader(y) = "Column - " & y.ToString
            Next
            For x = 0 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        End If

        AllCellsUseThisFont(gridfont)
        AllCellsUseThisForeColor(col)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()


    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
    ''' to name each column
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Integer)
        PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Cellfont</c> will be used as the font for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="CellFont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Integer, ByVal CellFont As Font)
        PopulateGridFromArray(arr, CellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Forecolor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Integer, ByVal Forecolor As Color)
        PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
    ''' for each new cells font
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Cellfont"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Integer, ByVal Cellfont As Font, ByVal ForeColor As Color)
        PopulateGridFromArray(arr, Cellfont, ForeColor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>col</c> will be used as the foreground color for each new cell and <c>gridfont</c> will be used if
    ''' <c>FirstRowHeader</c> is true the first row of data in the array will be used to name each column otherwise the columns will be
    ''' named Column - {ordinal}
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="FirstRowHeader"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Long, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean)

        Dim x, y As Integer
        Dim r, c As Integer

        r = arr.GetUpperBound(0) + 1
        c = arr.GetUpperBound(1) + 1

        If FirstRowHeader Then
            InitializeTheGrid(r - 1, c)
            For y = 0 To c - 1
                _GridHeader(y) = arr(0, y).ToString
            Next
            For x = 1 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        Else
            InitializeTheGrid(r, c)
            For y = 0 To c - 1
                _GridHeader(y) = "Column - " & y.ToString
            Next
            For x = 0 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        End If

        AllCellsUseThisFont(gridfont)
        AllCellsUseThisForeColor(col)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()


    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
    ''' to name each column
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Long)
        PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Cellfont</c> will be used as the font for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="CellFont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Long, ByVal CellFont As Font)
        PopulateGridFromArray(arr, CellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Forecolor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Long, ByVal Forecolor As Color)
        PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
    ''' for each new cells font
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Cellfont"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Long, ByVal Cellfont As Font, ByVal ForeColor As Color)
        PopulateGridFromArray(arr, Cellfont, ForeColor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of doubles <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>col</c> will be used as the foreground color for each new cell and <c>gridfont</c> will be used if
    ''' <c>FirstRowHeader</c> is true the first row of data in the array will be used to name each column otherwise the columns will be
    ''' named Column - {ordinal}
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="FirstRowHeader"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Double, ByVal gridfont As Font, ByVal col As Color, ByVal FirstRowHeader As Boolean)

        Dim x, y As Integer
        Dim r, c As Integer

        r = arr.GetUpperBound(0) + 1
        c = arr.GetUpperBound(1) + 1

        If FirstRowHeader Then
            InitializeTheGrid(r - 1, c)
            For y = 0 To c - 1
                _GridHeader(y) = arr(0, y).ToString
            Next
            For x = 1 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        Else
            InitializeTheGrid(r, c)
            For y = 0 To c - 1
                _GridHeader(y) = "Column - " & y.ToString
            Next
            For x = 0 To r - 1
                For y = 0 To c - 1
                    _grid(x, y) = arr(x, y).ToString
                Next
            Next
        End If

        AllCellsUseThisFont(gridfont)
        AllCellsUseThisForeColor(col)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()


    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
    ''' to name each column
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Double)
        PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Cellfont</c> will be used as the font for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="CellFont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Double, ByVal CellFont As Font)
        PopulateGridFromArray(arr, CellFont, _DefaultForeColor, True)
    End Sub

    ''' <summary>
    '''  Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Forecolor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Double, ByVal Forecolor As Color)
        PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, True)
    End Sub

    ''' <summary>
    ''' Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
    ''' to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
    ''' for each new cells font
    ''' </summary>
    ''' <param name="arr"></param>
    ''' <param name="Cellfont"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridFromArray(ByVal arr(,) As Double, ByVal Cellfont As Font, ByVal ForeColor As Color)
        PopulateGridFromArray(arr, Cellfont, ForeColor, True)
    End Sub

#End Region

#Region " PopulateGridWithDataAt Calls "

    ''' <summary>
    ''' Will allow a database populate of a grid within an already populated grid of data.
    ''' The effect will be to insert data from a carefully crafted query into a rectangular region of an
    ''' existing grid of data.
    ''' <list type="Bullet">
    ''' <item><c>ConnectionString</c> the database connection to be employed</item>
    ''' <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
    ''' <item><c>AtRow</c> the integer offset row to start populating the data at</item>
    ''' <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
    ''' <item><c>newheadercolor</c> the color to use for the header that will be created from the queried data</item>
    ''' <item><c>ColOffset</c> the column offset from the edge to start populating</item>
    ''' </list>
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Atrow"></param>
    ''' <param name="newbackcolor"></param>
    ''' <param name="newheadercolor"></param>
    ''' <param name="ColOffSet"></param>
    ''' <remarks></remarks>
    Public Sub PolulateGridWithDataAt(ByVal ConnectionString As String, _
                                            ByVal Sql As String, _
                                            ByVal Atrow As Integer, _
                                            ByVal newbackcolor As Color, _
                                            ByVal newheadercolor As Color, _
                                            ByVal ColOffSet As Integer)

        Dim cn As New SqlClient.SqlConnection(ConnectionString)
        Dim dbc As SqlClient.SqlCommand
        Dim dbc2 As SqlClient.SqlCommand
        Dim dbr As SqlClient.SqlDataReader
        Dim dbr2 As SqlClient.SqlDataReader

        Dim sql2 As String
        Dim t As Long
        Dim x, y, yy, xx As Integer

        RaiseEvent StartedDatabasePopulateOperation(Me)

        '_LastConnectionString = ConnectionString
        '_LastSQLString = Sql

        Try

            cn.Open()

            sql2 = Sql
            dbc2 = New SqlClient.SqlCommand(sql2, cn)
            dbc2.CommandTimeout = _dataBaseTimeOut
            dbr2 = dbc2.ExecuteReader
            y = 0
            yy = 0

            Do While dbr2.Read
                y = y + 1
                If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                    y = Me.MaxRowsSelected
                    yy = -1
                    Exit Do
                End If
            Loop

            dbr2.Close()
            dbc2.Dispose()
            cn.Close()

            Me.InsertRowsIntoGridAt(Atrow, y + 1)

            cn.Open()
            dbc = New SqlClient.SqlCommand(Sql, cn)
            dbc.CommandTimeout = _dataBaseTimeOut

            dbr = dbc.ExecuteReader

            ' Me.Cols = dbr.FieldCount

            ' InitializeTheGrid(y, dbr.FieldCount)

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If


            'AllCellsUseThisFont(Gridfont)
            'AllCellsUseThisForeColor(col)

            If dbr.FieldCount + ColOffSet < Me.Cols Then
                xx = dbr.FieldCount
            Else
                xx = Me.Cols - ColOffSet
            End If

            For x = 0 To xx - 1
                _grid(Atrow, x + ColOffSet) = dbr.GetName(x)
                Me.CellBackColor(Atrow, x + ColOffSet) = New SolidBrush(newheadercolor)
            Next

            'For x = 0 To Me.Cols - 1
            '    _GridHeader(x) = dbr.GetName(x)
            'Next

            t = Atrow + 1
            Do While dbr.Read
                ' Me.Rows = t + 1
                For x = 0 To xx - 1
                    If IsDBNull(dbr.Item(x)) Then

                        If _omitNulls Then
                            _grid(t, x + ColOffSet) = ""
                        Else
                            _grid(t, x + ColOffSet) = "{NULL}"
                        End If

                    Else
                        ' here we need to do some work on items of certain types
                        If dbr.Item(x).ToString = "System.Byte[]" Then
                            _grid(t, x + ColOffSet) = ReturnByteArrayAsHexString(dbr.Item(x))
                        Else
                            _grid(t, x + ColOffSet) = dbr.Item(x)
                        End If
                    End If
                    Me.CellBackColor(t, x + ColOffSet) = New SolidBrush(newbackcolor)
                Next
                t = t + 1

                If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                    Exit Do
                End If

                If _ShowProgressBar Then
                    pBar.Increment(1)
                    pBar.Refresh()
                End If
            Loop

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            pBar.Visible = False
            gb1.Visible = False

            Me.AutoSizeCellsToContents = True

            dbr.Close()

            dbc.Dispose()

            cn.Close()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()
        Catch ex As Exception
            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' Will allow a database populate of a grid within an already populated grid of data.
    ''' The effect will be to insert data from a carefully crafted query into a rectangular region of an
    ''' existing grid of data.
    ''' <list type="Bullet">
    ''' <item><c>ConnectionString</c> the database connection to be employed</item>
    ''' <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
    ''' <item><c>AtRow</c> the integer offset row to start populating the data at</item>
    ''' <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
    ''' <item><c>newheadercolor</c> the color to use for the header that will be created from the queried data</item>
    ''' </list>
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Atrow"></param>
    ''' <param name="newbackcolor"></param>
    ''' <param name="newheadercolor"></param>
    ''' <remarks></remarks>
    Public Sub PolulateGridWithDataAt(ByVal ConnectionString As String, _
                                        ByVal Sql As String, _
                                        ByVal Atrow As Integer, _
                                        ByVal newbackcolor As Color, _
                                        ByVal newheadercolor As Color)
        Dim cn As New SqlClient.SqlConnection(ConnectionString)
        Dim dbc As SqlClient.SqlCommand
        Dim dbc2 As SqlClient.SqlCommand
        Dim dbr As SqlClient.SqlDataReader
        Dim dbr2 As SqlClient.SqlDataReader

        Dim sql2 As String
        Dim t As Long
        Dim x, y, yy, xx As Integer

        RaiseEvent StartedDatabasePopulateOperation(Me)

        '_LastConnectionString = ConnectionString
        '_LastSQLString = Sql

        Try

            cn.Open()

            sql2 = Sql
            dbc2 = New SqlClient.SqlCommand(sql2, cn)
            dbc2.CommandTimeout = _dataBaseTimeOut
            dbr2 = dbc2.ExecuteReader
            y = 0
            yy = 0

            Do While dbr2.Read
                y = y + 1
                If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                    y = Me.MaxRowsSelected
                    yy = -1
                    Exit Do
                End If
            Loop

            dbr2.Close()
            dbc2.Dispose()
            cn.Close()

            Me.InsertRowsIntoGridAt(Atrow, y + 1)

            cn.Open()
            dbc = New SqlClient.SqlCommand(Sql, cn)
            dbc.CommandTimeout = _dataBaseTimeOut

            dbr = dbc.ExecuteReader

            ' Me.Cols = dbr.FieldCount

            ' InitializeTheGrid(y, dbr.FieldCount)

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If


            'AllCellsUseThisFont(Gridfont)
            'AllCellsUseThisForeColor(col)

            If dbr.FieldCount < Me.Cols Then
                xx = dbr.FieldCount
            Else
                xx = Me.Cols
            End If

            For x = 0 To xx - 1
                _grid(Atrow, x) = dbr.GetName(x)
                Me.CellBackColor(Atrow, x) = New SolidBrush(newheadercolor)
            Next

            'For x = 0 To Me.Cols - 1
            '    _GridHeader(x) = dbr.GetName(x)
            'Next

            t = Atrow + 1
            Do While dbr.Read
                ' Me.Rows = t + 1
                For x = 0 To xx - 1
                    If IsDBNull(dbr.Item(x)) Then

                        If _omitNulls Then
                            _grid(t, x) = ""
                        Else
                            _grid(t, x) = "{NULL}"
                        End If

                    Else
                        ' here we need to do some work on items of certain types
                        If dbr.Item(x).ToString = "System.Byte[]" Then
                            _grid(t, x) = ReturnByteArrayAsHexString(dbr.Item(x))
                        Else
                            _grid(t, x) = dbr.Item(x)
                        End If
                    End If
                    Me.CellBackColor(t, x) = New SolidBrush(newbackcolor)
                Next
                t = t + 1

                If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                    Exit Do
                End If

                If _ShowProgressBar Then
                    pBar.Increment(1)
                    pBar.Refresh()
                End If
            Loop

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            pBar.Visible = False
            gb1.Visible = False

            Me.AutoSizeCellsToContents = True

            dbr.Close()

            dbc.Dispose()

            cn.Close()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()
        Catch ex As Exception
            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' Will allow a database populate of a grid within an already populated grid of data.
    ''' The effect will be to insert data from a carefully crafted query into a rectangular region of an
    ''' existing grid of data. this call will omit the header for the inserted result...
    ''' <list type="Bullet">
    ''' <item><c>ConnectionString</c> the database connection to be employed</item>
    ''' <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
    ''' <item><c>AtRow</c> the integer offset row to start populating the data at</item>
    ''' <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
    ''' </list>
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Atrow"></param>
    ''' <param name="newbackcolor"></param>
    ''' <remarks></remarks>
    Public Sub PolulateGridWithDataAt(ByVal ConnectionString As String, _
                                        ByVal Sql As String, _
                                        ByVal Atrow As Integer, _
                                        ByVal newbackcolor As Color)


        PolulateGridWithDataAt(ConnectionString, Sql, Atrow, newbackcolor, True)


    End Sub

    ''' <summary>
    ''' Will allow a database populate of a grid within an already populated grid of data.
    ''' The effect will be to insert data from a carefully crafted query into a rectangular region of an
    ''' existing grid of data. this call will omit the header for the inserted result...
    ''' <list type="Bullet">
    ''' <item><c>ConnectionString</c> the database connection to be employed</item>
    ''' <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
    ''' <item><c>AtRow</c> the integer offset row to start populating the data at</item>
    ''' <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
    ''' <item><c>allowDups></c> Will not insert any rows that already exist in the grid if set to false</item>
    ''' </list>
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Atrow"></param>
    ''' <param name="newbackcolor"></param>
    ''' <param name="allowDups"></param>
    ''' <remarks></remarks>
    Public Sub PolulateGridWithDataAt(ByVal ConnectionString As String, _
                                        ByVal Sql As String, _
                                        ByVal Atrow As Integer, _
                                        ByVal newbackcolor As Color, _
                                        ByVal allowDups As Boolean)
        Dim cn As New SqlClient.SqlConnection(ConnectionString)
        Dim dbc As SqlClient.SqlCommand
        Dim dbc2 As SqlClient.SqlCommand
        Dim dbr As SqlClient.SqlDataReader
        Dim dbr2 As SqlClient.SqlDataReader

        Dim sql2 As String
        Dim t, tt As Long
        Dim x, y, yy, xx As Integer
        Dim fnd As Boolean = False
        Dim hst As String = ""
        Dim hst2 As String = ""

        RaiseEvent StartedDatabasePopulateOperation(Me)

        '_LastConnectionString = ConnectionString
        '_LastSQLString = Sql

        Try

            '' here we want to get whats in the grid already as a set of hashes for dup checking

            Dim ga As List(Of String) = New List(Of String)()

            Dim sb As New StringBuilder()

            For t = 0 To Me.Rows - 1
                sb = New StringBuilder()
                For x = 0 To Me.Cols - 1
                    sb.Append(x.ToString() + _grid(t, x).ToUpper() + "|")
                Next
                ga.Add(sb.ToString())
            Next

            cn.Open()

            sql2 = Sql
            dbc2 = New SqlClient.SqlCommand(sql2, cn)
            dbc2.CommandTimeout = _dataBaseTimeOut
            dbr2 = dbc2.ExecuteReader
            y = 0
            yy = 0

            If dbr2.FieldCount < Me.Cols Then
                xx = dbr2.FieldCount
            Else
                xx = Me.Cols
            End If

            Dim _ggrid(xx) As String

            Do While dbr2.Read

                hst = ""

                If Not allowDups Then
                    For x = 0 To xx - 1
                        If IsDBNull(dbr2.Item(x)) Then

                            If _omitNulls Then
                                _ggrid(x) = ""
                            Else
                                _ggrid(x) = "{NULL}"
                            End If

                        Else
                            ' here we need to do some work on items of certain types
                            If dbr2.Item(x).ToString = "System.Byte[]" Then
                                _ggrid(x) = ReturnByteArrayAsHexString(dbr2.Item(x))
                            Else
                                _ggrid(x) = dbr2.Item(x)
                            End If
                        End If

                    Next


                    For x = 0 To xx - 1
                        hst += x.ToString() + _ggrid(x).ToUpper() + "|"
                    Next

                    If ga.Contains(hst) Then
                        fnd = True
                    Else
                        fnd = False
                    End If


                    If Not fnd Then
                        y = y + 1
                        If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                            y = Me.MaxRowsSelected
                            yy = -1
                            Exit Do
                        End If
                    Else
                        Console.WriteLine("Dupe")
                    End If

                Else
                    y = y + 1
                    If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                        y = Me.MaxRowsSelected
                        yy = -1
                        Exit Do
                    End If
                End If

            Loop

            dbr2.Close()
            dbc2.Dispose()
            cn.Close()

            Me.InsertRowsIntoGridAt(Atrow, y)

            cn.Open()
            dbc = New SqlClient.SqlCommand(Sql, cn)
            dbc.CommandTimeout = _dataBaseTimeOut

            dbr = dbc.ExecuteReader

            ' Me.Cols = dbr.FieldCount

            ' InitializeTheGrid(y, dbr.FieldCount)

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If


            'AllCellsUseThisFont(Gridfont)
            'AllCellsUseThisForeColor(col)

            If dbr.FieldCount < Me.Cols Then
                xx = dbr.FieldCount
            Else
                xx = Me.Cols
            End If

            t = Atrow
            If (allowDups) Then
                Do While dbr.Read
                    ' Me.Rows = t + 1
                    For x = 0 To xx - 1
                        If IsDBNull(dbr.Item(x)) Then

                            If _omitNulls Then
                                _grid(t, x) = ""
                            Else
                                _grid(t, x) = "{NULL}"
                            End If

                        Else
                            ' here we need to do some work on items of certain types
                            If dbr.Item(x).ToString = "System.Byte[]" Then
                                _grid(t, x) = ReturnByteArrayAsHexString(dbr.Item(x))
                            Else
                                _grid(t, x) = dbr.Item(x)
                            End If
                        End If
                        Me.CellBackColor(t, x) = New SolidBrush(newbackcolor)
                    Next
                    t = t + 1

                    If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                        Exit Do
                    End If

                    If _ShowProgressBar Then
                        pBar.Increment(1)
                        pBar.Refresh()
                    End If
                Loop
            Else
                '' here we are gonna not import any duplicate rows

                tt = 0

                ''Dim _ggrid(xx) As String

                Do While dbr.Read
                    ' Me.Rows = t + 1
                    For x = 0 To xx - 1
                        If IsDBNull(dbr.Item(x)) Then

                            If _omitNulls Then
                                _ggrid(x) = ""
                            Else
                                _ggrid(x) = "{NULL}"
                            End If

                        Else
                            ' here we need to do some work on items of certain types
                            If dbr.Item(x).ToString = "System.Byte[]" Then
                                _ggrid(x) = ReturnByteArrayAsHexString(dbr.Item(x))
                            Else
                                _ggrid(x) = dbr.Item(x)
                            End If
                        End If
                        ''Me.CellBackColor(t, x) = New SolidBrush(newbackcolor)
                    Next

                    '' here we want to scan the current contents of the grid to see if these values are already in the thing

                    '' first we will build a giant hash string of what we are looking for

                    hst = ""
                    hst2 = ""

                    For x = 0 To xx - 1
                        hst += x.ToString() + _ggrid(x).ToUpper() + "|"
                    Next

                    If ga.Contains(hst) Then fnd = True Else fnd = False

                    If Not fnd Then
                        For x = 0 To xx - 1
                            _grid(t, x) = _ggrid(x)
                            Me.CellBackColor(t, x) = New SolidBrush(newbackcolor)
                        Next

                        t += 1

                    End If

                    If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                        Exit Do
                    End If

                    If _ShowProgressBar Then
                        pBar.Increment(1)
                        pBar.Refresh()
                    End If
                Loop

            End If

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            pBar.Visible = False
            gb1.Visible = False

            Me.AutoSizeCellsToContents = True

            dbr.Close()

            dbc.Dispose()

            cn.Close()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()
        Catch ex As Exception
            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            ''MsgBox(ex.Message)

        End Try

    End Sub


#End Region

#Region " Populate from SQL Server Calls "

    ' SQL Populate Data Calls

    ''' <summary>
    ''' Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
    ''' <c>col</c> for the foreground color and <c>gridfont</c> for the cells font style
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="col"></param>
    ''' <param name="gridfont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal col As Color, ByVal gridfont As Font)

        Dim x, numrows As Integer

        Try

            RaiseEvent StartedDatabasePopulateOperation(Me)

            numrows = 0

            InitializeTheGrid(0, SQLDR.FieldCount)

            For x = 0 To Me.Cols - 1
                _GridHeader(x) = SQLDR.GetName(x)
            Next

            Do While SQLDR.Read

                numrows += 1

                If Me.MaxRowsSelected > 0 And numrows > Me.MaxRowsSelected Then
                    RaiseEvent PartialSelection(Me)
                    Exit Do
                End If


                Me.Rows = numrows
                'For x = 0 To _Cols - 1
                For x = 0 To Me.Cols - 1
                    If IsDBNull(SQLDR.Item(x)) Then

                        If _omitNulls Then
                            _grid(numrows - 1, x) = ""
                        Else
                            _grid(numrows - 1, x) = "{NULL}"
                        End If

                    Else
                        ' here we need to do some work on items of certain types
                        If SQLDR.Item(x).ToString = "System.Byte[]" Then
                            _grid(numrows - 1, x) = ReturnByteArrayAsHexString(SQLDR.Item(x))
                        Else
                            'Console.WriteLine(SQLDR.Item(x).GetType.ToString())
                            If SQLDR.Item(x).GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then
                                If _ShowDatesWithTime Then
                                    Dim _dt As Date = DateTime.Parse(SQLDR.Item(x))

                                    _grid(numrows - 1, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                Else
                                    Dim _dt As Date = DateTime.Parse(SQLDR.Item(x))

                                    _grid(numrows - 1, x) = _dt.ToShortDateString()

                                End If
                            Else
                                If SQLDR.Item(x).GetType.ToString().ToUpper() = "SYSTEM.GUID" Then
                                    _grid(numrows - 1, x) = "This is a GUID"
                                Else
                                    _grid(numrows - 1, x) = SQLDR.Item(x)
                                End If
                            End If
                            ' _grid(numrows - 1, x) = SQLDR.Item(x)
                        End If
                    End If

                Next

            Loop

            Me.AllCellsUseThisForeColor(col)

            Me.AllCellsUseThisFont(gridfont)

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Me.Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
    ''' the grids default coloring and fonts for the cells content (settable using the propertries of the grid itself
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader)
        PopulateGridWithData(SQLDR, _DefaultForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
    ''' <c>ForeColor</c> for the foreground color and <c>gridfont</c> for the cells font style
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal ForeColor As Color)
        PopulateGridWithData(SQLDR, ForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
    ''' <c>GridFont</c> for the cells font style
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="GridFont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal GridFont As Font)
        PopulateGridWithData(SQLDR, _DefaultForeColor, GridFont)
    End Sub

    ''' <summary>
    ''' Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid 
    ''' with those results. <c>GridFont</c> and <c>col</c> be used to generate the font and the foreground color for the cell contents
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Gridfont As Font, ByVal col As Color)

        Dim dbc As SqlClient.SqlCommand
        Dim dbc2 As SqlClient.SqlCommand
        Dim sql2 As String
        Dim t As Long
        Dim y, yy As Integer

        RaiseEvent StartedDatabasePopulateOperation(Me)

        Try

            Cursor.Current = Cursors.WaitCursor
            y = 0
            yy = 0

            'Gets a count of the number of records returned by the query
            Using cn As New SqlClient.SqlConnection

                cn.ConnectionString = ConnectionString
                cn.Open()
                sql2 = Sql
                dbc2 = New SqlClient.SqlCommand(sql2, cn)
                dbc2.CommandTimeout = _dataBaseTimeOut

                Using dbr2 As SqlClient.SqlDataReader = dbc2.ExecuteReader

                    Do While dbr2.Read
                        y = y + 1
                        If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                            y = Me.MaxRowsSelected
                            yy = -1
                            Exit Do
                        End If
                    Loop

                End Using

            End Using

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                pBar.Style = ProgressBarStyle.Continuous
                gb1.Visible = True
                pBar.Step = 1
                pBar.Refresh()
                gb1.Refresh()
            End If

            Using cn2 As New SqlClient.SqlConnection

                cn2.ConnectionString = ConnectionString
                cn2.Open()
                dbc = New SqlClient.SqlCommand(Sql, cn2)
                dbc.CommandTimeout = _dataBaseTimeOut

                Using dbr As SqlClient.SqlDataReader = dbc.ExecuteReader

                    InitializeTheGrid(y, dbr.FieldCount)
                    AllCellsUseThisFont(Gridfont)
                    AllCellsUseThisForeColor(col)

                    For x As Integer = 0 To Me.Cols - 1
                        _GridHeader(x) = dbr.GetName(x)
                    Next

                    t = 0

                    Do While dbr.Read

                        Dim dbrRow As New List(Of Object)
                        Dim x As Integer = 0

                        For i As Integer = 0 To Me.Cols - 1
                            dbrRow.Add(dbr.Item(i))
                        Next

                        'Process the row item from the data reader
                        For Each o As Object In dbrRow

                            If o.Equals(DBNull.Value) Then

                                If _omitNulls Then
                                    _grid(t, x) = ""
                                Else
                                    _grid(t, x) = "{NULL}"
                                End If

                            ElseIf o.ToString = "System.Byte[]" Then

                                _grid(t, x) = ReturnByteArrayAsHexString(o)

                            ElseIf o.GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then

                                Dim _dt As Date = Convert.ToDateTime(o) 'DateTime.Parse(o)
                                If _ShowDatesWithTime Then
                                    _grid(t, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                Else
                                    _grid(t, x) = _dt.ToShortDateString()
                                End If

                            ElseIf o.GetType.ToString().ToUpper() = "SYSTEM.GUID" Then

                                Dim s As String = o.ToString()
                                _grid(t, x) = s

                            Else
                                _grid(t, x) = o

                            End If
                            'increment column index
                            x += 1
                        Next

                        'increment the row index
                        t = t + 1

                        If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                            Exit Do
                        End If

                        If _ShowProgressBar Then
                            pBar.PerformStep()
                        End If

                    Loop

                End Using

            End Using

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            If _ShowProgressBar Then
                pBar.PerformStep()
                pBar.Visible = False
                gb1.Visible = False
            End If

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        Finally

            'Set the cursor back to the default cursor
            Cursor.Current = Cursors.Default

        End Try

    End Sub

    ''' <summary>
    ''' Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid 
    ''' the grids defauls will be used for the cells fonts and coloring characteristics
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String)
        PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor)
    End Sub

    ''' <summary>
    ''' Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid 
    ''' the <c>col</c> parameter wwill be used for the cell foreground coloring 
    ''' the grids defauls will be used for the cells fonts and other coloring characteristics
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Col As Color)
        PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col)
    End Sub

    ''' <summary>
    ''' Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid 
    ''' the <c>fnt</c> parameter wwill be used for the cell fonts
    ''' the grids defauls will be used for the cells other coloring characteristics
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal fnt As Font)
        PopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="col"></param>
    ''' <param name="gridfont"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal col As Color, ByVal gridfont As Font)
        PopulateGridWithData(SQLDR, col, gridfont)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader)
        PopulateGridWithData(SQLDR, _DefaultForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal ForeColor As Color)
        PopulateGridWithData(SQLDR, ForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="SQLDR"></param>
    ''' <param name="GridFont"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByRef SQLDR As SqlClient.SqlDataReader, ByVal GridFont As Font)
        PopulateGridWithData(SQLDR, _DefaultForeColor, GridFont)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Gridfont As Font, ByVal col As Color)
        PopulateGridWithData(ConnectionString, Sql, Gridfont, col)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String)
        PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Col"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Col As Color)
        PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col)
    End Sub

    ''' <summary>
    ''' A synonym for the PopulateGridWithData method of the same signature
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub SQLPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal fnt As Font)
        PopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor)
    End Sub

#End Region

#Region " Populate from OLEDB Calls "

    ' OLE Populate Data Calls

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
    ''' </summary>
    ''' <param name="OLEDR"></param>
    ''' <param name="col"></param>
    ''' <param name="gridfont"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByRef OLEDR As OleDb.OleDbDataReader, ByVal col As Color, ByVal gridfont As Font)

        Dim x, numrows As Integer

        Try

            RaiseEvent StartedDatabasePopulateOperation(Me)

            numrows = 0

            InitializeTheGrid(1, OLEDR.FieldCount)

            For x = 0 To Me.Cols - 1
                _GridHeader(x) = OLEDR.GetName(x)
            Next

            Do While OLEDR.Read

                numrows += 1
                If Me.MaxRowsSelected > 0 And numrows < Me.MaxRowsSelected Then

                    Me.Rows = numrows

                    For x = 0 To _cols
                        If IsDBNull(OLEDR.Item(x)) Then

                            If _omitNulls Then
                                _grid(numrows - 1, x) = ""
                            Else
                                _grid(numrows - 1, x) = "{NULL}"
                            End If

                        Else
                            ' here we need to do some work on items of certain types
                            If OLEDR.Item(x).ToString = "System.Byte[]" Then
                                _grid(numrows - 1, x) = ReturnByteArrayAsHexString(OLEDR.Item(x))
                            Else
                                If OLEDR.Item(x).GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then
                                    If _ShowDatesWithTime Then
                                        Dim _dt As Date = DateTime.Parse(OLEDR.Item(x))

                                        _grid(numrows - 1, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                    Else
                                        Dim _dt As Date = DateTime.Parse(OLEDR.Item(x))

                                        _grid(numrows - 1, x) = _dt.ToShortDateString()

                                    End If
                                Else
                                    _grid(numrows - 1, x) = OLEDR.Item(x)
                                End If
                            End If
                        End If

                    Next
                Else
                    RaiseEvent PartialSelection(Me)
                    Exit Do
                End If
            Loop

            Me.AllCellsUseThisForeColor(col)

            Me.AllCellsUseThisFont(gridfont)

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
    ''' </summary>
    ''' <param name="OLEDR"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByRef OLEDR As OleDb.OleDbDataReader)
        OLEPopulateGridWithData(OLEDR, _DefaultForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
    ''' </summary>
    ''' <param name="OLEDR"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByRef OLEDR As OleDb.OleDbDataReader, ByVal ForeColor As Color)
        OLEPopulateGridWithData(OLEDR, ForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
    ''' </summary>
    ''' <param name="OLEDR"></param>
    ''' <param name="GridFont"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByRef OLEDR As OleDb.OleDbDataReader, ByVal GridFont As Font)
        OLEPopulateGridWithData(OLEDR, _DefaultForeColor, GridFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for OLE data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Gridfont As Font, ByVal col As Color)

        Dim cn As New OleDb.OleDbConnection(ConnectionString)
        Dim dbc As OleDb.OleDbCommand
        Dim dbc2 As OleDb.OleDbCommand
        Dim dbr As OleDb.OleDbDataReader
        Dim dbr2 As OleDb.OleDbDataReader

        Dim sql2 As String
        Dim t As Long
        Dim x, y, yy As Integer

        RaiseEvent StartedDatabasePopulateOperation(Me)

        '_LastConnectionString = ConnectionString
        '_LastSQLString = Sql


        Try

            cn.Open()

            sql2 = Sql
            dbc2 = New OleDb.OleDbCommand(sql2, cn)
            dbc2.CommandTimeout = _dataBaseTimeOut
            dbr2 = dbc2.ExecuteReader
            y = 0
            yy = 0

            Do While dbr2.Read
                y = y + 1
                If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                    y = Me.MaxRowsSelected
                    yy = -1
                    Exit Do
                End If
            Loop

            dbr2.Close()
            dbc2.Dispose()
            cn.Close()

            cn.Open()
            dbc = New OleDb.OleDbCommand(Sql, cn)
            dbc.CommandTimeout = _dataBaseTimeOut

            dbr = dbc.ExecuteReader

            'Me.Cols = dbr.FieldCount

            InitializeTheGrid(y, dbr.FieldCount)

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If


            AllCellsUseThisFont(Gridfont)
            AllCellsUseThisForeColor(col)

            For x = 0 To Me.Cols - 1
                _GridHeader(x) = dbr.GetName(x)
            Next

            t = 0
            Do While dbr.Read
                ' Me.Rows = t + 1
                For x = 0 To Me.Cols - 1
                    If IsDBNull(dbr.Item(x)) Then

                        If _omitNulls Then
                            _grid(t, x) = ""
                        Else
                            _grid(t, x) = "{NULL}"
                        End If

                    Else
                        ' here we need to do some work on items of certain types
                        If dbr.Item(x).ToString = "System.Byte[]" Then
                            _grid(t, x) = ReturnByteArrayAsHexString(dbr.Item(x))
                        Else
                            If dbr.Item(x).GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then
                                If _ShowDatesWithTime Then
                                    Dim _dt As Date = DateTime.Parse(dbr.Item(x))

                                    _grid(t, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                Else
                                    Dim _dt As Date = DateTime.Parse(dbr.Item(x))

                                    _grid(t, x) = _dt.ToShortDateString()

                                End If
                            Else
                                If dbr.Item(x).GetType.ToString().ToUpper() = "SYSTEM.GUID" Then
                                    _grid(t, x) = dbr.Item(x).ToString()
                                Else
                                    _grid(t, x) = dbr.Item(x)
                                End If
                                ''_grid(t, x) = dbr.Item(x)
                            End If
                        End If
                    End If
                Next
                t = t + 1

                If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                    Exit Do
                End If

                If _ShowProgressBar Then
                    pBar.Increment(1)
                    pBar.Refresh()
                End If
            Loop

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            pBar.Visible = False
            gb1.Visible = False

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            dbr.Close()

            dbc.Dispose()

            cn.Close()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try


    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for OLE data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String)
        OLEPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for OLE data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Col"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Col As Color)
        OLEPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for OLE data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub OLEPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal fnt As Font)
        OLEPopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor)
    End Sub


#End Region

#Region " Populate from ODBC Calls "

    ' ODBC Populate Data Calls

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
    ''' </summary>
    ''' <param name="OdbcDR"></param>
    ''' <param name="col"></param>
    ''' <param name="gridfont"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByRef OdbcDR As Odbc.OdbcDataReader, ByVal col As Color, ByVal gridfont As Font)

        Dim x, numrows As Integer

        Try

            RaiseEvent StartedDatabasePopulateOperation(Me)

            numrows = 0

            InitializeTheGrid(1, OdbcDR.FieldCount)

            For x = 0 To Me.Cols - 1
                _GridHeader(x) = OdbcDR.GetName(x)
            Next

            Do While OdbcDR.Read

                numrows += 1
                If Me.MaxRowsSelected > 0 And numrows < Me.MaxRowsSelected Then
                    Me.Rows = numrows

                    For x = 0 To _cols
                        If IsDBNull(OdbcDR.Item(x)) Then

                            If _omitNulls Then
                                _grid(numrows - 1, x) = ""
                            Else
                                _grid(numrows - 1, x) = "{NULL}"
                            End If

                        Else
                            ' here we need to do some work on items of certain types
                            If OdbcDR.Item(x).ToString = "System.Byte[]" Then
                                _grid(numrows - 1, x) = ReturnByteArrayAsHexString(OdbcDR.Item(x))
                            Else
                                If OdbcDR.Item(x).GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then
                                    If _ShowDatesWithTime Then
                                        Dim _dt As Date = DateTime.Parse(OdbcDR.Item(x))

                                        _grid(numrows - 1, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                    Else
                                        Dim _dt As Date = DateTime.Parse(OdbcDR.Item(x))

                                        _grid(numrows - 1, x) = _dt.ToShortDateString()

                                    End If
                                Else
                                    _grid(numrows - 1, x) = OdbcDR.Item(x)
                                End If
                            End If
                        End If

                    Next
                Else
                    RaiseEvent PartialSelection(Me)
                    Exit Do
                End If

            Loop

            Me.AllCellsUseThisForeColor(col)

            Me.AllCellsUseThisFont(gridfont)

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try

    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
    ''' </summary>
    ''' <param name="OdbcDR"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByRef OdbcDR As Odbc.OdbcDataReader)
        ODBCPopulateGridWithData(OdbcDR, _DefaultForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
    ''' </summary>
    ''' <param name="OdbcDR"></param>
    ''' <param name="ForeColor"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByRef OdbcDR As Odbc.OdbcDataReader, ByVal ForeColor As Color)
        ODBCPopulateGridWithData(OdbcDR, ForeColor, _DefaultCellFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
    ''' </summary>
    ''' <param name="OdbcDR"></param>
    ''' <param name="GridFont"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByRef OdbcDR As Odbc.OdbcDataReader, ByVal GridFont As Font)
        ODBCPopulateGridWithData(OdbcDR, _DefaultForeColor, GridFont)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for ODBC data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Gridfont As Font, ByVal col As Color)

        Dim cn As New Odbc.OdbcConnection(ConnectionString)
        Dim dbc As Odbc.OdbcCommand
        Dim dbc2 As Odbc.OdbcCommand
        Dim dbr As Odbc.OdbcDataReader
        Dim dbr2 As Odbc.OdbcDataReader

        Dim sql2 As String
        Dim t As Long
        Dim x, y, yy As Integer

        RaiseEvent StartedDatabasePopulateOperation(Me)

        '_LastConnectionString = ConnectionString
        '_LastSQLString = Sql


        Try

            cn.Open()

            sql2 = Sql
            dbc2 = New Odbc.OdbcCommand(sql2, cn)
            dbc2.CommandTimeout = _dataBaseTimeOut
            dbr2 = dbc2.ExecuteReader
            y = 0
            yy = 0

            Do While dbr2.Read
                y = y + 1
                If y > Me.MaxRowsSelected And Me.MaxRowsSelected > 0 Then
                    y = Me.MaxRowsSelected
                    yy = -1
                    Exit Do
                End If
            Loop

            dbr2.Close()
            dbc2.Dispose()
            cn.Close()

            cn.Open()
            dbc = New Odbc.OdbcCommand(Sql, cn)
            dbc.CommandTimeout = _dataBaseTimeOut

            dbr = dbc.ExecuteReader

            'Me.Cols = dbr.FieldCount

            InitializeTheGrid(y, dbr.FieldCount)

            If _ShowProgressBar Then
                pBar.Maximum = y
                pBar.Minimum = 0
                pBar.Value = 0
                pBar.Visible = True
                gb1.Visible = True
                pBar.Refresh()
                gb1.Refresh()
            End If

            AllCellsUseThisFont(Gridfont)
            AllCellsUseThisForeColor(col)

            For x = 0 To Me.Cols - 1
                _GridHeader(x) = dbr.GetName(x)
            Next

            t = 0
            Do While dbr.Read
                ' Me.Rows = t + 1
                For x = 0 To Me.Cols - 1
                    If IsDBNull(dbr.Item(x)) Then

                        If _omitNulls Then
                            _grid(t, x) = ""
                        Else
                            _grid(t, x) = "{NULL}"
                        End If

                    Else
                        ' here we need to do some work on items of certain types
                        If dbr.Item(x).ToString = "System.Byte[]" Then
                            _grid(t, x) = ReturnByteArrayAsHexString(dbr.Item(x))
                        Else
                            If dbr.Item(x).GetType.ToString().ToUpper() = "SYSTEM.DATETIME" Then
                                If _ShowDatesWithTime Then
                                    Dim _dt As Date = DateTime.Parse(dbr.Item(x))

                                    _grid(t, x) = _dt.ToShortDateString() + " " + _dt.ToShortTimeString()
                                Else
                                    Dim _dt As Date = DateTime.Parse(dbr.Item(x))

                                    _grid(t, x) = _dt.ToShortDateString()

                                End If
                            Else
                                If dbr.Item(x).GetType.ToString().ToUpper() = "SYSTEM.GUID" Then
                                    _grid(t, x) = dbr.Item(x).ToString()
                                Else
                                    _grid(t, x) = dbr.Item(x)
                                End If
                                ''_grid(t, x) = dbr.Item(x)
                            End If
                        End If
                    End If
                Next
                t = t + 1

                If Me.MaxRowsSelected > 0 And t >= Me.MaxRowsSelected Then
                    Exit Do
                End If

                If _ShowProgressBar Then
                    pBar.Increment(1)
                    pBar.Refresh()
                End If

            Loop

            If yy = -1 Then
                RaiseEvent PartialSelection(Me)
            End If

            pBar.Visible = False
            gb1.Visible = False

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            dbr.Close()

            dbc.Dispose()

            cn.Close()

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            Refresh()

            NormalizeTearaways()

        Catch ex As Exception

            pBar.Visible = False
            gb1.Visible = False

            RaiseEvent FinishedDatabasePopulateOperation(Me)

            MsgBox(ex.Message)

        End Try


    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for ODBC data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String)
        ODBCPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for ODBC data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="Col"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal Col As Color)
        ODBCPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col)
    End Sub

    ''' <summary>
    ''' As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
    ''' syntax for ODBC data access
    ''' </summary>
    ''' <param name="ConnectionString"></param>
    ''' <param name="Sql"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub ODBCPopulateGridWithData(ByVal ConnectionString As String, ByVal Sql As String, ByVal fnt As Font)
        ODBCPopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor)
    End Sub

#End Region

#Region " Populate from our WebService Return Strings "

    ' webservice populate calls

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String)

        PopulateViaWebServiceString(WebServiceResults, _DefaultForeColor, _DefaultCellFont)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal col As Color)

        PopulateViaWebServiceString(WebServiceResults, col, _DefaultCellFont)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal fnt As Font)

        PopulateViaWebServiceString(WebServiceResults, _DefaultForeColor, fnt)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="col"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal col As Color, ByVal fnt As Font)

        Dim argarray() As String = WebServiceResults.Split("|")
        Dim x, y As Integer
        Dim r, c As Integer
        r = Val(argarray(argarray.GetUpperBound(0)))
        c = Val(argarray(argarray.GetUpperBound(0) - 1) + 1) - 1

        'Me.Rows = Val(argarray(argarray.GetUpperBound(0))) + 1
        'Me.Cols = Val(argarray(argarray.GetUpperBound(0) - 1) + 1) - 1

        InitializeTheGrid(r, c)

        For y = 0 To Cols - 1
            _GridHeader(y) = argarray(y)
        Next

        If Me.OmitNulls Then
            For x = 1 To _rows
                For y = 0 To _cols - 1
                    If UCase(argarray((x * _cols) + y)) = "{NULL}" Then
                        argarray((x * _cols) + y) = ""
                    End If
                Next
            Next
        End If

        For x = 1 To _rows
            For y = 0 To _cols - 1
                _grid(x - 1, y) = argarray((x * _cols) + y)
            Next
        Next

        AllCellsUseThisForeColor(col)
        AllCellsUseThisFont(fnt)
        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="delimiter"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal delimiter As String)

        PopulateViaWebServiceString(WebServiceResults, delimiter, _DefaultForeColor, _DefaultCellFont)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="delimiter"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal delimiter As String, ByVal col As Color)

        PopulateViaWebServiceString(WebServiceResults, delimiter, col, _DefaultCellFont)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="delimiter"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal delimiter As String, ByVal fnt As Font)

        PopulateViaWebServiceString(WebServiceResults, delimiter, _DefaultForeColor, fnt)

    End Sub

    ''' <summary>
    ''' The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
    ''' The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
    ''' able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
    ''' of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
    ''' the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
    ''' </summary>
    ''' <param name="WebServiceResults"></param>
    ''' <param name="Delimiter"></param>
    ''' <param name="col"></param>
    ''' <param name="fnt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateViaWebServiceString(ByVal WebServiceResults As String, ByVal Delimiter As String, ByVal col As Color, ByVal fnt As Font)
        'parse the rows and colmuns off the string
        Dim rows As Integer = WebServiceResults.Substring(WebServiceResults.LastIndexOf(Delimiter) + Delimiter.Length)
        WebServiceResults = WebServiceResults.Substring(0, WebServiceResults.LastIndexOf(Delimiter))
        Dim cols As Integer = WebServiceResults.Substring(WebServiceResults.LastIndexOf(Delimiter) + Delimiter.Length)
        WebServiceResults = WebServiceResults.Substring(0, WebServiceResults.LastIndexOf(Delimiter))

        Dim argarray(,) As String = ReturnDelimitedStringAsArray(WebServiceResults, cols, rows, Delimiter)

        InitializeTheGrid(rows - 1, cols)
        Dim x, y As Integer

        For y = 0 To cols - 1
            _GridHeader(y) = argarray(0, y)
        Next

        If Me.OmitNulls Then
            For x = 1 To _rows
                For y = 0 To _cols - 1
                    If UCase(argarray(x, y)) = "{NULL}" Then
                        argarray(x, y) = ""
                    End If
                Next
            Next
        End If


        For x = 1 To _rows - 1
            For y = 0 To _cols - 1
                _grid(x, y) = argarray(x, y)
            Next
        Next

        Me.AllCellsUseThisFont(fnt)
        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()
        Me.AllCellsUseThisForeColor(col)

        Me.Refresh()

        NormalizeTearaways()

    End Sub

#End Region

#Region " Populate from other grids (The Awesome PIVOT POPULATES) "

    ' Pivot Populate Calls

    ''' <summary>
    ''' The pivot populate calls simulate pivot table functionality in excel.
    ''' The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
    ''' The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
    ''' then for each unique value set in each of the two columns the
    ''' </summary>
    ''' <param name="sgrid"></param>
    ''' <param name="xcol"></param>
    ''' <param name="ycol"></param>
    ''' <param name="scol"></param>
    ''' <param name="FormatSpec"></param>
    ''' <remarks></remarks>
    Public Sub PivotPopulate(ByVal sgrid As TAIGridControl, ByVal xcol As Integer, ByVal ycol As Integer, ByVal scol As Integer, ByVal FormatSpec As String)

        PivotPopulate(sgrid, xcol, ycol, scol, FormatSpec, _DefaultForeColor, _DefaultCellFont)

    End Sub

    Public Sub PivotPopulate(ByVal sgrid As TAIGridControl, ByVal xcol As Integer, ByVal ycol As Integer, ByVal scol As Integer)

        PivotPopulate(sgrid, xcol, ycol, scol, "0.0000", _DefaultForeColor, _DefaultCellFont)

    End Sub

    Public Sub PivotPopulate(ByVal sgrid As TAIGridControl, ByVal xcol As Integer, _
                             ByVal ycol As Integer, ByVal scol As Integer, _
                             ByVal formatspec As String, ByVal col As System.Drawing.Color)

        PivotPopulate(sgrid, xcol, ycol, scol, formatspec, col, _DefaultCellFont)

    End Sub

    Public Sub PivotPopulate(ByVal sgrid As TAIGridControl, ByVal xcol As Integer, _
                             ByVal ycol As Integer, ByVal scol As Integer, _
                             ByVal formatspec As String, ByVal col As System.Drawing.Color, _
                             ByVal fnt As System.Drawing.Font)

        Dim x, y, xxx As Integer
        Dim a, b, c As String
        Dim aa As Double
        Dim sx As Integer = sgrid.Cols - 1
        Dim sy As Integer = sgrid.Rows - 1
        Dim uniquerows(sy) As String
        Dim uniquecols(sx) As String
        ' Dim formatspec As String = "0.0000"

        Dim u As New ArrayList
        Dim uu As New ArrayList

        u.Clear()
        uu.Clear()


        ' how many unique vals do we have in the Xcol

        For x = 0 To sy
            a = sgrid.item(x, xcol)
            If Not u.Contains(a) Then
                u.Add(a)
            End If
        Next

        Me.Cols = u.Count + 1

        ' how many unique vals do we have in the Ycol

        For x = 0 To sy
            a = sgrid.item(x, ycol)
            If Not uu.Contains(a) Then
                uu.Add(a)
            End If
        Next

        Me.Rows = uu.Count

        ' here we will populate the header and the y column with the values being rolled up
        For x = 1 To u.Count
            Me.HeaderLabel(x) = u.Item(x - 1)
            'Me.item(0, x) = uniquecols(x)
        Next

        Me.HeaderLabel(0) = sgrid.HeaderLabel(ycol)

        For y = 0 To uu.Count - 1
            Me.item(y, 0) = uu.Item(y)
        Next

        ' here we will actually populate the values

        For x = 0 To u.Count - 1
            b = u.Item(x)
            For y = 0 To uu.Count - 1
                c = uu.Item(y)
                aa = 0
                For xxx = 0 To sy
                    If sgrid.item(xxx, xcol) = b And sgrid.item(xxx, ycol) = c Then
                        aa = aa + Val(sgrid.item(xxx, scol))
                    End If
                Next

                Me.item(y, x + 1) = Format(aa, formatspec)

            Next
        Next

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()
        Me.AllCellsUseThisForeColor(col)
        Me.AllCellsUseThisFont(fnt)
        Me.Refresh()

        NormalizeTearaways()

    End Sub

    Public Sub PivotPopulate(ByVal sgrid As TAIGridControl, ByVal xcol As Integer, _
                             ByVal ycol As Integer, ByVal scol As Integer, _
                             ByVal formatspec As String, _
                             ByVal fnt As System.Drawing.Font)

        PivotPopulate(sgrid, xcol, ycol, scol, formatspec, _DefaultForeColor, fnt)

    End Sub

    Public Sub FrequencyDistribution(ByVal sgrid As TAIGridControl, ByVal ColForFrequency As Integer)

        Dim codes As New ArrayList

        Dim t As Integer
        Dim tt As Integer

        For t = 0 To sgrid.Rows - 1
            Dim cd As String = sgrid.item(t, ColForFrequency)
            If Not codes.Contains(cd) Then
                codes.Add(cd)
            End If
        Next

        Me.Rows = codes.Count
        Me.Cols = 2
        Me.HeaderLabel(0) = sgrid.HeaderLabel(ColForFrequency)
        Me.HeaderLabel(1) = "Frequency"

        For t = 0 To codes.Count - 1
            Me.item(t, 0) = codes(t)

            Dim result As Integer = 0

            For tt = 0 To sgrid.Rows - 1
                If sgrid.item(tt, ColForFrequency) = codes(t) Then
                    result += 1
                End If
            Next

            Me.item(t, 1) = result.ToString()

        Next

        Me.AutoSizeCellsToContents = True
        Me.Refresh()

        NormalizeTearaways()


    End Sub

    Public Sub FrequencyDistribution(ByVal sgrid As TAIGridControl, _
                                     ByVal ColForFrequency As Integer, _
                                     ByVal SortDescending As Boolean)


        FrequencyDistribution(sgrid, ColForFrequency)

        SortGridOnColumnNumeric(1, SortDescending)

        Me.AutoSizeCellsToContents = True
        Me.Refresh()

    End Sub

#End Region

#Region " Populate from DataTables "

    ' Populate from a datatable

    ''' <summary>
    ''' Will take the supplied dataSet and extract the first table from that dataset and populate the grid with the
    ''' contents of that datatable
    ''' </summary>
    ''' <param name="dset"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithADataTable(ByVal dset As System.Data.DataSet)
        PopulateGridWithADataTable(dset.Tables(0))
    End Sub

    ''' <summary>
    ''' Will take thge supplied datatable and populate the grid with the contents oif that datatable
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks></remarks>
    Public Sub PopulateGridWithADataTable(ByVal dt As System.Data.DataTable)

        Dim t As Integer = 0
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim typ As String

        InitializeTheGrid(dt.Rows.Count, dt.Columns.Count)

        For t = 0 To dt.Columns.Count - 1
            Me.HeaderLabel(t) = dt.Columns.Item(t).ColumnName
        Next

        For x = 0 To dt.Rows.Count - 1
            For y = 0 To dt.Columns.Count - 1
                typ = dt.Rows.Item(x).Item(y).GetType.ToString.ToUpper

                If typ = "SYSTEM.STRING" Then
                    _grid(x, y) = Convert.ToString(dt.Rows.Item(x).Item(y))
                ElseIf typ = "SYSTEM.DBNULL" Then
                    If _omitNulls Then
                        _grid(x, y) = ""
                    Else
                        _grid(x, y) = "{NULL}"
                    End If
                ElseIf typ = "SYSTEM.DATETIME" Then
                    If _ShowDatesWithTime Then
                        _grid(x, y) = Convert.ToDateTime(dt.Rows.Item(x).Item(y)).ToString
                    Else
                        _grid(x, y) = Convert.ToDateTime(dt.Rows.Item(x).Item(y)).ToShortDateString
                    End If
                ElseIf typ = "SYSTEM.SINGLE" Then
                    _grid(x, y) = Convert.ToSingle(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.INT32" Then
                    _grid(x, y) = Convert.ToInt32(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.INT16" Then
                    _grid(x, y) = Convert.ToInt16(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.INT64" Then
                    _grid(x, y) = Convert.ToInt64(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.BOOLEAN" Then
                    If Convert.ToBoolean(dt.Rows.Item(x).Item(y)) Then
                        _grid(x, y) = "TRUE"
                    Else
                        _grid(x, y) = "FALSE"
                    End If
                ElseIf typ = "SYSTEM.DECIMAL" Then
                    _grid(x, y) = Convert.ToDecimal(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.DOUBLE" Then
                    _grid(x, y) = Convert.ToDouble(dt.Rows.Item(x).Item(y)).ToString
                ElseIf typ = "SYSTEM.RUNTIMETYPE" Then
                    _grid(x, y) = Convert.ToString(dt.Rows.Item(x).Item(y))
                Else
                    _grid(x, y) = dt.Rows.Item(x).Item(y).GetType.ToString
                End If
            Next
        Next

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

        NormalizeTearaways()


    End Sub

#End Region

#Region " Populate from a Directory "

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
    ''' The grid will then be cleared and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String)

        PopulateFromADirectory(Dirname, _DefaultCellFont, _DefaultForeColor, "*")

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
    ''' The grid will then be cleared and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>gridfont</c> font will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="Gridfont"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal Gridfont As Font)

        PopulateFromADirectory(Dirname, Gridfont, _DefaultForeColor, "*")

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
    ''' The grid will then be cleared and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>col</c> color will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal col As Color)

        PopulateFromADirectory(Dirname, _DefaultCellFont, col, "*")

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
    ''' The grid will then be cleared and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>col</c> color and <c>gridfont</c> font will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal Gridfont As Font, ByVal col As Color)

        PopulateFromADirectory(Dirname, Gridfont, col, "*")

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
    ''' pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="Pattern"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal Pattern As String)

        PopulateFromADirectory(Dirname, _DefaultCellFont, _DefaultForeColor, Pattern)

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
    ''' pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>gridfont</c> font will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="Gridfont"></param>
    ''' <param name="Pattern"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal Gridfont As Font, ByVal Pattern As String)

        PopulateFromADirectory(Dirname, Gridfont, _DefaultForeColor, Pattern)

    End Sub

    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
    ''' pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>col</c> color will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="col"></param>
    ''' <param name="Pattern"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal col As Color, ByVal Pattern As String)

        PopulateFromADirectory(Dirname, _DefaultCellFont, col, Pattern)

    End Sub


    ''' <summary>
    ''' Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
    ''' pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
    ''' contents showing the FileName, Last Update Time, and physical size. The supplied
    ''' <c>col</c> color and <c>gridfont</c> font will be used to show the content generated. 
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="gridfont"></param>
    ''' <param name="col"></param>
    ''' <param name="Pattern"></param>
    ''' <remarks></remarks>
    Public Sub PopulateFromADirectory(ByVal Dirname As String, ByVal gridfont As Font, ByVal col As Color, ByVal Pattern As String)

        Try
            Dim dinf As New System.IO.DirectoryInfo(Dirname)

            Dim finf As System.IO.FileInfo() = dinf.GetFiles(Pattern)
            Dim y As Integer
            Dim r, c As Integer

            r = finf.GetUpperBound(0)
            c = 3

            InitializeTheGrid(r + 1, c)

            _GridTitle = "Files in " + Dirname

            _GridHeader(0) = "File Name"
            _GridHeader(1) = "File Time"
            _GridHeader(2) = "File Size"

            For y = 0 To r
                _grid(y, 0) = finf(y).FullName
                _grid(y, 1) = finf(y).LastAccessTime
                _grid(y, 2) = finf(y).Length.ToString()
            Next

            AllCellsUseThisFont(gridfont)
            AllCellsUseThisForeColor(col)

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

            Me.Refresh()

            NormalizeTearaways()
        Catch ex As Exception

            InitializeTheGrid(1, 3)
            _GridTitle = "Files in " + "We got a problem..."

            _GridHeader(0) = "File Name"
            _GridHeader(1) = "File Time"
            _GridHeader(2) = "File Size"

            Me.Refresh()

            NormalizeTearaways()

        End Try



        'r = arr.GetUpperBound(0) + 1
        'c = arr.GetUpperBound(1) + 1

        'If FirstRowHeader Then
        '    InitializeTheGrid(r - 1, c)
        '    For y = 0 To c - 1
        '        _GridHeader(y) = arr(0, y).ToString
        '    Next
        '    For x = 1 To r - 1
        '        For y = 0 To c - 1
        '            _grid(x, y) = arr(x, y).ToString
        '        Next
        '    Next
        'Else
        '    InitializeTheGrid(r, c)
        '    For y = 0 To c - 1
        '        _GridHeader(y) = "Column - " & y.ToString
        '    Next
        '    For x = 0 To r - 1
        '        For y = 0 To c - 1
        '            _grid(x, y) = arr(x, y).ToString
        '        Next
        '    Next
        'End If

        'AllCellsUseThisFont(gridfont)
        'AllCellsUseThisForeColor(col)

        'Me.AutoSizeCellsToContents = True
        '_colEditRestrictions.Clear()

        'Me.Refresh()

        'NormalizeTearaways()



    End Sub

    ''' <summary>
    ''' Attempts to open the directory specfied by <c>Dirname</c> and enumerate the entire contents 
    ''' The results are the appended to the current grids contents 
    ''' The FileName, Last Update Time, and physical size are enumerated.
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <remarks></remarks>
    Public Sub AppendPopulate(ByVal Dirname As String)

        AppendPopulate(Dirname, "*")

    End Sub

    ''' <summary>
    ''' Attempts to open the directory specfied by <c>Dirname</c> and enumerate the contents via the supplied <c>Pattern</c>
    ''' The results are the appended to the current grids contents 
    ''' The FileName, Last Update Time, and physical size are enumerated.
    ''' </summary>
    ''' <param name="Dirname"></param>
    ''' <param name="Pattern"></param>
    ''' <remarks></remarks>
    Public Sub AppendPopulate(ByVal Dirname As String, ByVal Pattern As String)
        Dim dinf As New System.IO.DirectoryInfo(Dirname)

        Dim finf As System.IO.FileInfo() = dinf.GetFiles(Pattern)
        Dim y As Integer
        Dim r As Integer

        r = finf.GetUpperBound(0)

        _GridTitle = "Files in..."

        If _cols <> 3 Then
            _cols = 3

            _GridHeader(0) = "File Name"
            _GridHeader(1) = "File Time"
            _GridHeader(2) = "File Size"

        End If

        Dim oldrows As Integer = _rows - 1

        Rows += r + 1

        For y = 0 To r
            _grid(y + oldrows, 0) = finf(y).FullName
            _grid(y + oldrows, 1) = finf(y).LastAccessTime
            _grid(y + oldrows, 2) = finf(y).Length.ToString()
        Next

        AllCellsUseThisFont(_DefaultCellFont)
        AllCellsUseThisForeColor(_DefaultForeColor)

        Me.AutoSizeCellsToContents = True
        _colEditRestrictions.Clear()

        Me.Refresh()

    End Sub
#End Region

#End Region

    ''' <summary>
    ''' Will fire the CellClicked event from the outside world
    ''' </summary>
    ''' <param name="Row"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub RaiseCellClickedEvent(ByVal Row As Integer, ByVal col As Integer)
        RaiseEvent CellClicked(Me, Row, col)
    End Sub

    ''' <summary>
    ''' Will fire the CellDoubleClicked event from the outside world
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub RaiseCellDoubleClickedEvent(ByVal row As Integer, ByVal col As Integer)
        RaiseEvent CellDoubleClicked(Me, row, col)
    End Sub

    ''' <summary>
    ''' Will fire the GridHover event from the outside world
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="rowid"></param>
    ''' <param name="colid"></param>
    ''' <param name="textvalue"></param>
    ''' <remarks></remarks>
    Public Sub RaiseGridHoverEvents(ByVal sender As Object, ByVal rowid As Integer, ByVal colid As Integer, ByVal textvalue As String)

        If Not _TearAwayWork Then

            RaiseEvent GridHover(sender, rowid, colid, textvalue)

        End If

    End Sub

    ''' <summary>
    ''' Will remove a specified row of data from the grids contents
    ''' If rowid is greater than rows in the grid nothing will be removed
    ''' </summary>
    ''' <param name="rowid"></param>
    ''' <remarks></remarks>
    Public Sub RemoveRowFromGrid(ByVal rowid As Integer)

        Dim hdr() As String = _GridHeader

        If rowid < 0 Or rowid > _rows - 1 Then
            Exit Sub
        End If

        _Painting = True    ' stop the painting operations

        Dim tmpgrid(,) As String = _grid    ' get the current grid
        Dim newtmpgrid(_rows - 2, _cols - 1) As String  ' a temporary storage area
        Dim x As Integer = 0
        Dim xx As Integer = 0
        Dim y As Integer = 0


        For x = 0 To _rows - 1
            If x <> rowid Then
                For y = 0 To _cols - 1
                    newtmpgrid(xx, y) = tmpgrid(x, y)
                Next
                xx += 1
            End If
        Next


        'For x = 0 To rowid - 1  ' loop through all the rows up to the one we want to get rid of
        '    For y = 0 To _cols - 1
        '        newtmpgrid(x, y) = tmpgrid(x, y)    ' all the colums
        '    Next
        'Next

        'For x = rowid + 1 To _rows - 1  ' loop 
        '    For y = 0 To _cols - 1
        '        newtmpgrid(x - 1, y) = tmpgrid(x, y)
        '    Next
        'Next

        ' Me.PopulateGridFromArray(newtmpgrid)

        PopulateGridFromArray(newtmpgrid, _DefaultCellFont, _DefaultForeColor, False, False, hdr)

        _Painting = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Will attempt to remove specific rowws from the grid contained in the supplied
    ''' arraylist of integers
    ''' </summary>
    ''' <param name="ListOfRows"></param>
    ''' <remarks></remarks>
    Public Sub RemoveRowsFromGrid(ByVal ListOfRows As ArrayList)
        ' will take arraylist rows and purge them from the _grid(x,y) array

        If ListOfRows.Count = 0 Then
            ' I see nothing... I Do Nothing...
            Exit Sub
        End If

        Dim hdr() As String = _GridHeader
        Dim x, y, t As Integer

        _Painting = True    ' stop the painting operations

        ' calculate final row count

        x = 0
        For t = 0 To _rows - 1
            If Not ListOfRows.Contains(t) Then
                x += 1
            End If
        Next

        Dim finalgrid(x - 1, _cols - 1) As String

        x = 0
        For t = 0 To _rows - 1

            If Not ListOfRows.Contains(t) Then
                ' we have a row to go
                For y = 0 To _cols - 1
                    finalgrid(x, y) = _grid(t, y)
                Next
                x += 1
            End If
        Next

        ' finalgrid should have what we need now

        _SelectedRow = -1
        _SelectedRows.Clear()

        Dim colp() As String = _colPasswords
        Dim colmaxchars() As Integer = _colMaxCharacters
        Dim coledit() As Boolean = _colEditable
        Dim rowedit() As Boolean = _rowEditable
        Dim colhid() As Boolean = _colhidden

        PopulateGridFromArray(finalgrid, _DefaultCellFont, _DefaultForeColor, False, False, hdr)

        _colPasswords = colp
        _colMaxCharacters = colmaxchars
        _colEditable = coledit
        _colhidden = colhid

        _rowEditable = rowedit

        _Painting = False

        Me.Refresh()
        NormalizeTearaways()

    End Sub

    ''' <summary>
    ''' Will attempt to remove a specific column from the grids contents
    ''' If colid is greater than the number of columns in the grid nothing will
    ''' be removed
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <remarks></remarks>
    Public Sub RemoveColFromGrid(ByVal colid As Integer)
        If colid < 0 Or colid > _cols - 1 Then
            Exit Sub
        End If

        _Painting = True    ' stop the painting operations

        Dim tmpgrid(,) As String = Me.GetGridAsArray    ' get the current grid
        Dim newtmpgrid(_rows - 1, _cols - 2) As String  ' a temporary storage area
        Dim x As Integer = 0
        Dim y As Integer = 0

        If colid <> 0 Then
            ' we are not skipping the first row
            For y = 0 To colid - 1
                For x = 0 To _rows - 1
                    newtmpgrid(x, y) = tmpgrid(x, y)
                Next
            Next
        End If

        If colid < _cols - 1 Then
            ' we are not skipping the laast row
            For y = colid + 1 To _cols - 1
                For x = 0 To _rows - 1
                    newtmpgrid(x, y - 1) = tmpgrid(x, y)
                Next
            Next
        End If

        Me.PopulateGridFromArray(newtmpgrid)
        'PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)

        CheckGridTearAways(colid)

        _Painting = False

        Me.Refresh()

    End Sub

    ''' <summary>
    ''' Will attempt to walk the contents of a column for integers 1 through 12
    ''' on finding a 1 through 12 it will replace the integers with the name of that
    ''' month number IE 1 = January, 2 = February...
    ''' </summary>
    ''' <param name="columnid"></param>
    ''' <remarks></remarks>
    Public Sub ReplaceColMonthNumericWithMonthName(ByVal columnid As Integer)
        Dim y As Integer
        Dim a As String

        For y = 0 To _rows - 1
            If _grid(y, columnid) Is Nothing Then

            Else
                a = _grid(y, columnid)
                If IsNumeric(a) Then
                    ' its at least a number
                    Select Case CInt(a)
                        Case 1
                            a = "January"
                        Case 2
                            a = "February"
                        Case 3
                            a = "March"
                        Case 4
                            a = "April"
                        Case 5
                            a = "May"
                        Case 6
                            a = "June"
                        Case 7
                            a = "July"
                        Case 8
                            a = "August"
                        Case 9
                            a = "September"
                        Case 10
                            a = "October"
                        Case 11
                            a = "November"
                        Case 12
                            a = "December"
                        Case Else
                            ' leave a alone
                    End Select

                    _grid(y, columnid) = a

                End If
            End If
        Next

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Will walk the list of columns tornaway and will set the windows to be siz width
    ''' </summary>
    ''' <param name="siz"></param>
    ''' <remarks></remarks>
    Public Sub ResizeTearawayColumnsHorizontally(ByVal siz As Integer)

        If TearAways.Count = 0 Then
            ' we ain't got any stinking tearaways so lets bail
            Exit Sub
        End If

        Dim t As Integer

        Dim tear As TearAwayWindowEntry

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.Width = siz
            'Application.DoEvents()
        Next

        Me.ArrangeTearAwayWindows()

    End Sub

    ''' <summary>
    ''' Will walk the list of torn away columns and will set each one to siz height
    ''' </summary>
    ''' <param name="siz"></param>
    ''' <remarks></remarks>
    Public Sub ResizeTearawayColumnsVertically(ByVal siz As Integer)

        If TearAways.Count = 0 Then
            ' we ain't got any stinking tearaways so lets bail
            Exit Sub
        End If

        If _TearAwayWork Then
            Exit Sub
        End If
        _TearAwayWork = True

        Dim t As Integer

        Dim tear As TearAwayWindowEntry

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.SuspendLayout()
            tear.Winform.Height = siz
            'Application.DoEvents()
        Next

        Me.ArrangeTearAwayWindows()

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.ResumeLayout()
            'tear.Winform.Height = siz
            'Application.DoEvents()
        Next

        _TearAwayWork = False


    End Sub

    ''' <summary>
    ''' Will walk the list of torn away columns and will set each to sizx width andf sizy height
    ''' </summary>
    ''' <param name="sizx"></param>
    ''' <param name="sizy"></param>
    ''' <remarks></remarks>
    Public Sub ResizeTearawayColumnsVerticallyAndHorizontally(ByVal sizx As Integer, ByVal sizy As Integer)

        If TearAways.Count = 0 Then
            ' we ain't got any stinking tearaways so lets bail
            Exit Sub
        End If

        Dim t As Integer

        Dim tear As TearAwayWindowEntry

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.Height = sizy
            tear.Winform.Width = sizx
            'Application.DoEvents()
        Next

        Me.ArrangeTearAwayWindows()


    End Sub

    ''' <summary>
    ''' Will set the column at colid edit restrictions to the list contained in the CaretDelimitedString
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <param name="CaretDelimitedString"></param>
    ''' <remarks></remarks>
    Public Sub RestrictColumnEditsTo(ByVal colid As Integer, ByVal CaretDelimitedString As String)
        Dim reslist As New EditColumnRestrictor

        reslist.ColumnID = colid
        reslist.RestrictedList = CaretDelimitedString

        For Each it As EditColumnRestrictor In _colEditRestrictions
            If it.ColumnID = colid Then
                _colEditRestrictions.Remove(it)
            End If
        Next

        _colEditRestrictions.Add(reslist)

    End Sub

    ''' <summary>
    ''' Will set the column at colid edit restrictions list to the supplied ArrayListOfStrings
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <param name="ArrayListOfStrings"></param>
    ''' <remarks></remarks>
    Public Sub RestrictColumnEditsTo(ByVal colid As Integer, ByVal ArrayListOfStrings As ArrayList)
        Dim reslist As New EditColumnRestrictor

        Dim CaretDelimitedString As String = ""

        For Each ar As String In ArrayListOfStrings
            ar = ar.Replace("^", "+")
            CaretDelimitedString += ar + "^"
        Next

        If CaretDelimitedString.EndsWith("^") Then
            CaretDelimitedString = CaretDelimitedString.Substring(0, CaretDelimitedString.Length - 1)
        End If

        reslist.ColumnID = colid
        reslist.RestrictedList = CaretDelimitedString

        For Each it As EditColumnRestrictor In _colEditRestrictions
            If it.ColumnID = colid Then
                _colEditRestrictions.Remove(it)
            End If
        Next

        _colEditRestrictions.Add(reslist)

    End Sub

    Public Function ReturnDelimitedStringAsArray(ByVal StringToParse As String, ByVal Columns As Integer, ByVal rows As Integer, ByVal Delimiter As String) As Array
        'add a delimiter to the begning and the end of the string
        If Not StringToParse.StartsWith(Delimiter) Then
            StringToParse = Delimiter & StringToParse
        End If
        If Not StringToParse.EndsWith(Delimiter) Then
            StringToParse = StringToParse & Delimiter
        End If

        Dim mc As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(StringToParse, Delimiter)
        Dim RegExCounter As Integer = 0
        Dim rowCounter As Integer = 0
        Dim argarray(rows, Columns - 1) As String

        Do Until RegExCounter >= mc.Count - 1
            Dim intCol As Integer = 0
            'parse string
            argarray(rowCounter, intCol) = StringToParse.Substring(mc.Item(RegExCounter).Index + Delimiter.Length, mc.Item(RegExCounter + 1).Index - mc.Item(RegExCounter).Index - Delimiter.Length)
            RegExCounter += 1
            intCol += 1
            Do Until intCol = Columns
                If (mc.Item(RegExCounter).Index + Delimiter.Length) <> StringToParse.Length Then
                    argarray(rowCounter, intCol) = StringToParse.Substring(mc.Item(RegExCounter).Index + Delimiter.Length, mc.Item(RegExCounter + 1).Index - mc.Item(RegExCounter).Index - Delimiter.Length)
                End If
                intCol += 1
                RegExCounter += 1
            Loop
            rowCounter += 1
        Loop
        'clear out some memory
        mc = Nothing
        Return argarray
    End Function

    ''' <summary>
    ''' Will attempt to render the grids surface onto the supplied graphics context GR, The grid will be rendered into the rectangle denoted by
    ''' xloc,yloc and width and height
    ''' </summary>
    ''' <param name="gr"></param>
    ''' <param name="xloc"></param>
    ''' <param name="yloc"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub PlaceGridOnGraphicsContext(ByVal gr As Graphics, ByVal xloc As Integer, ByVal yloc As Integer, ByVal width As Integer, ByVal height As Integer)
        Dim cr As New Rectangle(xloc, yloc, width, height)

        RenderGridToGraphicsContext(gr, cr)

    End Sub

    ''' <summary>
    ''' Walks the list of open tear away columns and sets the to be on top of all windows
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PullAllTearAwaysToTheFront()
        If TearAways.Count = 0 Then
            Exit Sub
        End If

        If _TearAwayWork Then
            Exit Sub
        End If

        _TearAwayWork = True

        Dim t As Integer
        Dim tear As TearAwayWindowEntry

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.TopMost = True
            tear.Winform.BringToFront()
        Next

        _TearAwayWork = False
    End Sub

    ''' <summary>
    ''' Walks the list of tear away columns and sets them to be behind all open windows.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PushAllTearAwaysToTheBack()
        If TearAways.Count = 0 Then
            Exit Sub
        End If


        If _TearAwayWork Then
            Exit Sub
        End If

        _TearAwayWork = True


        Dim t As Integer
        Dim tear As TearAwayWindowEntry

        For t = 0 To TearAways.Count - 1
            tear = TearAways.Item(t)
            tear.Winform.TopMost = False
            tear.Winform.SendToBack()
        Next

        _TearAwayWork = False

    End Sub

    ''' <summary>
    ''' Selects all rows in the current grid
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SelectAllRows()
        Dim aList As New ArrayList
        Dim i As Integer
        For i = 0 To Me.Rows - 1
            aList.Add(i)
        Next

        Me.SelectedRows = aList

    End Sub

    ''' <summary>
    ''' Selects rows in the current grid from an arraylist of row IDs
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SelectRows(ByVal rowIDs As ArrayList)

        Try
            Me.SelectedRows = rowIDs
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.SelectRows Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Sets the background color for all cells in the grid to be the supplied color
    ''' </summary>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetAllCellBackcolors(ByVal color As Color)
        Dim x, y As Integer

        Dim colentry As Integer = GetGridBackColorListEntry(New SolidBrush(color))

        For x = 0 To _cols - 1
            For y = 0 To _rows - 1
                _gridBackColor(y, x) = colentry
            Next
        Next

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Sets the foreground color for all cells in the grid to be the supplied color
    ''' </summary>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetAllCellForecolors(ByVal color As Color)
        Dim x, y As Integer

        Dim colentry As Integer = GetGridForeColorListEntry(New Pen(color))

        For x = 0 To _cols - 1
            For y = 0 To _rows - 1
                _gridForeColor(y, x) = colentry
            Next
        Next

        Me.Invalidate()

    End Sub

    ''' <summary>
    ''' Sets all the cells in the specific column Col to be the the supplied color in the background 
    ''' </summary>
    ''' <param name="Col"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetColBackColor(ByVal Col As Integer, ByVal color As Color)

        Try

            If Col >= 0 And Col < _cols Then

                Dim iCol As Integer

                For iCol = 0 To _rows - 1

                    _gridBackColor(iCol, Col) = GetGridBackColorListEntry(New SolidBrush(color))

                Next

            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.SetColBackColor Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Sets all the cells in the specific column Col to be the the supplied color in the foreground 
    ''' </summary>
    ''' <param name="Col"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetColForeColor(ByVal Col As Integer, ByVal color As Color)

        Try
            If Col >= 0 And Col < _cols Then
                Dim iCol As Integer

                For iCol = 0 To _rows - 1

                    _gridForeColor(iCol, Col) = GetGridForeColorListEntry(New Pen(color))

                Next
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.SetColForeColor Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Applied the specified enumerated colorscheme to the grids contents
    ''' </summary>
    ''' <param name="Scheme"></param>
    ''' <remarks></remarks>
    Public Sub SetColorScheme(ByVal Scheme As TaiGridColorSchemes)

        Select Case Scheme
            Case TaiGridColorSchemes._Default
                _GridTitleBackcolor = Color.Blue
                _GridTitleForeColor = Color.White
                _GridHeaderBackcolor = Color.LightBlue
                _GridHeaderForecolor = Color.Black
                _CellOutlineColor = Color.Black
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.AntiqueWhite
                DefaultForegroundColor = System.Drawing.Color.Black
                _RowHighLiteBackColor = System.Drawing.Color.Blue
                _RowHighLiteForeColor = System.Drawing.Color.White
                _ColHighliteBackColor = System.Drawing.Color.MediumSlateBlue
                _ColHighliteForeColor = System.Drawing.Color.LightGray
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
            Case TaiGridColorSchemes._Technical
                _GridTitleBackcolor = Color.DarkBlue
                _GridTitleForeColor = Color.GhostWhite
                _GridHeaderBackcolor = Color.LightBlue
                _GridHeaderForecolor = Color.Black
                _CellOutlineColor = Color.Black
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.LightYellow
                DefaultForegroundColor = System.Drawing.Color.Black
                _RowHighLiteBackColor = System.Drawing.Color.LightSlateGray
                _RowHighLiteForeColor = System.Drawing.Color.Black
                _ColHighliteBackColor = System.Drawing.Color.MediumSpringGreen
                _ColHighliteForeColor = System.Drawing.Color.Black
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
            Case TaiGridColorSchemes._Colorful1
                _GridTitleBackcolor = Color.Blue
                _GridTitleForeColor = Color.Yellow
                _GridHeaderBackcolor = Color.Violet
                _GridHeaderForecolor = Color.Yellow
                _CellOutlineColor = Color.White
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.MediumPurple
                DefaultForegroundColor = Color.Yellow
                _RowHighLiteBackColor = System.Drawing.Color.Blue
                _RowHighLiteForeColor = System.Drawing.Color.White
                _ColHighliteBackColor = System.Drawing.Color.MediumSlateBlue
                _ColHighliteForeColor = System.Drawing.Color.LightGray
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
            Case TaiGridColorSchemes._Colorful2
                _GridTitleBackcolor = Color.Violet
                _GridTitleForeColor = Color.White
                _GridHeaderBackcolor = Color.Blue
                _GridHeaderForecolor = Color.White
                _CellOutlineColor = Color.Black
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.AntiqueWhite
                DefaultForegroundColor = System.Drawing.Color.Black
                _RowHighLiteBackColor = System.Drawing.Color.Blue
                _RowHighLiteForeColor = System.Drawing.Color.White
                _ColHighliteBackColor = System.Drawing.Color.MediumSlateBlue
                _ColHighliteForeColor = System.Drawing.Color.LightGray
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
            Case TaiGridColorSchemes._Fancy
                _GridTitleBackcolor = Color.Blue
                _GridTitleForeColor = Color.White
                _GridHeaderBackcolor = Color.LightBlue
                _GridHeaderForecolor = Color.Black
                _CellOutlineColor = Color.Black
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.AntiqueWhite
                DefaultForegroundColor = System.Drawing.Color.Black
                _RowHighLiteBackColor = System.Drawing.Color.Blue
                _RowHighLiteForeColor = System.Drawing.Color.White
                _ColHighliteBackColor = System.Drawing.Color.MediumSlateBlue
                _ColHighliteForeColor = System.Drawing.Color.LightGray
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
            Case Else
                _GridTitleBackcolor = Color.Blue
                _GridTitleForeColor = Color.White
                _GridHeaderBackcolor = Color.LightBlue
                _GridHeaderForecolor = Color.Black
                _CellOutlineColor = Color.Black
                _alternateColorationALTColor = Color.MediumSpringGreen
                _alternateColorationBaseColor = Color.AntiqueWhite
                DefaultBackgroundColor = Color.AntiqueWhite
                DefaultForegroundColor = System.Drawing.Color.Black
                _RowHighLiteBackColor = System.Drawing.Color.Blue
                _RowHighLiteForeColor = System.Drawing.Color.White
                _ColHighliteBackColor = System.Drawing.Color.MediumSlateBlue
                _ColHighliteForeColor = System.Drawing.Color.LightGray
                _BorderColor = Color.Black
                _excelAlternateRowColor = Color.FromArgb(204, 255, 204)
                Me.Refresh()
        End Select
    End Sub

    ''' <summary>
    ''' Will apply the supplied <c>ItemToSet</c> to the cell currently being edited in the grid
    ''' </summary>
    ''' <param name="ItemToSet"></param>
    ''' <remarks></remarks>
    Public Sub SetEditItemText(ByVal ItemToSet As String)
        If _EditMode Then
            If txtInput.Visible Then
                txtInput.Text = ItemToSet
            Else
                cmboInput.Text = ItemToSet
            End If
        End If
    End Sub

    ''' <summary>
    ''' Attempts to set the cell at <c>row</c> and <c>col</c> to be in edit mode
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <remarks></remarks>
    Public Sub SetEditItem(ByVal row As Integer, ByVal col As Integer)
        ' lets do some sanity checking here
        If row > -1 And row < _rows And _rowEditable(row) Then
            ' the rows are in range
            If col > -1 And col < _cols Then
                ' the cols are in range
                ' is that column editable
                If _colEditable(col) Then
                    ' aye it be editable

                    Me.Focus()

                    Dim xoff, yoff, r, c As Integer

                    _RowClicked = row
                    _ColClicked = col

                    If _RowClicked > -1 And _RowClicked < _rows Then
                        If _ColClicked > -1 And _ColClicked < _cols And _colEditable(_ColClicked) And _AllowInGridEdits Then

                            If IsColumnRestricted(_ColClicked) Then

                                Dim it As EditColumnRestrictor = GetColumnRestriction(_ColClicked)

                                cmboInput.Items.Clear()

                                Dim s() As String = it.RestrictedList.Split("^".ToCharArray)
                                For Each ss As String In s
                                    cmboInput.Items.Add(ss)
                                Next

                                ' we have selected a row and col lets move the txtinput there and bring it to the front
                                xoff = 0
                                yoff = 0

                                If _RowClicked > 0 Then
                                    For r = 0 To _RowClicked - 1
                                        yoff = yoff + RowHeight(r)
                                    Next
                                End If

                                If GridheaderVisible Then
                                    yoff = yoff + _GridHeaderHeight
                                End If

                                If _GridTitleVisible Then
                                    yoff = yoff + _GridTitleHeight
                                End If

                                If _ColClicked > 0 Then
                                    For c = 0 To _ColClicked - 1
                                        xoff = xoff + ColWidth(c)
                                    Next
                                End If

                                If vs.Visible And vs.Value > 0 Then
                                    yoff = yoff - GimmeYOffset(vs.Value)
                                End If

                                If hs.Visible And hs.Value > 0 Then
                                    xoff = xoff - GimmeXOffset(hs.Value)
                                End If

                                If _CellOutlines Then
                                    cmboInput.Top = yoff + 1
                                    cmboInput.Left = xoff + 1
                                    cmboInput.Width = ColWidth(_ColClicked) - 1
                                    cmboInput.Height = RowHeight(_RowClicked) - 2
                                    cmboInput.BackColor = _colEditableTextBackColor
                                Else
                                    cmboInput.Top = yoff
                                    cmboInput.Left = xoff
                                    cmboInput.Width = ColWidth(_ColClicked)
                                    cmboInput.Height = RowHeight(_RowClicked)
                                    cmboInput.BackColor = _colEditableTextBackColor
                                End If

                                cmboInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                                cmboInput.Text = _grid(_RowClicked, _ColClicked)

                                cmboInput.Visible = True
                                cmboInput.BringToFront()
                                cmboInput.DroppedDown = True
                                _EditModeCol = _ColClicked
                                _EditModeRow = _RowClicked
                                _EditMode = True

                                cmboInput.Focus()
                            Else
                                ' we have selected a row and col lets move the txtinput there and bring it to the front
                                xoff = 0
                                yoff = 0

                                If _RowClicked > 0 Then
                                    For r = 0 To _RowClicked - 1
                                        yoff = yoff + RowHeight(r)
                                    Next
                                End If

                                If GridheaderVisible Then
                                    yoff = yoff + _GridHeaderHeight
                                End If

                                If _GridTitleVisible Then
                                    yoff = yoff + _GridTitleHeight
                                End If

                                If _ColClicked > 0 Then
                                    For c = 0 To _ColClicked - 1
                                        xoff = xoff + ColWidth(c)
                                    Next
                                End If

                                If vs.Visible And vs.Value > 0 Then
                                    yoff = yoff - GimmeYOffset(vs.Value)
                                End If

                                If hs.Visible And hs.Value > 0 Then
                                    xoff = xoff - GimmeXOffset(hs.Value)
                                End If

                                If _CellOutlines Then
                                    txtInput.Top = yoff + 1
                                    txtInput.Left = xoff + 1
                                    txtInput.Width = ColWidth(_ColClicked) - 1
                                    txtInput.Height = RowHeight(_RowClicked) - 2
                                    txtInput.BackColor = _colEditableTextBackColor
                                Else
                                    txtInput.Top = yoff
                                    txtInput.Left = xoff
                                    txtInput.Width = ColWidth(_ColClicked)
                                    txtInput.Height = RowHeight(_RowClicked)
                                    txtInput.BackColor = _colEditableTextBackColor
                                End If

                                txtInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                                txtInput.Text = _grid(_RowClicked, _ColClicked)

                                txtInput.Visible = True
                                txtInput.BringToFront()
                                _EditModeCol = _ColClicked
                                _EditModeRow = _RowClicked
                                _EditMode = True

                                txtInput.Focus()
                            End If

                        End If
                    End If

                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Sets the row at <c>row</c> to be the corresponging <c>color</c> background color
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetRowBackColor(ByVal row As Integer, ByVal color As Color)

        Try

            If row >= 0 And row < _rows Then

                Dim iCol As Integer

                For iCol = 0 To Me.Cols

                    _gridBackColor(row, iCol) = GetGridBackColorListEntry(New SolidBrush(color))

                Next

            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.SetRowBackColor Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Sets the row at <c>row</c> to have the corresponding <c>color</c> foreground color
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="color"></param>
    ''' <remarks></remarks>
    Public Sub SetRowForeColor(ByVal row As Integer, ByVal color As Color)

        Try
            If row >= 0 And row < _rows Then
                Dim iCol As Integer

                For iCol = 0 To Me.Cols

                    _gridForeColor(row, iCol) = GetGridForeColorListEntry(New Pen(color))

                Next
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.SetRowForeColor Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will attempt to sort the contents of the current grid on <c>col</c> 
    ''' If <c>Descending</c> is true or false will dictate the order of the sort
    ''' </summary>
    ''' <param name="col"></param>
    ''' <param name="Descending"></param>
    ''' <remarks></remarks>
    Public Sub SortGridOnColumn(ByVal col As Integer, ByVal Descending As Boolean)

        If col < 0 Or col > _cols - 1 Then
            Exit Sub
        End If

        Dim oldcol As Integer = _ColOverOnMenuButton

        _ColOverOnMenuButton = col

        If Descending Then

            miSortDescending_Click(Me, New EventArgs)

        Else

            miSortAscending_Click(Me, New EventArgs)

        End If

        _ColOverOnMenuButton = oldcol

    End Sub

    ''' <summary>
    ''' Will atempt to sort the grids contents on <c>col</c> treating the column contents as dates. 
    ''' The <c>Descending</c> parameter will distate the order of the sort
    ''' </summary>
    ''' <param name="col"></param>
    ''' <param name="Descending"></param>
    ''' <remarks></remarks>
    Public Sub SortGridOnColumnDate(ByVal col As Integer, ByVal Descending As Boolean)

        If col < 0 Or col > _cols - 1 Then
            Exit Sub
        End If

        Dim oldcol As Integer = _ColOverOnMenuButton

        _ColOverOnMenuButton = col

        If Descending Then

            miDateDesc_Click(Me, New EventArgs)

        Else

            miDateAsc_Click(Me, New EventArgs)

        End If

        _ColOverOnMenuButton = oldcol

    End Sub

    ''' <summary>
    ''' Will atempt to sort the grids contents on <c>col</c> treating the column contents as numbers. 
    ''' The <c>Descending</c> parameter will distate the order of the sort
    ''' </summary>
    ''' <param name="col"></param>
    ''' <param name="Descending"></param>
    ''' <remarks></remarks>
    Public Sub SortGridOnColumnNumeric(ByVal col As Integer, ByVal Descending As Boolean)

        If col < 0 Or col > _cols - 1 Then
            Exit Sub
        End If

        Dim oldcol As Integer = _ColOverOnMenuButton

        _ColOverOnMenuButton = col

        If Descending Then

            miSortNumericDesc_Click(Me, New EventArgs)

        Else

            miSortNumericAsc_Click(Me, New EventArgs)

        End If

        _ColOverOnMenuButton = oldcol

    End Sub

    ''' <summary>
    ''' Will attempt to add all the values in a column denoted by <c>colnum</c> and return
    ''' the sum as a double
    ''' </summary>
    ''' <param name="colnum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SumUpColumn(ByVal colnum As Integer) As Double

        Dim t As Integer
        Dim a As String
        Dim result As Double = 0.0

        If colnum >= _cols Or _rows < 1 Then
            Return result
        End If

        Try
            For t = 0 To _rows - 1
                a = _grid(t, colnum).Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")

                result = result + Val(a)
            Next
        Catch ex As Exception
            ' this is really a cludge to not bomb on 
            ' certain format conversion errors

        End Try

        Return result

    End Function

    ''' <summary>
    ''' Manually sets the position of the verticle scrollbar of the grid if the contents are larger 
    ''' than the physical grid window.
    ''' </summary>
    ''' <param name="sb"></param>
    ''' <remarks></remarks>
    Public Sub SetVertScrollbarPosition(ByVal sb As Integer)
        If vs.Visible Then
            If sb <= vs.Maximum And sb >= vs.Minimum Then
                vs.Value = sb
            End If
        End If
    End Sub

    ''' <summary>
    ''' Will attempt to wrap the text data in a specified <c>col</c> at <c>wraplen</c> length.
    ''' The wrap is smat in that it tries to wrap on whitespace boundaries
    ''' </summary>
    ''' <param name="col"></param>
    ''' <param name="wraplen"></param>
    ''' <remarks></remarks>
    Public Sub WordWrapColumn(ByVal col As Integer, ByVal wraplen As Integer)
        Dim t As Integer

        For t = 0 To _rows - 1
            _grid(t, col) = SplitLongString(_grid(t, col), wraplen)
        Next

        Me.AutoSizeCellsToContents = True

        Me.Invalidate()

    End Sub

#Region " Column Math Functions Statistical"

    ''' <summary>
    ''' Will return the computed STDEV of all the numbers contained in the column denoted by <c>colid</c>
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnSTDEV(ByVal colid As Integer) As Double
        Dim arl As ArrayList = Me.GetColAsCleanedArrayList(colid)
        Dim avg As Double = 0
        Dim res As Double = 0

        Dim t As Integer

        If arl.Count = 0 Then
            Return res
            Exit Function
        End If


        For t = 0 To arl.Count - 1
            avg += Convert.ToDouble(arl.Item(t))
        Next

        ' now we can get the average

        avg = avg / arl.Count

        ' now to subtract the aaverage from each element in the array and square it
        ' giving us our squared deviations

        For t = 0 To arl.Count - 1
            arl.Item(t) = Convert.ToDouble(Convert.ToDouble(arl.Item(t)) - avg) ^ 2
        Next

        ' now lets get the sum of the squared deviations

        Dim sum As Double = 0

        For t = 0 To arl.Count - 1
            sum += Convert.ToDouble(arl.Item(t))
        Next

        ' Finally lets get the square root of the (sum / number in the set-1)

        res = Math.Sqrt(sum / (arl.Count - 1))

        Return res

    End Function

    ''' <summary>
    ''' Will return the computed STDEVP of all the numbers contained in the column denoted by <c>colid</c> 
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnSTDEVP(ByVal colid As Integer) As Double

        Dim arl As ArrayList = Me.GetColAsCleanedArrayList(colid)
        Dim avg As Double = 0
        Dim res As Double = 0

        Dim t As Integer

        If arl.Count = 0 Then
            Return res
            Exit Function
        End If


        For t = 0 To arl.Count - 1
            avg += Convert.ToDouble(arl.Item(t))
        Next

        ' now we can get the average

        avg = avg / arl.Count

        ' now to subtract the aaverage from each element in the array and square it
        ' giving us our squared deviations

        For t = 0 To arl.Count - 1
            arl.Item(t) = Convert.ToDouble(Convert.ToDouble(arl.Item(t)) - avg) ^ 2
        Next

        ' now lets get the sum of the squared deviations

        Dim sum As Double = 0

        For t = 0 To arl.Count - 1
            sum += Convert.ToDouble(arl.Item(t))
        Next

        ' Finally lets get the square root of the (sum / number in the set)

        res = Math.Sqrt(sum / arl.Count)

        Return res

    End Function

#End Region

#Region " Fuzzy Math Functions "

    ''' <summary>
    ''' Will calculate the fuzzy membership of the values at <c>colid</c> beyond <c>targetval</c> from the direction of <c>outlier</c>
    ''' Values will be between 0 and 1 where beyond <c>targetval</c> is 1 and between <c>outlier</c> and target are some portion
    ''' of 0 to 1. Will use a Liner function between <c>outlier</c> and <c>targetval</c> 
    ''' </summary>
    ''' <param name="colid"></param>
    ''' <param name="targetval"></param>
    ''' <param name="outlier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FuzzyColumnMembership(ByVal colid As Integer, ByVal targetval As Double, ByVal outlier As Double) As ArrayList
        Dim arl As New ArrayList
        Dim t, y As Integer
        Dim tv, tv2 As Double
        Dim tstr As String


        Dim lessthan As Boolean = True

        If outlier <= targetval Then
            lessthan = True
        Else
            lessthan = False
        End If

        For t = 0 To _rows - 1
            arl.Add(Convert.ToDouble("0.0"))
        Next

        For y = 0 To _rows - 1
            If Not IsNothing(_grid(y, colid)) Then
                ' get the string contained at that grids cell coordinates stripped of some money crap
                tstr = _grid(y, colid).Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "")
                If IsNumeric(tstr) Then
                    ' is it still a number? Yes it is 
                    tv = Convert.ToDouble(tstr)

                    ' Ok we got the number whats the direction of the check

                    If lessthan Then
                        ' we are checking up
                        If tv >= targetval Then
                            arl.Item(y) = Convert.ToDouble("1")
                        Else
                            If tv >= outlier Then
                                ' here we compute the weight

                                tv2 = targetval - outlier ' range of values

                                arl.Item(y) = 1 / (tv2 / (tv - outlier))

                            End If
                        End If
                    Else
                        ' we are checking down
                        If tv <= targetval Then
                            arl.Item(y) = Convert.ToDouble("1")
                        Else
                            If tv <= outlier Then
                                ' here we compute the weight

                                tv2 = outlier - targetval ' range of values

                                arl.Item(y) = 1 / (tv2 / (outlier - tv))

                            End If
                        End If
                    End If
                End If
            End If
        Next

        Return arl

    End Function

    Public Function FuzzyColumnCombine(ByVal colids As ArrayList) As ArrayList
        Dim arl As New ArrayList
        Dim bail As Boolean = False
        Dim t, x As Integer
        Dim sum As Double
        Dim s As String

        ' Initialize our resultset
        For t = 0 To _rows - 1
            arl.Add(Convert.ToDouble("0"))
        Next

        ' ensure all the ccols we are asking for membership values in are actually in the grid
        For t = 0 To colids.Count - 1
            x = Convert.ToInt32(colids.Item(t))
            If x < 0 Or x > _cols - 1 Then
                ' tyhis one is not there so lets setup to bail out of this 
                bail = True
                Exit For
            End If
        Next

        If bail Then
            ' time to bail out just return the all 0 membership set we crafted at he start
            Return arl
            Exit Function
        End If

        ' conditions are right for a test so lets check it out

        For t = 0 To _rows - 1
            sum = 0
            For x = 0 To colids.Count - 1

                s = _grid(t, Convert.ToInt32(colids.Item(x)))

                s = s.Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "")

                sum += Convert.ToDouble(s)

            Next

            arl.Item(t) = (sum / colids.Count)

        Next

        Return arl

    End Function

    Public Sub InsertFuzzyMathResultSet(ByVal arl As ArrayList, ByVal ColName As String)

        If arl.Count <> _rows Then
            Exit Sub
        End If

        Me.Cols += 1

        Me.HeaderLabel(_cols - 1) = ColName

        Dim y As Integer

        For y = 0 To arl.Count - 1
            Me.item(y, _cols - 1) = Convert.ToDouble(arl.Item(y)).ToString()
        Next

        Me.AutoSizeCellsToContents = True
        Me.Refresh()

    End Sub

#End Region

#Region " RollupFunctions "

    ''' <summary>
    ''' Will take the values at the specified <c>row</c> and starting at the specified <c>col</c> to the 
    ''' last column in the existing grid and add them up returning the result as a double.
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RollupColumn(ByVal row As Integer, ByVal col As Integer) As Double
        Dim result As Double = 0.0
        Dim t As Integer
        Dim a As String

        ' do some bounds checking
        If row > _rows - 1 Then
            Return result
            Exit Function
        End If

        If col > _cols - 1 Then
            Return result
            Exit Function
        End If

        For t = col To _cols - 1
            a = Me.item(row, t).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "")
            result = result + Val(a)
        Next

        Return result

    End Function

    ''' <summary>
    ''' will take the values at the specified <c>row</c> and <c>col</c> continuing to the edge of the grid
    ''' and will add them up returning the result as a double
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RollupCube(ByVal row As Integer, ByVal col As Integer) As Double
        Dim result As Double = 0.0
        Dim t, tt As Integer
        Dim a As String

        ' do some bounds checking
        If row > _rows - 1 Then
            Return result
            Exit Function
        End If

        If col > _cols - 1 Then
            Return result
            Exit Function
        End If

        For t = col To _cols - 1
            For tt = row To _rows - 1
                a = Me.item(tt, t).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "")
                result = result + Val(a)
            Next
        Next

        Return result
    End Function

    ''' <summary>
    ''' will take the values in a specified <c>col</c> and will take all the valies from the specfied <c>row</c>
    ''' until the last row in the grid and will add them up returning the result as a double
    ''' </summary>
    ''' <param name="row"></param>
    ''' <param name="col"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RollupRow(ByVal row As Integer, ByVal col As Integer) As Double
        Dim result As Double = 0.0
        Dim t As Integer
        Dim a As String

        ' do some bounds checking
        If row > _rows - 1 Then
            Return result
            Exit Function
        End If

        If col > _cols - 1 Then
            Return result
            Exit Function
        End If

        For t = row To _rows - 1
            a = Me.item(t, col).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "")
            result = result + Val(a)
        Next

        Return result
    End Function

#End Region

#Region " Import Methods "

    ''' <summary>
    ''' Will populate the grid from the supplied <c>sFilename</c>
    ''' The call will assume that the caller wants the first set of items from the supplied xml file
    ''' </summary>
    ''' <param name="sFilename"></param>
    ''' <remarks></remarks>
    Public Sub ImportFromXML(ByVal sFilename As String)

        Try

            Dim _ds As New DataSet

            _ds.ReadXml(sFilename, XmlReadMode.Auto)

            ' determine how many rows and columns
            Dim iRows As Integer = _ds.Tables(0).Rows.Count
            Dim iCols As Integer = _ds.Tables(0).Columns.Count

            Me.Rows = iRows
            Me.Cols = iCols

            ' fill in the column names
            Dim row As DataRow
            Dim iCol As Integer = 0
            Dim iRow As Integer = 0

            For iCol = 0 To iCols - 1
                _GridHeader(iCol) = _ds.Tables(0).Columns(iCol).ColumnName
            Next

            For iRow = 0 To iRows - 1
                row = _ds.Tables(0).Rows(iRow)
                For iCol = 0 To iCols - 1
                    If Not IsDBNull(row(iCol)) Then
                        _grid(iRow, iCol) = row(iCol)
                    Else
                        _grid(iRow, iCol) = "{NULL}"
                    End If

                Next
            Next

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ImportFromXML Error...")
        End Try

    End Sub

    ''' <summary>
    ''' Will populate the grid from the supplied <c>sFilename</c>
    ''' The call will attempt to get the data from the designated <c>tblnum</c> in the supplied xml file
    ''' </summary>
    ''' <param name="sFilename"></param>
    ''' <param name="tblnum"></param>
    ''' <remarks></remarks>
    Public Sub ImportFromXML(ByVal sFilename As String, ByVal tblnum As Integer)

        Try

            Dim _ds As New DataSet

            _ds.ReadXml(sFilename, XmlReadMode.Auto)

            If _ds.Tables.Count - 1 < tblnum Then
                tblnum = _ds.Tables.Count - 1
            End If

            ' determine how many rows and columns
            Dim iRows As Integer = _ds.Tables(tblnum).Rows.Count
            Dim iCols As Integer = _ds.Tables(tblnum).Columns.Count

            Me.Rows = iRows
            Me.Cols = iCols

            ' fill in the column names
            Dim row As DataRow
            Dim iCol As Integer = 0
            Dim iRow As Integer = 0

            For iCol = 0 To iCols - 1
                _GridHeader(iCol) = _ds.Tables(tblnum).Columns(iCol).ColumnName
            Next

            For iRow = 0 To iRows - 1
                row = _ds.Tables(tblnum).Rows(iRow)
                For iCol = 0 To iCols - 1
                    If Not IsDBNull(row(iCol)) Then
                        _grid(iRow, iCol) = row(iCol)
                    Else
                        _grid(iRow, iCol) = "{NULL}"
                    End If

                Next
            Next

            Me.AutoSizeCellsToContents = True
            _colEditRestrictions.Clear()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "TAIGRIDControl.ImportFromXML Error...")
        End Try

    End Sub

#End Region

#Region " Grid Printing Functions "

    ''' <summary>
    ''' Will attempt to print the contents of the grid to the default printer in the system
    ''' The grids own properties for Outlining printed cells, Printing page numbers,
    ''' Previewing the output first, Page orientation will be employed in the resulting process
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub PrintTheGrid()

        PrintTheGrid("", _gridReportMatchColors, _
                    _gridReportOutlineCells, _
                    _gridReportNumberPages, _
                    _gridReportPreviewFirst, _
                    _gridReportOrientLandscape)

    End Sub

    ''' <summary>
    ''' Will attempt to print the contents of the grid to the default printer in the system
    ''' The grids own properties for Outlining printed cells, Printing page numbers,
    ''' Previewing the output first, Page orientation will be employed in the resulting process
    ''' The supplied <c>Title</c> will be used to lable the pages of output
    ''' </summary>
    ''' <param name="Title"></param>
    ''' <remarks></remarks>
    Public Sub PrintTheGrid(ByVal Title As String)

        PrintTheGrid(Title, _gridReportMatchColors, _
                        _gridReportOutlineCells, _
                        _gridReportNumberPages, _
                        _gridReportPreviewFirst, _
                        _gridReportOrientLandscape)

    End Sub

    ''' <summary>
    ''' Will attempt to print the contents of the grid to the default printer in the system
    ''' The grids own properties for Outlining printed cells, Previewing the output first, 
    ''' The supplied <c>Title</c> will be used to lable the pages of output as well as the
    ''' supplied values for <c>NumberPages</c> and <c>Landscapemode</c> will override those setup 
    ''' in the grid properties
    ''' </summary>
    ''' <param name="Title"></param>
    ''' <param name="NumberPages"></param>
    ''' <param name="Landscapemode"></param>
    ''' <remarks></remarks>
    Public Sub PrintTheGrid(ByVal Title As String, ByVal NumberPages As Boolean, ByVal Landscapemode As Boolean)

        PrintTheGrid(Title, _gridReportMatchColors, _
                        _gridReportOutlineCells, _
                        NumberPages, _
                        _gridReportPreviewFirst, _
                        Landscapemode)

    End Sub

    ''' <summary>
    ''' Will attempt to print the contents of the grid to the default printer in the system
    ''' using supplied values for
    ''' <list type="Bullet">
    ''' <item> <c>Title</c> will use thee supplied strin g to title the resulting output</item>
    ''' <item> <c>MatchColors</c> attempting to match the colors on the grid with printed output</item>
    ''' <item> <c>OutlineCells</c> will draw an outline around each cell of output on the printed page</item>
    ''' <item> <c>NumberPages</c> will number each page as its printed</item>
    ''' <item> <c>PreviewFirst</c> will show the print preview windows forst before sending the results to the printer</item>
    ''' <item> <c>Landscapemode</c> will dictate that the resulting output be in landscape mode</item>
    ''' </list>
    ''' </summary>
    ''' <param name="Title"></param>
    ''' <param name="MatchColors"></param>
    ''' <param name="OutlineCells"></param>
    ''' <param name="NumberPages"></param>
    ''' <param name="PreviewFirst"></param>
    ''' <param name="LandscapeMode"></param>
    ''' <remarks></remarks>
    Public Sub PrintTheGrid(ByVal Title As String, _
                            ByVal MatchColors As Boolean, _
                            ByVal OutlineCells As Boolean, _
                            ByVal NumberPages As Boolean, _
                            ByVal PreviewFirst As Boolean, _
                            ByVal LandscapeMode As Boolean)

        If _psets Is Nothing Then
            _psets = New System.Drawing.Printing.PageSettings
        End If

        _gridReportMatchColors = MatchColors
        _gridReportNumberPages = NumberPages
        _gridReportOutlineCells = OutlineCells
        _gridReportPreviewFirst = PreviewFirst
        _gridReportOrientLandscape = LandscapeMode
        _gridReportTitle = Title

        Try

            If _psets.PrinterSettings.PrintRange = Printing.PrintRange.AllPages Then
                _gridReportPageNumbers = 1
                _gridReportCurrentrow = 0
                _gridReportCurrentColumn = 0
                _gridReportPrintedOn = Now
            Else
                CalculatePageRange()
                _gridReportPageNumbers = _gridStartPage
                _gridReportCurrentrow = _gridStartPageRow
                _gridReportCurrentColumn = 0
                _gridReportPrintedOn = Now
            End If

            If LandscapeMode Then
                _psets.Landscape = True
            Else
                _psets.Landscape = False
            End If

            If _psets.PrinterSettings.PrinterName <> _OriginalPrinterName Then
                ' We changed the printer invoke the Set default printer via the System.Management class

                Dim moReturn As Management.ManagementObjectCollection

                Dim moSearch As Management.ManagementObjectSearcher

                Dim mo As Management.ManagementObject

                moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Printer")

                moReturn = moSearch.Get

                For Each mo In moReturn
                    Dim objReturn As New Object()
                    Console.WriteLine(mo("Name"))
                    If mo("Name") = _psets.PrinterSettings.PrinterName Then
                        mo.InvokeMethod("SetDefaultPrinter", objReturn)
                    End If
                Next
            End If


            'pdoc.DefaultPageSettings.Landscape = LandscapeMode
            pdoc.DefaultPageSettings = _psets

            If _gridReportPreviewFirst Then

                Dim pview As New PrintPreviewDialog
                pview.Document = pdoc
                pview.WindowState = FormWindowState.Maximized
                pview.ShowDialog()
            Else
                pdoc.Print()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Print Error")
        Finally
            If _psets.PrinterSettings.PrinterName <> _OriginalPrinterName Then
                ' We changed the printer invoke the Set default printer via the System.Management class

                Dim moReturn As Management.ManagementObjectCollection

                Dim moSearch As Management.ManagementObjectSearcher

                Dim mo As Management.ManagementObject

                moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Printer")

                moReturn = moSearch.Get

                For Each mo In moReturn
                    Dim objReturn As New Object
                    ' Console.WriteLine(mo("Name"))
                    If mo("Name") = _OriginalPrinterName Then
                        mo.InvokeMethod("SetDefaultPrinter", objReturn)
                    End If
                Next
            End If
        End Try
    End Sub




#End Region

#Region " Suspending and Resuming Drawing Methods "

    ''' <summary>
    ''' Instructs the grid to stop is continuous redrawing 
    ''' Can be used to speed up population oiperations that are being performed manually
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SuspendGridPaintOperations()
        _Painting = True
    End Sub

    ''' <summary>
    ''' Will resume the grid automatic redrawing operations
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ResumeGridPaintOperations()
        _Painting = False
        Me.Refresh()
    End Sub

#End Region

#End Region

#Region " Event Handlers "

    Private Sub TAIGRIDv2_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RenderGrid(e.Graphics)

        If TearAways.Count <> 0 Then
            Dim t As Integer
            For t = 0 To TearAways.Count - 1
                Dim tear As TearAwayWindowEntry = TearAways.Item(t)
                tear.Winform.SelectedRow = _SelectedRow
            Next
        End If
    End Sub

    Private Sub TAIGRIDv2_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
        ClearToBackgroundColor()
        Me.Invalidate()
    End Sub

    Private Sub hs_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles hs.ValueChanged
        Me.Invalidate()

        If hs.Visible And _EditMode Then
            Dim xxoff As Integer = GimmeXOffset(hs.Value)
            Dim xxxoff As Integer = GimmeXOffset(_EditModeCol)
            Dim xoff As Integer = xxxoff - xxoff

            cmboInput.Left = xoff
            txtInput.Left = xoff

        End If
    End Sub

    Private Sub vs_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vs.ValueChanged

        Dim t As Integer

        Me.Invalidate()

        If TearAways.Count > 0 Then
            ' we have some tear sway windows open so lets set their verticle scrollers
            For t = 0 To TearAways.Count - 1
                Dim ta As TearAwayWindowEntry = TearAways.Item(t)
                ta.SetTearAwayScrollIndex(vs.Value)
            Next
        End If

        If vs.Visible And _EditMode Then

            Dim yyoff As Integer = GimmeYOffset(vs.Value)

            Dim yyyoff As Integer = GimmeYOffset(_EditModeRow)

            Dim yoff As Integer = yyyoff - yyoff

            If _GridTitleVisible And _GridHeaderVisible Then
                yoff += _GridTitleHeight + _GridHeaderHeight

                If yoff < _GridTitleHeight + _GridHeaderHeight Then

                    yoff = -100
                    
                End If
            Else
                If _GridTitleVisible Then
                    yoff += _GridTitleHeight

                    If yoff < _GridTitleHeight Then

                        yoff = -100

                    End If


                Else
                    If GridheaderVisible Then
                        yoff += GridHeaderHeight

                        If yoff < _GridHeaderHeight Then

                            yoff = -100

                        End If

                    End If
                End If
            End If

            cmboInput.Top = yoff
            txtInput.Top = yoff

        End If




    End Sub

    Private Sub MouseEnterHandler(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MouseEnter
        If _AutoFocus Then
            Me.Focus()
        End If
    End Sub

    Private Sub MouseWheelHandler(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseWheel
        Dim v As Integer
        Dim del As Integer = e.Delta

        If del < 0 Then
            del = -_MouseWheelScrollAmount
        Else
            If del > 0 Then
                del = _MouseWheelScrollAmount
            End If
        End If

        If vs.Visible Then
            v = vs.Value - del

            If v < 0 Then v = 0
            If v > _rows Then v = _rows

            '_SelectedRow = -1
            vs.Value = v

        Else
            If hs.Visible Then
                v = hs.Value - del

                If v < 0 Then v = 0
                If v > _cols Then v = _cols

                hs.Value = v

            End If
        End If
    End Sub

    Private Sub MouseUpHandler(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp

        Dim xoff, yoff, r, c, t As Integer
        Dim ss As String

        'Console.WriteLine("MouseUP")

        If _DoubleClickSemaphore Then
            _DoubleClickSemaphore = False
            Exit Sub
        End If

        If Not (_OldContextMenu Is Nothing) And Me.ContextMenu Is Nothing Then
            Me.ContextMenu = _OldContextMenu
        End If


        If _MouseDownOnHeader Then
            ' we were in column resize mode so lets clear all than and blow this pop stand
            ' that should prevent the unwanted selection of the top visible row for folks with
            ' shakey mouse control like yours truely...

            _LastMouseX = -1
            _LastMouseY = -1
            _MouseDownOnHeader = False
            _ColOverOnMouseDown = -1
            _RowOverOnMouseDown = -1
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        If e.Button = MouseButtons.Right Then
            ' bail on a right mousebutton
            Exit Sub
        End If

        Me.txtHandler.Focus()

        yoff = 0

        If _GridTitleVisible Then
            yoff = yoff + _GridTitleHeight
        End If

        If _GridHeaderVisible Then
            yoff = yoff + _GridHeaderHeight
        End If

        If e.Y < yoff Then
            ' we have clicked on the header or the title

            If _GridHeaderVisible Then
                ' have we clicked on the header
                If _GridTitleVisible Then

                    If hs.Visible Then
                        xoff = GimmeXOffset(hs.Value) + e.X
                    Else
                        xoff = e.X
                    End If

                    If e.Y > _GridTitleHeight Then
                        r = 0
                        c = 0
                        For c = 0 To _cols - 1
                            r = r + ColWidth(c)
                            If r > xoff Then
                                ' we got the column
                                If _SelectedColumn = c Then
                                    RaiseEvent ColumnDeSelected(Me, _SelectedColumn)
                                    _SelectedColumn = -1
                                    Me.Invalidate()
                                Else
                                    _SelectedColumn = c
                                    RaiseEvent ColumnSelected(Me, _SelectedColumn)
                                    Me.Invalidate()
                                End If
                                Exit For
                            End If
                        Next
                    Else
                        r = 0
                        c = 0
                        For c = 0 To _cols - 1
                            r = r + ColWidth(c)
                            If r > xoff Then
                                ' we got the column
                                If _SelectedColumn = c Then
                                    RaiseEvent ColumnDeSelected(Me, _SelectedColumn)
                                    _SelectedColumn = -1
                                    Me.Invalidate()
                                Else
                                    _SelectedColumn = c
                                    RaiseEvent ColumnSelected(Me, _SelectedColumn)
                                    Me.Invalidate()
                                End If
                                Exit For
                            End If
                        Next

                    End If

                End If
            End If
            Exit Sub
        End If

        If vs.Visible Then
            yoff = GimmeYOffset(vs.Value) + e.Y
        Else
            yoff = e.Y
        End If

        If _GridTitleVisible Then
            yoff = yoff - _GridTitleHeight
        End If

        If _GridHeaderVisible Then
            yoff = yoff - _GridHeaderHeight
        End If

        If hs.Visible Then
            xoff = GimmeXOffset(hs.Value) + e.X
        Else
            xoff = e.X
        End If

        _RowClicked = -1
        _ColClicked = -1
        r = 0
        c = 0
        If yoff < 0 Or Not _AllowRowSelection Then
            ' we have clicked on the header or the title area so we should skip the row section
        Else
            For r = 0 To _rows - 1
                c = c + _rowheights(r)
                If c > yoff Then
                    ' we got the row
                    _RowClicked = r

                    If Not _AllowMultipleRowSelections Then
                        ' handle like a regular selection 

                        If _SelectedRow > -1 And _SelectedRow <> r Then
                            RaiseEvent RowDeSelected(Me, _SelectedRow)
                        End If

                        'If _SelectedRow = r Then
                        '    '_SelectedRow = -1
                        '    '_SelectedRows.Clear()
                        '    '_SelectedRows.Add(_SelectedRow)
                        '    RaiseEvent RowDeSelected(Me, _SelectedRow)
                        '    Me.Invalidate()
                        '    Exit For
                        'Else
                        _SelectedRow = r
                        _SelectedRows.Clear()
                        _SelectedRows.Add(_SelectedRow)
                        '_SelectedRows.Add(_SelectedRow)
                        RaiseEvent RowSelected(Me, _SelectedRow)
                        Me.Invalidate()
                        Exit For
                        'End If

                    Else

                        If Control.ModifierKeys = Keys.Control Then

                            If _SelectedRows.Contains(r) And Control.ModifierKeys = Keys.Control Then
                                ' we need to de select that row here
                                If _SelectedRow = r Then
                                    _SelectedRow = -1
                                End If
                                _SelectedRows.Remove(r)
                                RaiseEvent RowDeSelected(Me, r)
                                Me.Invalidate()
                                Exit For
                            Else
                                If Control.ModifierKeys = Keys.Control Then
                                    _SelectedRow = r
                                    _SelectedRows.Add(r)
                                    RaiseEvent RowSelected(Me, _SelectedRow)
                                    Me.Invalidate()
                                    Exit For
                                Else
                                    If _SelectedRow = r Then
                                        RaiseEvent RowDeSelected(Me, _SelectedRow)
                                        _SelectedRow = -1
                                        _SelectedRows.Clear()
                                        Me.Invalidate()
                                        Exit For
                                    Else
                                        _SelectedRow = r
                                        _SelectedRows.Clear()
                                        _SelectedRows.Add(_SelectedRow)
                                        RaiseEvent RowSelected(Me, _SelectedRow)
                                        Me.Invalidate()
                                        Exit For
                                    End If
                                End If
                            End If

                        Else
                            If Control.ModifierKeys = Keys.Shift Then

                                If _ShiftMultiSelectSelectedRowCrap > -1 Then
                                    ' we have a row selected already

                                    _SelectedRows.Clear()

                                    If _ShiftMultiSelectSelectedRowCrap > r Then

                                        For t = r To _ShiftMultiSelectSelectedRowCrap
                                            _SelectedRows.Add(t)
                                        Next

                                        _SelectedRow = r
                                        RaiseEvent RowSelected(Me, _SelectedRow)
                                        Me.Invalidate()
                                        Exit For

                                    Else

                                        For t = _ShiftMultiSelectSelectedRowCrap To r
                                            _SelectedRows.Add(t)
                                        Next

                                        _SelectedRow = r
                                        RaiseEvent RowSelected(Me, _SelectedRow)
                                        Me.Invalidate()
                                        Exit For

                                    End If

                                Else
                                    ' we dont have a selectedrow already so lets haandle this like a ragular selection
                                    If _SelectedRow > -1 And _SelectedRow <> r Then
                                        RaiseEvent RowDeSelected(Me, _SelectedRow)
                                    Else
                                        If SelectedRow = -1 Then
                                            _ShiftMultiSelectSelectedRowCrap = r
                                        End If
                                    End If

                                    If _SelectedRow = r Then
                                        _SelectedRow = -1
                                        _ShiftMultiSelectSelectedRowCrap = -1
                                        _SelectedRows.Clear()
                                        '_SelectedRows.Add(_SelectedRow)
                                        RaiseEvent RowDeSelected(Me, _SelectedRow)
                                        Me.Invalidate()
                                        Exit For
                                    Else
                                        _SelectedRow = r
                                        _ShiftMultiSelectSelectedRowCrap = r
                                        _SelectedRows.Clear()
                                        _SelectedRows.Add(_SelectedRow)
                                        RaiseEvent RowSelected(Me, _SelectedRow)
                                        Me.Invalidate()
                                        Exit For
                                    End If
                                End If
                            Else
                                ' handle like a regular selection
                                If _SelectedRow > -1 And _SelectedRow <> r Then
                                    RaiseEvent RowDeSelected(Me, _SelectedRow)
                                Else
                                    If SelectedRow = -1 Then
                                        _ShiftMultiSelectSelectedRowCrap = r
                                    End If
                                End If

                                'If _SelectedRow = r Then
                                '    _SelectedRow = -1
                                '    _ShiftMultiSelectSelectedRowCrap = -1
                                '    _SelectedRows.Clear()
                                '    '_SelectedRows.Add(_SelectedRow)
                                '    RaiseEvent RowDeSelected(Me, _SelectedRow)
                                '    Me.Invalidate()
                                '    Exit For
                                'Else
                                _SelectedRow = r
                                _ShiftMultiSelectSelectedRowCrap = r
                                _SelectedRows.Clear()
                                _SelectedRows.Add(_SelectedRow)
                                RaiseEvent RowSelected(Me, _SelectedRow)
                                Me.Invalidate()
                                Exit For
                                'End If
                            End If
                        End If
                    End If
                End If
            Next
        End If

        If _RowClicked = -1 Then
            ' we did not click on a row so we should bail
            Exit Sub
        End If

        r = 0
        c = 0
        For c = 0 To _cols - 1
            r = r + ColWidth(c)
            If r > xoff Then
                ' we got the column
                _ColClicked = c
                Exit For
            End If
        Next

        If _ColClicked = -1 Then
            ' we did not click on a a column so lets bail
            Exit Sub
        End If

        If Me.Visible Then
            RaiseEvent CellClicked(Me, _RowClicked, _ColClicked)
            If _RowClicked > -1 And _RowClicked < _rows Then
                If _ColClicked > -1 And _ColClicked < _cols And _colEditable(_ColClicked) And _AllowInGridEdits Then

                    If IsColumnRestricted(_ColClicked) Then

                        Dim it As EditColumnRestrictor = GetColumnRestriction(_ColClicked)

                        cmboInput.Items.Clear()

                        Dim s() As String = it.RestrictedList.Split("^".ToCharArray)
                        For Each ss1 As String In s
                            cmboInput.Items.Add(ss1)
                        Next

                        ' we have selected a row and col lets move the txtinput there and bring it to the front
                        xoff = 0
                        yoff = 0

                        If _RowClicked > 0 Then
                            For r = 0 To _RowClicked - 1
                                yoff = yoff + RowHeight(r)
                            Next
                        End If

                        If GridheaderVisible Then
                            yoff = yoff + _GridHeaderHeight
                        End If

                        If _GridTitleVisible Then
                            yoff = yoff + _GridTitleHeight
                        End If

                        If _ColClicked > 0 Then
                            For c = 0 To _ColClicked - 1
                                xoff = xoff + ColWidth(c)
                            Next
                        End If

                        If vs.Visible And vs.Value > 0 Then
                            yoff = yoff - GimmeYOffset(vs.Value)
                        End If

                        If hs.Visible And hs.Value > 0 Then
                            xoff = xoff - GimmeXOffset(hs.Value)
                        End If

                        If _CellOutlines Then
                            cmboInput.Top = yoff + 1
                            cmboInput.Left = xoff + 1
                            cmboInput.Width = ColWidth(_ColClicked) - 1
                            cmboInput.Height = RowHeight(_RowClicked) - 2
                            cmboInput.BackColor = _colEditableTextBackColor
                        Else
                            cmboInput.Top = yoff
                            cmboInput.Left = xoff
                            cmboInput.Width = ColWidth(_ColClicked)
                            cmboInput.Height = RowHeight(_RowClicked)
                            cmboInput.BackColor = _colEditableTextBackColor
                        End If

                        cmboInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                        cmboInput.Text = _grid(_RowClicked, _ColClicked)

                        cmboInput.Visible = True
                        cmboInput.BringToFront()
                        cmboInput.DroppedDown = True
                        _EditModeCol = _ColClicked
                        _EditModeRow = _RowClicked
                        _EditMode = True

                        cmboInput.Focus()
                    Else
                        If _colboolean(_ColClicked) Then
                            ' we have clicked on a boolean editable cell lets flip those bits baby

                            ss = Trim(UCase(_grid(_RowClicked, _ColClicked)))

                            Select Case ss

                                Case "TRUE"
                                    ss = "FALSE"
                                Case "FALSE"
                                    ss = "TRUE"
                                Case "YES"
                                    ss = "NO"
                                Case "NO"
                                    ss = "YES"
                                Case "1"
                                    ss = "0"
                                Case "0"
                                    ss = "1"
                                Case "Y"
                                    ss = "N"
                                Case "N"
                                    ss = "Y"
                                Case Else
                                    ss = ""

                            End Select

                            _grid(_RowClicked, _ColClicked) = ss

                            Me.Refresh()

                        Else
                            ' we have selected a row and col lets move the txtinput there and bring it to the front
                            xoff = 0
                            yoff = 0

                            If _RowClicked > 0 Then
                                For r = 0 To _RowClicked - 1
                                    yoff = yoff + RowHeight(r)
                                Next
                            End If

                            If GridheaderVisible Then
                                yoff = yoff + _GridHeaderHeight
                            End If

                            If _GridTitleVisible Then
                                yoff = yoff + _GridTitleHeight
                            End If

                            If _ColClicked > 0 Then
                                For c = 0 To _ColClicked - 1
                                    xoff = xoff + ColWidth(c)
                                Next
                            End If

                            If vs.Visible And vs.Value > 0 Then
                                yoff = yoff - GimmeYOffset(vs.Value)
                            End If

                            If hs.Visible And hs.Value > 0 Then
                                xoff = xoff - GimmeXOffset(hs.Value)
                            End If

                            If _CellOutlines Then
                                txtInput.Top = yoff + 1
                                txtInput.Left = xoff + 1
                                txtInput.Width = ColWidth(_ColClicked) - 1
                                txtInput.Height = RowHeight(_RowClicked) - 2
                                txtInput.BackColor = _colEditableTextBackColor
                            Else
                                txtInput.Top = yoff
                                txtInput.Left = xoff
                                txtInput.Width = ColWidth(_ColClicked)
                                txtInput.Height = RowHeight(_RowClicked)
                                txtInput.BackColor = _colEditableTextBackColor
                            End If

                            txtInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                            txtInput.Text = _grid(_RowClicked, _ColClicked)

                            txtInput.Visible = True
                            txtInput.BringToFront()
                            _EditModeCol = _ColClicked
                            _EditModeRow = _RowClicked
                            _EditMode = True

                            txtInput.Focus()
                        End If
                    End If

                End If
            End If
        End If

    End Sub

    Private Function IsColumnRestricted(ByVal colid As Integer) As Boolean
        Dim ret As Boolean = False

        For Each it As EditColumnRestrictor In _colEditRestrictions
            If it.ColumnID = colid Then
                ret = True
                Exit For
            End If
        Next

        Return ret
    End Function

    Private Function GetColumnRestriction(ByVal colid As Integer) As EditColumnRestrictor
        Dim ret As New EditColumnRestrictor()

        For Each it As EditColumnRestrictor In _colEditRestrictions
            If it.ColumnID = colid Then
                ret = it
                Exit For
            End If
        Next

        Return ret

    End Function

    Private Sub DoubleClickHandler(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DoubleClick

        'Console.WriteLine("MouseDoubleClick")
        If _RowClicked <> -1 And _ColClicked <> -1 And Me.Visible Then
            RaiseEvent CellDoubleClicked(Me, _RowClicked, _ColClicked)
            _DoubleClickSemaphore = True
        End If
    End Sub

    Private Sub txtHandler_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHandler.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtHandler_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHandler.KeyDown

        ' we are NOT in editmode so lets handle this like any other

        RaiseEvent KeyPressedInGrid(Me, e.KeyCode)

        If e.KeyCode = Keys.Return Or e.KeyCode = Keys.Enter Then
            If _SelectedRow <> -1 And Me.Visible Then
                RaiseEvent CellDoubleClicked(Me, _SelectedRow, 0)
            End If
        End If

        If e.KeyCode = Keys.Left Then
            If hs.Visible Then
                Dim v As Integer
                v = hs.Value
                v -= 1
                If v < 0 Then v = 0
                hs.Value = v
            End If
        End If

        If e.KeyCode = Keys.Right Then
            If hs.Visible Then
                Dim v As Integer
                v = hs.Value
                v += 1
                If v >= hs.Maximum Then v = hs.Maximum - 1
                hs.Value = v
            End If
        End If

        If e.KeyCode = Keys.PageDown Then

            RaiseEvent RowDeSelected(Me, _SelectedRow)

            _SelectedRow = _SelectedRow + 10

            If _SelectedRow >= _rows Then
                _SelectedRow = _rows - 1
            End If

            _SelectedRows.Clear()
            _SelectedRows.Add(_SelectedRow)

            If vs.Visible Then
                Dim flag As Boolean = True
                Dim x, xx As Integer
                Do While flag
                    x = GimmeYOffset(_SelectedRow)
                    x = x - GimmeYOffset(vs.Value)

                    If _GridTitleVisible Then
                        xx = _GridTitleHeight
                    Else
                        xx = 0
                    End If

                    If _GridHeaderVisible Then
                        xx = xx + _GridHeaderHeight
                    End If

                    If hs.Visible Then
                        xx = xx + hs.Height
                    End If

                    xx = xx + _rowheights(_SelectedRow)

                    If x < (Me.Height - xx) Then
                        flag = False
                    Else
                        vs.Value = vs.Value + 1
                    End If

                Loop
            End If

            Me.Invalidate()

            RaiseEvent RowSelected(Me, _SelectedRow)

        End If


        If e.KeyCode = Keys.Down Then

            RaiseEvent RowDeSelected(Me, _SelectedRow)

            _SelectedRow = _SelectedRow + 1

            If _SelectedRow >= _rows Then
                _SelectedRow = _rows - 1
            End If

            _SelectedRows.Clear()
            _SelectedRows.Add(_SelectedRow)

            If vs.Visible Then
                Dim flag As Boolean = True
                Dim x, xx As Integer
                Do While flag
                    x = GimmeYOffset(_SelectedRow)
                    x = x - GimmeYOffset(vs.Value)

                    If _GridTitleVisible Then
                        xx = _GridTitleHeight
                    Else
                        xx = 0
                    End If

                    If _GridHeaderVisible Then
                        xx = xx + _GridHeaderHeight
                    End If

                    If hs.Visible Then
                        xx = xx + hs.Height
                    End If

                    xx = xx + _rowheights(_SelectedRow)

                    If x < (Me.Height - xx) Then
                        flag = False
                    Else
                        vs.Value = vs.Value + 1
                    End If

                Loop
            End If

            Me.Invalidate()

            RaiseEvent RowSelected(Me, _SelectedRow)

        End If

        If e.KeyCode = Keys.PageUp Then

            RaiseEvent RowDeSelected(Me, _SelectedRow)

            _SelectedRow = _SelectedRow - 10
            If _SelectedRow < 0 Then
                _SelectedRow = 0
            End If

            _SelectedRows.Clear()
            _SelectedRows.Add(_SelectedRow)

            If vs.Visible Then
                Dim flag As Boolean = True
                Do While flag

                    If _SelectedRow >= vs.Value Then
                        flag = False
                    Else
                        vs.Value = vs.Value - 1
                    End If
                Loop
            End If

            Me.Invalidate()

            RaiseEvent RowSelected(Me, _SelectedRow)

        End If

        If e.KeyCode = Keys.Up Then

            RaiseEvent RowDeSelected(Me, _SelectedRow)

            _SelectedRow = _SelectedRow - 1

            If _SelectedRow < 0 Then
                _SelectedRow = 0
            End If

            _SelectedRows.Clear()
            _SelectedRows.Add(_SelectedRow)

            If vs.Visible Then
                Dim flag As Boolean = True
                Do While flag

                    If _SelectedRow >= vs.Value Then
                        flag = False
                    Else
                        vs.Value = vs.Value - 1
                    End If

                Loop
            End If

            Me.Invalidate()

            RaiseEvent RowSelected(Me, _SelectedRow)

        End If

        If e.Modifiers = Keys.Control Then
            If e.KeyCode = Keys.F Then
                If _LastSearchText <> "" And _LastSearchColumn <> -1 And _SelectedRow <> -1 Then
                    Dim t As Integer

                    If _SelectedRow + 1 >= _rows Then
                        ' we are on the last row already flip to the first row
                        _SelectedRow = 0
                    End If

                    Dim found As Boolean = False
                    For t = _SelectedRow + 1 To _rows - 1
                        If InStr(UCase(_grid(t, _LastSearchColumn)), UCase(_LastSearchText), CompareMethod.Text) <> 0 Then
                            ' we have a match
                            If vs.Visible Then
                                vs.Value = t
                                found = True
                                _SelectedRow = t
                                Me.Invalidate()
                                e.Handled = True
                                Exit For
                            Else
                                found = True
                                _SelectedRow = t
                                Me.Invalidate()
                                e.Handled = True
                                Exit For
                            End If
                        End If
                    Next
                    If Not found Then
                        e.Handled = True
                        MsgBox("Search found nothing further")
                    End If
                End If
            End If

        End If

    End Sub

    Private Sub vs_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles vs.Scroll
        '_SelectedRow = -1
    End Sub

    Private Sub MouseDownHandler(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        Dim p As Point
        Dim x, y, xx, yy, r, c, rr, cc, yoff As Integer


        'Console.WriteLine("MouseDOWN")

        xx = -1
        yy = -1

        ' handler for leftmousebuttons on column resizing
        If e.Button = MouseButtons.Left And _AllowUserColumnResizing And _GridHeaderVisible Then

            p = Me.PointToClient(Control.MousePosition)
            x = p.X '+ Math.Abs(Me.Left)
            y = p.Y '+ Math.Abs(Me.Top)

            ' is the title visible
            If _GridTitleVisible Then
                rr = _GridTitleHeight
            Else
                rr = 0
            End If

            ' adjust for the hs scroll bar
            If hs.Visible Then
                x = x + GimmeXOffset(hs.Value)
            End If

            ' first of all are we actually in the header

            If y >= rr And y <= rr + _GridHeaderHeight Then
                ' yes we are lets setup for resizing

                cc = 0

                For c = 0 To _cols - 1
                    If cc <= x And cc + _colwidths(c) >= x Then
                        ' we have the column
                        xx = c
                        Exit For
                    Else
                        cc = cc + _colwidths(c)
                    End If
                Next

                _ColOverOnMenuButton = -1
                _ColOverOnMouseDown = xx

                _MouseDownOnHeader = True

                _LastMouseX = Control.MousePosition.X
                _LastMouseY = Control.MousePosition.Y
                _AutosizeCellsToContents = False
                Me.Cursor = System.Windows.Forms.Cursors.SizeWE
            Else
                ' No we aren't so lets NOT resize

                _MouseDownOnHeader = False
                _LastMouseX = -1
                _LastMouseY = -1
                _ColOverOnMenuButton = -1
                _ColOverOnMouseDown = -1

            End If

        End If

        ' handler for rightmousebuttons and popupmenus and allowing / disallowing ctrl key menus

        If e.Button = MouseButtons.Right And _
            (_AllowPopupMenu Or (Control.ModifierKeys = Keys.Control And _AllowControlKeyMenuPopup)) Then

            If Not (Me.ContextMenu Is Nothing) Then
                _OldContextMenu = Me.ContextMenu
                'Me.ContextMenu = Nothing
            End If

            p = Me.PointToClient(Control.MousePosition)
            x = p.X + Math.Abs(Me.Left)
            y = p.Y + Math.Abs(Me.Top)

            If vs.Visible Then
                yoff = GimmeYOffset(vs.Value) + e.Y
            Else
                yoff = e.Y
            End If

            If _GridTitleVisible Then
                yoff = yoff - _GridTitleHeight
            End If

            If _GridHeaderVisible Then
                yoff = yoff - _GridHeaderHeight
            End If

            If yoff < 0 Then
                ' we have clicked on the header or the title area so we should skip the row section
                _RowOverOnMenuButton = -1
            Else
                For r = 0 To _rows - 1
                    c = c + _rowheights(r)
                    If c > yoff Then
                        ' we got the row
                        _RowOverOnMenuButton = r
                        Exit For
                    End If
                Next
            End If

            ' adjust for the hs scroll bar
            If hs.Visible Then
                x = x + GimmeXOffset(hs.Value)
            End If

            ' adjust for the vs scroll bar
            If vs.Visible Then
                y = y + GimmeYOffset(vs.Value)
            End If

            cc = Math.Abs(Me.Left)

            For c = 0 To _cols - 1
                If cc <= x And cc + _colwidths(c) >= x Then
                    ' we have the column
                    xx = c
                    Exit For
                Else
                    cc = cc + _colwidths(c)
                End If
            Next

            'If _GridTitleVisible Then
            '    If _GridHeaderVisible Then
            '        ' we have the header and the titlevisible

            '    End If
            'End If

            _ColOverOnMenuButton = xx

            If Me.ContextMenu Is Nothing Then

                miStats.Text = "Rows = " + _rows.ToString() + " : Cols = " + _cols.ToString()

                menu.Show(Me, p)
            Else
                Me.ContextMenu.Show(Me, p)
            End If
            'menu.Show(Me, p)
            Exit Sub
        Else
            If e.Button = MouseButtons.Right Then
                RaiseEvent RightMouseButtonInGrid(Me)
                Exit Sub
            End If
        End If
    End Sub

    Private Sub MouseMoveHandler(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove

        Dim x, delta As Integer

        If _LMouseX = e.X And _LMouseY = e.Y Then
            Exit Sub
        Else
            _LMouseX = e.X
            _LMouseY = e.Y
        End If

        If _MouseDownOnHeader And _ColOverOnMouseDown > -1 And _AllowUserColumnResizing And _ColOverOnMouseDown < _cols Then
            x = Control.MousePosition.X

            ' calculate Deltas
            If x >= _LastMouseX Then
                delta = x - _LastMouseX
            Else
                delta = -(_LastMouseX - x)
            End If
            _LastMouseX = x

            _colwidths(_ColOverOnMouseDown) = _colwidths(_ColOverOnMouseDown) + delta

            If _colwidths(_ColOverOnMouseDown) < _UserColResizeMinimum Then
                _colwidths(_ColOverOnMouseDown) = _UserColResizeMinimum
            End If

            RaiseEvent ColumnResized(Me, _ColOverOnMouseDown)

            _AutoSizeAlreadyCalculated = False

            Me.Invalidate()

        Else

            MouseHoverHandler(sender, New System.EventArgs)

        End If

    End Sub

    Private Sub MouseHoverHandler(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim mp As Point = PointToClient(Control.MousePosition)

        Dim xoff, yoff, r, c, row, col As Integer

        If _GridTitleVisible And mp.Y <= _GridTitleHeight Then
            ' we are hovering over the grid title
            RaiseEvent GridHoverleave(Me)
            Exit Sub
        End If

        If _GridTitleVisible And _GridHeaderVisible And mp.Y <= (_GridTitleHeight + _GridHeaderHeight) Then
            ' we are hovering over the Header or the title so lets bail
            RaiseEvent GridHoverleave(Me)
            Exit Sub
        End If

        If vs.Visible Then
            yoff = GimmeYOffset(vs.Value) + mp.Y
        Else
            yoff = mp.Y
        End If

        If _GridTitleVisible Then
            yoff = yoff - _GridTitleHeight
        End If

        If _GridHeaderVisible Then
            yoff = yoff - _GridHeaderHeight
        End If

        If hs.Visible Then
            xoff = GimmeXOffset(hs.Value) + mp.X
        Else
            xoff = mp.X
        End If

        ' here xoff and yoff are converted to real grid coordinates if the exist

        r = 0
        c = 0
        row = -1

        For r = 0 To _rows - 1
            c = c + _rowheights(r)
            If c > yoff Then
                ' we got the row
                row = r
                Exit For
            End If
        Next

        r = 0
        c = 0
        col = -1
        For c = 0 To _cols - 1
            r = r + ColWidth(c)
            If r > xoff Then
                ' we got the column
                col = c
                Exit For
            End If
        Next

        If row > -1 And col > -1 Then
            ' we gots a winner

            RaiseEvent GridHover(Me, row, col, _grid(row, col))
        Else
            ' we gots a loser
            RaiseEvent GridHoverleave(Me)

        End If

    End Sub

    Private Sub TAIGRIDControl_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        _gridCellFontsList(0) = _DefaultCellFont
        _gridForeColorList(0) = New Pen(_DefaultForeColor)
        _gridCellAlignmentList(0) = _DefaultStringFormat
        _gridBackColorList(0) = New SolidBrush(_DefaultBackColor)

    End Sub

    Private Sub pdoc_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles pdoc.PrintPage

        Dim x, y, xx, r, c As Integer
        Dim fnt As New System.Drawing.Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Regular, GraphicsUnit.Pixel)
        Dim fnt2 As New System.Drawing.Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Bold, GraphicsUnit.Pixel)

        Dim tfnt As New System.Drawing.Font("Courier New", 8, FontStyle.Regular, GraphicsUnit.Pixel)


        Dim m As Single

        Dim ft As Font

        Dim greypen As New Pen(Color.Gray)

        Dim pagewidth As Integer = e.PageSettings.Bounds.Size.Width
        Dim pageheight As Integer = e.PageSettings.Bounds.Size.Height

        Dim lrmargin As Integer = 40
        Dim tbmargin As Integer = 70

        Dim colprintedonpage As Boolean = False

        If (AllColWidths() * _gridReportScaleFactor) < pagewidth - (2 * lrmargin) Then
            xx = ((pagewidth - (2 * lrmargin)) - (AllColWidths() * _gridReportScaleFactor)) / 2
        Else
            xx = 0
        End If


        Dim rect As New RectangleF(0, 0, 1, 1)

        x = lrmargin
        y = tbmargin

        Dim coloffset As Integer = 0
        Dim morecols As Boolean = True
        Dim currow As Integer = _gridReportCurrentrow


        ft = _GridHeaderFont

        ft = New Font(_GridHeaderFont.FontFamily, _
                        (_GridHeaderFont.SizeInPoints - 1) * _gridReportScaleFactor, _
                        _GridHeaderFont.Style, _GridHeaderFont.Unit)

        ' calculate size and place the printed on date on the page

        m = e.Graphics.MeasureString(_gridReportPrintedOn.ToLongDateString + vbCrLf _
                                    + _gridReportPrintedOn.ToLongTimeString, fnt).Width

        e.Graphics.DrawString(_gridReportPrintedOn.ToLongDateString + vbCrLf + _
                            _gridReportPrintedOn.ToLongTimeString, fnt, Brushes.Black, _
                            pagewidth - m - lrmargin, tbmargin / 2)

        If _gridReportTitle <> "" Then
            ' we want to title each page here

            Dim ttit As String = _gridReportTitle

            If ttit.Length > 98 Then
                Dim ttitarray As String() = ttit.Split(" ".ToCharArray())

                Dim ttitidx, curlen As Integer
                ttit = ""
                curlen = 0

                For ttitidx = 0 To ttitarray.GetUpperBound(0)
                    ttit += ttitarray(ttitidx) + " "

                    curlen += ttitarray(ttitidx).Length

                    If curlen > 98 Then
                        ttit += vbCrLf
                        curlen = 0
                    End If
                Next
            End If

            e.Graphics.DrawString(ttit, tfnt, Brushes.Black, lrmargin, tbmargin / 2)
        Else

            Dim ttit As String = _GridTitle

            If ttit.Length > 98 Then
                Dim ttitarray As String() = ttit.Split(" ".ToCharArray())

                Dim ttitidx, curlen As Integer
                ttit = ""
                curlen = 0

                For ttitidx = 0 To ttitarray.GetUpperBound(0)
                    ttit += ttitarray(ttitidx) + " "

                    curlen += ttitarray(ttitidx).Length

                    If curlen > 98 Then
                        ttit += vbCrLf
                        curlen = 0
                    End If
                Next
            End If

            e.Graphics.DrawString(ttit, tfnt, Brushes.Black, lrmargin, tbmargin / 2)

        End If

        If _gridReportNumberPages Then
            ' we want to number the pages here

            m = e.Graphics.MeasureString("Page " + _gridReportPageNumbers.ToString(), fnt).Height


            e.Graphics.DrawString("Page " + _gridReportPageNumbers.ToString(), fnt, _
                                    Brushes.Black, lrmargin, pageheight - tbmargin + m + 2)

        End If

        ' print the grid header

        For c = _gridReportCurrentColumn To Me.Cols - 1

            If x + _colwidths(c) + xx > pagewidth - lrmargin And colprintedonpage Then
                Exit For
            End If

            colprintedonpage = True

            If _gridReportMatchColors Then
                e.Graphics.FillRectangle(New SolidBrush(_GridHeaderBackcolor), x + xx, y, _colwidths(c), _GridHeaderHeight)
            End If

            rect.X = Convert.ToSingle(x + xx)
            rect.Y = Convert.ToSingle(y)
            rect.Width = Convert.ToSingle(_colwidths(c))
            rect.Height = Convert.ToSingle(_GridHeaderHeight)



            e.Graphics.DrawString(_GridHeader(c), ft, Brushes.Black, rect, _GridHeaderStringFormat)

            If _gridReportOutlineCells Then
                e.Graphics.DrawRectangle(greypen, x + xx, y, _colwidths(c), _GridHeaderHeight)
            End If

            x = x + _colwidths(c)
        Next


        y += _GridHeaderHeight
        x = lrmargin

        For r = _gridReportCurrentrow To Me.Rows - 1
            For c = _gridReportCurrentColumn To Me.Cols - 1
                If x + _colwidths(c) + xx > pagewidth - lrmargin And colprintedonpage Then
                    coloffset = c
                    morecols = True
                    Exit For
                Else
                    morecols = False
                End If

                colprintedonpage = True

                If _gridReportMatchColors Then
                    e.Graphics.FillRectangle(_gridBackColorList(_gridBackColor(r, c)), x + xx, y, _colwidths(c), _rowheights(r))
                End If

                rect.X = Convert.ToSingle(x + xx)
                rect.Y = Convert.ToSingle(y)
                rect.Width = Convert.ToSingle(_colwidths(c))
                rect.Height = Convert.ToSingle(_rowheights(r))

                ft = New Font(_gridCellFontsList(_gridCellFonts(r, c)).FontFamily, _
                              _gridCellFontsList(_gridCellFonts(r, c)).SizeInPoints - 1, _
                              _gridCellFontsList(_gridCellFonts(r, c)).Style, _
                              _gridCellFontsList(_gridCellFonts(r, c)).Unit)

                e.Graphics.DrawString(_grid(r, c), ft, _
                                      Brushes.Black, rect, _gridCellAlignmentList(_gridCellAlignment(r, c)))

                If _gridReportOutlineCells Then
                    e.Graphics.DrawRectangle(greypen, x + xx, y, _colwidths(c), _rowheights(r))
                End If

                x = x + _colwidths(c)

            Next
            x = lrmargin
            y += _rowheights(r)
            _gridReportCurrentrow += 1

            ' do we need to skip to next page here
            If y >= pageheight - tbmargin Then
                Exit For
            Else
                ' nope
            End If
        Next


        If _psets.PrinterSettings.PrintRange = Printing.PrintRange.SomePages Then

            If (_gridReportCurrentrow >= Me.Rows - 1 And Not morecols) Or _
                (_gridReportPageNumbers >= _gridEndPage) Then
                e.HasMorePages = False
                _gridReportPageNumbers = 1
                _gridReportCurrentrow = 0
                _gridReportCurrentColumn = 0
            Else

                If morecols Then
                    _gridReportCurrentColumn = coloffset
                    _gridReportCurrentrow = currow
                Else
                    _gridReportCurrentColumn = 0
                End If
                e.HasMorePages = True
                _gridReportPageNumbers += 1

            End If

        Else

            If _gridReportCurrentrow >= Me.Rows - 1 And Not morecols Then
                e.HasMorePages = False
                _gridReportPageNumbers = 1
                _gridReportCurrentrow = 0
                _gridReportCurrentColumn = 0
            Else

                If morecols Then
                    _gridReportCurrentColumn = coloffset
                    _gridReportCurrentrow = currow
                Else
                    _gridReportCurrentColumn = 0
                End If
                e.HasMorePages = True
                _gridReportPageNumbers += 1
            End If

        End If



    End Sub

    Private Sub PageOrientationChange(ByVal lsorientation As Boolean) Handles _PageSetupForm.OrientationChanged

        Dim oldorientation As Boolean = _psets.Landscape

        _psets.Landscape = lsorientation


        _PageSetupForm.MaxPage = CalculatePageRange()

        _psets.Landscape = oldorientation

    End Sub

    Private Sub PageSetupChange(ByVal psiz As System.Drawing.Printing.PaperSize) Handles _PageSetupForm.PageSizeChanged

        Try

            Dim ps As System.Drawing.Printing.PaperSize = _psets.PaperSize

            _psets.PaperSize = psiz

            _PageSetupForm.MaxPage = CalculatePageRange()

            _psets.PaperSize = ps

        Catch ex As Exception

            ' we might want to do something here



        End Try


    End Sub

    Private Sub PageMetricsChange(ByVal psiz As System.Drawing.Printing.PaperSize, ByVal lsorientation As Boolean) Handles _PageSetupForm.PaperMetricsHaveChanged

        Try

            LogThis("Inside the event handler for Page Meterics changing...")

            Dim ps As System.Drawing.Printing.PaperSize = _psets.PaperSize


            _psets.PaperSize = psiz
            _psets.Landscape = lsorientation
            _PageSetupForm.MaxPage = CalculatePageRange()
            _psets.PaperSize = ps

        Catch ex As Exception
            ' don't do a thing just exit with grace
        End Try

    End Sub

    Private Sub txtInput_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInput.Leave

        If _GridEditMode = GridEditModes.LostFocus Then

            If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
                Dim oldval As String = _grid(_RowClicked, _ColClicked)
                Dim newval As String = txtInput.Text
                _grid(_RowClicked, _ColClicked) = txtInput.Text
                RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            End If

            txtInput.SendToBack()
            txtInput.Visible = False

            Me.Invalidate()
            'e.Handled = False

        Else
            txtInput.Visible = False
            _EditModeCol = -1
            _EditModeRow = -1
            _EditMode = False
        End If

       
    End Sub

    Private Sub txtInput_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInput.KeyDown
        If e.KeyCode = Keys.Return And _GridEditMode = GridEditModes.KeyReturn Then

            If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
                Dim oldval As String = _grid(_RowClicked, _ColClicked)
                Dim newval As String = txtInput.Text
                _grid(_RowClicked, _ColClicked) = txtInput.Text
                RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            End If

            txtInput.SendToBack()
            txtInput.Visible = False

            Me.Invalidate()
            e.Handled = False

        End If

        If e.KeyCode = Keys.Tab Then
            Console.WriteLine("Tab Key Pressed")
            e.Handled = True
        End If
    End Sub

    Private Sub cmboInput_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmboInput.Leave

        If GridEditMode = GridEditModes.KeyReturn Then

            If _grid(_RowClicked, _ColClicked) <> cmboInput.Text Then
                Dim oldval As String = _grid(_RowClicked, _ColClicked)
                Dim newval As String = cmboInput.Text
                _grid(_RowClicked, _ColClicked) = cmboInput.Text
                RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            End If

            cmboInput.SendToBack()
            cmboInput.Visible = False

            Me.Invalidate()
            ' e.Handled = False


        Else
            cmboInput.Visible = False
            _EditModeCol = -1
            _EditModeRow = -1
            _EditMode = False
        End If

       
    End Sub

    Private Sub cmboInput_keyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmboInput.KeyDown
        If e.KeyCode = Keys.Return And _GridEditMode = GridEditModes.KeyReturn Then

            If _grid(_RowClicked, _ColClicked) <> cmboInput.Text Then
                Dim oldval As String = _grid(_RowClicked, _ColClicked)
                Dim newval As String = cmboInput.Text
                _grid(_RowClicked, _ColClicked) = cmboInput.Text
                RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            End If

            cmboInput.SendToBack()
            cmboInput.Visible = False

            Me.Invalidate()
            e.Handled = False

        End If
    End Sub

    Private Sub cmboInput_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmboInput.SelectedIndexChanged

        If _grid(_RowClicked, _ColClicked) <> cmboInput.Text Then
            Dim oldval As String = _grid(_RowClicked, _ColClicked)
            Dim newval As String = cmboInput.Text
            _grid(_RowClicked, _ColClicked) = cmboInput.Text
            RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
        End If

        cmboInput.SendToBack()
        cmboInput.Visible = False

        Me.Invalidate()

    End Sub

    Private Sub TAIGridControl_HandleDestroyed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.HandleDestroyed
        ' implemented to to destroy all tearaways if the parent grid gets destroyed

        KillAllTearAwayColumnWindows()

    End Sub

#Region " Popup Menu Handlers "

    Private Sub menu_Popup(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menu.Popup
        If Me.Antialias Then
            miSmoothing.Checked = True
        Else
            miSmoothing.Checked = False
        End If

        If Me.ExcelAutoFitColumn Then
            miAutoFitCols.Checked = True
        Else
            miAutoFitCols.Checked = False
        End If

        If Me.ExcelAutoFitRow Then
            miAutoFitRows.Checked = True
        Else
            miAutoFitRows.Checked = False
        End If

        If Me.ExcelUseAlternateRowColor Then
            Me.miALternateRowColors.Checked = True
        Else
            Me.miALternateRowColors.Checked = False
        End If

        If Me.ExcelOutlineCells Then
            Me.miOutlineExportedCells.Checked = True
        Else
            Me.miOutlineExportedCells.Checked = False
        End If

        If Me.ExcelMatchGridColorScheme Then
            Me.miMatchGridColors.Checked = True
        Else
            Me.miMatchGridColors.Checked = False
        End If

        If _ColOverOnMenuButton <> -1 Then
            miSearchInColumn.Enabled = True
        Else
            miSearchInColumn.Enabled = False
        End If

        If _AllowUserColumnResizing Then
            miAllowUserColumnResizing.Checked = True
        Else
            miAllowUserColumnResizing.Checked = False
        End If


        miExportToExcelMenu.Enabled = _AllowExcelFunctionality
        miTearColumnAway.Enabled = _AllowTearAwayFuncionality
        miArrangeTearAways.Enabled = _AllowTearAwayFuncionality
        miHideAllTearAwayColumns.Enabled = _AllowTearAwayFuncionality
        miHideColumnTearAway.Enabled = _AllowTearAwayFuncionality
        miMultipleColumnTearAway.Enabled = _AllowTearAwayFuncionality

        miExportToTextFile.Enabled = _AllowTextFunctionality
        miExportToHTMLTable.Enabled = _AllowHTMLFunctionality
        miExportToSQLScript.Enabled = _AllowSQLScriptFunctionality

        MenuItem5.Enabled = _AllowMathFunctionality
        miFormatStuff.Enabled = _AllowFormatFunctionality
        MenuItem2.Enabled = _AllowSettingsFunctionality
        MenuItem3.Enabled = _AllowSortFunctionality
        MenuItem4.Enabled = _AllowRowAndColumnFunctionality

    End Sub

    Private Sub miSmoothing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSmoothing.Click
        Me.Antialias = Not miSmoothing.Checked
    End Sub

    Private Sub miFontsLarger_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miFontsLarger.Click
        Dim fnt As System.Drawing.Font

        fnt = Me.DefaultCellFont

        If fnt.Size > 72 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit)

        Me.AllCellsUseThisFont(fnt2)

    End Sub

    Private Sub miFontsSmaller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miFontsSmaller.Click
        Dim fnt As System.Drawing.Font

        fnt = Me.DefaultCellFont

        If fnt.Size < 4 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit)

        Me.AllCellsUseThisFont(fnt2)

    End Sub

    Private Sub miFormatAsMoney_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miFormatAsMoney.Click

        Dim r As Integer
        Dim sf As New StringFormat
        Dim c As Integer = _ColOverOnMenuButton
        Dim a As String

        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        'sf.LineAlignment = StringAlignment.Far

        sf.LineAlignment = StringAlignment.Near
        sf.Alignment = StringAlignment.Far

        For r = 0 To _rows - 1
            If IsNumeric(_grid(r, c)) Then
                a = _grid(r, c)
                If a.StartsWith("$") Then
                    a = a.Substring(1)
                End If
                _grid(r, c) = Format(Val(a), "C")
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)
            End If
        Next

        Me.Refresh()

    End Sub

    Private Sub miFormatAsDecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miFormatAsDecimal.Click
        Dim r As Integer
        Dim sf As New StringFormat
        Dim c As Integer = _ColOverOnMenuButton
        Dim a As String

        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        'sf.LineAlignment = StringAlignment.Far

        sf.LineAlignment = StringAlignment.Near
        sf.Alignment = StringAlignment.Far

        For r = 0 To _rows - 1
            If IsNumeric(_grid(r, c)) Then
                a = _grid(r, c)
                If a.StartsWith("$") Then
                    a = a.Substring(1)
                End If
                _grid(r, c) = Format(Val(a), "G")
                _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)
            End If
        Next

        Me.Refresh()
    End Sub

    Private Sub miFormatAsText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miFormatAsText.Click
        Dim r As Integer
        Dim sf As New StringFormat(StringFormatFlags.FitBlackBox)
        Dim c As Integer = _ColOverOnMenuButton


        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        sf.LineAlignment = StringAlignment.Far

        For r = 0 To _rows - 1

            '_grid(r, c) = Format(Val(_grid(r, c)), "C")
            _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)

        Next

        Me.Refresh()
    End Sub

    Private Sub miCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miCenter.Click
        Dim r As Integer
        Dim sf As New StringFormat
        Dim c As Integer = _ColOverOnMenuButton


        sf.LineAlignment = StringAlignment.Center
        sf.Alignment = StringAlignment.Center

        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        For r = 0 To _rows - 1

            '_grid(r, c) = Format(Val(_grid(r, c)), "C")
            _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)

        Next

        Me.Refresh()
    End Sub

    Private Sub miLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miLeft.Click
        Dim r As Integer
        Dim sf As New StringFormat
        Dim c As Integer = _ColOverOnMenuButton


        sf.LineAlignment = StringAlignment.Near
        sf.Alignment = StringAlignment.Near

        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        For r = 0 To _rows - 1

            '_grid(r, c) = Format(Val(_grid(r, c)), "C")
            _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)

        Next

        Me.Refresh()
    End Sub

    Private Sub miRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miRight.Click
        Dim r As Integer
        Dim sf As New StringFormat
        Dim c As Integer = _ColOverOnMenuButton


        sf.LineAlignment = StringAlignment.Far
        sf.Alignment = StringAlignment.Far

        If c >= _cols Or _rows < 2 Or c < 0 Then
            Exit Sub
        End If

        For r = 0 To _rows - 1

            '_grid(r, c) = Format(Val(_grid(r, c)), "C")
            _gridCellAlignment(r, c) = GetGridCellAlignmentListEntry(sf)

        Next

        Me.Refresh()
    End Sub

    Private Sub miExportToExcel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miExportToExcel.Click
        Me.ExportToExcel()
    End Sub

    Private Sub miAutoFitCols_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miAutoFitCols.Click
        Me.ExcelAutoFitColumn = Not Me.ExcelAutoFitColumn
    End Sub

    Private Sub miAutoFitRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miAutoFitRows.Click
        Me.ExcelAutoFitRow = Not Me.ExcelAutoFitRow
    End Sub

    Private Sub miALternateRowColors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miALternateRowColors.Click
        Me.ExcelUseAlternateRowColor = Not Me.ExcelUseAlternateRowColor
    End Sub

    Private Sub miMatchGridColors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMatchGridColors.Click
        Me.ExcelMatchGridColorScheme = Not Me.ExcelMatchGridColorScheme
    End Sub

    Private Sub miOutlineExportedCells_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miOutlineExportedCells.Click
        Me.ExcelOutlineCells = Not Me.ExcelOutlineCells
    End Sub

    Private Sub miExportToTextFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miExportToTextFile.Click
        Me.ExportToText()
    End Sub

    Private Sub miHeaderFontSmaller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHeaderFontSmaller.Click
        Dim fnt As System.Drawing.Font

        fnt = _GridHeaderFont

        If fnt.Size < 4 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit)

        _GridHeaderFont = fnt2

        Me.Invalidate()

    End Sub

    Private Sub miHeaderFontLarger_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHeaderFontLarger.Click
        Dim fnt As System.Drawing.Font

        fnt = _GridHeaderFont

        If fnt.Size > 60 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit)

        _GridHeaderFont = fnt2

        Me.Invalidate()

    End Sub

    Private Sub miTitleFontSmaller_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miTitleFontSmaller.Click
        Dim fnt As System.Drawing.Font

        fnt = _GridTitleFont

        If fnt.Size < 4 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit)

        _GridTitleFont = fnt2

        Me.Invalidate()

    End Sub

    Private Sub miTitleFontLarger_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miTitleFontLarger.Click
        Dim fnt As System.Drawing.Font

        fnt = _GridTitleFont

        If fnt.Size > 60 Then

            Exit Sub

        End If

        Dim fnt2 As New System.Drawing.Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit)

        _GridTitleFont = fnt2

        Me.Invalidate()
    End Sub

    Private Sub miSearchInColumn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSearchInColumn.Click
        Dim frm As New frmSearchInColumn(Me.PointToScreen(Me.Location))

        frm.ColumnName = _GridHeader(_ColOverOnMenuButton)

        frm.ShowDialog()

        If Not frm.Canceled Then
            ' we wanna search
            Dim srch As String = frm.SearchText
            Dim t As Integer

            If srch = "" Then
                ' we have nothing to look for so lets bail
                Exit Sub
            End If

            _LastSearchText = srch
            _LastSearchColumn = _ColOverOnMenuButton

            For t = 0 To _rows - 1
                If InStr(UCase(_grid(t, _ColOverOnMenuButton)), UCase(srch), CompareMethod.Text) <> 0 Then
                    ' we have a match
                    If vs.Visible Then
                        vs.Value = t
                        _SelectedRow = t
                        Me.Invalidate()
                        Exit For
                    Else
                        _SelectedRow = t
                        Me.Invalidate()
                        Exit For
                    End If
                End If
            Next

        End If

    End Sub

    Private Sub miAutoSizeToContents_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miAutoSizeToContents.Click
        Me.AutoSizeCellsToContents = True
    End Sub

    Private Sub miAllowUserColumnResizing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miAllowUserColumnResizing.Click
        _AllowUserColumnResizing = Not _AllowUserColumnResizing
    End Sub

    Private Sub miSortAscending_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSortAscending.Click

        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer

        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)

        Me.Refresh()

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = _grid(y, _ColOverOnMenuButton)
            sitem.Iord = y
            lb.Items.Add(sitem)
            lb.DisplayMember = "Ivis"
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
        Next
        lb.Sorted = True

        For rr = 0 To lb.Items.Count - 1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid(rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
        Next


        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        _SelectedRow = -1
        _SelectedRows.Clear()
        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        Me.Invalidate()


        pBar.Visible = False
        gb1.Visible = False

    End Sub

    Private Sub miSortDescending_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSortDescending.Click

        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer
        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)



        Me.Refresh()

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = _grid(y, _ColOverOnMenuButton)
            sitem.Iord = y
            lb.Items.Add(sitem)
            lb.DisplayMember = "Ivis"
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
        Next
        lb.Sorted = True

        For rr = lb.Items.Count - 1 To 0 Step -1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid((lb.Items.Count - 1) - rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
        Next

        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        Me.Invalidate()

        _SelectedRow = -1
        _SelectedRows.Clear()
        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        pBar.Visible = False
        gb1.Visible = False

    End Sub

    Private Sub miDateAsc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miDateAsc.Click
        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer

        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)



        Me.Refresh()

        For y = 0 To _rows - 1
            If (Not IsDate(_grid(y, _ColOverOnMenuButton)) And _grid(y, _ColOverOnMenuButton).Trim() = "") Then

                _grid(y, _ColOverOnMenuButton) = DateTime.MinValue.ToString("MM/dd/yyyy")
            Else
                If (Not IsDate(_grid(y, _ColOverOnMenuButton) + "")) Then
                    MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a date", _
                             MsgBoxStyle.Critical, "Sort date descending message")
                    Exit Sub
                End If

            End If
        Next

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = Format(CDate(_grid(y, _ColOverOnMenuButton)), "yyyyMMdd")
            sitem.Iord = y
            lb.Items.Add(sitem)
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next
        lb.DisplayMember = "Ivis"
        lb.Sorted = True

        For rr = 0 To lb.Items.Count - 1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid(rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next

        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        For y = 0 To _rows - 1
            If _grid(y, _ColOverOnMenuButton).Trim = "01/01/0001" Then
                _grid(y, _ColOverOnMenuButton) = ""
            End If
        Next

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        Me.Refresh()

        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        pBar.Visible = False
        gb1.Visible = False

    End Sub

    Private Sub miDateDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miDateDesc.Click
        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer

        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)


        Me.Refresh()

        For y = 0 To _rows - 1
            If (Not IsDate(_grid(y, _ColOverOnMenuButton)) And _grid(y, _ColOverOnMenuButton).Trim() = "") Then

                _grid(y, _ColOverOnMenuButton) = DateTime.MinValue.ToString("MM/dd/yyyy")
            Else
                If (Not IsDate(_grid(y, _ColOverOnMenuButton) + "")) Then
                    MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a date", _
                             MsgBoxStyle.Critical, "Sort date descending message")
                    Exit Sub
                End If

            End If
        Next

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = Format(CDate(_grid(y, _ColOverOnMenuButton)), "yyyyMMdd")
            sitem.Iord = y
            lb.Items.Add(sitem)
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next
        lb.DisplayMember = "Ivis"
        lb.Sorted = True

        For rr = lb.Items.Count - 1 To 0 Step -1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid((lb.Items.Count - 1) - rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next

        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        For y = 0 To _rows - 1
            If _grid(y, _ColOverOnMenuButton).Trim = "01/01/0001" Then
                _grid(y, _ColOverOnMenuButton) = ""
            End If
        Next

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        Me.Refresh()

        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        pBar.Visible = False
        gb1.Visible = False

    End Sub

    Private Sub miSortNumericAsc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSortNumericAsc.Click
        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer

        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)



        Me.Refresh()

        For y = 0 To _rows - 1
            If Not IsNumeric(_grid(y, _ColOverOnMenuButton)) Then
                MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a Number", _
                         MsgBoxStyle.Critical, "Sort date ascending message")
                Exit Sub
            End If
        Next

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = Microsoft.VisualBasic.Right("000000000000000" + _grid(y, _ColOverOnMenuButton), 15)
            sitem.Iord = y
            lb.Items.Add(sitem)
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next
        lb.DisplayMember = "Ivis"
        lb.Sorted = True

        For rr = 0 To lb.Items.Count - 1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid(rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next

        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        pBar.Visible = False
        gb1.Visible = False
    End Sub

    Private Sub miSortNumericDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSortNumericDesc.Click
        Dim newgrid(_rows, _cols) As String
        Dim lb As New ListBox
        Dim y, rr, cc As Integer

        Dim presortcolwidths(_cols) As Integer
        Dim oldautosizecells As Boolean = _AutosizeCellsToContents

        Array.Copy(_colwidths, presortcolwidths, _cols)

        Me.Refresh()

        For y = 0 To _rows - 1
            If Not IsNumeric(_grid(y, _ColOverOnMenuButton)) Then
                MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a Number", _
                         MsgBoxStyle.Critical, "Sort date ascending message")
                Exit Sub
            End If
        Next

        If _ShowProgressBar Then
            pBar.Maximum = _rows * 2
            pBar.Minimum = 0
            pBar.Value = 0
            pBar.Visible = True
            gb1.Visible = True
            pBar.Refresh()
            gb1.Refresh()
        End If

        lb.Items.Clear()
        For y = 0 To _rows - 1
            Dim sitem As New SortItem
            sitem.Ivis = Microsoft.VisualBasic.Right("000000000000000" + _grid(y, _ColOverOnMenuButton), 15)
            sitem.Iord = y
            lb.Items.Add(sitem)
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next
        lb.DisplayMember = "Ivis"
        lb.Sorted = True


        For rr = lb.Items.Count - 1 To 0 Step -1
            ' loop through the current listbox
            For cc = 0 To _cols - 1
                newgrid((lb.Items.Count - 1) - rr, cc) = _grid(lb.Items(rr).iord, cc)
            Next
            If _ShowProgressBar Then
                pBar.Increment(1)
                pBar.Refresh()
            End If
            Application.DoEvents()
        Next

        PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, False)

        Array.Copy(presortcolwidths, _colwidths, _cols)

        _AutosizeCellsToContents = oldautosizecells

        RaiseEvent GridResorted(Me, _ColOverOnMenuButton)

        pBar.Visible = False
        gb1.Visible = False
    End Sub

    Private Sub miHideRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHideRow.Click
        If _RowOverOnMenuButton <> -1 And _RowOverOnMenuButton < _rows Then
            _rowheights(_RowOverOnMenuButton) = 0
            _AutosizeCellsToContents = False
            Me.Invalidate()
        End If
    End Sub

    Private Sub miHideColumn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHideColumn.Click
        If _ColOverOnMenuButton <> -1 And _ColOverOnMenuButton < _cols Then
            _colwidths(_ColOverOnMenuButton) = 0
            _AutosizeCellsToContents = False
            Me.Invalidate()
        End If
    End Sub

    Private Sub miShowAllRowsAndColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miShowAllRowsAndColumns.Click
        Me.AutoSizeCellsToContents = True
        Me.Invalidate()
    End Sub

    Private Sub miSetRowColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSetRowColor.Click
        Dim r As Integer
        Dim ccol As Integer
        If clrdlg.ShowDialog = DialogResult.OK Then
            ccol = GetGridBackColorListEntry(New SolidBrush(clrdlg.Color))
            For r = 0 To _cols - 1
                _gridBackColor(_RowOverOnMenuButton, r) = ccol
            Next
            Me.Invalidate()
        End If
    End Sub

    Private Sub miSetColumnColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSetColumnColor.Click
        Dim r As Integer
        Dim ccol As Integer
        If clrdlg.ShowDialog = DialogResult.OK Then
            ccol = GetGridBackColorListEntry(New SolidBrush(clrdlg.Color))
            For r = 0 To _rows - 1
                _gridBackColor(r, _ColOverOnMenuButton) = ccol
            Next
            Me.Invalidate()
        End If
    End Sub

    Private Sub miSetCellColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSetCellColor.Click
        If clrdlg.ShowDialog = DialogResult.OK Then

            _gridBackColor(_RowOverOnMenuButton, _ColOverOnMenuButton) = GetGridBackColorListEntry(New SolidBrush(clrdlg.Color))
            Me.Invalidate()

        End If
    End Sub

    Private Sub miSumColumn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSumColumn.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _rows - 1
            If IsNumeric(_grid(t, _ColOverOnMenuButton)) Then
                v += CDbl(_grid(t, _ColOverOnMenuButton))
                flag = True
            End If
        Next

        If flag Then
            MsgBox("The sum of the column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "is " & v.ToString, MsgBoxStyle.Information, "SUM COLUMN")
        Else
            MsgBox("The column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "Contains no numeric data...", MsgBoxStyle.Information, "SUM COLUMN")
        End If

    End Sub

    Private Sub miSumRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miSumRow.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _cols - 1
            If IsNumeric(_grid(_RowOverOnMenuButton, t)) Then
                v += CDbl(_grid(_RowOverOnMenuButton, t))
                flag = True
            End If
        Next

        If flag Then
            MsgBox("The sum of the row numbered " & _RowOverOnMenuButton & vbCrLf & "is " & v.ToString, MsgBoxStyle.Information, "SUM ROW")
        Else
            MsgBox("The row numbered " & _RowOverOnMenuButton & vbCrLf & "Contains no numeric data...", MsgBoxStyle.Information, "SUM ROW")
        End If
    End Sub

    Private Sub miMaxCol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMaxCol.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _rows - 1
            If IsNumeric(_grid(t, _ColOverOnMenuButton)) Then
                If Not flag Then
                    v = CDbl(_grid(t, _ColOverOnMenuButton))
                    flag = True
                Else
                    If CDbl(_grid(t, _ColOverOnMenuButton)) > v Then
                        v = CDbl(_grid(t, _ColOverOnMenuButton))
                    End If
                End If
            End If
        Next

        If flag Then
            MsgBox("The max value in the column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "is " & v.ToString, _
                        MsgBoxStyle.Information, "MAX IN COLUMN")
        Else
            MsgBox("The column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "Contains no numeric data...", _
                        MsgBoxStyle.Information, "MAX IN COLUMN")
        End If
    End Sub

    Private Sub miMaxRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMaxRow.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _cols - 1
            If IsNumeric(_grid(_RowOverOnMenuButton, t)) Then
                If Not flag Then
                    v = CDbl(_grid(_RowOverOnMenuButton, t))
                    flag = True
                Else
                    If CDbl(_grid(_RowOverOnMenuButton, t)) > v Then
                        v = CDbl(_grid(_RowOverOnMenuButton, t))
                    End If
                End If
            End If
        Next

        If flag Then
            MsgBox("The max value in the row numbered " & _RowOverOnMenuButton & vbCrLf & "is " & v.ToString, _
                    MsgBoxStyle.Information, "MAX IN ROW")
        Else
            MsgBox("The row numbered " & _RowOverOnMenuButton & vbCrLf & "Contains no numeric data...", _
                    MsgBoxStyle.Information, "MAX IN ROW")
        End If
    End Sub

    Private Sub miMinCol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMinCol.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _rows - 1
            If IsNumeric(_grid(t, _ColOverOnMenuButton)) Then
                If Not flag Then
                    v = CDbl(_grid(t, _ColOverOnMenuButton))
                    flag = True
                Else
                    If CDbl(_grid(t, _ColOverOnMenuButton)) < v Then
                        v = CDbl(_grid(t, _ColOverOnMenuButton))
                    End If
                End If
            End If
        Next

        If flag Then
            MsgBox("The min value in the column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "is " & v.ToString, _
                    MsgBoxStyle.Information, "MIN IN COLUMN")
        Else
            MsgBox("The column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "Contains no numeric data...", _
                    MsgBoxStyle.Information, "MIN IN COLUMN")
        End If
    End Sub

    Private Sub miMinRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMinRow.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _cols - 1
            If IsNumeric(_grid(_RowOverOnMenuButton, t)) Then
                If Not flag Then
                    v = CDbl(_grid(_RowOverOnMenuButton, t))
                    flag = True
                Else
                    If CDbl(_grid(_RowOverOnMenuButton, t)) < v Then
                        v = CDbl(_grid(_RowOverOnMenuButton, t))
                    End If
                End If
            End If
        Next

        If flag Then
            MsgBox("The min value in the row numbered " & _RowOverOnMenuButton & vbCrLf & "is " & v.ToString, _
                    MsgBoxStyle.Information, "MIN IN ROW")
        Else
            MsgBox("The row numbered " & _RowOverOnMenuButton & vbCrLf & "Contains no numeric data...", _
                    MsgBoxStyle.Information, "MIN IN ROW")
        End If
    End Sub

    Private Sub miColAverage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miColAverage.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _rows - 1
            If IsNumeric(_grid(t, _ColOverOnMenuButton)) Then
                v += CDbl(_grid(t, _ColOverOnMenuButton))
                flag = True
            End If
        Next

        If flag Then
            v = v / _rows
            MsgBox("The average value in the column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "is " & Format(v, "##,##0.00"), _
                    MsgBoxStyle.Information, "AVERAGE IN COLUMN")
        Else
            MsgBox("The column named " & _GridHeader(_ColOverOnMenuButton) & vbCrLf & "Contains no numeric data...", _
                    MsgBoxStyle.Information, "AVERAGE IN COLUMN")
        End If

    End Sub

    Private Sub miRowAverage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miRowAverage.Click
        Dim t As Integer
        Dim v As Double = 0
        Dim flag As Boolean = False

        For t = 0 To _cols - 1
            If IsNumeric(_grid(_RowOverOnMenuButton, t)) Then
                v += CDbl(_grid(_RowOverOnMenuButton, t))
                flag = True
            End If
        Next

        If flag Then
            v = v / _cols
            MsgBox("The average value in the row numbered " & _RowOverOnMenuButton & vbCrLf & "is " & Format(v, "##,##0.00"), _
                    MsgBoxStyle.Information, "AVERAGE IN ROW")
        Else
            MsgBox("The row numbered " & _RowOverOnMenuButton & vbCrLf & "Contains no numeric data...", _
                    MsgBoxStyle.Information, "AVERAGE IN ROW")
        End If
    End Sub

    Private Sub miCopyCellToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miCopyCellToClipboard.Click
        System.Windows.Forms.Clipboard.SetDataObject(_grid(_RowOverOnMenuButton, _ColOverOnMenuButton), True)
    End Sub

    Private Sub miPrintTheGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPrintTheGrid.Click
        Me.PrintTheGrid(_gridReportTitle, True, True, True, False, False)
    End Sub

    Private Sub miPreviewTheGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPreviewTheGrid.Click
        Me.PrintTheGrid(_gridReportTitle, True, True, True, True, False)
    End Sub

    Private Sub LogThis(ByVal str As String)

        'If _LoggingEnabled Then

        '    Dim r As System.IO.StreamWriter = System.IO.File.AppendText("C:\TAIGRIDLOG.TXT")

        '    r.WriteLine(str)

        '    r.Flush()
        '    r.Close()

        'End If


    End Sub

    Private Sub PurgeLog()
        'System.IO.File.Delete("C:\TAIGRIDLOG.TXT")
    End Sub

    Private Sub miPageSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPageSetup.Click

        If _rows = 0 Or _cols = 0 Then
            ' If we got nothing to print dont bring up the page setup 
            Exit Sub
        End If

        Try
            _psets = New System.Drawing.Printing.PageSettings

            'MsgBox(_psets.ToString())

            _OriginalPrinterName = _psets.PrinterSettings.PrinterName

            _PageSetupForm = New frmPageSetup(_psets)

        Catch ex As Exception

            miPreviewTheGrid.Enabled = False
            miPrintTheGrid.Enabled = False
            miPageSetup.Enabled = False
            Exit Sub

        End Try

        PurgeLog()

        LogThis("Dim ps As System.Drawing.Printing.PageSettings = _psets")

        Dim ps As System.Drawing.Printing.PageSettings = _psets

        LogThis("PageSetupForm = Nothing")

        _PageSetupForm = Nothing

        LogThis("_PageSetupForm = New frmPageSetup(_psets)")

        _PageSetupForm = New frmPageSetup(_psets)

        LogThis("PageSetupForm.MaxPage = CalculatePageRange()")

        _PageSetupForm.MaxPage = CalculatePageRange()

        LogThis(" _PageSetupForm.ShowDialog()")

        _PageSetupForm.ShowDialog()

        If _PageSetupForm.Canceled Then
            _psets = ps
        Else
            _psets = _PageSetupForm.Psets

            If _PageSetupForm.Print Then
                Me.PrintTheGrid(_gridReportTitle, True, True, True, False, _psets.Landscape)
            Else
                If _PageSetupForm.Preview Then
                    Me.PrintTheGrid(_gridReportTitle, True, True, True, True, _psets.Landscape)
                End If
            End If

        End If

        Me.Refresh()

    End Sub

    Private Sub miExportToSQLScript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miExportToSQLScript.Click
        Dim frm As New frmScriptToSQL(Me)

        frm.Location = Me.PointToScreen(Me.Location)

        frm.ShowDialog()

    End Sub

    Private Sub miExportToHTMLTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miExportToHTMLTable.Click
        Dim frm As New frmScriptToHTML(Me)

        frm.Location = Me.PointToScreen(Me.Location)

        frm.ShowDialog()

    End Sub

    Private Sub miDisplayFrequencyDistribution_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miDisplayFrequencyDistribution.Click
        Dim frm As New frmFreqDist(Me, _ColOverOnMenuButton)

        frm.ShowDialog()

    End Sub

    Private Sub miProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miProperties.Click
        Dim frm As New frmGridProperties(Me)

        frm.Location = Me.PointToScreen(Me.Location)

        frm.ShowDialog()

        Me.BringToFront()
    End Sub

    Private Sub miTearColumnAway_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miTearColumnAway.Click

        If _ColOverOnMenuButton = -1 Then
            Exit Sub
        End If

        _TearAwayWork = True

        TearAwayColumID(_ColOverOnMenuButton)

        _TearAwayWork = False

    End Sub

    Private Sub miHideColumnTearAway_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHideColumnTearAway.Click

        If TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)
            If ta.ColID = _ColOverOnMenuButton Then
                ' call into the child form to start the death spiral a happening
                ta.Winform.KillMe(_ColOverOnMenuButton)
            End If
        Next

    End Sub

    Private Sub miHideAllTearAwayColumns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miHideAllTearAwayColumns.Click
        If TearAways.Count = 0 Then
            Exit Sub
        End If

        Dim t As Integer

        For t = TearAways.Count - 1 To 0 Step -1
            Dim ta As TearAwayWindowEntry = TearAways.Item(t)
            ' call into the child form to start the death spiral a happening (one at a time as we
            ' be a killing em all
            ta.Winform.KillMe(ta.ColID)
        Next
    End Sub

    Private Sub miMultipleColumnTearAway_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miMultipleColumnTearAway.Click

        If _cols = 0 Then
            ' there be no columns to tear away  
            Exit Sub
        End If

        _TearAwayWork = True

        Dim frm As New frmMultipleColumnTearAway(_GridHeader)

        frm.Location = Me.PointToScreen(Me.Location)

        frm.ShowDialog()


        If Not frm.Canceled Then
            ' here we need to tear away the columns if they are selected

            If frm.SelectedIndices.Count > 0 Then

                Dim t As Integer

                For t = 0 To frm.SelectedIndices.Count - 1
                    TearAwayColumID(frm.SelectedIndices.Item(t))
                Next

            End If

            System.Threading.Thread.Sleep(100)
            Application.DoEvents()

            _TearAwayWork = False

            ArrangeTearAwayWindows()

        End If

        _TearAwayWork = False

    End Sub

    Private Sub miArrangeTearAways_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miArrangeTearAways.Click

        ArrangeTearAwayWindows()

    End Sub


#End Region

#End Region

#Region " Overrides "

    Protected Overrides Sub OnPaintBackground(ByVal pevent As System.Windows.Forms.PaintEventArgs)

    End Sub

    Protected Overrides Function ProcessDialogKey(ByVal kd As Keys) As Boolean
        If _EditMode And kd = Keys.Tab And _AllowInGridEdits Then
            ' we may need to bounce to the next col for edits or tab off the grid in the case of being at the end

            Dim x, y, xoff, yoff, r, c, nrc, ncc As Integer

            Dim flag As Boolean = False

            nrc = _RowClicked
            ncc = _ColClicked

            If ncc < _cols - 1 Then
                x = ncc + 1

            Else
                x = 0

                nrc += 1

            End If

            y = nrc

            If y > _rows - 1 Then
                flag = True
                x = -1
                y = -1

            Else
                Do While Not flag
                    If _colEditable(x) And Not (_colboolean(x)) Then
                        flag = True
                    Else
                        x += 1
                        If x > _cols - 1 Then
                            x = 0
                            y += 1

                            If y > _rows - 1 Then
                                flag = True
                                x = -1
                                y = -1

                            End If
                        End If
                    End If
                Loop
            End If

            ' if we get here and not flag  and x > -1 and y > -1 then we are at the end
            ' otherwise x = newcolumn for edit y = new row for edit
            ' we need to clean up the existing edit and jump to the new one

            ' who has focus

            If flag And x > -1 And y > -1 Then

                If txtInput.Visible Then
                    ' the txtinput has it

                    If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
                        Dim oldval As String = _grid(_RowClicked, _ColClicked)
                        Dim newval As String = txtInput.Text
                        _grid(_RowClicked, _ColClicked) = txtInput.Text
                        RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                    End If

                    txtInput.SendToBack()
                    txtInput.Visible = False

                Else
                    ' the cmboinput does

                    If _grid(_RowClicked, _ColClicked) <> cmboInput.Text And cmboInput.Text.Trim() <> "" Then
                        Dim oldval As String = _grid(_RowClicked, _ColClicked)
                        Dim newval As String = cmboInput.Text
                        _grid(_RowClicked, _ColClicked) = cmboInput.Text
                        RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                    End If

                    cmboInput.SendToBack()
                    cmboInput.Visible = False

                End If

                ' we now need to ensure that the selectedrows collection is clear

                _SelectedRows.Clear()

                ' now lets setup for the new edit

                If _SelectedRow <> y Then
                    Me.SelectedRow = y
                End If

                _RowClicked = y
                _ColClicked = x

                If _RowClicked > -1 And _RowClicked < _rows And _rowEditable(_RowClicked) Then
                    If _ColClicked > -1 And _ColClicked < _cols And _colEditable(_ColClicked) And _AllowInGridEdits Then

                        If IsColumnRestricted(_ColClicked) Then

                            Dim it As EditColumnRestrictor = GetColumnRestriction(_ColClicked)

                            cmboInput.Items.Clear()

                            Dim s() As String = it.RestrictedList.Split("^".ToCharArray)
                            For Each ss As String In s
                                cmboInput.Items.Add(ss)
                            Next

                            ' we have selected a row and col lets move the txtinput there and bring it to the front
                            xoff = 0
                            yoff = 0

                            If _RowClicked > 0 Then
                                For r = 0 To _RowClicked - 1
                                    yoff = yoff + RowHeight(r)
                                Next
                            End If

                            If GridheaderVisible Then
                                yoff = yoff + _GridHeaderHeight
                            End If

                            If _GridTitleVisible Then
                                yoff = yoff + _GridTitleHeight
                            End If

                            If _ColClicked > 0 Then
                                For c = 0 To _ColClicked - 1
                                    xoff = xoff + ColWidth(c)
                                Next
                            End If

                            If vs.Visible And vs.Value > 0 Then
                                yoff = yoff - GimmeYOffset(vs.Value)
                            End If

                            If hs.Visible And hs.Value > 0 Then
                                xoff = xoff - GimmeXOffset(hs.Value)
                            End If

                            If _CellOutlines Then
                                cmboInput.Top = yoff + 1
                                cmboInput.Left = xoff + 1
                                cmboInput.Width = ColWidth(_ColClicked) - 1
                                cmboInput.Height = RowHeight(_RowClicked) - 2
                                cmboInput.BackColor = _colEditableTextBackColor
                            Else
                                cmboInput.Top = yoff
                                cmboInput.Left = xoff
                                cmboInput.Width = ColWidth(_ColClicked)
                                cmboInput.Height = RowHeight(_RowClicked)
                                cmboInput.BackColor = _colEditableTextBackColor
                            End If

                            cmboInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                            cmboInput.Text = _grid(_RowClicked, _ColClicked)

                            cmboInput.Visible = True
                            cmboInput.BringToFront()
                            cmboInput.DroppedDown = True
                            _EditModeCol = _ColClicked
                            _EditModeRow = _RowClicked
                            _EditMode = True

                            cmboInput.Focus()
                        Else
                            ' we have selected a row and col lets move the txtinput there and bring it to the front
                            xoff = 0
                            yoff = 0

                            If _RowClicked > 0 Then
                                For r = 0 To _RowClicked - 1
                                    yoff = yoff + RowHeight(r)
                                Next
                            End If

                            If GridheaderVisible Then
                                yoff = yoff + _GridHeaderHeight
                            End If

                            If _GridTitleVisible Then
                                yoff = yoff + _GridTitleHeight
                            End If

                            If _ColClicked > 0 Then
                                For c = 0 To _ColClicked - 1
                                    xoff = xoff + ColWidth(c)
                                Next
                            End If

                            If vs.Visible And vs.Value > 0 Then
                                yoff = yoff - GimmeYOffset(vs.Value)
                            End If

                            If hs.Visible And hs.Value > 0 Then
                                xoff = xoff - GimmeXOffset(hs.Value)
                            End If

                            If _CellOutlines Then
                                txtInput.Top = yoff + 1
                                txtInput.Left = xoff + 1
                                txtInput.Width = ColWidth(_ColClicked) - 1
                                txtInput.Height = RowHeight(_RowClicked) - 2
                                txtInput.BackColor = _colEditableTextBackColor
                            Else
                                txtInput.Top = yoff
                                txtInput.Left = xoff
                                txtInput.Width = ColWidth(_ColClicked)
                                txtInput.Height = RowHeight(_RowClicked)
                                txtInput.BackColor = _colEditableTextBackColor
                            End If

                            txtInput.Font = _gridCellFontsList(_gridCellFonts(_RowClicked, _ColClicked))

                            txtInput.Text = _grid(_RowClicked, _ColClicked)

                            txtInput.Visible = True
                            txtInput.BringToFront()
                            _EditModeCol = _ColClicked
                            _EditModeRow = _RowClicked
                            _EditMode = True

                            txtInput.Focus()
                        End If

                    End If
                End If

                Return True
            Else ' If flag And x > -1 And y > -1 Then

                If txtInput.Visible Then
                    ' the txtinput has it

                    If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
                        Dim oldval As String = _grid(_RowClicked, _ColClicked)
                        Dim newval As String = txtInput.Text
                        _grid(_RowClicked, _ColClicked) = txtInput.Text
                        RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                    End If

                    txtInput.SendToBack()
                    txtInput.Visible = False

                End If


                If cmboInput.Visible Then
                    ' the cmboinput does

                    If _grid(_RowClicked, _ColClicked) <> cmboInput.Text And cmboInput.Text.Trim() <> "" Then
                        Dim oldval As String = _grid(_RowClicked, _ColClicked)
                        Dim newval As String = cmboInput.Text
                        _grid(_RowClicked, _ColClicked) = cmboInput.Text
                        RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                    End If

                    cmboInput.SendToBack()
                    cmboInput.Visible = False

                End If

                MyBase.ProcessDialogKey(kd)

                Return False

            End If ' If flag And x > -1 And y > -1 Then

        Else ' If _EditMode And kd = Keys.Tab And _AllowInGridEdits

            'If _EditMode And _AllowInGridEdits Then
            '    If txtInput.Visible Then
            '        ' the txtinput has it

            '        If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
            '            Dim oldval As String = _grid(_RowClicked, _ColClicked)
            '            Dim newval As String = txtInput.Text
            '            _grid(_RowClicked, _ColClicked) = txtInput.Text
            '            RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            '        End If

            '        txtInput.SendToBack()
            '        txtInput.Visible = False

            '    End If


            '    If cmboInput.Visible Then
            '        ' the cmboinput does

            '        If _grid(_RowClicked, _ColClicked) <> cmboInput.Text And cmboInput.Text.Trim() <> "" Then
            '            Dim oldval As String = _grid(_RowClicked, _ColClicked)
            '            Dim newval As String = cmboInput.Text
            '            _grid(_RowClicked, _ColClicked) = cmboInput.Text
            '            RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
            '        End If

            '        cmboInput.SendToBack()
            '        cmboInput.Visible = False

            '    End If

            'End If

            MyBase.ProcessDialogKey(kd)
            Return False
        End If ' If _EditMode And kd = Keys.Tab And _AllowInGridEdits
    End Function

#End Region

#Region " Internal Forms and other classes "

    Private Class frmExportToText
        Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents cmdOK As System.Windows.Forms.Button
        Friend WithEvents cmdCancel As System.Windows.Forms.Button
        Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
        Friend WithEvents rbTab As System.Windows.Forms.RadioButton
        Friend WithEvents rbSemicolon As System.Windows.Forms.RadioButton
        Friend WithEvents rbComma As System.Windows.Forms.RadioButton
        Friend WithEvents rbSpace As System.Windows.Forms.RadioButton
        Friend WithEvents rbOther As System.Windows.Forms.RadioButton
        Friend WithEvents txtOther As System.Windows.Forms.TextBox
        Friend WithEvents chkIncludeFieldNames As System.Windows.Forms.CheckBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents txtExportFile As System.Windows.Forms.TextBox
        Friend WithEvents cmdBrowse As System.Windows.Forms.Button
        Friend WithEvents chkIncludeLineTerminator As System.Windows.Forms.CheckBox
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.cmdOK = New System.Windows.Forms.Button
            Me.cmdCancel = New System.Windows.Forms.Button
            Me.GroupBox1 = New System.Windows.Forms.GroupBox
            Me.txtOther = New System.Windows.Forms.TextBox
            Me.rbOther = New System.Windows.Forms.RadioButton
            Me.rbSpace = New System.Windows.Forms.RadioButton
            Me.rbComma = New System.Windows.Forms.RadioButton
            Me.rbSemicolon = New System.Windows.Forms.RadioButton
            Me.rbTab = New System.Windows.Forms.RadioButton
            Me.chkIncludeFieldNames = New System.Windows.Forms.CheckBox
            Me.Label1 = New System.Windows.Forms.Label
            Me.txtExportFile = New System.Windows.Forms.TextBox
            Me.cmdBrowse = New System.Windows.Forms.Button
            Me.chkIncludeLineTerminator = New System.Windows.Forms.CheckBox
            Me.GroupBox1.SuspendLayout()
            Me.SuspendLayout()
            '
            'cmdOK
            '
            Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.cmdOK.Location = New System.Drawing.Point(472, 8)
            Me.cmdOK.Name = "cmdOK"
            Me.cmdOK.Size = New System.Drawing.Size(104, 24)
            Me.cmdOK.TabIndex = 0
            Me.cmdOK.Text = "OK"
            '
            'cmdCancel
            '
            Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.cmdCancel.Location = New System.Drawing.Point(472, 40)
            Me.cmdCancel.Name = "cmdCancel"
            Me.cmdCancel.Size = New System.Drawing.Size(104, 24)
            Me.cmdCancel.TabIndex = 1
            Me.cmdCancel.Text = "Cancel"
            '
            'GroupBox1
            '
            Me.GroupBox1.Controls.Add(Me.txtOther)
            Me.GroupBox1.Controls.Add(Me.rbOther)
            Me.GroupBox1.Controls.Add(Me.rbSpace)
            Me.GroupBox1.Controls.Add(Me.rbComma)
            Me.GroupBox1.Controls.Add(Me.rbSemicolon)
            Me.GroupBox1.Controls.Add(Me.rbTab)
            Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
            Me.GroupBox1.Name = "GroupBox1"
            Me.GroupBox1.Size = New System.Drawing.Size(456, 56)
            Me.GroupBox1.TabIndex = 2
            Me.GroupBox1.TabStop = False
            Me.GroupBox1.Text = "Choose the delimiter that separates your fields"
            '
            'txtOther
            '
            Me.txtOther.Location = New System.Drawing.Point(376, 22)
            Me.txtOther.Name = "txtOther"
            Me.txtOther.Size = New System.Drawing.Size(32, 20)
            Me.txtOther.TabIndex = 5
            Me.txtOther.Text = ""
            '
            'rbOther
            '
            Me.rbOther.Location = New System.Drawing.Point(312, 24)
            Me.rbOther.Name = "rbOther"
            Me.rbOther.Size = New System.Drawing.Size(56, 16)
            Me.rbOther.TabIndex = 4
            Me.rbOther.Text = "&Other"
            '
            'rbSpace
            '
            Me.rbSpace.Location = New System.Drawing.Point(248, 24)
            Me.rbSpace.Name = "rbSpace"
            Me.rbSpace.Size = New System.Drawing.Size(56, 16)
            Me.rbSpace.TabIndex = 3
            Me.rbSpace.Text = "S&pace"
            '
            'rbComma
            '
            Me.rbComma.Checked = True
            Me.rbComma.Location = New System.Drawing.Point(168, 24)
            Me.rbComma.Name = "rbComma"
            Me.rbComma.Size = New System.Drawing.Size(72, 16)
            Me.rbComma.TabIndex = 2
            Me.rbComma.TabStop = True
            Me.rbComma.Text = "&Comma"
            '
            'rbSemicolon
            '
            Me.rbSemicolon.Location = New System.Drawing.Point(80, 24)
            Me.rbSemicolon.Name = "rbSemicolon"
            Me.rbSemicolon.Size = New System.Drawing.Size(80, 16)
            Me.rbSemicolon.TabIndex = 1
            Me.rbSemicolon.Text = "&Semicolon"
            '
            'rbTab
            '
            Me.rbTab.Location = New System.Drawing.Point(24, 24)
            Me.rbTab.Name = "rbTab"
            Me.rbTab.Size = New System.Drawing.Size(48, 16)
            Me.rbTab.TabIndex = 0
            Me.rbTab.Text = "&Tab"
            '
            'chkIncludeFieldNames
            '
            Me.chkIncludeFieldNames.Checked = True
            Me.chkIncludeFieldNames.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkIncludeFieldNames.Location = New System.Drawing.Point(32, 72)
            Me.chkIncludeFieldNames.Name = "chkIncludeFieldNames"
            Me.chkIncludeFieldNames.Size = New System.Drawing.Size(200, 24)
            Me.chkIncludeFieldNames.TabIndex = 3
            Me.chkIncludeFieldNames.Text = "Include Field Names on First Row"
            '
            'Label1
            '
            Me.Label1.Location = New System.Drawing.Point(16, 104)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(128, 24)
            Me.Label1.TabIndex = 4
            Me.Label1.Text = "Export To File:"
            Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'txtExportFile
            '
            Me.txtExportFile.Location = New System.Drawing.Point(16, 128)
            Me.txtExportFile.Name = "txtExportFile"
            Me.txtExportFile.Size = New System.Drawing.Size(448, 20)
            Me.txtExportFile.TabIndex = 5
            Me.txtExportFile.Text = ""
            '
            'cmdBrowse
            '
            Me.cmdBrowse.Location = New System.Drawing.Point(472, 128)
            Me.cmdBrowse.Name = "cmdBrowse"
            Me.cmdBrowse.Size = New System.Drawing.Size(104, 24)
            Me.cmdBrowse.TabIndex = 6
            Me.cmdBrowse.Text = "Browse"
            '
            'chkIncludeLineTerminator
            '
            Me.chkIncludeLineTerminator.Checked = True
            Me.chkIncludeLineTerminator.CheckState = System.Windows.Forms.CheckState.Checked
            Me.chkIncludeLineTerminator.Location = New System.Drawing.Point(256, 72)
            Me.chkIncludeLineTerminator.Name = "chkIncludeLineTerminator"
            Me.chkIncludeLineTerminator.Size = New System.Drawing.Size(200, 24)
            Me.chkIncludeLineTerminator.TabIndex = 7
            Me.chkIncludeLineTerminator.Text = "Include Line Terminator"
            '
            'frmExportToText
            '
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.ClientSize = New System.Drawing.Size(584, 176)
            Me.ControlBox = False
            Me.Controls.Add(Me.chkIncludeLineTerminator)
            Me.Controls.Add(Me.cmdBrowse)
            Me.Controls.Add(Me.txtExportFile)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.chkIncludeFieldNames)
            Me.Controls.Add(Me.GroupBox1)
            Me.Controls.Add(Me.cmdCancel)
            Me.Controls.Add(Me.cmdOK)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MinimumSize = New System.Drawing.Size(590, 200)
            Me.Name = "frmExportToText"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "Export Grid Data To Text File..."
            Me.GroupBox1.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

#End Region

#Region " Object History "

        ' ----------|----------|----------|-----------------------------------------------------
        '    DATE   |   WHO    |   WHAT   |  Description of WHAT
        ' ----------|----------|----------|-----------------------------------------------------

#End Region

#Region " Declarations "

        Private _delimiter As String = ","

        Private _filename As String

        Private _includeFieldNames As Boolean = True
        Private _includeLineTerminator As Boolean = True

#End Region

#Region " Event Templates "

        Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click

            Try
                Dim openFile As New System.Windows.Forms.SaveFileDialog

                openFile.InitialDirectory = Environment.CurrentDirectory
                openFile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"

                openFile.DefaultExt = "txt"

                If openFile.ShowDialog() = DialogResult.OK Then

                    Me.txtExportFile.Text = openFile.FileName

                End If

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "frmExportToText.cmdBrowse_Click Error...")
            End Try
        End Sub

        Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click

            If Me.rbOther.Checked Then
                If Me.txtOther.Text.Trim = "" Then
                    MsgBox("You have selected Other as your delimiter but you did not specify what the delimiter should be! " & _
                           "Please correct this before proceeding!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, _
                           "Export To Text Error...")
                    Me.DialogResult = DialogResult.None
                    Exit Sub
                End If
            End If

            If Me.txtExportFile.Text = "" Then
                MsgBox("You must select the file to export the data to! " & _
                       "Please correct this before proceeding!", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, _
                       "Export To Text Error...")
                Me.DialogResult = DialogResult.None
                Exit Sub
            End If

            If Me.rbTab.Checked Then
                Me._delimiter = vbTab
            ElseIf Me.rbSemicolon.Checked Then
                Me._delimiter = ";"
            ElseIf Me.rbComma.Checked Then
                Me._delimiter = ","
            ElseIf Me.rbSpace.Checked Then
                Me._delimiter = " "
            ElseIf Me.rbOther.Checked Then
                Me._delimiter = Me.txtOther.Text
            End If

            Me._filename = Me.txtExportFile.Text
            Me._includeFieldNames = Me.chkIncludeFieldNames.CheckState
            Me._includeLineTerminator = Me.chkIncludeLineTerminator.CheckState

        End Sub

#End Region

#Region " Private Methods "

#End Region

#Region " Private Types"

#End Region

#Region " Public Methods "

#End Region

#Region " Public Properties "

        Public Property Delimiter() As String
            Get
                Return _delimiter
            End Get
            Set(ByVal Value As String)
                _delimiter = Value
            End Set
        End Property

        Public Property Filename() As String
            Get
                Return _filename
            End Get
            Set(ByVal Value As String)
                _filename = Value
            End Set
        End Property

        Public Property IncludeFieldNames() As Boolean
            Get
                Return _includeFieldNames
            End Get
            Set(ByVal Value As Boolean)
                _includeFieldNames = Value
            End Set
        End Property

        Public Property IncludeLineTerminator() As Boolean
            Get
                Return _includeLineTerminator
            End Get
            Set(ByVal Value As Boolean)
                _includeLineTerminator = Value
            End Set
        End Property

#End Region

#Region " Public Types "

#End Region

    End Class

    Private Class frmExportingToExcelWorking
        Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.Label1 = New System.Windows.Forms.Label
            Me.Label2 = New System.Windows.Forms.Label
            Me.SuspendLayout()
            '
            'Label1
            '
            Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label1.Location = New System.Drawing.Point(8, 8)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(288, 32)
            Me.Label1.TabIndex = 0
            Me.Label1.Text = "Sending data to Excel."
            '
            'Label2
            '
            Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.Label2.Location = New System.Drawing.Point(80, 60)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(376, 32)
            Me.Label2.TabIndex = 1
            Me.Label2.Text = "This may take a few moments..."
            Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'frmExportingToExcelWorking
            '
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.BackColor = System.Drawing.Color.AntiqueWhite
            Me.ClientSize = New System.Drawing.Size(516, 157)
            Me.ControlBox = False
            Me.Controls.Add(Me.Label2)
            Me.Controls.Add(Me.Label1)
            Me.Name = "frmExportingToExcelWorking"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Exporting....."
            Me.TopMost = True
            Me.ResumeLayout(False)

        End Sub

        Public Sub UpdateDisplay(ByVal msg As String)
            Label2.Text = msg
            Label2.Refresh()
            Application.DoEvents()
        End Sub

#End Region

    End Class

    Private Class frmSearchInColumn
        Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

        Public Sub New(ByVal Loc As Point)
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

            Me.StartPosition = FormStartPosition.Manual
            Me.Location = Loc

        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents btnCancel As System.Windows.Forms.Button
        Friend WithEvents btnSearch As System.Windows.Forms.Button
        Friend WithEvents txtSearchItem As System.Windows.Forms.TextBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.btnCancel = New System.Windows.Forms.Button
            Me.btnSearch = New System.Windows.Forms.Button
            Me.txtSearchItem = New System.Windows.Forms.TextBox
            Me.Label1 = New System.Windows.Forms.Label
            Me.SuspendLayout()
            '
            'btnCancel
            '
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.Location = New System.Drawing.Point(452, 12)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(88, 24)
            Me.btnCancel.TabIndex = 1
            Me.btnCancel.Text = "Cancel"
            '
            'btnSearch
            '
            Me.btnSearch.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnSearch.Location = New System.Drawing.Point(452, 44)
            Me.btnSearch.Name = "btnSearch"
            Me.btnSearch.Size = New System.Drawing.Size(88, 24)
            Me.btnSearch.TabIndex = 2
            Me.btnSearch.Text = "Search"
            '
            'txtSearchItem
            '
            Me.txtSearchItem.Location = New System.Drawing.Point(24, 28)
            Me.txtSearchItem.Name = "txtSearchItem"
            Me.txtSearchItem.Size = New System.Drawing.Size(360, 20)
            Me.txtSearchItem.TabIndex = 0
            Me.txtSearchItem.Text = ""
            '
            'Label1
            '
            Me.Label1.Location = New System.Drawing.Point(28, 52)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(284, 16)
            Me.Label1.TabIndex = 3
            Me.Label1.Text = "Enter the text you wish to search for..."
            '
            'frmSearchInColumn
            '
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.BackColor = System.Drawing.Color.AntiqueWhite
            Me.ClientSize = New System.Drawing.Size(552, 90)
            Me.Controls.Add(Me.Label1)
            Me.Controls.Add(Me.txtSearchItem)
            Me.Controls.Add(Me.btnSearch)
            Me.Controls.Add(Me.btnCancel)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
            Me.Name = "frmSearchInColumn"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
            Me.Text = "Search for something in this column named..."
            Me.ResumeLayout(False)

        End Sub

#End Region

        Private _Canceled As Boolean = True
        Private _SearchText As String = ""
        Private _ColumnName As String = ""

        Public Property Canceled() As Boolean
            Get
                Return _Canceled
            End Get
            Set(ByVal Value As Boolean)
                _Canceled = Value
            End Set
        End Property

        Public Property SearchText() As String
            Get
                Return _SearchText
            End Get
            Set(ByVal Value As String)
                _SearchText = Value
                Me.txtSearchItem.Text = Value
            End Set
        End Property

        Public Property ColumnName() As String
            Get
                Return _ColumnName
            End Get
            Set(ByVal Value As String)
                _ColumnName = Value
                Me.Text = "Search for something in this column named..." & Value
            End Set
        End Property

        Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
            _Canceled = True
            Me.Hide()
        End Sub

        Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
            _Canceled = False
            _SearchText = txtSearchItem.Text
            Me.Hide()
        End Sub

    End Class

    Private Class SortItem
        Private _ItemVisual As String
        Private _ItemOrdinal As Integer

        Public Property Ivis() As String
            Get
                Return _ItemVisual
            End Get
            Set(ByVal Value As String)
                _ItemVisual = Value
            End Set
        End Property

        Public Property Iord() As Integer
            Get
                Return _ItemOrdinal
            End Get
            Set(ByVal Value As Integer)
                _ItemOrdinal = Value
            End Set
        End Property

    End Class

    Private Class EditColumnRestrictor
        Private _ColumnID As Integer
        Private _RestrictorList As String

        Public Property ColumnID() As Integer
            Get
                Return _ColumnID
            End Get
            Set(ByVal Value As Integer)
                _ColumnID = Value
            End Set
        End Property

        Public Property RestrictedList() As String
            Get
                Return _RestrictorList
            End Get
            Set(ByVal Value As String)
                _RestrictorList = Value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return _ColumnID.ToString()
        End Function

    End Class

    Private Class TearAwayWindowEntry
        Private _columnID As Integer
        Private _Winform As frmColumnTearAway

        Public Property ColID() As Integer
            Get
                Return _columnID
            End Get
            Set(ByVal Value As Integer)
                _columnID = Value
            End Set
        End Property

        Public Property Winform() As frmColumnTearAway
            Get
                Return _Winform
            End Get
            Set(ByVal Value As frmColumnTearAway)
                _Winform = Value
            End Set
        End Property

        Public Sub KillTearAway()
            If _Winform Is Nothing Then
                Exit Sub
            End If

            _Winform.Close()

        End Sub

        Public Sub HideTearAway()
            If _Winform Is Nothing Then
                Exit Sub
            End If

            _Winform.Hide()
        End Sub

        Public Sub ShowTearAway()
            If _Winform Is Nothing Then
                Exit Sub
            End If

            _Winform.Show()
        End Sub

        Public Sub SetTearAwayScrollParameters(ByVal min As Integer, ByVal max As Integer, ByVal visible As Boolean)
            If _Winform Is Nothing Then
                Exit Sub
            End If

            _Winform.vscroller.Visible = visible
            _Winform.vscroller.Minimum = min
            _Winform.vscroller.Maximum = max

        End Sub

        Public Sub SetTearAwayScrollIndex(ByVal index As Integer)
            If _Winform Is Nothing Then
                Exit Sub
            End If

            If index >= _Winform.vscroller.Minimum And index <= _Winform.vscroller.Maximum Then

                _Winform.vscroller.Value = index

            End If

        End Sub

    End Class

#End Region

End Class

