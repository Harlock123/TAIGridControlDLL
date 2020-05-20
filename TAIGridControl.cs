using System.Data;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Collections;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;
using Excel = Microsoft.Office.Interop.Excel;
//using DocumentFormat;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;

using ClosedXML.Excel;
//using DocumentFormat.OpenXml.Spreadsheet;

namespace TAIGridControl2
{

    // 
    // TAIGRIDcontrol.cs
    // Lonnie Allen Watson
    // May 6th 2020
    // 
    // A Little Back Story from early 2000's
    //
    // Being sick and tired of the crappy implimentation of the grid control of VB.net
    // I developed this grid to allow easier programatic access, The current databound grid
    // works better if its used in a databound way, If you want to use the grid in a manner simillar
    // to the way it was used under VB6 you are out of luck. This grid will even expose doubleclick
    // events directly on the cell ( Wow what a concept )
    // 
    // Version 2.0.0.0 First Version in C#


    public class TAIGridControl : UserControl
    {
        public bool _LoggingEnabled = false;

        // Items for the grid Title
        private string _GridTitle = "Grid Title";
        private bool _GridTitleVisible = true;
        private Font _GridTitleFont = new Font("Arial", 16, FontStyle.Regular, GraphicsUnit.Point);
        private int _GridTitleHeight = 16;
        private Color _GridTitleBackcolor = Color.Blue;
        private Color _GridTitleForeColor = Color.White;
        private Point _GridSize;

        // Items for the grid Header
        private string[] _GridHeader = new string[2];
        private Font _GridHeaderFont = new Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point);
        private bool _GridHeaderVisible = true;
        private Color _GridHeaderBackcolor = Color.LightBlue;
        private Color _GridHeaderForecolor = Color.Black;
        private int _GridHeaderHeight = 16;
        private StringFormat _GridHeaderStringFormat = new StringFormat();

        private string[,] _grid = new string[2, 2];
        private int[,] _gridBackColor = new int[2, 2];
        private Brush[] _gridBackColorList = new Brush[2];
        private int[,] _gridCellAlignment = new int[2, 2];
        private StringFormat[] _gridCellAlignmentList = new StringFormat[2];
        private int[,] _gridCellFonts = new int[2, 2];
        private Font[] _gridCellFontsList = new Font[2];
        private int[,] _gridForeColor = new int[2, 2];
        private Pen[] _gridForeColorList = new Pen[2];
        private string[] _colPasswords = new string[2];
        private int[] _colMaxCharacters = new int[2];
        private bool _CellOutlines = true;
        private Color _CellOutlineColor = Color.Black;
        private ContextMenu _OldContextMenu;

        private int[] _colwidths = new int[2];
        private bool[] _colEditable = new bool[2];
        private bool[] _rowEditable = new bool[2];
        private Color _colEditableTextBackColor = Color.Yellow;
        private ArrayList _colEditRestrictions = new ArrayList();
        // ' Private _rowEditRestrictions As New ArrayList
        private int[] _coloffsets = new int[2];
        private bool[] _colhidden = new bool[2];
        private bool[] _colboolean = new bool[2];
        private int[] _rowheights = new int[2];
        private int[] _rowoffsets = new int[2];
        private bool[] _rowhidden = new bool[2];

        private int[] _ColWidthsBeforeAutoSize;
        private int[] _RowHeightsBeforeAutoSize;
        private int[] _ColWidthsAfterAutoSize;
        private int[] _RowHeightsAfterAutoSize;

        private int _rows = 0;
        private int _cols = 0;

        private int _LMouseX = -1;
        private int _LMouseY = -1;

        private bool _AllowPopupMenu = true;
        private bool _AllowInGridEdits = false;
        private bool _AllowRowSelection = true;

        private bool _AllowTearAwayFuncionality = true;
        private bool _AllowExcelFunctionality = true;
        private bool _AllowTextFunctionality = true;
        private bool _AllowHTMLFunctionality = true;
        private bool _AllowSQLScriptFunctionality = true;
        private bool _AllowPrintFunctionality = true;
        private bool _AllowMathFunctionality = true;
        private bool _AllowSettingsFunctionality = true;
        private bool _AllowSortFunctionality = true;
        private bool _AllowFormatFunctionality = true;
        private bool _AllowRowAndColumnFunctionality = true;

        private bool _EditMode = false;    // set whenever a edit session is possible on a textbox or combobox
                                           // cleared when that textbox or combobox loses focus
        private int _EditModeRow = -1;
        private int _EditModeCol = -1;
        private bool _AllowControlKeyMenuPopup = true;
        private bool _AllowColumnSelection = true;
        private bool _AllowMultipleRowSelections = true;
        private bool _AllowWhiteSpaceInCells = true;
        private bool _antialias = false;
        private Color _alternateColorationALTColor = Color.MediumSpringGreen;
        private Color _alternateColorationBaseColor = Color.AntiqueWhite;
        private bool _alternateColorationMode = false;
        private bool _AutoFocus = false;
        private int _DefaultColWidth = 50;
        private int _DefaultRowHeight = 14;
        private Color _DefaultBackColor = Color.AntiqueWhite;
        private Color _DefaultForeColor = Color.Black;
        private Font _DefaultCellFont = new Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Point);
        private StringFormat _DefaultStringFormat = new StringFormat();
        private bool _AutosizeCellsToContents = false;
        private bool _AutoSizeAlreadyCalculated = false;
        private bool _AutoSizeSemaphore = true;
        private bool _Painting = false;
        private bool _TearAwayWork = false;
        private int _dataBaseTimeOut = 500;
        private bool _omitNulls = false;
        private int _MouseWheelScrollAmount = 10;
        private int _RowClicked = -1;
        private int _ColClicked = -1;
        private int _SelectedRow = -1;
        private int _ShiftMultiSelectSelectedRowCrap = -1;
        private ArrayList _SelectedRows = new ArrayList();
        private int _SelectedColumn = -1;
        private bool _ShowDatesWithTime = false;
        private bool _ShowProgressBar = true;
        private bool _ShowExcelExportMessage = true;
        private Color _RowHighLiteBackColor = Color.Blue;
        private Color _RowHighLiteForeColor = Color.White;
        private Color _ColHighliteBackColor = Color.MediumSlateBlue;
        private Color _ColHighliteForeColor = Color.LightGray;
        private Color _BorderColor = Color.Black;
        private int _ScrollBarWeight = 14;
        private BorderStyle _BorderStyle = BorderStyle.FixedSingle;
        private int _MaxRowsSelected = 0;
        private int _PaginationSize;
        private int _scrollinterval = 5;
        private string _LastSearchText = "";
        private int _LastSearchColumn = -1;
        private bool _DoubleClickSemaphore = false;

        // items for the export to text
        private string _delimiter = ",";
        private bool _includeFieldNames = true;
        private bool _includeLineTerminator = true;

        // excel constants
        public const int xlPortrait = 1;
        public const int xlLandscape = 2;
        private const int xlAutomatic = -4105;
        private const int xlContinuous = 1;
        private const int xlThin = 2;
        private const int xlEdgeLeft = 7;
        private const int xlEdgeTop = 8;
        private const int xlEdgeBottom = 9;
        private const int xlEdgeRight = 10;
        private const int xlInsideVertical = 11;
        private const int xlInsideHorizontal = 12;
        private const int xlCenter = -4108;
        private const int xlTop = -4160;
        private const int xlToRight = -4161;
        private const int xlNormalView = 1;
        private const int xlPageBreakPreview = 2;
        private const int xlMaximized = -4137;

        // items for the export to excel
        private string _excelFilename = "";
        private string _excelWorkSheetName = "Grid Output";
        private bool _excelKeepAlive = true;
        private int _excelPageOrientation = xlPortrait;
        private bool _excelPageFit = true;
        private bool _excelIncludeColumnHeaders = true;
        private bool _excelShowBorders = false;
        private bool _excelMaximized = true;
        private bool _excelAutoFitRow = true;
        private bool _excelAutoFitColumn = true;
        private Color _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
        private bool _excelUseAlternateRowColor = true;
        private bool _excelMatchGridColorScheme = true;
        private bool _excelOutlineCells = true;
        private int _excelMaxRowsPerSheet = 30000;

        // items to support user resizing of columns 
        private bool _MouseDownOnHeader = false;
        private int _ColOverOnMouseDown = -1;
        private int _RowOverOnMouseDown = -1;
        private bool _AllowUserColumnResizing = true;
        private int _LastMouseY = 0;
        private int _LastMouseX = 0;
        private int _UserColResizeMinimum = 5;

        // items fo the export to xml
        private string _xmlFilename = "";
        private string _xmlDataSetName = "Grid_Output";
        private string _xmlNameSpace = "TAI_Grid_Ouptut";
        private string _xmlTableName = "Table";
        private bool _xmlIncludeSchema = false;

        // Support menuing
        private int _ColOverOnMenuButton = -1;
        private int _RowOverOnMenuButton = -1;

        // support for report printing

        private string _gridReportTitle = "";
        private bool _gridReportMatchColors = true;
        private bool _gridReportOutlineCells = true;
        private bool _gridReportPreviewFirst = true;
        private bool _gridReportNumberPages = true;
        private bool _gridReportOrientLandscape = false;
        private float _gridReportScaleFactor = Conversions.ToSingle(1.0);

        private int _gridReportPageNumbers = -1;
        private int _gridReportCurrentrow = -1;
        private int _gridReportCurrentColumn = -1;
        private DateTime _gridReportPrintedOn = DateAndTime.Now;

        private int _gridStartPage = -1;
        private int _gridEndPage = -1;
        private bool _gridPrintingAllPages = true;
        private int _gridStartPageRow = -1;

        // Private _psets As System.Drawing.Printing.PageSettings = New System.Drawing.Printing.PageSettings

        // Private _OriginalPrinterName As String = _psets.PrinterSettings.PrinterName

        // Private _image As System.Drawing.Bitmap

        // Private WithEvents _PageSetupForm As New frmPageSetup(_psets)

        private System.Drawing.Printing.PageSettings _psets;

        private string _OriginalPrinterName = "";

        private Bitmap _image;

        private frmPageSetup __PageSetupForm;

        private frmPageSetup _PageSetupForm
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return __PageSetupForm;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (__PageSetupForm != null)
                {
                    __PageSetupForm.OrientationChanged -= PageOrientationChange;
                    __PageSetupForm.PageSizeChanged -= PageSetupChange;
                    __PageSetupForm.PaperMetricsHaveChanged -= PageMetricsChange;
                }

                __PageSetupForm = value;
                if (__PageSetupForm != null)
                {
                    __PageSetupForm.OrientationChanged += PageOrientationChange;
                    __PageSetupForm.PageSizeChanged += PageSetupChange;
                    __PageSetupForm.PaperMetricsHaveChanged += PageMetricsChange;
                }
            }
        }

        // Private WithEvents _PageSetupForm As Object

        // 'Private WithEvents TearItem As New frmColumnTearAway

        private ArrayList TearAways = new ArrayList();
       
        /// <summary>
        /// Denotes the form of action necessary to be taken to have a cell in editmode actually have its value
        /// change. Fireing the cell edited event. Either having the user press the enter/return key or having the
        /// user shift focus to another control or cell in the grid itself.
        /// </summary>
        /// <remarks></remarks>
        public enum GridEditModes
        {
            KeyReturn = 0,
            LostFocus = 1
        }

        private GridEditModes _GridEditMode = GridEditModes.KeyReturn;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        public TAIGridControl() : base()
        {

            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call
            txtHandler.Height = 1;
            txtHandler.Width = 1;
            txtHandler.Left = 0;
            txtHandler.Top = 0;
            txtHandler.BackColor = _BorderColor;
        }

        // UserControl overrides dispose to clean up the component list.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!(components == null))
                    components.Dispose();
            }

            base.Dispose(disposing);
        }

        // Required by the Windows Form Designer
        private IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        private HScrollBar _hs;

        internal HScrollBar hs
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _hs;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_hs != null)
                {
                    _hs.ValueChanged -= hs_ValueChanged;
                }

                _hs = value;
                if (_hs != null)
                {
                    _hs.ValueChanged += hs_ValueChanged;
                }
            }
        }

        private VScrollBar _vs;

        internal VScrollBar vs
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _vs;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_vs != null)
                {
                    _vs.ValueChanged -= vs_ValueChanged;
                    _vs.Scroll -= vs_Scroll;
                }

                _vs = value;
                if (_vs != null)
                {
                    _vs.ValueChanged += vs_ValueChanged;
                    _vs.Scroll += vs_Scroll;
                }
            }
        }

        private TextBox _txtHandler;

        internal TextBox txtHandler
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtHandler;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtHandler != null)
                {
                    _txtHandler.KeyPress -= txtHandler_KeyPress;
                    _txtHandler.KeyDown -= txtHandler_KeyDown;
                }

                _txtHandler = value;
                if (_txtHandler != null)
                {
                    _txtHandler.KeyPress += txtHandler_KeyPress;
                    _txtHandler.KeyDown += txtHandler_KeyDown;
                }
            }
        }

        #region Context Menu Items

        private ContextMenu _menu;

        internal ContextMenu menu
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _menu;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_menu != null)
                {
                    _menu.Popup -= menu_Popup;
                }

                _menu = value;
                if (_menu != null)
                {
                    _menu.Popup += menu_Popup;
                }
            }
        }

        private MenuItem _miStats;

        internal MenuItem miStats
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miStats;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miStats != null)
                {
                }

                _miStats = value;
                if (_miStats != null)
                {
                }
            }
        }

        private MenuItem _miExportToExcelMenu;

        internal MenuItem miExportToExcelMenu
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miExportToExcelMenu;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miExportToExcelMenu != null)
                {
                }

                _miExportToExcelMenu = value;
                if (_miExportToExcelMenu != null)
                {
                }
            }
        }

        private MenuItem _miExportToExcel;

        internal MenuItem miExportToExcel
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miExportToExcel;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miExportToExcel != null)
                {
                    _miExportToExcel.Click -= miExportToExcel_Click_1;
                }

                _miExportToExcel = value;
                if (_miExportToExcel != null)
                {
                    _miExportToExcel.Click += miExportToExcel_Click_1;
                }
            }
        }

        private MenuItem _miAutoFitCols;

        internal MenuItem miAutoFitCols
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miAutoFitCols;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miAutoFitCols != null)
                {
                    _miAutoFitCols.Click -= miAutoFitCols_Click;
                }

                _miAutoFitCols = value;
                if (_miAutoFitCols != null)
                {
                    _miAutoFitCols.Click += miAutoFitCols_Click;
                }
            }
        }

        private MenuItem _miAutoFitRows;

        internal MenuItem miAutoFitRows
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miAutoFitRows;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miAutoFitRows != null)
                {
                    _miAutoFitRows.Click -= miAutoFitRows_Click;
                }

                _miAutoFitRows = value;
                if (_miAutoFitRows != null)
                {
                    _miAutoFitRows.Click += miAutoFitRows_Click;
                }
            }
        }

        private MenuItem _miALternateRowColors;

        internal MenuItem miALternateRowColors
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miALternateRowColors;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miALternateRowColors != null)
                {
                    _miALternateRowColors.Click -= miALternateRowColors_Click;
                }

                _miALternateRowColors = value;
                if (_miALternateRowColors != null)
                {
                    _miALternateRowColors.Click += miALternateRowColors_Click;
                }
            }
        }

        private MenuItem _miMatchGridColors;

        internal MenuItem miMatchGridColors
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMatchGridColors;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMatchGridColors != null)
                {
                    _miMatchGridColors.Click -= miMatchGridColors_Click;
                }

                _miMatchGridColors = value;
                if (_miMatchGridColors != null)
                {
                    _miMatchGridColors.Click += miMatchGridColors_Click;
                }
            }
        }

        private MenuItem _miOutlineExportedCells;

        internal MenuItem miOutlineExportedCells
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miOutlineExportedCells;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miOutlineExportedCells != null)
                {
                    _miOutlineExportedCells.Click -= miOutlineExportedCells_Click;
                }

                _miOutlineExportedCells = value;
                if (_miOutlineExportedCells != null)
                {
                    _miOutlineExportedCells.Click += miOutlineExportedCells_Click;
                }
            }
        }

        private MenuItem _miFormatStuff;

        internal MenuItem miFormatStuff
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFormatStuff;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFormatStuff != null)
                {
                }

                _miFormatStuff = value;
                if (_miFormatStuff != null)
                {
                }
            }
        }

        private MenuItem _miFormatAsMoney;

        internal MenuItem miFormatAsMoney
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFormatAsMoney;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFormatAsMoney != null)
                {
                    _miFormatAsMoney.Click -= miFormatAsMoney_Click;
                }

                _miFormatAsMoney = value;
                if (_miFormatAsMoney != null)
                {
                    _miFormatAsMoney.Click += miFormatAsMoney_Click;
                }
            }
        }

        private MenuItem _miFormatAsDecimal;

        internal MenuItem miFormatAsDecimal
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFormatAsDecimal;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFormatAsDecimal != null)
                {
                    _miFormatAsDecimal.Click -= miFormatAsDecimal_Click;
                }

                _miFormatAsDecimal = value;
                if (_miFormatAsDecimal != null)
                {
                    _miFormatAsDecimal.Click += miFormatAsDecimal_Click;
                }
            }
        }

        private MenuItem _miFormatAsText;

        internal MenuItem miFormatAsText
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFormatAsText;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFormatAsText != null)
                {
                    _miFormatAsText.Click -= miFormatAsText_Click;
                }

                _miFormatAsText = value;
                if (_miFormatAsText != null)
                {
                    _miFormatAsText.Click += miFormatAsText_Click;
                }
            }
        }

        private MenuItem _miCenter;

        internal MenuItem miCenter
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miCenter;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miCenter != null)
                {
                    _miCenter.Click -= miCenter_Click;
                }

                _miCenter = value;
                if (_miCenter != null)
                {
                    _miCenter.Click += miCenter_Click;
                }
            }
        }

        private MenuItem _miLeft;

        internal MenuItem miLeft
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miLeft;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miLeft != null)
                {
                    _miLeft.Click -= miLeft_Click;
                }

                _miLeft = value;
                if (_miLeft != null)
                {
                    _miLeft.Click += miLeft_Click;
                }
            }
        }

        private MenuItem _miRight;

        internal MenuItem miRight
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miRight;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miRight != null)
                {
                    _miRight.Click -= miRight_Click;
                }

                _miRight = value;
                if (_miRight != null)
                {
                    _miRight.Click += miRight_Click;
                }
            }
        }

        private MenuItem _miFontsSmaller;

        internal MenuItem miFontsSmaller
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFontsSmaller;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFontsSmaller != null)
                {
                    _miFontsSmaller.Click -= miFontsSmaller_Click;
                }

                _miFontsSmaller = value;
                if (_miFontsSmaller != null)
                {
                    _miFontsSmaller.Click += miFontsSmaller_Click;
                }
            }
        }

        private MenuItem _miFontsLarger;

        internal MenuItem miFontsLarger
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miFontsLarger;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miFontsLarger != null)
                {
                    _miFontsLarger.Click -= miFontsLarger_Click;
                }

                _miFontsLarger = value;
                if (_miFontsLarger != null)
                {
                    _miFontsLarger.Click += miFontsLarger_Click;
                }
            }
        }

        private MenuItem _miSmoothing;

        internal MenuItem miSmoothing
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSmoothing;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSmoothing != null)
                {
                    _miSmoothing.Click -= miSmoothing_Click;
                }

                _miSmoothing = value;
                if (_miSmoothing != null)
                {
                    _miSmoothing.Click += miSmoothing_Click;
                }
            }
        }

        private ProgressBar _pBar;

        internal ProgressBar pBar
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _pBar;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_pBar != null)
                {
                }

                _pBar = value;
                if (_pBar != null)
                {
                }
            }
        }

        private MenuItem _miExportToTextFile;

        internal MenuItem miExportToTextFile
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miExportToTextFile;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miExportToTextFile != null)
                {
                    _miExportToTextFile.Click -= miExportToTextFile_Click;
                }

                _miExportToTextFile = value;
                if (_miExportToTextFile != null)
                {
                    _miExportToTextFile.Click += miExportToTextFile_Click;
                }
            }
        }

        private MenuItem _MenuItem2;

        internal MenuItem MenuItem2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem2 != null)
                {
                }

                _MenuItem2 = value;
                if (_MenuItem2 != null)
                {
                }
            }
        }

        private MenuItem _miHeaderFontLarger;

        internal MenuItem miHeaderFontLarger
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHeaderFontLarger;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHeaderFontLarger != null)
                {
                    _miHeaderFontLarger.Click -= miHeaderFontLarger_Click;
                }

                _miHeaderFontLarger = value;
                if (_miHeaderFontLarger != null)
                {
                    _miHeaderFontLarger.Click += miHeaderFontLarger_Click;
                }
            }
        }

        private MenuItem _miHeaderFontSmaller;

        internal MenuItem miHeaderFontSmaller
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHeaderFontSmaller;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHeaderFontSmaller != null)
                {
                    _miHeaderFontSmaller.Click -= miHeaderFontSmaller_Click;
                }

                _miHeaderFontSmaller = value;
                if (_miHeaderFontSmaller != null)
                {
                    _miHeaderFontSmaller.Click += miHeaderFontSmaller_Click;
                }
            }
        }

        private MenuItem _miTitleFontLarger;

        internal MenuItem miTitleFontLarger
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miTitleFontLarger;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miTitleFontLarger != null)
                {
                    _miTitleFontLarger.Click -= miTitleFontLarger_Click;
                }

                _miTitleFontLarger = value;
                if (_miTitleFontLarger != null)
                {
                    _miTitleFontLarger.Click += miTitleFontLarger_Click;
                }
            }
        }

        private MenuItem _miTitleFontSmaller;

        internal MenuItem miTitleFontSmaller
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miTitleFontSmaller;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miTitleFontSmaller != null)
                {
                    _miTitleFontSmaller.Click -= miTitleFontSmaller_Click;
                }

                _miTitleFontSmaller = value;
                if (_miTitleFontSmaller != null)
                {
                    _miTitleFontSmaller.Click += miTitleFontSmaller_Click;
                }
            }
        }

        private GroupBox _gb1;

        internal GroupBox gb1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _gb1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_gb1 != null)
                {
                }

                _gb1 = value;
                if (_gb1 != null)
                {
                }
            }
        }

        private MenuItem _miSearchInColumn;

        internal MenuItem miSearchInColumn
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSearchInColumn;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSearchInColumn != null)
                {
                    _miSearchInColumn.Click -= miSearchInColumn_Click;
                }

                _miSearchInColumn = value;
                if (_miSearchInColumn != null)
                {
                    _miSearchInColumn.Click += miSearchInColumn_Click;
                }
            }
        }

        private MenuItem _miAutoSizeToContents;

        internal MenuItem miAutoSizeToContents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miAutoSizeToContents;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miAutoSizeToContents != null)
                {
                    _miAutoSizeToContents.Click -= miAutoSizeToContents_Click;
                }

                _miAutoSizeToContents = value;
                if (_miAutoSizeToContents != null)
                {
                    _miAutoSizeToContents.Click += miAutoSizeToContents_Click;
                }
            }
        }

        private MenuItem _miAllowUserColumnResizing;

        internal MenuItem miAllowUserColumnResizing
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miAllowUserColumnResizing;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miAllowUserColumnResizing != null)
                {
                    _miAllowUserColumnResizing.Click -= miAllowUserColumnResizing_Click;
                }

                _miAllowUserColumnResizing = value;
                if (_miAllowUserColumnResizing != null)
                {
                    _miAllowUserColumnResizing.Click += miAllowUserColumnResizing_Click;
                }
            }
        }

        private MenuItem _MenuItem3;

        internal MenuItem MenuItem3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem3 != null)
                {
                }

                _MenuItem3 = value;
                if (_MenuItem3 != null)
                {
                }
            }
        }

        private MenuItem _miSortAscending;

        internal MenuItem miSortAscending
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSortAscending;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSortAscending != null)
                {
                    _miSortAscending.Click -= miSortAscending_Click;
                }

                _miSortAscending = value;
                if (_miSortAscending != null)
                {
                    _miSortAscending.Click += miSortAscending_Click;
                }
            }
        }

        private MenuItem _miSortDescending;

        internal MenuItem miSortDescending
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSortDescending;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSortDescending != null)
                {
                    _miSortDescending.Click -= miSortDescending_Click;
                }

                _miSortDescending = value;
                if (_miSortDescending != null)
                {
                    _miSortDescending.Click += miSortDescending_Click;
                }
            }
        }

        private MenuItem _MenuItem4;

        internal MenuItem MenuItem4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem4 != null)
                {
                }

                _MenuItem4 = value;
                if (_MenuItem4 != null)
                {
                }
            }
        }

        private MenuItem _miHideRow;

        internal MenuItem miHideRow
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHideRow;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHideRow != null)
                {
                    _miHideRow.Click -= miHideRow_Click;
                }

                _miHideRow = value;
                if (_miHideRow != null)
                {
                    _miHideRow.Click += miHideRow_Click;
                }
            }
        }

        private MenuItem _miHideColumn;

        internal MenuItem miHideColumn
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHideColumn;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHideColumn != null)
                {
                    _miHideColumn.Click -= miHideColumn_Click;
                }

                _miHideColumn = value;
                if (_miHideColumn != null)
                {
                    _miHideColumn.Click += miHideColumn_Click;
                }
            }
        }

        private MenuItem _miSetRowColor;

        internal MenuItem miSetRowColor
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSetRowColor;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSetRowColor != null)
                {
                    _miSetRowColor.Click -= miSetRowColor_Click;
                }

                _miSetRowColor = value;
                if (_miSetRowColor != null)
                {
                    _miSetRowColor.Click += miSetRowColor_Click;
                }
            }
        }

        private MenuItem _miSetColumnColor;

        internal MenuItem miSetColumnColor
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSetColumnColor;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSetColumnColor != null)
                {
                    _miSetColumnColor.Click -= miSetColumnColor_Click;
                }

                _miSetColumnColor = value;
                if (_miSetColumnColor != null)
                {
                    _miSetColumnColor.Click += miSetColumnColor_Click;
                }
            }
        }

        private MenuItem _miSetCellColor;

        internal MenuItem miSetCellColor
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSetCellColor;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSetCellColor != null)
                {
                    _miSetCellColor.Click -= miSetCellColor_Click;
                }

                _miSetCellColor = value;
                if (_miSetCellColor != null)
                {
                    _miSetCellColor.Click += miSetCellColor_Click;
                }
            }
        }

        private MenuItem _miShowAllRowsAndColumns;

        internal MenuItem miShowAllRowsAndColumns
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miShowAllRowsAndColumns;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miShowAllRowsAndColumns != null)
                {
                    _miShowAllRowsAndColumns.Click -= miShowAllRowsAndColumns_Click;
                }

                _miShowAllRowsAndColumns = value;
                if (_miShowAllRowsAndColumns != null)
                {
                    _miShowAllRowsAndColumns.Click += miShowAllRowsAndColumns_Click;
                }
            }
        }

        private ColorDialog _clrdlg;

        internal ColorDialog clrdlg
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _clrdlg;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_clrdlg != null)
                {
                }

                _clrdlg = value;
                if (_clrdlg != null)
                {
                }
            }
        }

        private MenuItem _miDateAsc;

        internal MenuItem miDateAsc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miDateAsc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miDateAsc != null)
                {
                    _miDateAsc.Click -= miDateAsc_Click;
                }

                _miDateAsc = value;
                if (_miDateAsc != null)
                {
                    _miDateAsc.Click += miDateAsc_Click;
                }
            }
        }

        private MenuItem _miDateDesc;

        internal MenuItem miDateDesc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miDateDesc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miDateDesc != null)
                {
                    _miDateDesc.Click -= miDateDesc_Click;
                }

                _miDateDesc = value;
                if (_miDateDesc != null)
                {
                    _miDateDesc.Click += miDateDesc_Click;
                }
            }
        }

        private MenuItem _MenuItem5;

        internal MenuItem MenuItem5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem5 != null)
                {
                }

                _MenuItem5 = value;
                if (_MenuItem5 != null)
                {
                }
            }
        }

        private MenuItem _miSumColumn;

        internal MenuItem miSumColumn
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSumColumn;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSumColumn != null)
                {
                    _miSumColumn.Click -= miSumColumn_Click;
                }

                _miSumColumn = value;
                if (_miSumColumn != null)
                {
                    _miSumColumn.Click += miSumColumn_Click;
                }
            }
        }

        private MenuItem _miSumRow;

        internal MenuItem miSumRow
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSumRow;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSumRow != null)
                {
                    _miSumRow.Click -= miSumRow_Click;
                }

                _miSumRow = value;
                if (_miSumRow != null)
                {
                    _miSumRow.Click += miSumRow_Click;
                }
            }
        }

        private MenuItem _miMaxCol;

        internal MenuItem miMaxCol
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMaxCol;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMaxCol != null)
                {
                    _miMaxCol.Click -= miMaxCol_Click;
                }

                _miMaxCol = value;
                if (_miMaxCol != null)
                {
                    _miMaxCol.Click += miMaxCol_Click;
                }
            }
        }

        private MenuItem _miMaxRow;

        internal MenuItem miMaxRow
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMaxRow;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMaxRow != null)
                {
                    _miMaxRow.Click -= miMaxRow_Click;
                }

                _miMaxRow = value;
                if (_miMaxRow != null)
                {
                    _miMaxRow.Click += miMaxRow_Click;
                }
            }
        }

        private MenuItem _miMinCol;

        internal MenuItem miMinCol
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMinCol;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMinCol != null)
                {
                    _miMinCol.Click -= miMinCol_Click;
                }

                _miMinCol = value;
                if (_miMinCol != null)
                {
                    _miMinCol.Click += miMinCol_Click;
                }
            }
        }

        private MenuItem _miMinRow;

        internal MenuItem miMinRow
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMinRow;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMinRow != null)
                {
                    _miMinRow.Click -= miMinRow_Click;
                }

                _miMinRow = value;
                if (_miMinRow != null)
                {
                    _miMinRow.Click += miMinRow_Click;
                }
            }
        }

        private MenuItem _miColAverage;

        internal MenuItem miColAverage
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miColAverage;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miColAverage != null)
                {
                    _miColAverage.Click -= miColAverage_Click;
                }

                _miColAverage = value;
                if (_miColAverage != null)
                {
                    _miColAverage.Click += miColAverage_Click;
                }
            }
        }

        private MenuItem _miRowAverage;

        internal MenuItem miRowAverage
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miRowAverage;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miRowAverage != null)
                {
                    _miRowAverage.Click -= miRowAverage_Click;
                }

                _miRowAverage = value;
                if (_miRowAverage != null)
                {
                    _miRowAverage.Click += miRowAverage_Click;
                }
            }
        }

        private MenuItem _MenuItem7;

        internal MenuItem MenuItem7
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem7;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem7 != null)
                {
                }

                _MenuItem7 = value;
                if (_MenuItem7 != null)
                {
                }
            }
        }

        private MenuItem _MenuItem8;

        internal MenuItem MenuItem8
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem8;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem8 != null)
                {
                }

                _MenuItem8 = value;
                if (_MenuItem8 != null)
                {
                }
            }
        }

        private MenuItem _MenuItem9;

        internal MenuItem MenuItem9
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem9;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem9 != null)
                {
                }

                _MenuItem9 = value;
                if (_MenuItem9 != null)
                {
                }
            }
        }

        private MenuItem _MenuItem10;

        internal MenuItem MenuItem10
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem10;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem10 != null)
                {
                }

                _MenuItem10 = value;
                if (_MenuItem10 != null)
                {
                }
            }
        }

        private MenuItem _miCopyCellToClipboard;

        internal MenuItem miCopyCellToClipboard
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miCopyCellToClipboard;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miCopyCellToClipboard != null)
                {
                    _miCopyCellToClipboard.Click -= miCopyCellToClipboard_Click;
                }

                _miCopyCellToClipboard = value;
                if (_miCopyCellToClipboard != null)
                {
                    _miCopyCellToClipboard.Click += miCopyCellToClipboard_Click;
                }
            }
        }

        private MenuItem _MenuItem1;

        internal MenuItem MenuItem1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem1 != null)
                {
                }

                _MenuItem1 = value;
                if (_MenuItem1 != null)
                {
                }
            }
        }

        private MenuItem _miPrintTheGrid;

        internal MenuItem miPrintTheGrid
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miPrintTheGrid;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miPrintTheGrid != null)
                {
                    _miPrintTheGrid.Click -= miPrintTheGrid_Click;
                }

                _miPrintTheGrid = value;
                if (_miPrintTheGrid != null)
                {
                    _miPrintTheGrid.Click += miPrintTheGrid_Click;
                }
            }
        }

        private MenuItem _miPreviewTheGrid;

        internal MenuItem miPreviewTheGrid
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miPreviewTheGrid;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miPreviewTheGrid != null)
                {
                    _miPreviewTheGrid.Click -= miPreviewTheGrid_Click;
                }

                _miPreviewTheGrid = value;
                if (_miPreviewTheGrid != null)
                {
                    _miPreviewTheGrid.Click += miPreviewTheGrid_Click;
                }
            }
        }

        private MenuItem _miPageSetup;

        internal MenuItem miPageSetup
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miPageSetup;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miPageSetup != null)
                {
                    _miPageSetup.Click -= miPageSetup_Click;
                }

                _miPageSetup = value;
                if (_miPageSetup != null)
                {
                    _miPageSetup.Click += miPageSetup_Click;
                }
            }
        }

        private TextBox _txtInput;

        internal TextBox txtInput
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _txtInput;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_txtInput != null)
                {
                    _txtInput.Leave -= txtInput_Leave;
                    _txtInput.KeyDown -= txtInput_KeyDown;
                }

                _txtInput = value;
                if (_txtInput != null)
                {
                    _txtInput.Leave += txtInput_Leave;
                    _txtInput.KeyDown += txtInput_KeyDown;
                }
            }
        }

        private ComboBox _cmboInput;

        internal ComboBox cmboInput
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _cmboInput;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_cmboInput != null)
                {
                    _cmboInput.Leave -= cmboInput_Leave;
                    _cmboInput.KeyDown -= cmboInput_keyDown;
                    _cmboInput.SelectedIndexChanged -= cmboInput_SelectedIndexChanged;
                }

                _cmboInput = value;
                if (_cmboInput != null)
                {
                    _cmboInput.Leave += cmboInput_Leave;
                    _cmboInput.KeyDown += cmboInput_keyDown;
                    _cmboInput.SelectedIndexChanged += cmboInput_SelectedIndexChanged;
                }
            }
        }

        private MenuItem _miSortNumericAsc;

        internal MenuItem miSortNumericAsc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSortNumericAsc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSortNumericAsc != null)
                {
                    _miSortNumericAsc.Click -= miSortNumericAsc_Click;
                }

                _miSortNumericAsc = value;
                if (_miSortNumericAsc != null)
                {
                    _miSortNumericAsc.Click += miSortNumericAsc_Click;
                }
            }
        }

        private MenuItem _miSortNumericDesc;

        internal MenuItem miSortNumericDesc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miSortNumericDesc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miSortNumericDesc != null)
                {
                    _miSortNumericDesc.Click -= miSortNumericDesc_Click;
                }

                _miSortNumericDesc = value;
                if (_miSortNumericDesc != null)
                {
                    _miSortNumericDesc.Click += miSortNumericDesc_Click;
                }
            }
        }

        private MenuItem _miExportToSQLScript;

        internal MenuItem miExportToSQLScript
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miExportToSQLScript;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miExportToSQLScript != null)
                {
                    _miExportToSQLScript.Click -= miExportToSQLScript_Click;
                }

                _miExportToSQLScript = value;
                if (_miExportToSQLScript != null)
                {
                    _miExportToSQLScript.Click += miExportToSQLScript_Click;
                }
            }
        }

        private MenuItem _miExportToHTMLTable;

        internal MenuItem miExportToHTMLTable
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miExportToHTMLTable;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miExportToHTMLTable != null)
                {
                    _miExportToHTMLTable.Click -= miExportToHTMLTable_Click;
                }

                _miExportToHTMLTable = value;
                if (_miExportToHTMLTable != null)
                {
                    _miExportToHTMLTable.Click += miExportToHTMLTable_Click;
                }
            }
        }

        private MenuItem _miDisplayFrequencyDistribution;

        internal MenuItem miDisplayFrequencyDistribution
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miDisplayFrequencyDistribution;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miDisplayFrequencyDistribution != null)
                {
                    _miDisplayFrequencyDistribution.Click -= miDisplayFrequencyDistribution_Click;
                }

                _miDisplayFrequencyDistribution = value;
                if (_miDisplayFrequencyDistribution != null)
                {
                    _miDisplayFrequencyDistribution.Click += miDisplayFrequencyDistribution_Click;
                }
            }
        }

        private MenuItem _miProperties;

        internal MenuItem miProperties
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miProperties;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miProperties != null)
                {
                    _miProperties.Click -= miProperties_Click;
                }

                _miProperties = value;
                if (_miProperties != null)
                {
                    _miProperties.Click += miProperties_Click;
                }
            }
        }

        private MenuItem _MenuItem11;

        internal MenuItem MenuItem11
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem11;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem11 != null)
                {
                }

                _MenuItem11 = value;
                if (_MenuItem11 != null)
                {
                }
            }
        }

        private MenuItem _miTearColumnAway;

        internal MenuItem miTearColumnAway
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miTearColumnAway;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miTearColumnAway != null)
                {
                    _miTearColumnAway.Click -= miTearColumnAway_Click;
                }

                _miTearColumnAway = value;
                if (_miTearColumnAway != null)
                {
                    _miTearColumnAway.Click += miTearColumnAway_Click;
                }
            }
        }

        private MenuItem _MenuItem12;

        internal MenuItem MenuItem12
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _MenuItem12;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_MenuItem12 != null)
                {
                }

                _MenuItem12 = value;
                if (_MenuItem12 != null)
                {
                }
            }
        }

        private MenuItem _miHideColumnTearAway;

        internal MenuItem miHideColumnTearAway
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHideColumnTearAway;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHideColumnTearAway != null)
                {
                    _miHideColumnTearAway.Click -= miHideColumnTearAway_Click;
                }

                _miHideColumnTearAway = value;
                if (_miHideColumnTearAway != null)
                {
                    _miHideColumnTearAway.Click += miHideColumnTearAway_Click;
                }
            }
        }

        private MenuItem _miHideAllTearAwayColumns;

        internal MenuItem miHideAllTearAwayColumns
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miHideAllTearAwayColumns;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miHideAllTearAwayColumns != null)
                {
                    _miHideAllTearAwayColumns.Click -= miHideAllTearAwayColumns_Click;
                }

                _miHideAllTearAwayColumns = value;
                if (_miHideAllTearAwayColumns != null)
                {
                    _miHideAllTearAwayColumns.Click += miHideAllTearAwayColumns_Click;
                }
            }
        }
        
        private MenuItem _miMultipleColumnTearAway;

        internal MenuItem miMultipleColumnTearAway
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miMultipleColumnTearAway;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miMultipleColumnTearAway != null)
                {
                    _miMultipleColumnTearAway.Click -= miMultipleColumnTearAway_Click;
                }

                _miMultipleColumnTearAway = value;
                if (_miMultipleColumnTearAway != null)
                {
                    _miMultipleColumnTearAway.Click += miMultipleColumnTearAway_Click;
                }
            }
        }

        private MenuItem _miArrangeTearAways;

        internal MenuItem miArrangeTearAways
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _miArrangeTearAways;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_miArrangeTearAways != null)
                {
                    _miArrangeTearAways.Click -= miArrangeTearAways_Click;
                }

                _miArrangeTearAways = value;
                if (_miArrangeTearAways != null)
                {
                    _miArrangeTearAways.Click += miArrangeTearAways_Click;
                }
            }
        }

        #endregion

        private ToolTip __TTip;

        internal ToolTip _TTip
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return __TTip;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (__TTip != null)
                {
                }

                __TTip = value;
                if (__TTip != null)
                {
                }
            }
        }

        private System.Drawing.Printing.PrintDocument _pdoc;

        internal System.Drawing.Printing.PrintDocument pdoc
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _pdoc;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_pdoc != null)
                {
                    _pdoc.PrintPage -= pdoc_PrintPage;
                }

                _pdoc = value;
                if (_pdoc != null)
                {
                    _pdoc.PrintPage += pdoc_PrintPage;
                }
            }
        }


        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            components = new Container();
            _hs = new HScrollBar();
            _hs.ValueChanged += hs_ValueChanged;
            _vs = new VScrollBar();
            _vs.ValueChanged += vs_ValueChanged;
            _vs.Scroll += vs_Scroll;
            _txtHandler = new TextBox();
            _txtHandler.KeyPress += txtHandler_KeyPress;
            _txtHandler.KeyDown += txtHandler_KeyDown;
            _menu = new ContextMenu();
            _menu.Popup += menu_Popup;
            _miCopyCellToClipboard = new MenuItem();
            _miCopyCellToClipboard.Click += miCopyCellToClipboard_Click;
            _MenuItem7 = new MenuItem();
            _MenuItem5 = new MenuItem();
            _miSumColumn = new MenuItem();
            _miSumColumn.Click += miSumColumn_Click;
            _miSumRow = new MenuItem();
            _miSumRow.Click += miSumRow_Click;
            _miMaxCol = new MenuItem();
            _miMaxCol.Click += miMaxCol_Click;
            _miMaxRow = new MenuItem();
            _miMaxRow.Click += miMaxRow_Click;
            _miMinCol = new MenuItem();
            _miMinCol.Click += miMinCol_Click;
            _miMinRow = new MenuItem();
            _miMinRow.Click += miMinRow_Click;
            _miColAverage = new MenuItem();
            _miColAverage.Click += miColAverage_Click;
            _miRowAverage = new MenuItem();
            _miRowAverage.Click += miRowAverage_Click;
            _miDisplayFrequencyDistribution = new MenuItem();
            _miDisplayFrequencyDistribution.Click += miDisplayFrequencyDistribution_Click;
            _MenuItem9 = new MenuItem();
            _miExportToExcelMenu = new MenuItem();
            _miExportToExcel = new MenuItem();
            _miExportToExcel.Click += miExportToExcel_Click_1;
            _miAutoFitCols = new MenuItem();
            _miAutoFitCols.Click += miAutoFitCols_Click;
            _miAutoFitRows = new MenuItem();
            _miAutoFitRows.Click += miAutoFitRows_Click;
            _miALternateRowColors = new MenuItem();
            _miALternateRowColors.Click += miALternateRowColors_Click;
            _miMatchGridColors = new MenuItem();
            _miMatchGridColors.Click += miMatchGridColors_Click;
            _miOutlineExportedCells = new MenuItem();
            _miOutlineExportedCells.Click += miOutlineExportedCells_Click;
            _miExportToTextFile = new MenuItem();
            _miExportToTextFile.Click += miExportToTextFile_Click;
            _miExportToHTMLTable = new MenuItem();
            _miExportToHTMLTable.Click += miExportToHTMLTable_Click;
            _miExportToSQLScript = new MenuItem();
            _miExportToSQLScript.Click += miExportToSQLScript_Click;
            _MenuItem8 = new MenuItem();
            _miFormatStuff = new MenuItem();
            _miFormatAsMoney = new MenuItem();
            _miFormatAsMoney.Click += miFormatAsMoney_Click;
            _miFormatAsDecimal = new MenuItem();
            _miFormatAsDecimal.Click += miFormatAsDecimal_Click;
            _miFormatAsText = new MenuItem();
            _miFormatAsText.Click += miFormatAsText_Click;
            _miCenter = new MenuItem();
            _miCenter.Click += miCenter_Click;
            _miLeft = new MenuItem();
            _miLeft.Click += miLeft_Click;
            _miRight = new MenuItem();
            _miRight.Click += miRight_Click;
            _MenuItem2 = new MenuItem();
            _miFontsSmaller = new MenuItem();
            _miFontsSmaller.Click += miFontsSmaller_Click;
            _miFontsLarger = new MenuItem();
            _miFontsLarger.Click += miFontsLarger_Click;
            _miHeaderFontSmaller = new MenuItem();
            _miHeaderFontSmaller.Click += miHeaderFontSmaller_Click;
            _miHeaderFontLarger = new MenuItem();
            _miHeaderFontLarger.Click += miHeaderFontLarger_Click;
            _miTitleFontSmaller = new MenuItem();
            _miTitleFontSmaller.Click += miTitleFontSmaller_Click;
            _miTitleFontLarger = new MenuItem();
            _miTitleFontLarger.Click += miTitleFontLarger_Click;
            _miSmoothing = new MenuItem();
            _miSmoothing.Click += miSmoothing_Click;
            _miAutoSizeToContents = new MenuItem();
            _miAutoSizeToContents.Click += miAutoSizeToContents_Click;
            _miAllowUserColumnResizing = new MenuItem();
            _miAllowUserColumnResizing.Click += miAllowUserColumnResizing_Click;
            _MenuItem3 = new MenuItem();
            _miSortAscending = new MenuItem();
            _miSortAscending.Click += miSortAscending_Click;
            _miSortDescending = new MenuItem();
            _miSortDescending.Click += miSortDescending_Click;
            _miDateAsc = new MenuItem();
            _miDateAsc.Click += miDateAsc_Click;
            _miDateDesc = new MenuItem();
            _miDateDesc.Click += miDateDesc_Click;
            _miSortNumericAsc = new MenuItem();
            _miSortNumericAsc.Click += miSortNumericAsc_Click;
            _miSortNumericDesc = new MenuItem();
            _miSortNumericDesc.Click += miSortNumericDesc_Click;
            _MenuItem4 = new MenuItem();
            _miHideRow = new MenuItem();
            _miHideRow.Click += miHideRow_Click;
            _miHideColumn = new MenuItem();
            _miHideColumn.Click += miHideColumn_Click;
            _miShowAllRowsAndColumns = new MenuItem();
            _miShowAllRowsAndColumns.Click += miShowAllRowsAndColumns_Click;
            _miSetRowColor = new MenuItem();
            _miSetRowColor.Click += miSetRowColor_Click;
            _miSetColumnColor = new MenuItem();
            _miSetColumnColor.Click += miSetColumnColor_Click;
            _miSetCellColor = new MenuItem();
            _miSetCellColor.Click += miSetCellColor_Click;
            _MenuItem10 = new MenuItem();
            _miSearchInColumn = new MenuItem();
            _miSearchInColumn.Click += miSearchInColumn_Click;
            _MenuItem1 = new MenuItem();
            _miPrintTheGrid = new MenuItem();
            _miPrintTheGrid.Click += miPrintTheGrid_Click;
            _miPreviewTheGrid = new MenuItem();
            _miPreviewTheGrid.Click += miPreviewTheGrid_Click;
            _miPageSetup = new MenuItem();
            _miPageSetup.Click += miPageSetup_Click;
            _MenuItem12 = new MenuItem();
            _miTearColumnAway = new MenuItem();
            _miTearColumnAway.Click += miTearColumnAway_Click;
            _miMultipleColumnTearAway = new MenuItem();
            _miMultipleColumnTearAway.Click += miMultipleColumnTearAway_Click;
            _miArrangeTearAways = new MenuItem();
            _miArrangeTearAways.Click += miArrangeTearAways_Click;
            _miHideColumnTearAway = new MenuItem();
            _miHideColumnTearAway.Click += miHideColumnTearAway_Click;
            _miHideAllTearAwayColumns = new MenuItem();
            _miHideAllTearAwayColumns.Click += miHideAllTearAwayColumns_Click;
            _MenuItem11 = new MenuItem();
            _miProperties = new MenuItem();
            _miProperties.Click += miProperties_Click;
            _miStats = new MenuItem();
            _pBar = new ProgressBar();
            _gb1 = new GroupBox();
            _clrdlg = new ColorDialog();
            _pdoc = new System.Drawing.Printing.PrintDocument();
            _pdoc.PrintPage += pdoc_PrintPage;
            _txtInput = new TextBox();
            _txtInput.Leave += txtInput_Leave;
            _txtInput.KeyDown += txtInput_KeyDown;
            _cmboInput = new ComboBox();
            _cmboInput.Leave += cmboInput_Leave;
            _cmboInput.KeyDown += cmboInput_keyDown;
            _cmboInput.SelectedIndexChanged += cmboInput_SelectedIndexChanged;
            __TTip = new ToolTip(components);
            _gb1.SuspendLayout();
            SuspendLayout();
            // 
            // hs
            // 
            _hs.Dock = DockStyle.Bottom;
            _hs.Location = new Point(0, 138);
            _hs.Name = "hs";
            _hs.Size = new Size(528, 12);
            _hs.TabIndex = 1;
            _hs.Visible = false;
            // 
            // vs
            // 
            _vs.Dock = DockStyle.Right;
            _vs.Location = new Point(516, 0);
            _vs.Name = "vs";
            _vs.Size = new Size(12, 138);
            _vs.TabIndex = 0;
            _vs.Visible = false;
            // 
            // txtHandler
            // 
            _txtHandler.Location = new Point(0, 0);
            _txtHandler.Name = "txtHandler";
            _txtHandler.Size = new Size(12, 20);
            _txtHandler.TabIndex = 2;
            _txtHandler.Text = "TextBox1";
            // 
            // menu
            // 
            _menu.MenuItems.AddRange(new MenuItem[] { _miCopyCellToClipboard, _MenuItem7, _MenuItem5, _MenuItem9, _miExportToExcelMenu, _miExportToTextFile, _miExportToHTMLTable, _miExportToSQLScript, _MenuItem8, _miFormatStuff, _MenuItem2, _MenuItem3, _MenuItem4, _MenuItem10, _miSearchInColumn, _MenuItem1, _miPrintTheGrid, _miPreviewTheGrid, _miPageSetup, _MenuItem12, _miTearColumnAway, _miMultipleColumnTearAway, _miArrangeTearAways, _miHideColumnTearAway, _miHideAllTearAwayColumns, _MenuItem11, _miProperties, _miStats });
            // 
            // miCopyCellToClipboard
            // 
            _miCopyCellToClipboard.Index = 0;
            _miCopyCellToClipboard.Text = "Copy Cell To Clipboard";
            // 
            // MenuItem7
            // 
            _MenuItem7.Index = 1;
            _MenuItem7.Text = "-";
            // 
            // MenuItem5
            // 
            _MenuItem5.Index = 2;
            _MenuItem5.MenuItems.AddRange(new MenuItem[] { _miSumColumn, _miSumRow, _miMaxCol, _miMaxRow, _miMinCol, _miMinRow, _miColAverage, _miRowAverage, _miDisplayFrequencyDistribution });
            _MenuItem5.Text = "Math";
            // 
            // miSumColumn
            // 
            _miSumColumn.Index = 0;
            _miSumColumn.Text = "Sum Column";
            // 
            // miSumRow
            // 
            _miSumRow.Index = 1;
            _miSumRow.Text = "Sum Row";
            // 
            // miMaxCol
            // 
            _miMaxCol.Index = 2;
            _miMaxCol.Text = "Max In Column";
            // 
            // miMaxRow
            // 
            _miMaxRow.Index = 3;
            _miMaxRow.Text = "Max In Row";
            // 
            // miMinCol
            // 
            _miMinCol.Index = 4;
            _miMinCol.Text = "Min In Column";
            // 
            // miMinRow
            // 
            _miMinRow.Index = 5;
            _miMinRow.Text = "Min In Row";
            // 
            // miColAverage
            // 
            _miColAverage.Index = 6;
            _miColAverage.Text = "Column Average";
            // 
            // miRowAverage
            // 
            _miRowAverage.Index = 7;
            _miRowAverage.Text = "Row Average";
            // 
            // miDisplayFrequencyDistribution
            // 
            _miDisplayFrequencyDistribution.Index = 8;
            _miDisplayFrequencyDistribution.Text = "Frequency Distribution";
            // 
            // MenuItem9
            // 
            _MenuItem9.Index = 3;
            _MenuItem9.Text = "-";
            // 
            // miExportToExcelMenu
            // 
            _miExportToExcelMenu.Index = 4;
            _miExportToExcelMenu.MenuItems.AddRange(new MenuItem[] { _miExportToExcel, _miAutoFitCols, _miAutoFitRows, _miALternateRowColors, _miMatchGridColors, _miOutlineExportedCells });
            _miExportToExcelMenu.Text = "Export To Excel";
            // 
            // miExportToExcel
            // 
            _miExportToExcel.Index = 0;
            _miExportToExcel.Text = "Export To Excel";
            // 
            // miAutoFitCols
            // 
            _miAutoFitCols.Index = 1;
            _miAutoFitCols.Text = "AutoFit Excel Columns";
            // 
            // miAutoFitRows
            // 
            _miAutoFitRows.Index = 2;
            _miAutoFitRows.Text = "AutoFit Excel Rows";
            // 
            // miALternateRowColors
            // 
            _miALternateRowColors.Index = 3;
            _miALternateRowColors.Text = "Alternate Row Colors";
            // 
            // miMatchGridColors
            // 
            _miMatchGridColors.Index = 4;
            _miMatchGridColors.Text = "Match Grid Colors";
            // 
            // miOutlineExportedCells
            // 
            _miOutlineExportedCells.Index = 5;
            _miOutlineExportedCells.Text = "Outline Exported Grid Cells";
            // 
            // miExportToTextFile
            // 
            _miExportToTextFile.Index = 5;
            _miExportToTextFile.Text = "Export To a Text File";
            // 
            // miExportToHTMLTable
            // 
            _miExportToHTMLTable.Index = 6;
            _miExportToHTMLTable.Text = "Export To an HTML Table";
            // 
            // miExportToSQLScript
            // 
            _miExportToSQLScript.Index = 7;
            _miExportToSQLScript.Text = "Export To an SQL Script";
            // 
            // MenuItem8
            // 
            _MenuItem8.Index = 8;
            _MenuItem8.Text = "-";
            // 
            // miFormatStuff
            // 
            _miFormatStuff.Index = 9;
            _miFormatStuff.MenuItems.AddRange(new MenuItem[] { _miFormatAsMoney, _miFormatAsDecimal, _miFormatAsText, _miCenter, _miLeft, _miRight });
            _miFormatStuff.Text = "Format Functions";
            // 
            // miFormatAsMoney
            // 
            _miFormatAsMoney.Index = 0;
            _miFormatAsMoney.Text = "Format As Money";
            // 
            // miFormatAsDecimal
            // 
            _miFormatAsDecimal.Index = 1;
            _miFormatAsDecimal.Text = "Format as Decimal";
            // 
            // miFormatAsText
            // 
            _miFormatAsText.Index = 2;
            _miFormatAsText.Text = "Format as Text";
            // 
            // miCenter
            // 
            _miCenter.Index = 3;
            _miCenter.Text = "Center";
            // 
            // miLeft
            // 
            _miLeft.Index = 4;
            _miLeft.Text = "Left";
            // 
            // miRight
            // 
            _miRight.Index = 5;
            _miRight.Text = "Right";
            // 
            // MenuItem2
            // 
            _MenuItem2.Index = 10;
            _MenuItem2.MenuItems.AddRange(new MenuItem[] { _miFontsSmaller, _miFontsLarger, _miHeaderFontSmaller, _miHeaderFontLarger, _miTitleFontSmaller, _miTitleFontLarger, _miSmoothing, _miAutoSizeToContents, _miAllowUserColumnResizing });
            _MenuItem2.Text = "Settings Functions";
            // 
            // miFontsSmaller
            // 
            _miFontsSmaller.Index = 0;
            _miFontsSmaller.Text = "Grid Fonts Smaller";
            // 
            // miFontsLarger
            // 
            _miFontsLarger.Index = 1;
            _miFontsLarger.Text = "Grid Fonts Larger";
            // 
            // miHeaderFontSmaller
            // 
            _miHeaderFontSmaller.Index = 2;
            _miHeaderFontSmaller.Text = "Header Fonts Smaller";
            // 
            // miHeaderFontLarger
            // 
            _miHeaderFontLarger.Index = 3;
            _miHeaderFontLarger.Text = "Header Fonts Larger";
            // 
            // miTitleFontSmaller
            // 
            _miTitleFontSmaller.Index = 4;
            _miTitleFontSmaller.Text = "Title Font Smaller";
            // 
            // miTitleFontLarger
            // 
            _miTitleFontLarger.Index = 5;
            _miTitleFontLarger.Text = "Title Font Larger";
            // 
            // miSmoothing
            // 
            _miSmoothing.Index = 6;
            _miSmoothing.Text = "Smoothing";
            // 
            // miAutoSizeToContents
            // 
            _miAutoSizeToContents.Index = 7;
            _miAutoSizeToContents.Text = "Auto Size To Contents";
            // 
            // miAllowUserColumnResizing
            // 
            _miAllowUserColumnResizing.Index = 8;
            _miAllowUserColumnResizing.Text = "Allow User Column Resizing";
            // 
            // MenuItem3
            // 
            _MenuItem3.Index = 11;
            _MenuItem3.MenuItems.AddRange(new MenuItem[] { _miSortAscending, _miSortDescending, _miDateAsc, _miDateDesc, _miSortNumericAsc, _miSortNumericDesc });
            _MenuItem3.Text = "Sort";
            // 
            // miSortAscending
            // 
            _miSortAscending.Index = 0;
            _miSortAscending.Text = "Ascii Ascending";
            // 
            // miSortDescending
            // 
            _miSortDescending.Index = 1;
            _miSortDescending.Text = "Ascii Descending";
            // 
            // miDateAsc
            // 
            _miDateAsc.Index = 2;
            _miDateAsc.Text = "Date Ascending";
            // 
            // miDateDesc
            // 
            _miDateDesc.Index = 3;
            _miDateDesc.Text = "Date Descending";
            // 
            // miSortNumericAsc
            // 
            _miSortNumericAsc.Index = 4;
            _miSortNumericAsc.Text = "Numeric Ascending";
            // 
            // miSortNumericDesc
            // 
            _miSortNumericDesc.Index = 5;
            _miSortNumericDesc.Text = "Numeric Descending";
            // 
            // MenuItem4
            // 
            _MenuItem4.Index = 12;
            _MenuItem4.MenuItems.AddRange(new MenuItem[] { _miHideRow, _miHideColumn, _miShowAllRowsAndColumns, _miSetRowColor, _miSetColumnColor, _miSetCellColor });
            _MenuItem4.Text = "Row and Column Options";
            // 
            // miHideRow
            // 
            _miHideRow.Index = 0;
            _miHideRow.Text = "Hide Row";
            // 
            // miHideColumn
            // 
            _miHideColumn.Index = 1;
            _miHideColumn.Text = "Hide Column";
            // 
            // miShowAllRowsAndColumns
            // 
            _miShowAllRowsAndColumns.Index = 2;
            _miShowAllRowsAndColumns.Text = "Show All Rows and Columns";
            // 
            // miSetRowColor
            // 
            _miSetRowColor.Index = 3;
            _miSetRowColor.Text = "Color Row";
            // 
            // miSetColumnColor
            // 
            _miSetColumnColor.Index = 4;
            _miSetColumnColor.Text = "Color Column";
            // 
            // miSetCellColor
            // 
            _miSetCellColor.Index = 5;
            _miSetCellColor.Text = "Color Cell";
            // 
            // MenuItem10
            // 
            _MenuItem10.Index = 13;
            _MenuItem10.Text = "-";
            // 
            // miSearchInColumn
            // 
            _miSearchInColumn.Index = 14;
            _miSearchInColumn.Text = "Find In Column";
            // 
            // MenuItem1
            // 
            _MenuItem1.Index = 15;
            _MenuItem1.Text = "-";
            // 
            // miPrintTheGrid
            // 
            _miPrintTheGrid.Index = 16;
            _miPrintTheGrid.Text = "Print The Grids Contents";
            // 
            // miPreviewTheGrid
            // 
            _miPreviewTheGrid.Index = 17;
            _miPreviewTheGrid.Text = "Preview The Grids Contents";
            // 
            // miPageSetup
            // 
            _miPageSetup.Index = 18;
            _miPageSetup.Text = "Page and Printer Setup";
            // 
            // MenuItem12
            // 
            _MenuItem12.Index = 19;
            _MenuItem12.Text = "-";
            // 
            // miTearColumnAway
            // 
            _miTearColumnAway.Index = 20;
            _miTearColumnAway.Text = "Tear Column Away";
            // 
            // miMultipleColumnTearAway
            // 
            _miMultipleColumnTearAway.Index = 21;
            _miMultipleColumnTearAway.Text = "Tear Multiple Columns Away";
            // 
            // miArrangeTearAways
            // 
            _miArrangeTearAways.Index = 22;
            _miArrangeTearAways.Text = "Arrange Open Tear Away Columns";
            // 
            // miHideColumnTearAway
            // 
            _miHideColumnTearAway.Index = 23;
            _miHideColumnTearAway.Text = "Hide Column Tear Away";
            // 
            // miHideAllTearAwayColumns
            // 
            _miHideAllTearAwayColumns.Index = 24;
            _miHideAllTearAwayColumns.Text = "Hide All Tear Away Columns";
            // 
            // MenuItem11
            // 
            _MenuItem11.Index = 25;
            _MenuItem11.Text = "-";
            // 
            // miProperties
            // 
            _miProperties.Index = 26;
            _miProperties.Text = "Properties";
            // 
            // miStats
            // 
            _miStats.Index = 27;
            _miStats.Text = "Stats";
            // 
            // pBar
            // 
            _pBar.Anchor = AnchorStyles.Top | AnchorStyles.Left
                | AnchorStyles.Right;
            _pBar.Cursor = Cursors.WaitCursor;
            _pBar.Location = new Point(4, 16);
            _pBar.Name = "pBar";
            _pBar.Size = new Size(392, 16);
            _pBar.TabIndex = 3;
            _pBar.Visible = false;
            // 
            // gb1
            // 
            _gb1.Anchor = AnchorStyles.Top | AnchorStyles.Left
                | AnchorStyles.Right;
            _gb1.BackColor = SystemColors.ScrollBar;
            _gb1.Controls.Add(_pBar);
            _gb1.Location = new Point(56, 8);
            _gb1.Name = "gb1";
            _gb1.Size = new Size(400, 48);
            _gb1.TabIndex = 4;
            _gb1.TabStop = false;
            _gb1.Text = "Progress...";
            _gb1.Visible = false;
            // 
            // pdoc
            // 
            // 
            // txtInput
            // 
            _txtInput.BorderStyle = BorderStyle.None;
            _txtInput.Location = new Point(0, 20);
            _txtInput.Name = "txtInput";
            _txtInput.Size = new Size(12, 13);
            _txtInput.TabIndex = 6;
            _txtInput.Text = "TextBox1";
            _txtInput.Visible = false;
            // 
            // cmboInput
            // 
            _cmboInput.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmboInput.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, Conversions.ToByte(0));
            _cmboInput.ItemHeight = 13;
            _cmboInput.Location = new Point(0, 36);
            _cmboInput.Name = "cmboInput";
            _cmboInput.Size = new Size(36, 21);
            _cmboInput.TabIndex = 7;
            _cmboInput.Visible = false;
            // 
            // TAIGridControl
            // 
            BackColor = SystemColors.GradientActiveCaption;
            Controls.Add(_cmboInput);
            Controls.Add(_txtInput);
            Controls.Add(_gb1);
            Controls.Add(_txtHandler);
            Controls.Add(_vs);
            Controls.Add(_hs);
            Name = "TAIGridControl";
            Size = new Size(528, 150);
            base.Paint += TAIGRIDv2_Paint;
            base.SizeChanged += TAIGRIDv2_SizeChanged;
            base.MouseEnter += MouseEnterHandler;
            base.MouseWheel += MouseWheelHandler;
            base.MouseUp += MouseUpHandler;
            base.DoubleClick += DoubleClickHandler;
            base.MouseDown += MouseDownHandler;
            base.MouseMove += MouseMoveHandler;
            base.Load += TAIGRIDControl_Load;
            base.HandleDestroyed += TAIGridControl_HandleDestroyed;
            _gb1.ResumeLayout(false);
            base.Paint += TAIGRIDv2_Paint;
            base.SizeChanged += TAIGRIDv2_SizeChanged;
            base.MouseEnter += MouseEnterHandler;
            base.MouseWheel += MouseWheelHandler;
            base.MouseUp += MouseUpHandler;
            base.DoubleClick += DoubleClickHandler;
            base.MouseDown += MouseDownHandler;
            base.MouseMove += MouseMoveHandler;
            base.Load += TAIGRIDControl_Load;
            base.HandleDestroyed += TAIGridControl_HandleDestroyed;
            ResumeLayout(false);
            PerformLayout();
        }

        #region Events

        // CellClicked
        /// <summary>
        /// Raised whenever a cell is clicked. Coordinates designated by RowClicked/ColumnClicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="RowClicked"></param>
        /// <param name="ColumnClicked"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a cell is clicked designated by RowClicked, ColumnClicked")]
        public event CellClickedEventHandler CellClicked;

        public delegate void CellClickedEventHandler(object sender, int RowClicked, int ColumnClicked);

        // CellDoubleClicked
        /// <summary>
        /// Raised whenever a cell is doubleclicked. Coordinates designated by RowClicked/ColumnClicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="RowClicked"></param>
        /// <param name="ColumnClicked"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a cell is doubleclicked designated by RowClicked, ColumnClicked")]
        public event CellDoubleClickedEventHandler CellDoubleClicked;

        public delegate void CellDoubleClickedEventHandler(object sender, int RowClicked, int ColumnClicked);

        // CellEdited
        /// <summary>
        /// Raised whenever a cell is edited by the user, if cell editing is turned on correctly.
        /// RowClicked/ColumnClicked designated which cell was edited. oldval/newval designated the previous contents and
        /// the new contents respectively
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="RowClicked"></param>
        /// <param name="ColumnClicked"></param>
        /// <param name="oldval"></param>
        /// <param name="newval"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a cell Edited by the user")]
        public event CellEditedEventHandler CellEdited;

        public delegate void CellEditedEventHandler(object sender, int RowClicked, int ColumnClicked, string oldval, string newval);

        // RowSelected
        /// <summary>
        /// Raised whenever a row is selected with the mouse or the keyboard. RowSelected designated which row.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="RowSelected"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a row is selected with mouse or keyboard. Rowselected is returned")]
        public event RowSelectedEventHandler RowSelected;

        public delegate void RowSelectedEventHandler(object sender, int RowSelected);

        // RowDeSelected
        /// <summary>
        /// Raised whenever a row is deselected with the mouse or the kayboard. RowDeselected designated which row
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="RowDeselected"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a row is deselected with mouse or keyboard. RowSelected is returned")]
        public event RowDeSelectedEventHandler RowDeSelected;

        public delegate void RowDeSelectedEventHandler(object sender, int RowDeselected);

        // PartialSelection
        /// <summary>
        /// Raised whenever a populategrid from database call exceeded the set threshold of records
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a populategrid from database exceeds a set threshold of records")]
        public event PartialSelectionEventHandler PartialSelection;

        public delegate void PartialSelectionEventHandler(object sender);

        // TooManyRecords
        /// <summary>
        /// Raised whenever a populategrid from database call gets to many records that the grid cannot handle. After the rewrite
        /// in 2005 this event is exceedintgly difficult to fire as the grid can now handle millions of records at a time.
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a populategrid from database gets so many records that the bitmap becomes to big")]
        public event TooManyRecordsEventHandler TooManyRecords;

        public delegate void TooManyRecordsEventHandler(object sender);

        // TooManyFields
        /// <summary>
        /// Raised whenever a populategrid from database call gets to many records that the grid cannot handle. After the rewrite
        /// in 2005 this event is exceedintgly difficult to fire as the grid can now handle millions of records at a time.
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised whenever a populategrid from database gets so many records that the bitmap becomes to big")]
        public event TooManyFieldsEventHandler TooManyFields;

        public delegate void TooManyFieldsEventHandler(object sender);

        // StartedDatabasePopulateOperation
        /// <summary>
        /// Raised when the grid starts a PopulateGridWData call from a supported data source (SQL,OLE,ODBC,DATATABLE etc.)
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised whenever the grid is starting to PopulateGridWData")]
        public event StartedDatabasePopulateOperationEventHandler StartedDatabasePopulateOperation;

        public delegate void StartedDatabasePopulateOperationEventHandler(object sender);

        // FinishedDatabasePopulateOperation
        /// <summary>
        /// Raised when the grid finishes a PopulateGridWData call from a supported data source (SQL,OLE,ODBC,DATATABLE etc.)
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised when the grid is finished PopulatingGridWData")]
        public event FinishedDatabasePopulateOperationEventHandler FinishedDatabasePopulateOperation;

        public delegate void FinishedDatabasePopulateOperationEventHandler(object sender);

        // Column Resized
        /// <summary>
        /// Raised when the user resizes a column using the mouse. ColumnIndex designated the column being resized
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ColumnIndex"></param>
        /// <remarks></remarks>
        [Description("Raised when the user Resizes as column using the mouse")]
        public event ColumnResizedEventHandler ColumnResized;

        public delegate void ColumnResizedEventHandler(object sender, int ColumnIndex);

        // Column Selected
        /// <summary>
        /// Raised when the user selectes a column using the mouse. ColumnIndex designated the column being resized
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ColumnIndex"></param>
        /// <remarks></remarks>
        [Description("Raised when the user selects as column using the mouse")]
        public event ColumnSelectedEventHandler ColumnSelected;

        public delegate void ColumnSelectedEventHandler(object sender, int ColumnIndex);

        // Column DeSelected
        /// <summary>
        /// Raised when the user selectes a column using the mouse. ColumnIndex designated the column being resized
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="OldColumnIndex"></param>
        /// <remarks></remarks>
        [Description("Raised when the user deselects as column using the mouse")]
        public event ColumnDeSelectedEventHandler ColumnDeSelected;

        public delegate void ColumnDeSelectedEventHandler(object sender, int OldColumnIndex);


        // GridResorted
        /// <summary>
        /// Raised when the user resorts the grids contents on a chosen column index. ColumnIndex is the chosen column.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ColumnIndex"></param>
        /// <remarks></remarks>
        [Description("Raised when the user Sorts the grid on a given column ColumnIndex is that column")]
        public event GridResortedEventHandler GridResorted;

        public delegate void GridResortedEventHandler(object sender, int ColumnIndex);

        // KeypressedInGrid
        /// <summary>
        /// Raised when the user presses a key on the keyboard while the grid has focus and a cell is not being edited.
        /// The Keycode parameter is of the type Keys
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="keyCode"></param>
        /// <remarks></remarks>
        [Description("Raised when the user Presses any keyboard key returns a type of Keys")]
        public event KeyPressedInGridEventHandler KeyPressedInGrid;

        public delegate void KeyPressedInGridEventHandler(object sender, Keys keyCode);

        // RightMouseButtonInGrid
        /// <summary>
        /// Raised when the user presses the rightmousebutton in the grid and the grid is not doing any of its Popup context
        /// menus.
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised when the user selects the rightmousebutton in a grid and the grid is NOT doing POPUP menus")]
        public event RightMouseButtonInGridEventHandler RightMouseButtonInGrid;

        public delegate void RightMouseButtonInGridEventHandler(object sender);

        // GridHover
        /// <summary>
        /// Raised as the user loiters over the rendered grid contents with the mouse. Row/Col esignated the cell being
        /// hovered over, Item indicates the contents of 6the cell being hovered over.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="Item"></param>
        /// <remarks></remarks>
        [Description("Raised when the user is hovering over the grid itself Not the grid container just the rendered grid")]
        public event GridHoverEventHandler GridHover;

        public delegate void GridHoverEventHandler(object sender, int row, int col, string Item);

        // GridHoverLeave
        /// <summary>
        /// Raised as the user moved the mouse off of the grid rendered contens after hovering over those contents previously
        /// </summary>
        /// <param name="sender"></param>
        /// <remarks></remarks>
        [Description("Raised when the user is hovering over the grid itself Not the grid container just the rendered grid")]
        public event GridHoverleaveEventHandler GridHoverleave;

        public delegate void GridHoverleaveEventHandler(object sender);

        #endregion

        /// <summary>
        /// Enmeration for selecting preset color schemes used to configures the theme of the grids display
        /// </summary>
        /// <remarks></remarks>
        public enum TaiGridColorSchemes : int
        {
            _Default = 0,
            _Business,
            _Technical,
            _Fancy,
            _Colorful1,
            _Colorful2
        }



        // AllowTearAwayFunctionality
        /// <summary>
        /// Allow or Disallow the tearaway a column functionality within the grid itself.
        /// Column tearaways allow for removing a columns contents to a seperate window that floats outside the
        /// boundarys of the grids containers. This functionality might prove useful in some circumstances but may
        /// also confuse the display for some users. This setting will turn on or off the availability of this function.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Tearaway menu items of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowTearAwayFunctionality
        {
            get
            {
                return _AllowTearAwayFuncionality;
            }
            set
            {
                _AllowTearAwayFuncionality = value;
            }
        }

        // AllowExcelFunctionality
        /// <summary>
        /// Allow or disallow the ability to export the grids contents to excel via the Context menu.
        /// Heavily used with some reporting applications where numerics are displayed in aggregate,
        /// other uses of the grid in items where personal data are displayed might
        /// necessitate turning the functionality off for privacy/Hipaa reasons.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Excel menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowExcelFunctionality
        {
            get
            {
                return _AllowExcelFunctionality;
            }
            set
            {
                _AllowExcelFunctionality = value;
            }
        }

        // AllowTextFunctionality
        /// <summary>
        /// Allows or disallows the text menu functionality of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Text menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowTextFunctionality
        {
            get
            {
                return _AllowTextFunctionality;
            }
            set
            {
                _AllowTextFunctionality = value;
            }
        }

        // AllowHTMLFunctionality
        /// <summary>
        /// Allows or disallows the HTML menu on the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the HTML menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowHTMLFunctionality
        {
            get
            {
                return _AllowHTMLFunctionality;
            }
            set
            {
                _AllowHTMLFunctionality = value;
            }
        }

        // AllowSQLFunctionality
        /// <summary>
        /// Allows or disallows the SQL menu off of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the SQL menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowSQLFunctionality
        {
            get
            {
                return _AllowSQLScriptFunctionality;
            }
            set
            {
                _AllowSQLScriptFunctionality = value;
            }
        }

        // AllowMathFunctionality
        /// <summary>
        /// Allows or disallows the Math submenu off of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Math menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowMathFunctionality
        {
            get
            {
                return _AllowMathFunctionality;
            }
            set
            {
                _AllowMathFunctionality = value;
            }
        }

        // AllowFormatFunctionality
        /// <summary>
        /// Allows or disallows the Format submenu off of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Format menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowFormatFunctionality
        {
            get
            {
                return _AllowFormatFunctionality;
            }
            set
            {
                _AllowFormatFunctionality = value;
            }
        }

        // AllowSettingsFunctionality
        /// <summary>
        /// Allows or disallows the settings submenu off of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Settings menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowSettingsFunctionality
        {
            get
            {
                return _AllowSettingsFunctionality;
            }
            set
            {
                _AllowSettingsFunctionality = value;
            }
        }

        // AllowSortFunctionality
        /// <summary>
        /// Allows or disallows the Sort submenu off of the grids context menu.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the Sort menu of the builtin context menu")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowSortFunctionality
        {
            get
            {
                return _AllowSortFunctionality;
            }
            set
            {
                _AllowSortFunctionality = value;
            }
        }

        // AllowColumnSelection
        /// <summary>
        /// Allows or disallows the selection of a column by clicking the header of a column with the mouse
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow selection of a column visually by single clicking on the header of that column")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowColumnSelection
        {
            get
            {
                return _AllowColumnSelection;
            }
            set
            {
                _AllowColumnSelection = value;

                Invalidate();
            }
        }

        // AllowControlKeyMenuPopup
        /// <summary>
        /// Allows or disallows the pulling up of the grids context menu my pressintg the ctrl key while right mousebuttoning
        /// over the grid itself. This allows programs that are hosting the grids to create their own context menus but to
        /// still have the grids context menus available via the ctrl/right mousebutton combination.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the menu from poping up on a CTRL menubutton.")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowControlKeyMenuPopup
        {
            get
            {
                return _AllowControlKeyMenuPopup;
            }
            set
            {
                _AllowControlKeyMenuPopup = value;

                Invalidate();
            }
        }

        // AllowInGridEdits
        /// <summary>
        /// Allows or disallows the editing of grids contents. This is not an all or nothing process. The developer hase to turn this
        /// on and explicitly set the columns where they want to allow editing in order for in grid edits to function.
        /// Alternately they might elect to restrict editing of a cells contents to s list of available selections
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the editing of the grid contents at the column level")]
        [DefaultValue(typeof(bool), "False")]
        public bool AllowInGridEdits
        {
            get
            {
                return _AllowInGridEdits;
            }
            set
            {
                _AllowInGridEdits = value;
            }
        }

        // AllowMultipleRowSelections
        /// <summary>
        /// Allows or disallows the ability of the user to select more than a single row in the grid at one time
        /// via the standard CTRL/SHIFT key click mechanism used in the Windows OS. The rows selected will then
        /// be exposed via the <c>SelectedRows</c> collection
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the selection of Multiple rows in the grid with the CTRL key ")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowMultipleRowSelections
        {
            get
            {
                return _AllowMultipleRowSelections;
            }
            set
            {
                _AllowMultipleRowSelections = value;
            }
        }

        // AllowPopupMenu
        /// <summary>
        /// Allow or disallow the grids own context menu to appear via the right mousebutton
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the builtin popup menu for font selection and sizing to occur")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowPopupMenu
        {
            get
            {
                return _AllowPopupMenu;
            }
            set
            {
                _AllowPopupMenu = value;
            }
        }

        // AllowRowSelection
        /// <summary>
        /// Allow or disallow the ability to select a single or multiple rows in the grid with the mouse,
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow selection of a Row or Multiple Rows visually by single clicking in the Row itself")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowRowSelection
        {
            get
            {
                return _AllowRowSelection;
            }
            set
            {
                _AllowRowSelection = value;

                Invalidate();
            }
        }

        // AllowWhiteSpaceInCells
        /// <summary>
        /// Will allow/disallow Whitespace in cells (newlines and what not)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow Whitespace in cells (newlines and what not)")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowWhiteSpaceInCells
        {
            get
            {
                return _AllowWhiteSpaceInCells;
            }
            set
            {
                _AllowWhiteSpaceInCells = value;

                Invalidate();
            }
        }


        // AllowUserColumnResizing
        /// <summary>
        /// Allow or disallow user column resizing with the mouse
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the user to resize a column")]
        [DefaultValue(typeof(bool), "True")]
        public bool AllowUserColumnResizing
        {
            get
            {
                return _AllowUserColumnResizing;
            }
            set
            {
                _AllowUserColumnResizing = value;
            }
        }

        // AlternateColoration
        /// <summary>
        /// Turns on or off the Alternate coloration mode of the grids display where it will alternate the background
        /// color of the rows inserted between <c>AlternateColorationAltColor</c> and <c>AlternateColorationBaseColor</c>
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will trun on/off the alternate coloration mode of the grid. Rows will alternate between defined backcolor and the defined alternatecolor")]
        public bool AlternateColoration
        {
            get
            {
                return _alternateColorationMode;
            }
            set
            {
                _alternateColorationMode = value;
            }
        }

        // AlternateColorationAltColor
        /// <summary>
        /// One of the colors used when the grid is rendering in AlternateColoration mode
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the alternate color for the alternate Coloration mode of operation")]
        public Color AlternateColorationAltColor
        {
            get
            {
                return _alternateColorationALTColor;
            }
            set
            {
                _alternateColorationALTColor = value;
            }
        }

        // AlternateColorationBaseColor
        /// <summary>
        /// One of the colors used when the grid is rendering in AlternateColoration mode
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the base color for the alternate Coloration mode of operation")]
        public Color AlternateColorationBaseColor
        {
            get
            {
                return _alternateColorationBaseColor;
            }
            set
            {
                _alternateColorationBaseColor = value;
            }
        }

        // Antialias
        /// <summary>
        /// Turns on or off the smoothing mode of the grids textual rendering engine
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Turns on/off the antialias mode of the grids rendering engine")]
        [DefaultValue(typeof(bool), "False")]
        public bool Antialias
        {
            get
            {
                return _antialias;
            }
            set
            {
                _antialias = value;

                Invalidate();
            }
        }

        // AutoSizeCellsToContents
        /// <summary>
        /// Allows or disallows the grids rendering engine to automagically resize the grids row and column metrics
        /// to accomodate the contents being inserted into the grid manually or via one of the PopulateFromDatabase calls.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the grid to automatically adjust the grid cells heigh and width to match the textual contents of the cells")]
        [DefaultValue(typeof(bool), "False")]
        public bool AutoSizeCellsToContents
        {
            get
            {
                return _AutosizeCellsToContents;
            }
            set
            {
                _AutosizeCellsToContents = value;
                if (value)
                {
                    _AutoSizeAlreadyCalculated = false;
                    _AutoSizeSemaphore = true;
                    DoAutoSizeCheck(CreateGraphics());
                }
                Invalidate();
            }
        }

        // AutoFocus
        /// <summary>
        /// Allows or disallows the grids ability to automagically gain focus as the user mouses over the grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow/disallow the grid to automatically gain focus on mouseover")]
        [DefaultValue(typeof(bool), "False")]
        public bool AutoFocus
        {
            get
            {
                return _AutoFocus;
            }
            set
            {
                _AutoFocus = value;
            }
        }

        /// <summary>
        /// Gets or Sets the background color for the control
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the background color for the control")]
        [DefaultValue(typeof(Color), "GradientActiveCaption")]
        public override Color BackColor
        {
            get
            {
                return base.BackColor;
            }
            set
            {
                base.BackColor = value;
            }
        }

        // BorderColor
        /// <summary>
        /// The color used to render the border of the grid itself when <c>BorderStyle</c> = something
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the border color for the drawn border on Borderstyle = something")]
        public Color BorderColor
        {
            get
            {
                return _BorderColor;
            }
            set
            {
                _BorderColor = value;
            }
        }

        // BorderStyle
        /// <summary>
        /// The border style use to draw the grid border itself
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the border style for the grids container object or frame")]
        public new BorderStyle BorderStyle
        {
            get
            {
                return _BorderStyle;
            }
            set
            {
                _BorderStyle = value;
            }
        }

        // ShowProgressBar
        /// <summary>
        /// Allows or disallows the display f the progress bar across the top of the grid itself when long database population
        /// processes are underway.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow the grid to show a small progress bar along its top edges on long database populate methods")]
        [DefaultValue(typeof(bool), "True")]
        public bool ShowProgressBar
        {
            get
            {
                return _ShowProgressBar;
            }
            set
            {
                _ShowProgressBar = value;
            }
        }

        // ShowExcelExportMessage
        /// <summary>
        /// Allows or disallows the display of a topmost windows signaling to the end users that the grid
        /// is sending it's content to excel. Because messaging excel is sometime a lengthy process the display of
        /// the dialog might prove useful in those situations. Messaging excel though is an inherantly messy process
        /// where a user interacting with a different instance of excel might confuse the system and make the
        /// export process fail. In these cases the dialog might also prove useful in that the user can be instructed
        /// 'Hands off' with the dialog is visable. It can however be confusing when this dialog is on top of everything.
        /// As always its use might prove useful or it might not depending on environmental factors.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will allow the grid to show a small window topmost on exporting to excel with status information")]
        [DefaultValue(typeof(bool), "True")]
        public bool ShowExcelExportMessage
        {
            get
            {
                return _ShowExcelExportMessage;
            }
            set
            {
                _ShowExcelExportMessage = value;
            }
        }

        // Cols
        /// <summary>
        /// Get or Sets the number of columns in the grid itself
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("How many columns are in the current grid")]
        public int Cols
        {
            get
            {
                return _cols;
            }
            set
            {
                SetCols(value);
                Invalidate();
            }
        }

        // CellAlignment
        /// <summary>
        /// Get or sets the alignment of the textual element contained at Grid coordinates R (row) and C (col)
        /// </summary>
        /// <param name="r"></param>
        /// <param name="c"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public StringFormat get_CellAlignment(int r, int c)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
                return _DefaultStringFormat;
            else
                return _gridCellAlignmentList[_gridCellAlignment[r, c]];
        }

        public void set_CellAlignment(int r, int c, StringFormat value)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
            {
            }
            else
            {
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(value);
                Invalidate();
            }
        }

        // CellBackColor
        /// <summary>
        /// Gets or sets the background color of the specificied cell at R (row) and C (col)
        /// </summary>
        /// <param name="r"></param>
        /// <param name="c"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Brush get_CellBackColor(int r, int c)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
                return new SolidBrush(Color.AntiqueWhite);
            else
                return _gridBackColorList[_gridBackColor[r, c]];
        }

        public void set_CellBackColor(int r, int c, Brush value)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
            {
            }
            else
            {
                _gridBackColor[r, c] = GetGridBackColorListEntry(value);
                Invalidate();
            }
        }

        // ColBackColorEdit
        /// <summary>
        /// Gets or sets the background color used to render a cell when that cell is in edit mode
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the background color for the cell when its being edited")]
        public Color ColBackColorEdit
        {
            get
            {
                return _colEditableTextBackColor;
            }
            set
            {
                _colEditableTextBackColor = value;
                txtInput.BackColor = value;
            }
        }

        // CellFont
        /// <summary>
        /// Gets or sets the font used to render a the cells contents. The cell is designated by
        /// its cordinates R (row) and C (col)
        /// </summary>
        /// <param name="r"></param>
        /// <param name="c"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Font get_CellFont(int r, int c)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
                return _gridCellFontsList[0];
            else
                return _gridCellFontsList[_gridCellFonts[r, c]];
        }

        public void set_CellFont(int r, int c, Font value)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
            {
            }
            else
            {
                _gridCellFonts[r, c] = GetGridCellFontListEntry(value);
                Invalidate();
            }
        }

        // CellForeColor
        /// <summary>
        /// Gets or sets the foreground color used to render a cell at coordinated R (row) and C (col)
        /// </summary>
        /// <param name="r"></param>
        /// <param name="c"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Pen get_CellForeColor(int r, int c)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
                return _gridForeColorList[0]; // New Pen(_DefaultForeColor.Blue)
            else
                return _gridForeColorList[_gridForeColor[r, c]];
        }

        public void set_CellForeColor(int r, int c, Pen value)
        {
            if (r > _rows - 1 | c > _cols - 1 | r < 0 | c < 0)
            {
            }
            else
            {
                _gridForeColor[r, c] = GetGridForeColorListEntry(value);
                Invalidate();
            }
        }

        // CellOutlines
        /// <summary>
        /// Allows or disallows the rendering engins outlining of cells as it draws their contents
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Turns Cell outlining on or off")]
        public bool CellOutlines
        {
            get
            {
                return _CellOutlines;
            }
            set
            {
                _CellOutlines = value;
                Invalidate();
            }
        }

        // ColCheckBox
        /// <summary>
        /// Gets or sets the columns status as a boolean value where it will interpret the contents of a column
        /// as boolean values. 1,True,Y,y,Yes,yes and other variations will be rendered as a check checkbok
        /// 0,False,n,N,No,no and other variatons will be rendered as unchecked checkboxes. ALl other values will
        /// be rendered as disabled checkboxes that are unchecked. If the column is editable then the grid will
        /// manage checkbox state for you toggling the contents as the user interacts with the cells contents via
        /// the mouse.
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool get_ColCheckBox(int idx)
        {
            if (idx < 0 | idx > _cols)
                return false;
            else
                return _colboolean[idx];
        }

        public void set_ColCheckBox(int idx, bool value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
                _colboolean[idx] = value;
        }

        // ColEditable
        /// <summary>
        /// Gets or sets the editable status of a column at index idx.
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool get_ColEditable(int idx)
        {
            if (idx < 0 | idx > _cols)
                return false;
            else
                return _colEditable[idx];
        }

        public void set_ColEditable(int idx, bool value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
                _colEditable[idx] = value;
        }

        // RowEditable
        /// <summary>
        /// Gets or sets the editable status of a Row at index idx.
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool get_RowEditable(int idx)
        {
            if (idx < 0 | idx > _cols)
                return false;
            else
                return _rowEditable[idx];
        }

        public void set_RowEditable(int idx, bool value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
                _rowEditable[idx] = value;
        }

        // ColWidth
        /// <summary>
        /// Gets or sets the width of a column at index idx in pixels
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int get_ColWidth(int idx)
        {
            if (idx < 0 | idx > _cols)
                return 0;
            else
                return _colwidths[idx];
        }

        public void set_ColWidth(int idx, int value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
            {
                _AutosizeCellsToContents = false;
                _colwidths[idx] = value;
                Invalidate();
            }
        }

        // ColPassword
        /// <summary>
        /// Gets or sets the rendering text to be used for a column at index idx to be set as a password column.
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string get_ColPassword(int idx)
        {
            if (idx < 0 | idx > _cols)
                return "";
            else
                return _colPasswords[idx];
        }

        public void set_ColPassword(int idx, string value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
            {
                _colPasswords[idx] = value;
                Refresh();
            }
        }

        // ColMaxCharacters
        /// <summary>
        /// Gets or sets the number of characters that a clumn at index idx will display before the rendering engine will
        /// display the elipsis ... characters at the end.
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int get_ColMaxCharacters(int idx)
        {
            if (idx < 0 | idx > _cols)
                return 0;
            else
                return _colMaxCharacters[idx];
        }

        public void set_ColMaxCharacters(int idx, int value)
        {
            if (idx < 0 | idx > _cols)
            {
            }
            else
            {
                _colMaxCharacters[idx] = value;

                if (_AutosizeCellsToContents)
                {
                    _AutoSizeAlreadyCalculated = false;
                    _AutoSizeSemaphore = true;
                    DoAutoSizeCheck(CreateGraphics());
                }

                Refresh();
            }
        }

        // DataBaseTimeOut
        /// <summary>
        /// Gets or sets the time in seconds for a global database timeout value for all the Populate with database calls
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the database timeout value associated with the various populate from database methods of the grid")]
        [DefaultValue(typeof(int), "500")]
        public int DataBaseTimeOut
        {
            get
            {
                return _dataBaseTimeOut;
            }
            set
            {
                _dataBaseTimeOut = value;
            }
        }

        // DefaultCellFont
        /// <summary>
        /// Gets or sets the default font used to render cells.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the font used for cell additions by default")]
        public Font DefaultCellFont
        {
            get
            {
                return _DefaultCellFont;
            }
            set
            {
                _DefaultCellFont = value;
            }
        }

        // DefaultBackColor
        /// <summary>
        /// Gets or sets the default background color used to render cells
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the Background color use by default Cells in the grid")]
        public Color DefaultBackgroundColor
        {
            get
            {
                return _DefaultBackColor;
            }
            set
            {
                _DefaultBackColor = value;
                _gridBackColorList[0] = new SolidBrush(value);
            }
        }

        // DefaultForeColor
        /// <summary>
        /// Gets or sets the default foreground color user to render cells
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the color use by default for text added to grid")]
        public Color DefaultForegroundColor
        {
            get
            {
                return _DefaultForeColor;
            }
            set
            {
                _DefaultForeColor = value;
                _gridForeColorList[0] = new Pen(value);
            }
        }

        // Delimiter
        /// <summary>
        /// Gets or sets the default field delimiter used for the export to text methods
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field delimiter for export to text methods")]
        public string Delimiter
        {
            get
            {
                return _delimiter;
            }
            set
            {
                _delimiter = value;
            }
        }

        // GridEditMode
        /// <summary>
        /// Gets or sets the field forcing a return key on an edited cell to edit its contents or just losing focus will fire
        /// a cell edited event.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field forcing a return key to edit or kist losing focus to edit")]
        public GridEditModes GridEditMode
        {
            get
            {
                return _GridEditMode;
            }
            set
            {
                _GridEditMode = value;
            }
        }

        /// <summary>
        /// When the grid is in editmode this is the column that is currently being edited
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int EditModeCol
        {
            get
            {
                if (_EditMode)
                    return _EditModeCol;
                else
                    return -1;
            }
        }

        /// <summary>
        /// When the grid is in editmode this is the row currently being edited
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int EditModeRow
        {
            get
            {
                if (_EditMode)
                    return _EditModeRow;
                else
                    return -1;
            }
        }

        /// <summary>
        /// When the grid is maintaining i set of tearaway columns this will return True false otherwise
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool GridDoingTearAwayWork
        {
            get
            {
                return _TearAwayWork;
            }
        }

        /// <summary>
        /// When the user brings up the context menu this will retiurn the column they were over when the menu was called up
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int ColOverOnMenuButton
        {
            get
            {
                return _ColOverOnMenuButton;
            }
        }

        /// <summary>
        /// When the user brings up the context menu this will return the row they were over when the menu was called up
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int RowOverOnMenuButton
        {
            get
            {
                return _RowOverOnMenuButton;
            }
        }

        // GridHeaderFont
        /// <summary>
        /// Gets or sets the font used to render the column header
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the font used for the Grid Header by default")]
        public Font GridHeaderFont
        {
            get
            {
                return _GridHeaderFont;
            }
            set
            {
                _GridHeaderFont = value;
            }
        }

        // GridHeaderStringFormat
        /// <summary>
        /// Gets or sets the formatting characteristics of the grid header line. (left,right,centered etc.)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the stringformat object used for the Grid Header by default")]
        public StringFormat GridHeaderStringFormat
        {
            get
            {
                return _GridHeaderStringFormat;
            }
            set
            {
                _GridHeaderStringFormat = value;
                Invalidate();
            }
        }

        // GridHeaderVisible
        /// <summary>
        /// Allows or disallows the display of the grid header line
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("is the Gridheader visible or not")]
        public bool GridheaderVisible
        {
            get
            {
                return _GridHeaderVisible;
            }
            set
            {
                _GridHeaderVisible = value;
                Invalidate();
            }
        }

        // GridHeaderHeight
        /// <summary>
        /// Gets or sets the height of the grids header in pixels
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets how hight the grid header is drawn in pixels")]
        public int GridHeaderHeight
        {
            get
            {
                return _GridHeaderHeight;
            }
            set
            {
                _GridHeaderHeight = value;
                Invalidate();
            }
        }

        // GridHeaderBackColor
        /// <summary>
        /// Gets or sets the background color used to render the grids header
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the color used for the Grid Header background")]
        public Color GridHeaderBackColor
        {
            get
            {
                return _GridHeaderBackcolor;
            }
            set
            {
                _GridHeaderBackcolor = value;
                Invalidate();
            }
        }

        // GridHeaderForeColor
        /// <summary>
        /// Gets or sets the foreground color useed to render the grids header
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the foreground color used to draw the grid header")]
        public Color GridHeaderForeColor
        {
            get
            {
                return _GridHeaderForecolor;
            }
            set
            {
                _GridHeaderForecolor = value;
                Invalidate();
            }
        }

        // GridReportOrientLandscape
        /// <summary>
        /// Gets or sets the grids output for reporting to be landscape mode or not.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report output to landscape mode")]
        [DefaultValue(typeof(bool), "False")]
        public bool GridReportOrientLandscape
        {
            get
            {
                return _gridReportOrientLandscape;
            }
            set
            {
                _gridReportOrientLandscape = value;
            }
        }

        // GridReportOutlineCells
        /// <summary>
        /// Gets or sets the grids reporting engine to outline cells or not
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report to outline printed cells")]
        [DefaultValue(typeof(bool), "True")]
        public bool GridReportOutlineCells
        {
            get
            {
                return _gridReportOutlineCells;
            }
            set
            {
                _gridReportOutlineCells = value;
            }
        }

        // GridReportNumberPages
        /// <summary>
        /// Gets or sets the grids reporting engine to number the output pages or not
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report generation to number pages")]
        [DefaultValue(typeof(bool), "True")]
        public bool GridReportNumberPages
        {
            get
            {
                return _gridReportNumberPages;
            }
            set
            {
                _gridReportNumberPages = value;
            }
        }

        // GridReportMatchColors
        /// <summary>
        /// Gets or sets the grids reporting engine to attempt to match reported output coloration with the onscreen
        /// display engines coloration scheme. Some on-screen colors dont look all that well when printed to paper,
        /// this is especially true wjen the printer is a black and white printer and the screen representation is
        /// full of various colors.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report generation to match grid colors")]
        [DefaultValue(typeof(bool), "True")]
        public bool GridReportMatchColors
        {
            get
            {
                return _gridReportMatchColors;
            }
            set
            {
                _gridReportMatchColors = value;
            }
        }

        // GridReportTitle
        /// <summary>
        /// Gets or sets the textual tile to apply to reported output
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report generation Title")]
        [DefaultValue(typeof(string), "")]
        public string GridReportTitle
        {
            get
            {
                return _gridReportTitle;
            }
            set
            {
                _gridReportTitle = value;
            }
        }

        // GridReportPreviewFirst
        /// <summary>
        /// Allows or disallows the preview display when printing output from the grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will set grid auto report generation to preview output first")]
        [DefaultValue(typeof(bool), "True")]
        public bool GridReportPreviewFirst
        {
            get
            {
                return _gridReportPreviewFirst;
            }
            set
            {
                _gridReportPreviewFirst = value;
            }
        }


        // AutoFitColumn
        /// <summary>
        /// Allows or disallows the export to excel engines resizing excel columns to fit the contents
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the columns are autosized")]
        public bool ExcelAutoFitColumn
        {
            get
            {
                return _excelAutoFitColumn;
            }
            set
            {
                _excelAutoFitColumn = value;
            }
        }

        // AutoFitRow
        /// <summary>
        /// Allows or disallows the export to excel engins resize rows to fit the contents
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the rows are autosized")]
        public bool ExcelAutoFitRow
        {
            get
            {
                return _excelAutoFitRow;
            }
            set
            {
                _excelAutoFitRow = value;
            }
        }

        // ExcelAlternateColoration
        /// <summary>
        /// Gets or sets the color used to decorate alternate rows on the excel output when matching grid color
        /// scheme is turned off.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the Color used by the to populoate each alternate row on exports to excel ")]
        public Color ExcelAlternateColoration
        {
            get
            {
                return _excelAlternateRowColor;
            }
            set
            {
                _excelAlternateRowColor = value;
            }
        }

        // ExcelMatchGridColorScheme
        /// <summary>
        /// Allows or disallows the export to excel engine ability to attemt to color the excel output to match the
        /// colors used on the onscreen grid. Not all screen colors convert to excel cleanly, and different versions
        /// of excel interpret colors differently. The export engine attempts to match but those matches are not always
        /// perfect.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Forces the Export to Excel function to attempt to match the grids color scheme on export ")]
        public bool ExcelMatchGridColorScheme
        {
            get
            {
                return _excelMatchGridColorScheme;
            }
            set
            {
                _excelMatchGridColorScheme = value;
            }
        }

        // Filename
        /// <summary>
        /// Gets or sets the name of the excel spreadsheet that is generated when the export operation is complete
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the name of the excel spreadsheet")]
        public string ExcelFilename
        {
            get
            {
                return _excelFilename;
            }
            set
            {
                _excelFilename = value;
            }
        }

        // IncludeColumnHeaders
        /// <summary>
        /// Allows or disallows the insertion of the header column of the grid into the excel output.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the grid column header row is included in the spreadsheet")]
        public bool ExcelIncludeColumnHeaders
        {
            get
            {
                return _excelIncludeColumnHeaders;
            }
            set
            {
                _excelIncludeColumnHeaders = value;
            }
        }

        // Keep Alive
        /// <summary>
        /// Gets or sets the setting to keep excell alive and kicking after the grid has been sent to excel.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the spreadsheet should remain open after it is filled")]
        public bool ExcelKeepAlive
        {
            get
            {
                return _excelKeepAlive;
            }
            set
            {
                _excelKeepAlive = value;
            }
        }

        // Maximized
        /// <summary>
        /// Gets or sets the making excel opening up maximized or not when its instances and
        /// messaged during the export process
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the spreadsheet should be maximized")]
        public bool ExcelMaximized
        {
            get
            {
                return _excelMaximized;
            }
            set
            {
                _excelMaximized = value;
            }
        }

        // Page Orientation
        /// <summary>
        /// Gets or sets the orientation of the excel output for printing
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls the orientation of the spreadsheet output")]
        public int ExcelPageOrientation
        {
            get
            {
                return _excelPageOrientation;
            }
            set
            {
                if (value == xlPortrait | value == xlLandscape)
                    _excelPageOrientation = value;
            }
        }

        // OutlineCells
        /// <summary>
        /// Turns on or off the outlining of cells that are populated during the export process
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Turns on or off the outlineing of all exported grid cells in a solid line")]
        public bool ExcelOutlineCells
        {
            get
            {
                return _excelOutlineCells;
            }
            set
            {
                _excelOutlineCells = value;
            }
        }

        // Show Borders
        /// <summary>
        /// gets or sets the showborders cetting of cells that are populated during the export to excel process
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls whether the cells of the spreadsheet should be outlined")]
        public bool ExcelShowBorders
        {
            get
            {
                return _excelShowBorders;
            }
            set
            {
                _excelShowBorders = value;
            }
        }

        // UseAlternateRowColor
        /// <summary>
        /// Gets or sets the using of the alternat coloring scheme for excel output
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the property which determines if the spreadsheet uses alternating color scheme")]
        public bool ExcelUseAlternateRowColor
        {
            get
            {
                return _excelUseAlternateRowColor;
            }
            set
            {
                _excelUseAlternateRowColor = value;
            }
        }

        // Workbook Name
        /// <summary>
        /// Gets or sets the name of the worksheet name to be used when the export to excel process is underway
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the field that controls the name of the sheet")]
        public string ExcelWorksheetName
        {
            get
            {
                return _excelWorkSheetName;
            }
            set
            {
                if (value.Length > 31)
                    _excelWorkSheetName = value.Substring(0, 31);
                else
                    _excelWorkSheetName = value;
            }
        }

        // MaxrowsperSheet 
        /// <summary>
        /// Gets or sets the maximum rows to be sent to excel before anoher worksheet will be created to carry the remainder
        /// of data during the export to excel process. This Add a new worksheet and continue will take place until all the
        /// data in the grid is sent to excel. Excel has a limit of 65535 rows per worksheet.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the maximum number of rows that will be put on a single worksheet during excel export")]
        public int ExcelMaxRowsPerSheet
        {
            get
            {
                return _excelMaxRowsPerSheet;
            }
            set
            {
                if (value > 65535)
                    _excelMaxRowsPerSheet = 65535;
                else if (value < 100)
                    _excelMaxRowsPerSheet = 100;
                else
                    _excelMaxRowsPerSheet = value;
            }
        }


        // HealerLabel
        /// <summary>
        /// Gets or sets the column header label use for the column at index idx
        /// </summary>
        /// <param name="columnID"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string get_HeaderLabel(int columnID)
        {
            if (columnID < 0 | columnID > _cols - 1)
                return "";
            else if (_GridHeader[columnID] == null)
                return "";
            else
                return _GridHeader[columnID];
        }

        public void set_HeaderLabel(int columnID, string value)
        {
            if (columnID < 0 | columnID > _cols - 1)
            {
            }
            else
            {
                _GridHeader[columnID] = value;
                Invalidate();
            }
        }

        // IncludeFieldNames
        /// <summary>
        /// Allows or disallows the inclusion of the header lable on grids outoput when exporting to text
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets whether the first row of the grid with the column names should be " + "included in the export to text file")]
        [DefaultValue(typeof(bool), "True")]
        public bool IncludeFieldNames
        {
            get
            {
                return _includeFieldNames;
            }
            set
            {
                _includeFieldNames = value;
            }
        }

        // IncludeLineTerminator
        /// <summary>
        /// Allows or disallows the inclusion of line termination characters when exporting the grids contents to text
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets whether the export should add a line terminator at the end of each row")]
        [DefaultValue(typeof(bool), "True")]
        public bool IncludeLineTerminator
        {
            get
            {
                return _includeLineTerminator;
            }
            set
            {
                _includeLineTerminator = value;
            }
        }

        /// <summary>
        /// Gets contents of the grid cell at row R and col C
        /// </summary>
        /// <param name="r">The Row number to fetch contents from</param>
        /// <param name="c">The Col number to fetch contents from</param>
        /// <value></value>
        /// <returns>The String representation of the cells contents located at <c>r</c> Row and <c>c</c> Col. Returns an empty string if out of bounds or the contents of the cell are NULL</returns>
        /// <remarks></remarks>
        public string get_item(int r, int c)
        {
            if (r < 0 | c < 0 | r > _rows - 1 | c > _cols - 1)
                return "";
            else if (_grid[r, c] == null)
                return "";
            else
                return _grid[r, c];
        }
        /// <summary>
        /// will set the value of a cell located as <c>r</c> and <c>c</c> to the value of <c>value</c>
        /// </summary>
        /// <param name="r">the row number to set</param>
        /// <param name="c">the col number to set</param>
        /// <param name="value">the value to set for the cell as row, col</param>
        public void set_item(int r, int c, string value)
        {
            if (r < 0 | c < 0 | r > _rows - 1 | c > _cols - 1)
            {
            }
            else
            {
                _grid[r, c] = value;
                Invalidate();
            }
        }

        /// <summary>
        /// Gets contents of the grid cell at row R and col colname
        /// </summary>
        /// <param name="r">The Row number to fetch contents from</param>
        /// <param name="colname">The string name of the column number to fetch contents from (Taken from the Header of that column)</param>
        /// <value></value>
        /// <returns>The String representation of the cells contents located at <c>r</c> Row and <c>colname</c> Col. Returns an empty string if out of bounds or the contents of the cell are NULL</returns>
        /// <remarks></remarks>
        public string get_item(int r, string colname)
        {
            int c = -1;

            c = GetColumnIDByName(colname);

            if (r < 0 | c < 0 | r > _rows - 1 | c > _cols - 1)
                return "";
            else if (_grid[r, c] == null)
                return "";
            else
                return _grid[r, c];
        }
        /// <summary>
        /// will set the value of a cell located as <c>r</c> and <c>colname</c> the columns name based on its header to the value of <c>value</c>
        /// </summary>
        /// <param name="r">the row number to set</param>
        /// <param name="colname">the column name to set (based on the columns header)</param>
        /// <param name="value">the value to set at the cell indicated by r,colname</param>
        public void set_item(int r, string colname, string value)
        {
            int c = -1;

            c = GetColumnIDByName(colname);

            if (r < 0 | c < 0 | r > _rows - 1 | c > _cols - 1)
            {
            }
            else
            {
                _grid[r, c] = value;
                Invalidate();
            }
        }

        // MaxRowsSelected
        /// <summary>
        /// Gets or sets the maximum number of rows to populate the grid with when using the various database populate
        /// calls. Set to 0 to have the parameter unbounded. With th rewrite in 2005 the grid can accomodate millions of
        /// rows of data so this setting is largely unnecessary now.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the limit of the number or rows filled in a grid by the various populate from database methods. Will raise an event if this is set and the database fill operation exceeds this threshold. If set to 0 the ALL rows will be selected and no events will fire")]
        [DefaultValue(typeof(int), "0")]
        public int MaxRowsSelected
        {
            get
            {
                return _MaxRowsSelected;
            }
            set
            {
                _MaxRowsSelected = value;
            }
        }

        // OmitNulls
        /// <summary>
        /// Allows or disallows the rendering of the work (NULL) on reading nulls from the varous database population methods
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will turn on or off the rendering of the work (NULL) on reading nulls from the database with the various populate from database methods of the grid")]
        [DefaultValue(typeof(bool), "False")]
        public bool OmitNulls
        {
            get
            {
                return _omitNulls;
            }
            set
            {
                _omitNulls = value;
            }
        }

        // PaginationSize
        /// <summary>
        /// Gets or sets the number of rows to scroll up or down on the pageup and pagedown keys
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("How many rows to scroll on a Page up or Down")]
        public int PaginationSize
        {
            get
            {
                return _PaginationSize;
            }
            set
            {
                _PaginationSize = value;
            }
        }

        // PageSettings
        /// <summary>
        /// Gets or sets the PageSettings object used print the grids contents to windows printer devices
        /// allows for the developer to hand into the grid preconfigured print environments to support their
        /// special needs.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("The System.Drawing.Printing.PageSettings object used to print the grid to windows printers")]
        public System.Drawing.Printing.PageSettings PageSettings
        {
            get
            {
                return _psets;
            }
            set
            {
                _psets = value;
            }
        }

        // Rows
        /// <summary>
        /// Get to sets the number or rows of data contained in the current grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("How many rows in the current grid control")]
        public int Rows
        {
            get
            {
                return _rows;
            }
            set
            {
                SetRows(value);
                Invalidate();
            }
        }

        // RowHeight
        /// <summary>
        /// Gets or sets the height of the row at index idx in pixels
        /// </summary>
        /// <param name="idx"></param>
        /// <value></value>
        /// <returns> Integer in pixels of the <c>idx</c> row height currently</returns>
        /// <remarks></remarks>
        public int get_RowHeight(int idx)
        {
            if (idx < 0 | idx > _rows)
                return 0;
            else
                return _rowheights[idx];
        }

        /// <summary>
        /// Will explicitly set the hight of <c>idx</c> row to be <c>value</c> pixels in height.
        /// </summary>
        /// <param name="idx"></param>
        /// <param name="value"></param>
        public void set_RowHeight(int idx, int value)
        {
            if (idx < 0 | idx > _rows)
            {
            }
            else
            {
                _AutosizeCellsToContents = false;
                _rowheights[idx] = value;
                Invalidate();
            }
        }

        // SCrollBarWeight
        /// <summary>
        /// Gets or sets the height or width of the horizontal and verticle scroll bars on the surface of the grid itself.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets Height or the Horizontal Scroll Bar or Width of the Vertical Scroll bar in Pixels")]
        [DefaultValue(typeof(int), "14")]
        public int ScrollBarWeight
        {
            get
            {
                return _ScrollBarWeight;
            }
            set
            {
                _ScrollBarWeight = value;
                Invalidate();
            }
        }

        // ScrollInterval
        /// <summary>
        /// Gets or sets the amount of screen scroll that the scroll bars will move in pixels
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the amount of screen that a scroll operation will make in Pixels")]
        public int ScrollInterval
        {
            get
            {
                return _scrollinterval;
            }
            set
            {
                _scrollinterval = value;
                Invalidate();
            }
        }

        // SelectedColumn
        /// <summary>
        /// Gets or sets the currently selected column ID. If more than one column is selected then the
        /// <c>SelectedColumns</c> arraylist will contain the set of IDs representative of the selected columns
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the currently selected Column on the grid control")]
        public int SelectedColumn
        {
            get
            {
                return _SelectedColumn;
            }
            set
            {
                if (!(value > _cols - 1 | value < 0))
                {
                    _SelectedColumn = value;
                    Invalidate();
                }
            }
        }

        // SelectedRow
        /// <summary>
        /// Gets or sets the currently selected row ID. If more than one row is selected then the <c>SelectedRows</c>
        /// Arraylist will contain the set of IDs representative of the selected rows
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the currently selected row on the grid control")]
        public int SelectedRow
        {
            get
            {
                return _SelectedRow;
            }
            set
            {
                if (!(value > _rows - 1 | value < -1))
                {
                    _SelectedRow = value;

                    _SelectedRows.Clear();
                    _SelectedRows.Add(_SelectedRow);

                    if (_SelectedRow != -1 & vs.Visible)
                        vs.Value = _SelectedRow;

                    Invalidate();
                }
            }
        }

        // SelectedRows
        /// <summary>
        /// Gets or sets the currently selected row list in the current grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the currently selected row list on the grid control")]
        public ArrayList SelectedRows
        {
            get
            {
                return _SelectedRows;
            }
            set
            {
                _SelectedRows = value;

                Invalidate();
            }
        }

        // SelectedColBackColor
        /// <summary>
        /// Gets or sets the background color for the currently selected column in the grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the background color for the highlighted column in the grid")]
        public Color SelectedColBackColor
        {
            get
            {
                return _ColHighliteBackColor;
            }
            set
            {
                _ColHighliteBackColor = value;
            }
        }

        // SelectedColForeColor
        /// <summary>
        /// Gets or sets the currently selected column foreground color in the grid
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the foreground color for the selected column in the grid")]
        public Color SelectedColForeColor
        {
            get
            {
                return _ColHighliteForeColor;
            }
            set
            {
                _ColHighliteForeColor = value;
            }
        }

        // SelectedRowBackColor
        /// <summary>
        /// Gets or sets the currently selected row background color
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the background color for the highlighted row in the grid")]
        public Color SelectedRowBackColor
        {
            get
            {
                return _RowHighLiteBackColor;
            }
            set
            {
                _RowHighLiteBackColor = value;
            }
        }

        // SelectedRowForeColor
        /// <summary>
        /// Gets or sets the currently selected row foreground color
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Sets the foreground color for the selected row in the grid")]
        public Color SelectedRowForeColor
        {
            get
            {
                return _RowHighLiteForeColor;
            }
            set
            {
                _RowHighLiteForeColor = value;
            }
        }

        // ShowDatesWithTime
        /// <summary>
        /// Allows or disallows the display of the time portion of datetime values read from the various database populators
        /// if disallowed just the date portions of these datatypes will be displayed.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Will Will Turn on or off the expansion of Dates to include time values if they are present")]
        [DefaultValue(typeof(bool), "False")]
        public bool ShowDatesWithTime
        {
            get
            {
                return _ShowDatesWithTime;
            }
            set
            {
                _ShowDatesWithTime = value;
                Refresh();
            }
        }

        // TitleBackColor
        /// <summary>
        /// Gets or sets the background color of the title bar in the displayed grid control
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the background color for the title")]
        public Color TitleBackColor
        {
            get
            {
                return _GridTitleBackcolor;
            }
            set
            {
                _GridTitleBackcolor = value;
                Invalidate();
            }
        }

        // TitleFont
        /// <summary>
        /// Gets or sets the font used to display the title of the grid control
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the font for the title")]
        public Font TitleFont
        {
            get
            {
                return _GridTitleFont;
            }
            set
            {
                _GridTitleFont = value;
                _GridTitleHeight = Conversions.ToInteger(CreateGraphics().MeasureString("Yy", value).Height);
                Invalidate();
            }
        }

        // TitleForeColor
        /// <summary>
        /// Gets or sets the foreground color used to render the title of the grid control
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the color of the text for the title")]
        public Color TitleForeColor
        {
            get
            {
                return _GridTitleForeColor;
            }
            set
            {
                _GridTitleForeColor = value;
                Invalidate();
            }
        }

        // TitleText
        /// <summary>
        /// Gets or sets the actual title text displayed in the grid controls title bar
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the text for the title")]
        public string TitleText
        {
            get
            {
                return _GridTitle;
            }
            set
            {
                _GridTitle = value;
                Invalidate();
            }
        }

        // TitleVisible
        /// <summary>
        /// Allows or disallows the display of the title bar on the grid control itself
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets whether the title should be displayed")]
        public bool TitleVisible
        {
            get
            {
                return _GridTitleVisible;
            }
            set
            {
                _GridTitleVisible = value;
                Invalidate();
            }
        }

        // UserColResizeMinimum
        /// <summary>
        /// Gets or sets the minimum size in pixels the user will be allowed to resize columns to if user column
        /// resizeing is enabled. This settiing prevents users from resizeing columns to 0 pixels in width making them
        /// difficult to make visable again.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the Minimum column width for a user to be able to resize a column to in Px.")]
        [DefaultValue(typeof(int), "5")]
        public int UserColResizeMinimum
        {
            get
            {
                return _UserColResizeMinimum;
            }
            set
            {
                _UserColResizeMinimum = value;
            }
        }

        // VisibleHeight
        /// <summary>
        /// Gets the height of the grid portion of the visable grid in pixels. Minus the height of the visible scrollbars
        /// if the scrollbars are visible
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets the height of the grid minus the horizontal scrollbars height if its visible")]
        public int VisibleHeight
        {
            get
            {
                if (hs.Visible)
                    return Height - hs.Height;
                else
                    return Height;
            }
        }

        // VisibleWidth
        /// <summary>
        /// Gets the width in pixels of the grid area minus the width of the verticle scrollbar if its visible
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets the width of the grid minus the verticle scrollbars width if its visible")]
        public int VisibleWidth
        {
            get
            {
                if (vs.Visible)
                    return Width - vs.Width;
                else
                    return Width;
            }
        }

        // XMLDataSetName
        /// <summary>
        /// Gets or sets he name of the dataset used during the export to XML of the grids contents
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the dataset name used during the export of the grid to XML.")]
        public string XMLDataSetName
        {
            get
            {
                return _xmlDataSetName;
            }
            set
            {
                _xmlDataSetName = value;
            }
        }

        // XMLFileName
        /// <summary>
        /// Gets or sets the filename used to export the contents of the grid to or read from during an XML import operation
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the file name used during the export/import of the grid to XML.")]
        public string XMLFileName
        {
            get
            {
                return _xmlFilename;
            }
            set
            {
                _xmlFilename = value;
            }
        }

        // XMLIncludeSchema
        /// <summary>
        /// Allows or disallows the exporting of the sceme defination in the resulting xml output
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the flag which is used during the export of the grid to XML.")]
        public bool XMLIncludeSchema
        {
            get
            {
                return _xmlIncludeSchema;
            }
            set
            {
                _xmlIncludeSchema = value;
            }
        }

        // XMLNameSpace
        /// <summary>
        /// Gets or sets the namespace used to embed the contents of the grid into when exporting to XML
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the name space used during the export of the grid to XML.")]
        public string XMLNameSpace
        {
            get
            {
                return _xmlNameSpace;
            }
            set
            {
                _xmlNameSpace = value;
            }
        }

        // XMLTableName
        /// <summary>
        /// Gets or sets the table named used to export the grids content into during xml export
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        [Description("Gets or Sets the table name used during the export of the grid to XML.")]
        public string XMLTableName
        {
            get
            {
                return _xmlTableName;
            }
            set
            {
                _xmlTableName = value;
            }
        }
              

        /// <summary>
        /// Takes the DelimitedSTringArray string and splits it up on the Delimiter. Then adds a row to the grids contents
        /// filling the newly added row with the split fields from the DelimitedStringArray.
        /// </summary>
        /// <param name="DelimitedStringArray"></param>
        /// <param name="Delimiter"></param>
        /// <remarks></remarks>
        public void AddRowToGrid(string DelimitedStringArray, string Delimiter)
        {
            bool _oldPainting;

            _oldPainting = _Painting;


            _Painting = true;


            // added Aug 2, 2004
            // Larry found an oddity in this routine when he was adding to the grid that had no columns
            // we decided to make it generate columns by default if none where already in the grid
            // 
            if (_cols == 0)
            {
                // we have no columns in the grid so we need to add some now
                var b = DelimitedStringArray.Split(Conversions.ToChar(Delimiter));
                int xx = 0;

                Cols = b.GetUpperBound(0) + 1;
                var loopTo = _cols - 1;
                for (xx = 0; xx <= loopTo; xx++)
                    set_HeaderLabel(xx, "COLUMN " + xx.ToString());

                AutoSizeCellsToContents = true;

                Refresh();
            }

            // end of Aug 2, 2004

            var a = DelimitedStringArray.Split(Delimiter.ToCharArray(), _cols);
            int x = 0;

            Rows = _rows + 1;
            var loopTo1 = _cols;
            for (x = 0; x <= loopTo1; x++)
                _grid[_rows - 1, x] = "";
            var loopTo2 = a.GetUpperBound(0);
            for (x = 0; x <= loopTo2; x++)
                _grid[_rows - 1, x] = a[x];

            _AutoSizeAlreadyCalculated = false;

            _Painting = _oldPainting;

            Invalidate();
        }

        /// <summary>
        /// Takes the DelimitedSTringArray string and splits it up on the default delimiter of '|'.
        /// Then adds a row to the grids contents filling the newly added row with the split fields
        /// from the DelimitedStringArray.
        /// </summary>
        /// <param name="DelimitedStringArray"></param>
        /// <remarks></remarks>
        public void AddRowToGrid(string DelimitedStringArray)
        {
            AddRowToGrid(DelimitedStringArray, "|");
        }

        /// <summary>
        /// Sets all cells in the grid to be rendered using the Font style specified by fnt
        /// </summary>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void AllCellsUseThisFont(Font fnt)
        {
            int r, c;
            int selrow = -1;
            var loopTo = _rows - 1;

            // If _MouseRow >= 0 And _MouseRow <= _rows - 1 Then
            // selrow = _MouseRow
            // End If

            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = _cols - 1;
                for (c = 0; c <= loopTo1; c++)
                    _gridCellFonts[r, c] = 0;// fnt
            }

            // Me.txtKeyHandler.Top = Me.TAIGCanvas.Top
            // Me.txtKeyHandler.Left = Me.TAIGCanvas.Left
            // TAIGPanel.ScrollControlIntoView(txtKeyHandler)

            _DefaultCellFont = fnt;
            _gridCellFontsList[0] = fnt;
            _AutoSizeAlreadyCalculated = false;

            Invalidate();
        }

        /// <summary>
        /// Sets all cells in the grid to use the forgroundcolor specified by fcol
        /// </summary>
        /// <param name="fcol"></param>
        /// <remarks></remarks>
        public void AllCellsUseThisForeColor(Color fcol)
        {
            int r, c;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = _cols - 1;
                for (c = 0; c <= loopTo1; c++)
                    _gridForeColor[r, c] = 0;
            }

            _DefaultForeColor = fcol;
            _gridForeColorList[0] = new Pen(fcol);

            Invalidate();
        }

        /// <summary>
        /// Will decrease the size of all displayed fonts in the grid by a single point
        /// </summary>
        /// <remarks></remarks>
        public void AllFontsSmaller()
        {
            miFontsSmaller_Click(this, new EventArgs());
            miHeaderFontSmaller_Click(this, new EventArgs());
            miTitleFontSmaller_Click(this, new EventArgs());
        }

        /// <summary>
        /// Will increase the size of all displayed fonts in the grid by a single point
        /// </summary>
        /// <remarks></remarks>
        public void AllFontsLarger()
        {
            miFontsLarger_Click(this, new EventArgs());
            miHeaderFontLarger_Click(this, new EventArgs());
            miTitleFontLarger_Click(this, new EventArgs());
        }

        /// <summary>
        /// Will set all the rows in the grid to use the background color specified by startcolor
        /// </summary>
        /// <param name="startcolor"></param>
        /// <remarks></remarks>
        public void AllRowsThisColor(Color startcolor)
        {
            int r, c;

            _Painting = true;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = Cols - 1;
                for (c = 0; c <= loopTo1; c++)
                    _gridBackColor[r, c] = GetGridBackColorListEntry(new SolidBrush(startcolor));
            }

            _Painting = false;

            Invalidate();
        }

        /// <summary>
        /// Will instruct all data represented in the grid to be colored using the specified color Startcolor
        /// </summary>
        /// <param name="startcolor"></param>
        /// <remarks></remarks>
        public void AllTextThisColor(Color startcolor)
        {
            int r, c;

            _Painting = true;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = Cols - 1;
                for (c = 0; c <= loopTo1; c++)
                    _gridForeColor[r, c] = GetGridForeColorListEntry(new Pen(startcolor));
            }

            _Painting = false;

            Invalidate();
        }

        /// <summary>
        /// Will take the parmeters startcolor and alternatecolor and color every other row in the grid using these two
        /// colors
        /// </summary>
        /// <param name="startcolor"></param>
        /// <param name="alternatecolor"></param>
        /// <remarks></remarks>
        public void AlternateRowColoration(Color startcolor, Color alternatecolor)
        {
            int r, c;
            bool flag = false;

            if (_rows < 2)
                // we ain't got enough rows to alternate colorize
                return;

            _Painting = true;

            _alternateColorationMode = true;
            _alternateColorationBaseColor = startcolor;
            _alternateColorationALTColor = alternatecolor;
            var loopTo = _rows - 1;
            for (r = 1; r <= loopTo; r++)
            {
                var loopTo1 = Cols - 1;
                for (c = 0; c <= loopTo1; c++)
                {
                    if (flag)
                        _gridBackColor[r, c] = GetGridBackColorListEntry(new SolidBrush(alternatecolor));
                    else
                        _gridBackColor[r, c] = GetGridBackColorListEntry(new SolidBrush(startcolor));
                }
                flag = !flag;
            }

            _Painting = false;

            Invalidate();
        }

        /// <summary>
        /// will take the property defined basecolor and altcolor and apply the alternaterowcoloration
        /// process to the contents of the grid
        /// </summary>
        /// <remarks></remarks>
        public void AlternateRowColoration()
        {
            AlternateRowColoration(_alternateColorationBaseColor, _alternateColorationALTColor);
        }

        /// <summary>
        /// Will iterate through the maintained list of tearaway windows attemptiong to place
        /// them on the screen so that they dont overlap each other. Simillar to the old windows arrange windows
        /// menu item from the wfw 1.1 and windows 95/98 days
        /// </summary>
        /// <remarks></remarks>
        public void ArrangeTearAwayWindows()
        {
            int maxy = 0;

            int t;
            var rect = SystemInformation.WorkingArea;
            int x, y;
            TearAwayWindowEntry tear;


            if (TearAways.Count == 0)
                // we ain't got any tearaways so lets bail
                return;

            if (_TearAwayWork)
                return;

            _TearAwayWork = true;
            var loopTo = TearAways.Count - 1;

            // lets see if we can minimize the windows size first here

            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                y = tear.Winform.MaxRenderHeight();

                if (maxy < y)
                    maxy = y;
            }

            // now maxy is the largest windows maximum render height so lets compare

            if (maxy < 350)
            {
                var loopTo1 = TearAways.Count - 1;
                for (t = 0; t <= loopTo1; t++)
                {
                    tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                    tear.Winform.Height = maxy + 10;
                }
            }

            maxy = 0;
            var loopTo2 = TearAways.Count - 1;

            // now to so the moving about


            // first we need to get the height of the largest window
            for (t = 0; t <= loopTo2; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                if (maxy < tear.Winform.Height)
                    maxy = tear.Winform.Height;
            }

            x = 0;
            y = 0;
            var loopTo3 = TearAways.Count - 1;

            // ok now we have the height so  lets start arranging them
            for (t = 0; t <= loopTo3; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];

                if (x + tear.Winform.Width > rect.Width)
                {
                    // that window is off screen so lets organize it down a bit
                    x = 0;
                    if (y + maxy * 2 > rect.Height)
                        y = 0;
                    else
                        y += maxy;
                }

                var loc = new Point(x, y);

                tear.Winform.Location = loc;

                x += tear.Winform.Width;
            }

            _TearAwayWork = false;
        }

        /// <summary>
        /// Erase's the contents of the grid and sets it up to contain 1 row and 1 column
        /// </summary>
        /// <remarks></remarks>
        public void BlankTheGrid()
        {
            Cols = 1;
            Rows = 1;

            ClearAllText();

            set_ColWidth(0, Width);
            set_RowHeight(0, Height);

            _AutoSizeAlreadyCalculated = false;

            Refresh();
        }

        /// <summary>
        /// resets all the columns of the grid to not be displaying boolean datatypes ( CheckBoxes )
        /// </summary>
        /// <remarks></remarks>
        public void ClearAllGridCheckboxStates()
        {
            int c;
            var loopTo = _cols - 1;
            for (c = 0; c <= loopTo; c++)
                _colboolean[c] = false;
            Refresh();
        }

        /// <summary>
        /// Clears the text in the grid but leaves the columns and rows in place
        /// </summary>
        /// <remarks></remarks>
        public void ClearAllText()
        {
            _grid = new string[_rows + 1, _cols + 1];
            _AutoSizeAlreadyCalculated = false;
            Invalidate();
        }

        /// <summary>
        /// Clears the internal column restriction list allows all editable columns to contain any arbritrary
        /// textual data.
        /// </summary>
        /// <remarks></remarks>
        public void ClearAllColumnEditRestrictionLists()
        {
            _colEditRestrictions.Clear();
        }

        /// <summary>
        /// Removes the column edit restrinctions from the columnid designated by colid
        /// </summary>
        /// <param name="colid"></param>
        /// <remarks></remarks>
        public void ClearSpecificColumnEditRestrictionList(int colid)
        {
            foreach (EditColumnRestrictor it in _colEditRestrictions)
            {
                if (it.ColumnID == colid)
                    _colEditRestrictions.Remove(it);
            }
        }

        /// <summary>
        /// Takes the contents of the grid and copys it to the clipboard as a Tab delimited array of text elements
        /// suitable for pasting into excel or word.
        /// </summary>
        /// <remarks></remarks>
        public void CopyGridToClipboard()
        {
            int x, y;
            string s = "";
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                var loopTo1 = _cols - 1;
                for (x = 0; x <= loopTo1; x++)
                {
                    s = s + _grid[y, x];
                    if (x == _cols - 1)
                        s = s + Constants.vbCrLf;
                    else
                        s = s + Constants.vbTab;
                }
            }

            Clipboard.SetDataObject(s);
        }

        /// <summary>
        /// Applys the designated stringformat object sf to the contents of the colum designated as c
        /// </summary>
        /// <param name="c"></param>
        /// <param name="sf"></param>
        /// <remarks></remarks>
        public void ColumnFormat(int c, StringFormat sf)
        {
            int r;

            if (c >= _cols | _rows < 1 | c < 0)
                return;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);

            _AutoSizeAlreadyCalculated = false;

            Refresh();
        }

        /// <summary>
        /// applys the standard format specification of currency to the designated column at colume id C
        /// </summary>
        /// <param name="c"></param>
        /// <remarks></remarks>
        public void ColumnFormatasMoney(int c)
        {
            int r;
            var sf = new StringFormat();

            if (c >= _cols | _rows < 1 | c < 0)
                return;

            // sf.LineAlignment = StringAlignment.Far
            sf.LineAlignment = StringAlignment.Near;
            sf.Alignment = StringAlignment.Far;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                if (Information.IsNumeric(_grid[r, c]))
                {
                    _grid[r, c] = Strings.Format(Conversion.Val(_grid[r, c].Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")), "C");
                    _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);
                }
            }

            _AutoSizeAlreadyCalculated = false;

            Refresh();
        }

        /// <summary>
        /// applys the standard format specification of numbers to the designated column at colume id C
        /// </summary>
        /// <param name="c"></param>
        /// <param name="sFormat"></param>
        /// <remarks></remarks>
        public void ColumnFormatasNumber(int c, string sFormat)
        {
            int r;
            var sf = new StringFormat();

            if (c >= _cols | _rows < 1 | c < 0)
                return;

            // sf.LineAlignment = StringAlignment.Far
            sf.LineAlignment = StringAlignment.Near;
            sf.Alignment = StringAlignment.Far;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                if (Information.IsNumeric(_grid[r, c]))
                {
                    _grid[r, c] = Strings.Format(Conversion.Val(_grid[r, c].Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")), sFormat);
                    _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);
                }
            }

            _AutoSizeAlreadyCalculated = false;

            Refresh();
        }

        /// <summary>
        /// applys the standard format specification of short date to the designated column at colume id C
        /// </summary>
        /// <param name="c"></param>
        /// <remarks></remarks>
        public void ColumnFormatasShortDate(int c)
        {
            int r;

            if (c >= _cols | _rows < 1 | c < 0)
                return;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                if (Information.IsDate(_grid[r, c]))
                    _grid[r, c] = Strings.Format(_grid[r, c], "Short Date");
            }

            _AutoSizeAlreadyCalculated = false;

            Refresh();
        }

        /// <summary>
        /// Will return a string that is representative of an SQL script that will write the gridshape and its contents
        /// to a table in a database that the resulting script is handed to.
        /// </summary>
        /// <param name="tname"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public string CreatePersistanceScript(string tname)
        {
            string result = "";
            string fname = "";
            var maxl = new int[_cols + 1];

            var sb = new StringBuilder();

            int m, r, c;
            var loopTo = _cols - 1;

            // figure out how large each column needs to be to persist the grids contents

            for (c = 0; c <= loopTo; c++)
            {
                m = 0;
                var loopTo1 = _rows - 1;
                for (r = 0; r <= loopTo1; r++)
                {
                    if (_grid[r, c] == null)
                    {
                    }
                    else if (_grid[r, c].Length > m)
                        m = _grid[r, c].Length;
                }
                // range check the size of the result to the maximum varchar size
                if (m > 8000)
                    m = 8000;

                // if we got no data then set the filed to hold something
                if (m == 0)
                    m = 10;

                // set the size in the array
                maxl[c] = m;
            }

            result = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" + tname + "]') " + "and OBJECTPROPERTY(id, N'IsUserTable') = 1)" + Constants.vbCrLf + "drop table [dbo].[" + tname + "] " + Constants.vbCrLf + "GO" + Constants.vbCrLf + Constants.vbCrLf;



            result += "CREATE TABLE [dbo].[" + tname + "] (" + Constants.vbCrLf;

            result += Constants.vbTab + "[ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," + Constants.vbCrLf;
            var loopTo2 = _cols - 1;
            for (c = 0; c <= loopTo2; c++)
            {
                fname = get_HeaderLabel(c).ToUpper();

                if ((fname ?? "") == "ID")
                    fname = "ID_DATA";

                result += Constants.vbTab + "[" + fname + "] [VARCHAR] (" + maxl[c].ToString() + ") NULL ," + Constants.vbCrLf;
            }

            result += ") ON [PRIMARY]" + Constants.vbCrLf + "GO " + Constants.vbCrLf + Constants.vbCrLf;

            sb.Append(result);

            result = "";
            var loopTo3 = _rows - 1;
            for (r = 0; r <= loopTo3; r++)
            {
                result = "INSERT INTO [" + tname + "] (";
                var loopTo4 = _cols - 1;
                for (c = 0; c <= loopTo4; c++)
                {
                    fname = get_HeaderLabel(c).ToUpper();

                    if ((fname ?? "") == "ID")
                        fname = "ID_DATA";
                    result += "[" + fname + "],";
                }

                result = result.Substring(0, result.Length - 1) + ") VALUES (";
                var loopTo5 = _cols - 1;
                for (c = 0; c <= loopTo5; c++)
                {
                    if (_grid[r, c] == null)
                        fname = "{null}";
                    else
                        fname = _grid[r, c];

                    if (fname.Length > maxl[c])
                        fname = fname.Substring(0, maxl[c]).Replace("'", "''");
                    else
                        fname = fname.Replace("'", "''");

                    if ((fname ?? "") == "{null}")
                        result += "NULL";
                    else
                        result += "'" + fname + "'";

                    if (c < Cols - 1)
                        result += ",";
                    else
                        result += ")" + Constants.vbCrLf + "GO" + Constants.vbCrLf + Constants.vbCrLf;
                }

                sb.Append(result);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Will return a string containing an HTML table representation of the grids contents
        /// Borderval is the size parameter of the tables borders
        /// Matchcolors will turn on or off the attempt to set the colors of the table to match thos of the grid itself
        /// OmitNulls will have the rendering of empty cells in the grid or not. (creating holes in the resuting
        /// html output)
        /// </summary>
        /// <param name="BorderVal"></param>
        /// <param name="MatchColors"></param>
        /// <param name="OmitNulls"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public string CreateHTMLTableScript(int BorderVal, bool MatchColors, bool OmitNulls)
        {
            var rb = new StringBuilder();
            string result = "";
            string rr = "";

            int r, c;

            if (MatchColors)
            {
                if (BorderVal > 0)
                    rb.Append("<TABLE BGCOLOR = " + ReturnHTMLColor(BackColor) + " " + "BORDER = " + Conversions.ToString((char)34) + BorderVal.ToString() + Conversions.ToString((char)34) + ">" + Constants.vbCrLf);
                else
                    rb.Append("<TABLE BGCOLOR = " + ReturnHTMLColor(BackColor) + " " + ">" + Constants.vbCrLf);

                rb.Append("<TR><TD BGCOLOR = " + ReturnHTMLColor(TitleBackColor) + " COLSPAN =" + Conversions.ToString((char)34) + _cols.ToString() + Conversions.ToString((char)34) + ">");
                rb.Append(TitleText + "</TD></TR>" + Constants.vbCrLf);

                rb.Append("<TR>");
                var loopTo = _cols - 1;
                for (c = 0; c <= loopTo; c++)
                    rb.Append("<TH BGCOLOR=" + ReturnHTMLColor(GridHeaderBackColor) + ">" + get_HeaderLabel(c) + "</TH>");

                rb.Append("</TR>" + Constants.vbCrLf);
                var loopTo1 = _rows - 1;
                for (r = 0; r <= loopTo1; r++)
                {
                    rb.Append("<TR>");
                    var loopTo2 = _cols - 1;
                    for (c = 0; c <= loopTo2; c++)
                    {
                        SolidBrush sb = (System.Drawing.SolidBrush)_gridBackColorList[_gridBackColor[r, c]];
                        string txt = _grid[r, c] + "";

                        if (OmitNulls & (txt ?? "") == "{null}")
                            txt = "";

                        rb.Append("<TD BGCOLOR =" + ReturnHTMLColor(sb.Color) + ">" + txt + "</TD>");
                    }
                    rb.Append("</TR>" + Constants.vbCrLf);
                }

                rb.Append("</TABLE>" + Constants.vbCrLf);
            }
            else
            {
                if (BorderVal > 0)
                    rb.Append("<TABLE BORDER = " + Conversions.ToString((char)34) + BorderVal.ToString() + Conversions.ToString((char)34) + ">" + Constants.vbCrLf);
                else
                    rb.Append("<TABLE>" + Constants.vbCrLf);

                rb.Append("<TR><TD COLSPAN =" + Conversions.ToString((char)34) + _cols.ToString() + Conversions.ToString((char)34) + ">");
                rb.Append(TitleText + "</TD></TR>" + Constants.vbCrLf);

                result += "<TR>";
                var loopTo3 = _cols - 1;
                for (c = 0; c <= loopTo3; c++)
                    rb.Append("<TH>" + get_HeaderLabel(c) + "</TH>");

                rb.Append("</TR>" + Constants.vbCrLf);
                var loopTo4 = _rows - 1;
                for (r = 0; r <= loopTo4; r++)
                {
                    rb.Append("<TR>");
                    var loopTo5 = _cols - 1;
                    for (c = 0; c <= loopTo5; c++)
                    {
                        string txt = _grid[r, c] + "";

                        if (OmitNulls & (txt ?? "") == "{null}")
                            txt = "";

                        rb.Append("<TD>" + txt + "</TD>");
                    }
                    rb.Append("</TR>" + Constants.vbCrLf);
                }

                rb.Append("</TABLE>" + Constants.vbCrLf);
            }

            return rb.ToString();
        }

        /// <summary>
        /// Overload that will set the border to 1 pixel, Matchcolors, and omitnulls
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public string CreateHTMLTableScript()
        {
            // Will use the default border of 1
            // Will Matchcolors
            // Will OmitNulls

            return CreateHTMLTableScript(1, true, true);
        }

        /// <summary>
        /// Will deslect all rows in the grid if any are selected
        /// </summary>
        /// <remarks></remarks>
        public void DeSelectAllRows()
        {
            _SelectedRows.Clear();
            _SelectedRow = -1;
            Refresh();
        }

        /// <summary>
        /// Will set the tooltip on the mouse to match the cells contents that its hovering over
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ttText"></param>
        /// <remarks></remarks>
        public void DisplayGridToolTip(object sender, string ttText)
        {
            if (_TearAwayWork)
                return;


            if (sender is TAIGridControl)
            {
                _TTip.SetToolTip((System.Windows.Forms.Control)sender, ttText);
                _TTip.ShowAlways = true;
                _TTip.Active = true;
            }
            else if (sender is frmColumnTearAway)
            {
                // we are a hovering on a tearaway form so lets pass it to that form

                int t;
                var f = (frmColumnTearAway)sender;
                var loopTo = TearAways.Count - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    var ti = (TearAwayWindowEntry)TearAways[t];

                    if (f.Colid == ti.ColID)
                    {
                        ti.Winform.ShowToolTipOnForm(ttText);
                        break;
                    }
                }
            }
            else
                _TTip.SetToolTip((System.Windows.Forms.Control)sender, ttText);
        }

        /// <summary>
        /// Will hide the tooltip on the mouse pointer if its visible
        /// </summary>
        /// <remarks></remarks>
        public void HideGridToolTip()
        {
            if (_TearAwayWork)
                return;

            _TTip.Active = false;

            if (TearAways.Count == 0)
                return;

            int t;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                TearAwayWindowEntry ti;

                ti = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];

                ti.Winform.HideToolTipOnForm();
            }
        }

        /// <summary>
        /// The <c>DoControlBreakProcessing</c> method will as its name indicates conduct a good oldfasioned sub-total
        /// and grand-total parse on the contents of the grid using the old style Cobol rules for control break processing.
        /// <list type="bullet">
        /// <item>
        /// <c>BreakColIntArrayList</c> list needs to contain the column IDs that will be looked at to determine
        /// where to break and subtotal.
        /// </item>
        /// <item>
        /// <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
        /// be calculated.
        /// </item>
        /// <item>
        /// <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
        /// before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
        /// </item>
        /// <item>
        /// <c>ColumnToPlaceSubTotalTextIn</c> indicates the column to reiterate the criteria for the break into.
        /// Think of this as the label to apply to the subtotal rows.
        /// </item>
        /// <item>
        /// <c>SubtotalText</c> is an arbritrary string of text to be appended to the lable defined above.
        /// </item>
        /// <item>
        /// <c>RightAlignSubTotalText</c> will allow or disallow the right aligning of the resulting subtotal lables.
        /// </item>
        /// <item>
        /// <c>ColorForSubTotalRows</c> defines the backgroundcolor to use when inserting a subtotal row into the
        /// resulting output.
        /// </item>
        /// <item>
        /// <c>BlankSeperateBreaks</c> will allow or disallow the insertion of an additional blank row after a
        /// subtotal operation.
        /// </item>
        /// <item>
        /// <c>EchoBreakFieldsOnSubTotalLines</c> will allow or disallow the echoing of the criteria for the
        /// above subtotal operation on the line with the subtotal figures itself.
        /// </item>
        /// <item>
        /// <c>TreatBlanksAsSame</c> will force the parser to treat a blank field in a row to be treated as the
        /// most recent previous non blank field for the purposes of determining that the break is necessary.
        /// </item>
        /// </list>
        /// Note:
        /// Because the control break process works from top to bottom on the current contents of the grid
        /// those contents should be sorted as the results will not have any real meaning if the grids contents
        /// are not sorted before the call to this method.
        /// </summary>
        /// <param name="BreakColIntArrayList"></param>
        /// <param name="SumColumnIntegerArraylist"></param>
        /// <param name="IgnoreCase"></param>
        /// <param name="ColumnToPlaceSubtotalTextIn"></param>
        /// <param name="SubtotalText"></param>
        /// <param name="RightAlignSubTotalText"></param>
        /// <param name="ColorForSubTotalRows"></param>
        /// <param name="BlankSeperateBreaks"></param>
        /// <param name="EchoBreakFieldsOnSubtotalLines"></param>
        /// <param name="TreatBlanksAsSame"></param>
        /// <remarks></remarks>
        public void DoControlBreakProcessing(ArrayList BreakColIntArrayList, ArrayList SumColumnIntegerArraylist, bool IgnoreCase, int ColumnToPlaceSubtotalTextIn, string SubtotalText, bool RightAlignSubTotalText, Color ColorForSubTotalRows, bool BlankSeperateBreaks, bool EchoBreakFieldsOnSubtotalLines, bool TreatBlanksAsSame)
        {
            string breakstring = "";
            string oldbreak = "";
            string gvalue = "";
            var scols = new double[SumColumnIntegerArraylist.Count];
            var tscols = new double[SumColumnIntegerArraylist.Count];

            var sfmt = new StringFormat();

            sfmt.Alignment = StringAlignment.Far;
            sfmt.LineAlignment = StringAlignment.Center;


            var hdr = _GridHeader;

            int x, y, xx = default(int), ngridcurrow = default(int);
            var loopTo = scols.GetUpperBound(0);
            for (x = 0; x <= loopTo; x++)
            {
                // init the sums colums here to 0
                scols[x] = 0;
                tscols[x] = 0;
            }

            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var loopTo2 = BreakColIntArrayList.Count - 1;

                // lets calculate how many breaks we are going to have

                for (x = 0; x <= loopTo2; x++)
                {
                    // start off the Break

                    if (string.IsNullOrEmpty(_grid[y, Conversions.ToInteger(BreakColIntArrayList[x])]) & TreatBlanksAsSame)
                    {
                        var sss = oldbreak.Split("|".ToCharArray());
                        if (sss.GetUpperBound(0) >= x)
                            breakstring += sss[x] + "|";
                        else
                            breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
                    }
                    else
                        breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
                }

                if (IgnoreCase)
                    breakstring = breakstring.ToUpper();

                if ((breakstring ?? "") != (oldbreak ?? ""))
                {
                    xx += 1;
                    if (BlankSeperateBreaks)
                        xx += 1;
                    oldbreak = breakstring;
                }

                breakstring = "";
            }

            // dimension our new grid to get it ready to hold the manipulated data

            var ngrid = new string[_rows + xx + 1, _cols];

            // calculate the first breakstring

            breakstring = "";
            var loopTo3 = BreakColIntArrayList.Count - 1;
            for (x = 0; x <= loopTo3; x++)
            {
                // start off the break
                if (string.IsNullOrEmpty(_grid[y, Conversions.ToInteger(BreakColIntArrayList[x])]) & TreatBlanksAsSame)
                {
                    var sss = oldbreak.Split("|".ToCharArray());
                    if (sss.GetUpperBound(0) >= x)
                        breakstring += sss[x] + "|";
                    else
                        breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
                }
                else
                    breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
            }

            if (IgnoreCase)
                breakstring = breakstring.ToUpper();

            oldbreak = "";
            var loopTo4 = _rows - 1;
            for (y = 0; y <= loopTo4; y++)
            {
                // our main loop here

                breakstring = "";
                var loopTo5 = BreakColIntArrayList.Count - 1;
                for (x = 0; x <= loopTo5; x++)
                {
                    // start off the break
                    if (string.IsNullOrEmpty(_grid[y, Conversions.ToInteger(BreakColIntArrayList[x])]) & TreatBlanksAsSame)
                    {
                        var sss = oldbreak.Split("|".ToCharArray());
                        if (sss.GetUpperBound(0) >= x)
                            breakstring += sss[x] + "|";
                        else
                            breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
                    }
                    else
                        breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";
                }

                if (IgnoreCase)
                    breakstring = breakstring.ToUpper();

                if ((breakstring ?? "") != (oldbreak ?? "") | string.IsNullOrEmpty(oldbreak))
                {
                    // we have a break
                    // are we on the first break
                    if (string.IsNullOrEmpty(oldbreak))
                    {
                        var loopTo6 = BreakColIntArrayList.Count - 1;
                        // yes we is
                        for (x = 0; x <= loopTo6; x++)
                            // breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                            ngrid[ngridcurrow, Conversions.ToInteger(BreakColIntArrayList[x])] = _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])];

                        oldbreak = breakstring;
                    }
                    else
                    {
                        var loopTo7 = SumColumnIntegerArraylist.Count - 1;

                        // we have a real break here so we need to display the subtotal lines

                        for (x = 0; x <= loopTo7; x++)
                            ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = scols[x].ToString();

                        if (EchoBreakFieldsOnSubtotalLines)
                        {
                            if (oldbreak.EndsWith("|"))
                                oldbreak = oldbreak.Substring(0, oldbreak.Length - 1);

                            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = oldbreak + " " + SubtotalText;
                        }
                        else
                            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = SubtotalText;
                        var loopTo8 = scols.GetUpperBound(0);


                        // clear the sum columns now
                        for (x = 0; x <= loopTo8; x++)
                            // init the sums colums here to 0
                            scols[x] = 0;

                        ngridcurrow += 1;

                        if (BlankSeperateBreaks)
                            ngridcurrow += 1;
                        var loopTo9 = BreakColIntArrayList.Count - 1;
                        for (x = 0; x <= loopTo9; x++)
                            // breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                            ngrid[ngridcurrow, Conversions.ToInteger(BreakColIntArrayList[x])] = _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])];

                        oldbreak = breakstring;
                    }
                }

                var loopTo10 = SumColumnIntegerArraylist.Count - 1;

                // here we have to sum up the selected columns

                for (x = 0; x <= loopTo10; x++)
                {

                    // lets sanitize the grid value to make it convert numerically
                    gvalue = _grid[y, Conversions.ToInteger(SumColumnIntegerArraylist[x])].Replace("$", "").Replace("(", "-").Replace(")", "");

                    if (Information.IsNumeric(gvalue))
                    {
                        if ((gvalue.Split(".".ToCharArray())[0] ?? "") == (gvalue ?? ""))
                        {
                            // singe a split returns the same value in position 0 then we dont have a decimal
                            // treat it as an integer

                            scols[x] += Conversions.ToInteger(gvalue);
                            tscols[x] += Conversions.ToInteger(gvalue);
                        }
                        else
                        {
                            // its got a decimal in it so treat it as a double

                            scols[x] += Conversions.ToDouble(gvalue);
                            tscols[x] += Conversions.ToDouble(gvalue);
                        }
                    }
                }

                var loopTo11 = _cols - 1;
                for (x = 0; x <= loopTo11; x++)
                {
                    // here we have to echo the Non Break columns and their values into the results grid
                    if (BreakColIntArrayList.Contains(x))
                    {
                    }
                    else
                        ngrid[ngridcurrow, x] = _grid[y, x];
                }

                ngridcurrow += 1;
            }

            var loopTo12 = SumColumnIntegerArraylist.Count - 1;

            // now lets do the final break here

            for (x = 0; x <= loopTo12; x++)
                ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = scols[x].ToString();

            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = SubtotalText;

            ngridcurrow += 1;

            if (BlankSeperateBreaks)
                ngridcurrow += 1;
            var loopTo13 = SumColumnIntegerArraylist.Count - 1;
            for (x = 0; x <= loopTo13; x++)
                ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = tscols[x].ToString();

            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = "Grand Total";


            // now lets push the new grid into the old grids contents

            PopulateGridFromArray(ngrid, _DefaultCellFont, _DefaultForeColor, false, false, hdr);
            var loopTo14 = _rows - 1;
            for (y = 0; y <= loopTo14; y++)
            {
                if (_grid[y, ColumnToPlaceSubtotalTextIn] == null)
                {
                }
                else if (RightAlignSubTotalText & _grid[y, ColumnToPlaceSubtotalTextIn].EndsWith(SubtotalText))
                {
                    set_CellAlignment(y, ColumnToPlaceSubtotalTextIn, sfmt);
                    var loopTo15 = _cols - 1;
                    for (xx = 0; xx <= loopTo15; xx++)
                        set_CellBackColor(y, xx, new SolidBrush(ColorForSubTotalRows));
                }
                else if (_grid[y, ColumnToPlaceSubtotalTextIn].EndsWith(SubtotalText))
                {
                    var loopTo16 = _cols - 1;
                    for (xx = 0; xx <= loopTo16; xx++)
                        set_CellBackColor(y, xx, new SolidBrush(ColorForSubTotalRows));
                }
            }

            Invalidate();
        }

        /// <summary>
        /// The <c>DoControlBreakProcessing</c> method will as its name indicates conduct a good oldfasioned sub-total
        /// and grand-total parse on the contents of the grid using the old style Cobol rules for control break processing.
        /// <list type="bullet">
        /// <item>
        /// <c>BreakColIntArrayList</c> list needs to contain the column IDs that will be looked at to determine
        /// where to break and subtotal.
        /// </item>
        /// <item>
        /// <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
        /// be calculated.
        /// </item>
        /// <item>
        /// <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
        /// before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
        /// </item>
        /// <item>
        /// <c>ColumnToPlaceSubTotalTextIn</c> indicates the column to reiterate the criteria for the break into.
        /// Think of this as the label to apply to the subtotal rows.
        /// </item>
        /// <item>
        /// <c>SubtotalText</c> is an arbritrary string of text to be appended to the lable defined above.
        /// </item>
        /// <item>
        /// <c>RightAlignSubTotalText</c> will allow or disallow the right aligning of the resulting subtotal lables.
        /// </item>
        /// <item>
        /// <c>ColorForSubTotalRows</c> defines the backgroundcolor to use when inserting a subtotal row into the
        /// resulting output.
        /// </item>
        /// <item>
        /// <c>BlankSeperateBreaks</c> will allow or disallow the insertion of an additional blank row after a
        /// subtotal operation.
        /// </item>
        /// <item>
        /// <c>EchoBreakFieldsOnSubTotalLines</c> will allow or disallow the echoing of the criteria for the
        /// above subtotal operation on the line with the subtotal figures itself.
        /// </item>
        /// 
        /// </list>
        /// Note:
        /// Because the control break process works from top to bottom on the current contents of the grid
        /// those contents should be sorted as the results will not have any real meaning if the grids contents
        /// are not sorted before the call to this method.
        /// </summary>
        /// <param name="BreakColIntArrayList"></param>
        /// <param name="SumColumnIntegerArraylist"></param>
        /// <param name="IgnoreCase"></param>
        /// <param name="ColumnToPlaceSubtotalTextIn"></param>
        /// <param name="SubtotalText"></param>
        /// <param name="RightAlignSubTotalText"></param>
        /// <param name="ColorForSubTotalRows"></param>
        /// <param name="BlankSeperateBreaks"></param>
        /// <param name="EchoBreakFieldsOnSubtotalLines"></param>
        /// <remarks></remarks>
        public void DoControlBreakProcessing(ArrayList BreakColIntArrayList, ArrayList SumColumnIntegerArraylist, bool IgnoreCase, int ColumnToPlaceSubtotalTextIn, string SubtotalText, bool RightAlignSubTotalText, Color ColorForSubTotalRows, bool BlankSeperateBreaks, bool EchoBreakFieldsOnSubtotalLines)
        {
            string breakstring = "";
            string oldbreak = "";
            string gvalue = "";
            var scols = new double[SumColumnIntegerArraylist.Count];
            var tscols = new double[SumColumnIntegerArraylist.Count];

            var sfmt = new StringFormat();

            sfmt.Alignment = StringAlignment.Far;
            sfmt.LineAlignment = StringAlignment.Center;


            var hdr = _GridHeader;

            int x, y, xx = default(int), ngridcurrow = default(int);
            var loopTo = scols.GetUpperBound(0);
            for (x = 0; x <= loopTo; x++)
            {
                // init the sums colums here to 0
                scols[x] = 0;
                tscols[x] = 0;
            }

            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var loopTo2 = BreakColIntArrayList.Count - 1;

                // lets calculate how many breaks we are going to have

                for (x = 0; x <= loopTo2; x++)
                    // start off the Break
                    breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";

                if (IgnoreCase)
                    breakstring = breakstring.ToUpper();

                if ((breakstring ?? "") != (oldbreak ?? ""))
                {
                    xx += 1;
                    if (BlankSeperateBreaks)
                        xx += 1;
                    oldbreak = breakstring;
                }

                breakstring = "";
            }

            // dimension our new grid to get it ready to hold the manipulated data

            var ngrid = new string[_rows + xx + 1, _cols];

            // calculate the first breakstring

            breakstring = "";
            var loopTo3 = BreakColIntArrayList.Count - 1;
            for (x = 0; x <= loopTo3; x++)
                // start off the Break
                breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";

            if (IgnoreCase)
                breakstring = breakstring.ToUpper();

            oldbreak = "";
            var loopTo4 = _rows - 1;
            for (y = 0; y <= loopTo4; y++)
            {
                // our main loop here

                breakstring = "";
                var loopTo5 = BreakColIntArrayList.Count - 1;
                for (x = 0; x <= loopTo5; x++)
                    // start off the Break
                    breakstring += _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])] + "|";

                if (IgnoreCase)
                    breakstring = breakstring.ToUpper();

                if ((breakstring ?? "") != (oldbreak ?? "") | string.IsNullOrEmpty(oldbreak))
                {
                    // we have a break
                    // are we on the first break
                    if (string.IsNullOrEmpty(oldbreak))
                    {
                        var loopTo6 = BreakColIntArrayList.Count - 1;
                        // yes we is
                        for (x = 0; x <= loopTo6; x++)
                            // breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                            ngrid[ngridcurrow, Conversions.ToInteger(BreakColIntArrayList[x])] = _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])];

                        oldbreak = breakstring;
                    }
                    else
                    {
                        var loopTo7 = SumColumnIntegerArraylist.Count - 1;

                        // we have a real break here so we need to display the subtotal lines

                        for (x = 0; x <= loopTo7; x++)
                            ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = scols[x].ToString();

                        if (EchoBreakFieldsOnSubtotalLines)
                        {
                            if (oldbreak.EndsWith("|"))
                                oldbreak = oldbreak.Substring(0, oldbreak.Length - 1);

                            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = oldbreak + " " + SubtotalText;
                        }
                        else
                            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = SubtotalText;
                        var loopTo8 = scols.GetUpperBound(0);


                        // clear the sum columns now
                        for (x = 0; x <= loopTo8; x++)
                            // init the sums colums here to 0
                            scols[x] = 0;

                        ngridcurrow += 1;

                        if (BlankSeperateBreaks)
                            ngridcurrow += 1;
                        var loopTo9 = BreakColIntArrayList.Count - 1;
                        for (x = 0; x <= loopTo9; x++)
                            // breakstring += _grid(y, CInt(BreakColIntArrayList.Item(x))) + "|"
                            ngrid[ngridcurrow, Conversions.ToInteger(BreakColIntArrayList[x])] = _grid[y, Conversions.ToInteger(BreakColIntArrayList[x])];

                        oldbreak = breakstring;
                    }
                }

                var loopTo10 = SumColumnIntegerArraylist.Count - 1;

                // here we have to sum up the selected columns

                for (x = 0; x <= loopTo10; x++)
                {

                    // lets sanitize the grid value to make it convert numerically
                    gvalue = _grid[y, Conversions.ToInteger(SumColumnIntegerArraylist[x])].Replace("$", "").Replace("(", "-").Replace(")", "");

                    if (Information.IsNumeric(gvalue))
                    {
                        if ((gvalue.Split(".".ToCharArray())[0] ?? "") == (gvalue ?? ""))
                        {
                            // singe a split returns the same value in position 0 then we dont have a decimal
                            // treat it as an integer

                            scols[x] += Conversions.ToInteger(gvalue);
                            tscols[x] += Conversions.ToInteger(gvalue);
                        }
                        else
                        {
                            // its got a decimal in it so treat it as a double

                            scols[x] += Conversions.ToDouble(gvalue);
                            tscols[x] += Conversions.ToDouble(gvalue);
                        }
                    }
                }

                var loopTo11 = _cols - 1;
                for (x = 0; x <= loopTo11; x++)
                {
                    // here we have to echo the Non Break columns and their values into the results grid
                    if (BreakColIntArrayList.Contains(x))
                    {
                    }
                    else
                        ngrid[ngridcurrow, x] = _grid[y, x];
                }

                ngridcurrow += 1;
            }

            var loopTo12 = SumColumnIntegerArraylist.Count - 1;

            // now lets do the final break here

            for (x = 0; x <= loopTo12; x++)
                ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = scols[x].ToString();

            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = SubtotalText;

            ngridcurrow += 1;

            if (BlankSeperateBreaks)
                ngridcurrow += 1;
            var loopTo13 = SumColumnIntegerArraylist.Count - 1;
            for (x = 0; x <= loopTo13; x++)
                ngrid[ngridcurrow, Conversions.ToInteger(SumColumnIntegerArraylist[x])] = tscols[x].ToString();

            ngrid[ngridcurrow, ColumnToPlaceSubtotalTextIn] = "Grand Total";


            // now lets push the new grid into the old grids contents

            PopulateGridFromArray(ngrid, _DefaultCellFont, _DefaultForeColor, false, false, hdr);
            var loopTo14 = _rows - 1;
            for (y = 0; y <= loopTo14; y++)
            {
                if (_grid[y, ColumnToPlaceSubtotalTextIn] == null)
                {
                }
                else if (RightAlignSubTotalText & (_grid[y, ColumnToPlaceSubtotalTextIn].EndsWith(SubtotalText) | _grid[y, ColumnToPlaceSubtotalTextIn].EndsWith("Grand Total")))
                {
                    set_CellAlignment(y, ColumnToPlaceSubtotalTextIn, sfmt);
                    var loopTo15 = _cols - 1;
                    for (xx = 0; xx <= loopTo15; xx++)
                        set_CellBackColor(y, xx, new SolidBrush(ColorForSubTotalRows));
                }
                else if (_grid[y, ColumnToPlaceSubtotalTextIn].EndsWith(SubtotalText) | _grid[y, ColumnToPlaceSubtotalTextIn].EndsWith("Grand Total"))
                {
                    var loopTo16 = _cols - 1;
                    for (xx = 0; xx <= loopTo16; xx++)
                        set_CellBackColor(y, xx, new SolidBrush(ColorForSubTotalRows));
                }
            }

            Invalidate();
        }

        /// <summary>
        /// This will take the supplied <c>BreakColIntArrayList</c> and do a cell colorization operation
        /// on the current grids contents alternating between <c>StartColor</c> and <c>AltColor</c> on a change
        /// in the cols id'd in the supplied arraylist.
        /// </summary>
        /// <param name="BreakColIntArrayList"></param>
        /// <param name="StartColor"></param>
        /// <param name="AltColor"></param>
        /// <remarks></remarks>
        public void DoControlBreakColorization(ArrayList BreakColIntArrayList, Color StartColor, Color AltColor)
        {
            bool start = false;
            string a = "--------------------------------------------------------";

            int t = 0;
            var loopTo = Rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                string aa = "";

                int tt = 0;
                var loopTo1 = BreakColIntArrayList.Count - 1;
                for (tt = 0; tt <= loopTo1; tt++)
                    aa += _grid[t, tt];

                if ((aa ?? "") != (a ?? ""))
                {
                    start = !start;
                    a = aa;
                }

                var loopTo2 = Cols - 1;
                for (tt = 0; tt <= loopTo2; tt++)
                {
                    if (start)
                        set_CellBackColor(t, tt, new SolidBrush(StartColor));
                    else
                        set_CellBackColor(t, tt, new SolidBrush(AltColor));
                }
            }

            Refresh();
        }

        /// <summary>
        /// Another control break process,
        /// <list type="bullet">
        /// <item>
        /// <c>BreakColIntArrayValues</c> list needs to contain the column IDs that will be looked at to determine
        /// where to break and subtotal.
        /// </item>
        /// <item>
        /// <c>IgnoreCase</c> will insruct the parser to convert everything to uppercase
        /// before it determines if a transition is occuring and thus a break and subtotal operation is necessary.
        /// </item>
        /// <item>
        /// <c>SumColumnIntegerArrayList</c> a list of column ids on which the sums will
        /// be calculated.
        /// </item>
        /// <item>
        /// <c>ColorForBreakSubtotals</c> defines the backgroundcolor to use when inserting a subtotal row into the
        /// resulting output.
        /// </item>
        /// <item>
        /// <c>CutoffRow</c> The maximum row to search for in the grid for processing.
        /// will stop processing at <c>CutoffRow</c>
        /// </item>
        /// 
        /// </list>
        /// Note:
        /// Because the control break process works from top to bottom on the current contents of the grid
        /// those contents should be sorted as the results will not have any real meaning if the grids contents
        /// are not sorted before the call to this method.
        /// </summary>
        /// <param name="BreakColArrayValues"></param>
        /// <param name="ColToFindValues"></param>
        /// <param name="IgnoreCase"></param>
        /// <param name="SumColumnIntegerArrayList"></param>
        /// <param name="ColorForBreakSubtotals"></param>
        /// <param name="CutoffRow"></param>
        /// <remarks></remarks>
        public void DoControlBreakSubTotals(ArrayList BreakColArrayValues, int ColToFindValues, bool IgnoreCase, ArrayList SumColumnIntegerArrayList, Color ColorForBreakSubtotals, int CutoffRow)
        {
            int r = _rows + 1;
            int x, y, xx;
            string s, ss;

            var scols = new double[SumColumnIntegerArrayList.Count];

            Rows += BreakColArrayValues.Count + 1;
            var loopTo = BreakColArrayValues.Count - 1;
            for (x = 0; x <= loopTo; x++)
            {
                var loopTo1 = scols.GetUpperBound(0);
                // lets iterate through the BreakItems

                // clear the sum columns first
                for (xx = 0; xx <= loopTo1; xx++)
                    scols[xx] = 0;

                s = Conversions.ToString(BreakColArrayValues[x]);
                var loopTo2 = CutoffRow;
                for (xx = 0; xx <= loopTo2; xx++)
                {
                    if (IgnoreCase)
                    {
                        if ((Strings.UCase(s) ?? "") == (Strings.UCase(_grid[xx, ColToFindValues]) ?? ""))
                        {
                            var loopTo3 = SumColumnIntegerArrayList.Count - 1;
                            for (y = 0; y <= loopTo3; y++)
                            {
                                ss = _grid[xx, Conversions.ToInteger(SumColumnIntegerArrayList[y])] + "";
                                ss = ss.Replace("(", "-").Replace("$", "").Replace(")", "");

                                if (Information.IsNumeric(ss))
                                    scols[y] += Convert.ToDouble(ss);
                            }
                        }
                    }
                    else if ((s ?? "") == (_grid[xx, ColToFindValues] ?? ""))
                    {
                        var loopTo4 = SumColumnIntegerArrayList.Count - 1;
                        for (y = 0; y <= loopTo4; y++)
                        {
                            ss = _grid[xx, Conversions.ToInteger(SumColumnIntegerArrayList[y])] + "";
                            ss = ss.Replace("(", "-").Replace("$", "").Replace(")", "");

                            if (Information.IsNumeric(ss))
                                scols[y] += Convert.ToDouble(ss);
                        }
                    }
                }

                _grid[r + x, ColToFindValues] = s;
                var loopTo5 = SumColumnIntegerArrayList.Count - 1;
                for (xx = 0; xx <= loopTo5; xx++)
                    _grid[r + x, Conversions.ToInteger(SumColumnIntegerArrayList[xx])] = scols[xx].ToString();

                SetRowBackColor(r + x, ColorForBreakSubtotals);
            }

            Refresh();
        }

        public void DoSelectedRowHighlight()
        {
        }

        /// <summary>
        /// Will select and highlight the indicated rownum as if the user had selected it with the mouse
        /// </summary>
        /// <param name="rownum"></param>
        /// <remarks></remarks>
        public void DoSelectedRowHighlight(int rownum)
        {
            if (rownum < 0 | rownum >= _rows)
                return;

            if (vs.Visible)
                vs.Value = 0;

            SelectedRow = rownum;
            // _SelectedRow = rownum

            Invalidate();
        }

        /// <summary>
        /// Will instance Microsoft excel and place the contents of the grid on the first worksheet in the excel application
        /// </summary>
        /// <remarks></remarks>
        public void ExportToExcel()
        {

            try
            {
                string filetogenerate = ""; // System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EXCELOUTPUT.xlsx");

                if (_OldContextMenu != null)
                {
                    SendKeys.Send("{ESC}");
                    SendKeys.Send("{ESC}");
                }

                var frm1 = new frmExcelOutput(this);

                frm1.ShowDialog();

                if (frm1.FRMOK)
                {
                    filetogenerate = frm1.SELECTEDPATH;

                    var frm = new frmExportingToExcelWorking();

                    if (_ShowExcelExportMessage)
                    {
                        frm.Show();
                        frm.Refresh();
                    }

                    IXLWorkbook workbook = new XLWorkbook();

                    //IXLWorksheet worksheet = workbook.Worksheets.Add(_excelWorkSheetName);
                    IXLWorksheet worksheet = workbook.Worksheets.Add(frm1.SELECTEDWORKBOOKNAME);

                    for (int h = 0; h < _cols; h++)
                    {
                        worksheet.Cell(1, h + 1).Value = _GridHeader[h];

                        if (_excelMatchGridColorScheme)
                        {

                            worksheet.Cell(1, h + 1).Style.Font.Bold = _GridHeaderFont.Bold;
                            worksheet.Cell(1, h + 1).Style.Fill.BackgroundColor = XLColor.FromColor(_GridHeaderBackcolor);
                            worksheet.Cell(1, h + 1).Style.Font.FontColor = XLColor.FromColor(_GridHeaderForecolor);

                            worksheet.Cell(1, h + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(1, h + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(1, h + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            worksheet.Cell(1, h + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;


                        }
                    }

                    for (int r = 0; r < _rows; r++)
                    {

                        for (int c = 0; c < _cols; c++)
                        {
                            string Gval = _grid[r, c];

                            if (frm1.OMITNULLS && Gval.ToUpper().Trim() == "{NULL}")
                                Gval = "";

                            worksheet.Cell(r + 2, c + 1).Value = Gval;

                            if (_excelMatchGridColorScheme)
                            {

                                worksheet.Cell(r + 2, c + 1).Style.Font.Bold = get_CellFont(r, c).Bold;

                                SolidBrush bb = (SolidBrush)get_CellBackColor(r, c);

                                worksheet.Cell(r + 2, c + 1).Style.Fill.BackgroundColor = XLColor.FromColor(bb.Color);

                                worksheet.Cell(r + 2, c + 1).Style.Font.FontColor = XLColor.FromColor(get_CellForeColor(r, c).Color);

                                worksheet.Cell(r + 2, c + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(r + 2, c + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(r + 2, c + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                worksheet.Cell(r + 2, c + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                            }

                        }

                    }

                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filetogenerate);

                    if (_ShowExcelExportMessage)
                    {
                        frm.Hide();
                        frm = null;
                    }
                }


                //using ( DocumentFormat.OpenXml.Packaging.SpreadsheetDocument document = SpreadsheetDocument.Create(filetogenerate, SpreadsheetDocumentType.Workbook))
                //{
                //    WorkbookPart workbookPart = document.AddWorkbookPart();
                //    workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                //    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                //    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                //    worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                //    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
                //    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = 
                //        new DocumentFormat.OpenXml.Spreadsheet.Sheet() 
                //        { 
                //            Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = _excelWorkSheetName 
                //        };

                //    sheets.Append(sheet);

                //    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                //    List<String> columns = new List<string>();

                //    foreach (string s in _GridHeader)
                //    {
                //        columns.Add(s);

                //        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                //        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                //        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(s);

                //        if (_excelMatchGridColorScheme)
                //        {

                //        }

                //        headerRow.AppendChild(cell);
                //    }

                //    //List<String> columns = new List<string>();
                //    //foreach (System.Data.DataColumn column in table.Columns)
                //    //{
                //    //    columns.Add(column.ColumnName);

                //    //    Cell cell = new Cell();
                //    //    cell.DataType = CellValues.String;
                //    //    cell.CellValue = new CellValue(column.ColumnName);
                //    //    headerRow.AppendChild(cell);
                //    //}

                //    sheetData.AppendChild(headerRow);

                //    for(int r = 0;r<_rows;r++)
                //    {
                //        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                //        for (int c=0;c<_cols;c++)
                //        {
                //            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                //            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                //            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(_grid[r,c]);
                //            newRow.AppendChild(cell);
                //        }

                //        sheetData.AppendChild(newRow);
                //    }

                //    //foreach (DataRow dsrow in table.Rows)
                //    //{
                //    //    Row newRow = new Row();
                //    //    foreach (String col in columns)
                //    //    {
                //    //        Cell cell = new Cell();
                //    //        cell.DataType = CellValues.String;
                //    //        cell.CellValue = new CellValue(dsrow[col].ToString());
                //    //        newRow.AppendChild(cell);
                //    //    }

                //    //    sheetData.AppendChild(newRow);
                //    //}

                //    workbookPart.Workbook.Save();
                //}
            }

            //Excel.Application _excel;
            //Excel.Workbook _workbook = new Excel.Workbook();

            //try
            //{
            //    _excel = (Excel.Application)Interaction.CreateObject("Excel.Application");
            //    // _excel.Visible = True

            //    // _excel.ScreenUpdating = False

            //    // _workbook = _excel.Workbooks.Add()

            //    ExportToExcel(_excel, _workbook, _excelWorkSheetName);
            //}
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToExcel Error...");
                return;
            }
        }

        /// <summary>
        /// Will take the supplied instance of Microsoft excel <c>_excel</c> and place the contents of the grid on the
        /// worksheet named <c>wsname</c> in the supplied workbook <c>_Workbook</c>
        /// </summary>
        /// <param name="_excel"></param>
        /// <param name="_WorkBook"></param>
        /// <param name="wsname"></param>
        /// <remarks></remarks>
        public void ExportToExcel(Excel.Application _excel, Excel.Workbook _WorkBook, string wsname)
        {
            var frm = new frmExportingToExcelWorking();
            string lastsheetname;

            if (_ShowExcelExportMessage)
            {
                frm.Show();
                frm.Refresh();
            }

            Refresh();
            Application.DoEvents();

            int TotalRows = -1;
            Excel.Worksheet sh;
            try
            {
                SolidBrush _br;
                string rng;
                int idx = 0;
                string FirstColumn = "A";
                string LastColumn = ReturnExcelColumn(Cols - 1);
                int CurrentSHidx = 1;
                int ccidx = 0;

                int rows = Rows;
                _WorkBook = _excel.Workbooks.Add();

                if (rows > _excelMaxRowsPerSheet)
                {
                    idx = Conversions.ToInteger(rows / (double)_excelMaxRowsPerSheet);

                    // get around the Int round up problem of VB

                    if (idx * _excelMaxRowsPerSheet > rows)
                        idx -= 1;
                }

                // _WorkBook.Worksheets("Sheet2").Delete()
                // _WorkBook.Worksheets("Sheet3").Delete()

                while (idx != -1)
                {
                    if (idx > 0)
                        rows = _excelMaxRowsPerSheet;
                    else
                        rows = Rows - TotalRows;

                    if (CurrentSHidx == 1)
                    {
                        lastsheetname = wsname + " " + CurrentSHidx.ToString();
                        Excel.Worksheet sheet = (Excel.Worksheet)_WorkBook.Worksheets["Sheet1"];

                        sheet.Name = lastsheetname;

                        sh = (Excel.Worksheet)_WorkBook.ActiveSheet;
                    }
                    else
                    {
                        _WorkBook.Worksheets.Add();
                        sh = (Excel.Worksheet)_WorkBook.ActiveSheet;

                        // Dim cmd As String = "After:=Sheets(" + Chr(34) + lastsheetname + Chr(34) + ")"

                        // sh.move(cmd)
                        lastsheetname = wsname + " " + CurrentSHidx.ToString();
                        sh.Name = lastsheetname;
                    }

                    _excel.MaxChange = 0.001;

                    _excel.ActiveWorkbook.PrecisionAsDisplayed = false;
                    var arr = new object[rows + 1 + 1, Cols + 1];
                    int r, c, rmod;

                    rmod = 0;
                    r = 0;
                    c = 0;

                    while (r != rows + 1)
                    {
                        while (c != Cols)
                        {
                            if (r > 0)
                                arr[r, c] = get_item(TotalRows, c);
                            else
                                arr[r, c] = get_HeaderLabel(c);
                            c += 1;
                        }
                        r += 1;
                        c = 0;
                        TotalRows += 1;
                    }

                    rng = FirstColumn + "1:" + LastColumn + (rows + 1).ToString();

                    sh.Range[rng].NumberFormat = "General";

                    sh.Range[rng].Value = arr;

                    if (_excelMatchGridColorScheme)
                    {
                        // header always on row 1 of the sheet
                        c = 1;
                        rng = FirstColumn + "1:" + LastColumn + "1";
                        sh.Range[rng].Interior.Color = Information.RGB(_GridHeaderBackcolor.R, _GridHeaderBackcolor.G, _GridHeaderBackcolor.B);

                        r = (CurrentSHidx - 1) * _excelMaxRowsPerSheet;

                        int rmax = r + _excelMaxRowsPerSheet;

                        if (rmax > Rows)
                            rmax = Rows;

                        c = 0;

                        // here we will blast the first range of standard color in a single shot

                        rng = FirstColumn + "2:" + LastColumn + (rows + 1).ToString();
                        _br = (System.Drawing.SolidBrush)_gridBackColorList[0]; // element 0 is the default/first backcolor

                        // now lets blast this color into the background of the grid
                        sh.Range[rng].Interior.Color = Information.RGB(_br.Color.R, _br.Color.G, _br.Color.B);

                        ccidx = 1;

                        while (ccidx != _gridBackColorList.GetUpperBound(0))
                        {
                            if (_gridBackColorList[ccidx] == null)
                            {
                            }
                            else
                            {
                                r = (CurrentSHidx - 1) * _excelMaxRowsPerSheet;
                                rmod = (CurrentSHidx - 1) * _excelMaxRowsPerSheet;
                                while (r != rmax)
                                {
                                    c = 0;
                                    while (c != Cols)
                                    {
                                        if (_gridBackColor[r, c] == ccidx)
                                        {
                                            // we got a color thats different than the blasted backcolor


                                            rng = ReturnExcelColumn(c) + (r - rmod + 2).ToString();
                                            rng = rng + ":" + rng;
                                            _br = (System.Drawing.SolidBrush)_gridBackColorList[ccidx];
                                            //sh.Range[rng]
                                            sh.Range[rng].Interior.Color = Information.RGB(_br.Color.R, _br.Color.G, _br.Color.B);
                                        }
                                        c += 1;
                                    }
                                    r += 1;
                                }
                            }
                            ccidx += 1;
                        }
                    }
                    else
                    {
                        r = 2;
                        if (ExcelUseAlternateRowColor)
                        {
                            // 38 seconds to load 5000 lines of claims data same as previous loop
                            while (r < rows)
                            {
                                rng = FirstColumn + r.ToString() + ":" + LastColumn + r.ToString();
                                // there are 56 possible colors, Lonnie uses number 35 in the grid.  Maybe 57, I didn't try index# 0...
                                sh.Range[rng].Interior.ColorIndex = 35;
                                r += 2;
                            }
                        }
                    }

                    if (_excelIncludeColumnHeaders)
                    {

                        sh.Range["1:1"].Font.Bold = true;
                        sh.Range["1:1"].HorizontalAlignment = xlCenter;
                    }

                    if (_excelShowBorders | _excelOutlineCells)
                    {

                        // try to catch errors here just in case someone is putting massive amounts of text into
                        // cells and the range selection diddys are breaking in excel versions less then 2007

                        try
                        {
                            FirstColumn = "A";
                            LastColumn = ReturnExcelColumn(Cols - 1);
                            rng = FirstColumn + "1:" + LastColumn + (rows + 1).ToString();

                            Excel.Range tRange = sh.Range[rng];

                            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                            tRange.Borders.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;


                            //sh.Range[rng].Borders[xlEdgeRight].LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlEdgeRight).Weight = xlThin;
                            //sh.Range[rng].Borders(xlEdgeRight).ColorIndex = xlAutomatic;

                            //sh.Range[rng].Borders(xlEdgeLeft).LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlEdgeLeft).Weight = xlThin;
                            //sh.Range[rng].Borders(xlEdgeLeft).ColorIndex = xlAutomatic;

                            //sh.Range[rng].Borders(xlEdgeTop).LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlEdgeTop).Weight = xlThin;
                            //sh.Range[rng].Borders(xlEdgeTop).ColorIndex = xlAutomatic;

                            //sh.Range[rng].Borders(xlEdgeBottom).LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlEdgeBottom).Weight = xlThin;
                            //sh.Range[rng].Borders(xlEdgeBottom).ColorIndex = xlAutomatic;

                            //sh.Range[rng].Borders(xlInsideVertical).LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlInsideVertical).Weight = xlThin;
                            //sh.Range[rng].Borders(xlInsideVertical).ColorIndex = xlAutomatic;

                            //sh.Range[rng].Borders(xlInsideHorizontal).LineStyle = xlContinuous;
                            //sh.Range[rng].Borders(xlInsideHorizontal).Weight = xlThin;
                            //sh.Range[rng].Borders(xlInsideHorizontal).ColorIndex = xlAutomatic;
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    if (_excelAutoFitColumn)
                    {
                        Excel.Range tRange = sh.Range[rng];
                        tRange.EntireColumn.AutoFit();
                        //sh.Range(rng).EntireColumn.Autofit();
                    }
                    if (_excelAutoFitRow)
                    {
                        Excel.Range tRange = sh.Range[rng];
                        tRange.EntireRow.AutoFit();
                        //sh.Range(rng).EntireRow.Autofit();
                    }



                    // If Me._excelAutoFitColumn Then
                    // sh.Range(rng).EntireColumn.Autofit()
                    // End If

                    sh.PageSetup.Orientation = (Excel.XlPageOrientation)_excelPageOrientation;

                    _excel.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;

                    // save the spreadsheet
                    _excel.AlertBeforeOverwriting = false;
                    _excel.DisplayAlerts = false;

                    _excel.ScreenUpdating = true;

                    r = 0;
                    idx -= 1;
                    CurrentSHidx += 1;
                    TotalRows -= 1; // this subtracts 1 row incase there is another worksheet that is needed.  WIthout this the first row will be skipped
                }

                if (_ShowExcelExportMessage)
                {
                    frm.Hide();
                    frm = null;
                }

                _excel.Visible = true;
                _WorkBook = null;
                sh = null;
                _excel = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                if (_ShowExcelExportMessage)
                {
                    frm.Hide();
                    frm = null;
                }

                _WorkBook = null;
                sh = null;
                _excel = null;

                GC.Collect();
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRiDControl.ExportToExcel Error...");
            }
        }

        /// <summary>
        /// Will open the internal export filename dialog querying the user for the filename to export to
        /// Will then export the contents of the grid to a textfile employing the properties setup by
        /// displayed dialog, Filename, Include column headers as fieldname, the field terminator, and the line
        /// terminator...
        /// </summary>
        /// <remarks></remarks>
        public void ExportToText()
        {
            try
            {
                var frmExport = new frmExportToText();

                if ((int)frmExport.ShowDialog() == (int)DialogResult.OK)
                {
                    string sDelimiter = frmExport.Delimiter;
                    string sFilename = frmExport.Filename;
                    bool bIncludeFieldNames = frmExport.IncludeFieldNames;
                    bool bIncludeLineTerminator = frmExport.IncludeLineTerminator;

                    ExportToText(sFilename, sDelimiter, bIncludeFieldNames, bIncludeLineTerminator);
                }

                frmExport = null;
            }

            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
            }
        }

        /// <summary>
        /// Will export the contents of the grid to the supplied filename <c>sFilename</c> employing the internally set properties
        /// to control the field delimiters, line terminators, and inclusion of column headers as field names...
        /// </summary>
        /// <param name="sFilename"></param>
        /// <remarks></remarks>
        public void ExportToText(string sFilename)
        {
            try
            {
                ExportToText(sFilename, _delimiter, _includeFieldNames, _includeLineTerminator);
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
            }
        }

        /// <summary>
        /// Will export the contens of the grid to a file named <c>sFilename</c> with the field delimiters of <c>sDelimiter</c>
        /// The internal properties will control the inclusion of column headers as field names and the characters used to terminate
        /// the lines of output
        /// </summary>
        /// <param name="sFilename"></param>
        /// <param name="sDelimiter"></param>
        /// <remarks></remarks>
        public void ExportToText(string sFilename, string sDelimiter)
        {
            try
            {
                ExportToText(sFilename, sDelimiter, _includeFieldNames, _includeLineTerminator);
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
            }
        }

        /// <summary>
        /// Will export the contents of the grid to a textfile name <c>sFilename</c> using <c>sDelimiter</c> for
        /// field delimiters and using <c>bIncludeFieldNames</c> to control inclusion of the column headers as field names
        /// the interna;l property for the end of the lines termination will be employed.
        /// </summary>
        /// <param name="sFilename"></param>
        /// <param name="sDelimiter"></param>
        /// <param name="bIncludeFieldNames"></param>
        /// <remarks></remarks>
        public void ExportToText(string sFilename, string sDelimiter, bool bIncludeFieldNames)
        {
            try
            {
                ExportToText(sFilename, sDelimiter, bIncludeFieldNames, _includeLineTerminator);
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
            }
        }

        /// <summary>
        /// Will export the contents of the grid to <c>sFilename</c> employing <c>sDelimiter</c> for the field delimiters
        /// <c>bIncludeFieldNames</c> to control inclusion of the column headers as field names, and the <c>bIncludeLineTerminator</c>
        /// to control the CRLF at the end of the lines of output
        /// </summary>
        /// <param name="sFileName"></param>
        /// <param name="sDelimiter"></param>
        /// <param name="bIncludeFieldNames"></param>
        /// <param name="bIncludeLineTerminator"></param>
        /// <remarks></remarks>
        public void ExportToText(string sFileName, string sDelimiter, bool bIncludeFieldNames, bool bIncludeLineTerminator)
        {
            try
            {
                if (string.IsNullOrEmpty(sFileName.Trim()))
                {
                    Interaction.MsgBox("You need to specify the filename before I can continue!", (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
                    return;
                }

                var sw = System.IO.File.CreateText(sFileName);

                int iCols = Cols;
                int iCol;
                int iRows = Rows;
                int iRow;
                int iRowStart = 0;

                string sLine;

                if (bIncludeFieldNames)
                {
                    if (bIncludeLineTerminator)
                    {
                        sLine = "";
                        var loopTo = iCols - 1;
                        for (iCol = 0; iCol <= loopTo; iCol++)
                            sLine += get_HeaderLabel(iCol) + sDelimiter;
                        sw.WriteLine(sLine);
                    }
                    else
                    {
                        sLine = "";
                        var loopTo1 = iCols - 1;
                        for (iCol = 0; iCol <= loopTo1; iCol++)
                            sLine += get_HeaderLabel(iCol) + sDelimiter;
                        sw.Write(sLine);
                    }
                }

                if (bIncludeLineTerminator)
                {
                    var loopTo2 = iRows - 1;
                    for (iRow = iRowStart; iRow <= loopTo2; iRow++)
                    {
                        sLine = "";
                        var loopTo3 = iCols - 1;
                        for (iCol = 0; iCol <= loopTo3; iCol++)
                            sLine += get_item(iRow, iCol) + sDelimiter;
                        sw.WriteLine(sLine);
                    }
                }
                else
                {
                    var loopTo4 = iRows - 1;
                    for (iRow = iRowStart; iRow <= loopTo4; iRow++)
                    {
                        sLine = "";
                        var loopTo5 = iCols - 1;
                        for (iCol = 0; iCol <= loopTo5; iCol++)
                            sLine += get_item(iRow, iCol) + sDelimiter;
                        sw.Write(sLine);
                    }
                }

                sw.Close();
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToText Error...");
            }
        }

        /// <summary>
        /// Will return a list(of string) of the unique values contained in ColId of the current grid contents
        /// </summary>
        /// <param name="ColId"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<string> DistinctInColumn(int ColId)
        {
            var ret = new List<string>();

            int t = 0;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (!ret.Contains(get_item(t, ColId)))
                    ret.Add(get_item(t, ColId));
            }

            return ret;
        }

        /// <summary>
        /// will do a case insensitive rip through grid col colvalue searching for strvalue
        /// on finding it will return id value of the row where the search was successful
        /// -1 otherwise
        /// </summary>
        /// <param name="strValue"></param>
        /// <param name="colvalue"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public int FindInColumn(string strValue, int colvalue)
        {
            // will do a case insensative rip through grid col colvalue searching for strvalue
            // on finding it will return id value of the row where the search was successful
            // -1 otherwise

            int t;
            int ret = -1;

            if (colvalue <= _cols - 1 & colvalue > -1)
            {
                var loopTo = Rows - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    if ((Strings.UCase(Strings.Trim(strValue)) ?? "") == (Strings.UCase(Strings.Trim(_grid[t, colvalue])) ?? ""))
                    {
                        ret = t;
                        break;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// will do a case insensative or sensitive  (depending on the CaseSensitive parameter)
        /// rip through grid col colvalue searching for strvalue
        /// on finding it will return id value of the row where the search was successful
        /// -1 otherwise
        /// </summary>
        /// <param name="strValue"></param>
        /// <param name="colvalue"></param>
        /// <param name="CaseSensitive"></param>
        /// <returns>Will return the first row ID of the search or -1 if the search is unsuccessful</returns>
        /// <remarks></remarks>
        public int FindInColumn(string strValue, int colvalue, bool CaseSensitive)
        {
            // will do a eithyer a case insensative  or case sensative rip through grid col colvalue searching for strvalue
            // on finding it will return id value of the row where the search was successful
            // -1 otherwise

            int t;
            int ret = -1;

            if (colvalue <= _cols - 1 & colvalue > -1)
            {
                if (CaseSensitive)
                {
                    var loopTo = Rows - 1;
                    for (t = 0; t <= loopTo; t++)
                    {
                        if ((Strings.Trim(strValue) ?? "") == (Strings.Trim(_grid[t, colvalue]) ?? ""))
                        {
                            ret = t;
                            break;
                        }
                    }
                }
                else
                    ret = FindInColumn(strValue, colvalue);
            }

            return ret;
        }

        /// <summary>
        /// will free the memory associated with the internal captured image of the grids contents
        /// </summary>
        /// <remarks></remarks>
        public void FreeGridContentImage()
        {
            _image.Dispose();
        }

        /// <summary>
        /// Will return the column ID of the first column matching name. CaseSensative will toggle the search matching
        /// the case of the searched name.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="CaseSensitive"></param>
        /// <returns>Column ID of the first match, -1 if not found</returns>
        /// <remarks></remarks>
        public int GetColumnIDByName(string name, bool CaseSensitive)
        {
            // returns column ID of NAME string from current grid contents
            // -1 if not found

            int ret = -1;
            int t;
            if (CaseSensitive)
            {
                var loopTo = _cols - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    if ((Strings.Trim(name) ?? "") == (Strings.Trim(_GridHeader[t]) ?? ""))
                    {
                        // we have a match
                        ret = t;
                        break;
                    }
                }
            }
            else
                ret = GetColumnIDByName(name);

            return ret;
        }

        /// <summary>
        /// Will return the column ID of the first column matching name. The search is not case sensative.
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Column ID of the first match, -1 if not found</returns>
        /// <remarks></remarks>
        public int GetColumnIDByName(string name)
        {
            // returns column ID of NAME string from current grid contents
            // -1 if not found

            int ret = -1;
            int t;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if ((Strings.Trim(Strings.UCase(name)) ?? "") == (Strings.Trim(Strings.UCase(_GridHeader[t])) ?? ""))
                {
                    // we have a match
                    ret = t;
                    break;
                }
            }

            return ret;
        }

        /// <summary>
        /// Will return an ArrayList of distinct values contained in a given Column within the grids
        /// current contents. Column searched indicated by colid. Search will ignore case differences.
        /// </summary>
        /// <param name="colid"></param>
        /// <returns>ArrayList of distinct values sorted</returns>
        /// <remarks></remarks>
        public ArrayList GetDistinctColumnEntries(int colid)
        {
            var exl = new ArrayList();

            return GetDistinctColumnEntries(colid, exl, true);
        }

        /// <summary>
        /// Will return an ArrayList of distinct values contained in a given Column within the grids
        /// current contents. Column searched indicated by colid. The ignorecase parameter will
        /// allow or disallow the differences in case to be taken into account.
        /// </summary>
        /// <param name="colid"></param>
        /// <param name="ignorecase"></param>
        /// <returns>ArrayList of distinct values sorted</returns>
        /// <remarks></remarks>
        public ArrayList GetDistinctColumnEntries(int colid, bool ignorecase)
        {
            var exl = new ArrayList();

            return GetDistinctColumnEntries(colid, exl, ignorecase);
        }

        /// <summary>
        /// Will return an ArrayList of distinct values contained in a given Column within the grids
        /// current contents. Column searched indicated by colid. The parameter for the Exclusionlist will
        /// contain any values you wish to omit from the results. The ignorecase parameter will
        /// allow or disallow the differences in case to be taken into account.
        /// </summary>
        /// <param name="colid"></param>
        /// <param name="exclusionlist"></param>
        /// <param name="ignorecase"></param>
        /// <returns>ArrayList of distinct values sorted</returns>
        /// <remarks></remarks>
        public ArrayList GetDistinctColumnEntries(int colid, ArrayList exclusionlist, bool ignorecase)
        {
            var retval = new ArrayList();
            string s;

            int r;

            if (exclusionlist == null)
                exclusionlist.Add("KJKJKJJHKJHKJHKJHKJH");// some dummy value here
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                s = _grid[r, colid];
                if (ignorecase)
                {
                    if (!exclusionlist.Contains(Strings.UCase(s)))
                    {
                        if (retval == null)
                            retval.Add(Strings.UCase(s));
                        else if (!retval.Contains(Strings.UCase(s)))
                            retval.Add(Strings.UCase(s));
                    }
                }
                else if (!exclusionlist.Contains(s))
                {
                    if (retval == null)
                        retval.Add(s);
                    else if (!retval.Contains(s))
                        retval.Add(s);
                }
            }

            retval.Sort();

            return retval;
        }

        /// <summary>
        /// Returns the grids contents as a two dimensional array of string values
        /// </summary>
        /// <returns>2 dimensional array of strings representative of the ccurrent contents of the grid</returns>
        /// <remarks></remarks>
        public string[,] GetGridAsArray()
        {
            var result = new string[_rows + 1, _cols];
            int r, c;
            var loopTo = _cols - 1;
            for (c = 0; c <= loopTo; c++)
                result[0, c] = _GridHeader[c];
            var loopTo1 = _rows;
            for (r = 1; r <= loopTo1; r++)
            {
                var loopTo2 = _cols - 1;
                for (c = 0; c <= loopTo2; c++)
                    result[r, c] = _grid[r - 1, c];
            }

            return result;
        }

        /// <summary>
        /// Scans ColID for Values in ColVals and colors each row where ColID contains ColVal with corresponding ColorVal
        /// </summary>
        /// <param name="Colid"></param>
        /// <param name="Colvals"></param>
        /// <param name="ColorVals"></param>
        /// <remarks></remarks>
        public void SetRowBackgroundsBasedOnValue(int Colid, List<string> Colvals, List<Color> ColorVals)
        {
            int t;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                string v = get_item(t, Colid);

                int i = 0;
                var loopTo1 = Colvals.Count - 1;
                for (i = 0; i <= loopTo1; i++)
                {
                    if ((Colvals[i] ?? "") == (v ?? ""))
                    {
                        SetRowBackColor(t, ColorVals[i]);
                        i = Colvals.Count;
                    }
                }
            }
        }

        /// <summary>
        /// Returns the contents of the grid as a 24 bit bitmat System.Drawing.Bitmap
        /// Useful for when the contents of the grid need to be image mapped onto another
        /// surface like a 3D rotating cube (believe it or not) or perhaps a printer page context.
        /// </summary>
        /// <returns>24 Bit System.Drawing.Bitmap</returns>
        /// <remarks></remarks>
        public Image GetGridContentsAsImage()
        {

            // here we will get a picture of the attached canvas 
            int h;
            int w;

            h = AllRowHeights();
            w = AllColWidths();

            // If _GridTitleVisible Then
            // h = h + _GridTitleHeight
            // End If

            if (_GridHeaderVisible)
                h = h + _GridHeaderHeight;

            if (!(_image == null))
            {
                _image.Dispose();
                _image = null;    // clear and release the last image gathered 
            }

            _image = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format24bppRgb);

            var g1 = Graphics.FromImage(_image);

            OleRenderGrid(g1);

            return _image;
        }

        /// <summary>
        /// Will return a row in the grid as a string with '|' character delimitring the columns.
        /// </summary>
        /// <param name="rowid"></param>
        /// <returns>String representation of the row in the grid at rowid with '|' characters between the fields</returns>
        /// <remarks></remarks>
        public string GetRowAsString(int rowid)
        {
            int c = 0;
            string result = "";

            if (rowid <= _rows - 1)
            {
                var loopTo = _cols - 1;
                for (c = 0; c <= loopTo; c++)
                {
                    result = result + _grid[rowid, c];
                    if (c < _cols - 1)
                        result = result + "|";
                }
            }

            return result;
        }

        /// <summary>
        /// Will return the indicated column at col as an Arraylist of values
        /// </summary>
        /// <param name="col"></param>
        /// <returns>ArrayList indicating the contents of the column at col</returns>
        /// <remarks></remarks>
        public ArrayList GetColAsArrayList(int col)
        {
            // if col is illegal then will return an empty arraylist
            var ar = new ArrayList();
            int r;

            if (col >= 0 & col < _cols)
            {
                var loopTo = _rows - 1;
                for (r = 0; r <= loopTo; r++)
                    ar.Add(_grid[r, col]);
            }
            return ar;
        }

        /// <summary>
        /// Will return the indicated column at col as an Arraylist of values. The values will be cleaned
        /// of specific formatting prior to being returned. Dollar values will have the $() and ,'s removed
        /// numeric values will be converted represent the numeric string representation. This is useful to
        /// subsequent insertion into an excel spreadsheet for example.
        /// </summary>
        /// <param name="col"></param>
        /// <returns>ArrayList indicating the contents of the column at col</returns>
        /// <remarks></remarks>
        public ArrayList GetColAsCleanedArrayList(int col)
        {
            // if col is illegal then will return an empty arraylist
            var ar = new ArrayList();
            int r;
            string s;

            if (col >= 0 & col < _cols)
            {
                var loopTo = _rows - 1;
                for (r = 0; r <= loopTo; r++)
                {
                    if (!(_grid[r, col] == null))
                    {
                        if (!string.IsNullOrEmpty(_grid[r, col].Trim()))
                        {
                            s = _grid[r, col];
                            s = CleanMoneyString(s);
                            if (Information.IsNumeric(s))
                            {
                                try
                                {
                                    ar.Add(Convert.ToDouble(s));
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
            }
            return ar;
        }

        /// <summary>
        /// Will return the indicated row as an arraylist of values
        /// </summary>
        /// <param name="row"></param>
        /// <returns>Arraylist of values at row</returns>
        /// <remarks></remarks>
        public ArrayList GetRowAsArrayList(int row)
        {
            // if row is illegal then will return an empty arraylist
            var ar = new ArrayList();
            int c;

            if (row >= 0 & row < _rows)
            {
                var loopTo = _cols - 1;
                for (c = 0; c <= loopTo; c++)
                    ar.Add(_grid[row, c]);
            }
            return ar;
        }

        /// <summary>
        /// Will insert numrows of blank space into the grid at row atrow.
        /// </summary>
        /// <param name="atrow"></param>
        /// <param name="numrows"></param>
        /// <remarks></remarks>
        public void InsertRowsIntoGridAt(int atrow, int numrows)
        {
            _Painting = true;

            // do the math here

            var oldrowhidden = new bool[_rowhidden.GetUpperBound(0) + 1];
            var oldcolhidden = new bool[_colhidden.GetUpperBound(0) + 1];
            var oldcoleditable = new bool[_colEditable.GetUpperBound(0) + 1];
            var oldroweditable = new bool[_rowEditable.GetUpperBound(0) + 1];
            var oldcolwidths = new int[_colwidths.GetUpperBound(0) + 1];
            var oldrowheights = new int[_rowheights.GetUpperBound(0) + 1];
            var oldgridheader = new string[_GridHeader.GetUpperBound(0) + 1];
            var oldgrid = new string[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridbcolor = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridfcolor = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridfonts = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridcolpasswords = new string[_colPasswords.GetUpperBound(0) + 1];
            var oldgridcellalignment = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            int r, c;
            int x, y;

            x = oldgrid.GetUpperBound(0);
            y = oldgrid.GetUpperBound(1);
            var loopTo = x;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = y;
                for (c = 0; c <= loopTo1; c++)
                {
                    oldgrid[r, c] = _grid[r, c];
                    oldgridbcolor[r, c] = _gridBackColor[r, c];
                    oldgridfcolor[r, c] = _gridForeColor[r, c];
                    oldgridfonts[r, c] = _gridCellFonts[r, c];
                    oldgridcellalignment[r, c] = _gridCellAlignment[r, c];
                }
            }

            var loopTo2 = Math.Min(_GridHeader.GetUpperBound(0), _colwidths.GetUpperBound(0));
            for (c = 0; c <= loopTo2; c++)
            {
                oldgridheader[c] = _GridHeader[c];
                oldcolwidths[c] = _colwidths[c];
                oldgridcolpasswords[c] = _colPasswords[c];
                oldcolhidden[c] = _colhidden[c];
                oldcoleditable[c] = _colEditable[c];
            }

            var loopTo3 = _rowheights.GetUpperBound(0);
            for (r = 0; r <= loopTo3; r++)
                oldrowheights[r] = _rowheights[r];
            var loopTo4 = _rowhidden.GetUpperBound(0);
            for (r = 0; r <= loopTo4; r++)
                oldrowhidden[r] = _rowhidden[r];
            var loopTo5 = _rowEditable.GetUpperBound(0);
            for (r = 0; r <= loopTo5; r++)
                oldroweditable[r] = _rowEditable[r];

            // we have the state

            _rows += numrows;

            _rowhidden = new bool[_rows + 1];
            _colhidden = new bool[_cols + 1];
            _colEditable = new bool[_cols + 1];
            _rowEditable = new bool[_rows + 1];
            _rowheights = new int[_rows + 1];
            _colwidths = new int[_cols + 1];
            _GridHeader = new string[_cols + 1];
            _grid = new string[_rows + 1, _cols + 1];
            _gridBackColor = new int[_rows + 1, _cols + 1];
            _gridForeColor = new int[_rows + 1, _cols + 1];
            _gridCellFonts = new int[_rows + 1, _cols + 1];
            _gridCellAlignment = new int[_rows + 1, _cols + 1];
            _colPasswords = new string[_cols + 1];
            var loopTo6 = y;

            // columns aren't changing so we can just do the column only stuff here
            for (c = 0; c <= loopTo6; c++)
            {
                _colPasswords[c] = oldgridcolpasswords[c];
                _GridHeader[c] = oldgridheader[c];
                _colwidths[c] = oldcolwidths[c];
                _colhidden[c] = oldcolhidden[c];
                _colEditable[c] = oldcoleditable[c];
            }

            if (atrow == 0)
            {
                var loopTo7 = x;
                // we are just moving rows with an offset
                for (r = 0; r <= loopTo7; r++)
                {
                    var loopTo8 = y;
                    for (c = 0; c <= loopTo8; c++)
                    {
                        _grid[r + numrows, c] = oldgrid[r, c];
                        _gridBackColor[r + numrows, c] = oldgridbcolor[r, c];
                        _gridForeColor[r + numrows, c] = oldgridfcolor[r, c];
                        _gridCellFonts[r + numrows, c] = oldgridfonts[r, c];
                        _gridCellAlignment[r + numrows, c] = oldgridcellalignment[r, c];
                    }
                    _rowheights[r + numrows] = oldrowheights[r];
                    _rowhidden[r + numrows] = oldrowhidden[r];
                }

                var loopTo9 = numrows - 1;
                for (r = 0; r <= loopTo9; r++)
                {
                    var loopTo10 = y;
                    for (c = 0; c <= loopTo10; c++)
                    {
                        _grid[r, c] = "";
                        _gridBackColor[r, c] = GetGridBackColorListEntry(new SolidBrush(_DefaultBackColor));
                        _gridForeColor[r, c] = GetGridForeColorListEntry(new Pen(_DefaultForeColor));
                        _gridCellFonts[r, c] = GetGridCellFontListEntry(_DefaultCellFont);
                        _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(_DefaultStringFormat);
                    }
                    _rowheights[r] = _DefaultRowHeight;
                    _rowEditable[r] = true;
                    _rowhidden[r] = false;
                }
            }
            else
            {
                var loopTo11 = atrow - 1;
                for (r = 0; r <= loopTo11; r++)
                {
                    var loopTo12 = y;
                    for (c = 0; c <= loopTo12; c++)
                    {
                        _grid[r, c] = oldgrid[r, c];
                        _gridBackColor[r, c] = oldgridbcolor[r, c];
                        _gridForeColor[r, c] = oldgridfcolor[r, c];
                        _gridCellFonts[r, c] = oldgridfonts[r, c];
                        _gridCellAlignment[r, c] = oldgridcellalignment[r, c];
                    }
                    _rowheights[r] = oldrowheights[r];
                    _rowEditable[r] = oldroweditable[r];
                    _rowhidden[r] = oldrowhidden[r];
                }

                var loopTo13 = x;
                for (r = atrow; r <= loopTo13; r++)
                {
                    var loopTo14 = y;
                    for (c = 0; c <= loopTo14; c++)
                    {
                        _grid[r + numrows, c] = oldgrid[r, c];
                        _gridBackColor[r + numrows, c] = oldgridbcolor[r, c];
                        _gridForeColor[r + numrows, c] = oldgridfcolor[r, c];
                        _gridCellFonts[r + numrows, c] = oldgridfonts[r, c];
                        _gridCellAlignment[r + numrows, c] = oldgridcellalignment[r, c];
                    }
                    _rowheights[r + numrows] = oldrowheights[r];
                    _rowEditable[r + numrows] = true;
                    _rowhidden[r + numrows] = oldrowhidden[r];
                }

                var loopTo15 = numrows - 1;
                for (r = 0; r <= loopTo15; r++)
                {
                    var loopTo16 = y;
                    for (c = 0; c <= loopTo16; c++)
                    {
                        _grid[r + atrow, c] = "";
                        _gridBackColor[r + atrow, c] = GetGridBackColorListEntry(new SolidBrush(_DefaultBackColor));
                        _gridForeColor[r + atrow, c] = GetGridForeColorListEntry(new Pen(_DefaultForeColor));
                        _gridCellFonts[r + atrow, c] = GetGridCellFontListEntry(_DefaultCellFont);
                        _gridCellAlignment[r + atrow, c] = GetGridCellAlignmentListEntry(_DefaultStringFormat);
                    }
                    _rowheights[r + atrow] = _DefaultRowHeight;
                    _rowEditable[r + atrow] = true;
                    _rowhidden[r + atrow] = false;
                }
            }

            var loopTo17 = _cols - 1;
            for (c = 0; c <= loopTo17; c++)
            {
                if (_colwidths[c] == 0 & !_colhidden[c])
                    _colwidths[c] = _DefaultColWidth;
            }

            var loopTo18 = _rows - 1;
            for (r = 0; r <= loopTo18; r++)
            {
                if (_rowheights[r] == 0 & !_rowhidden[r])
                    _rowheights[r] = _DefaultRowHeight;
            }

            _Painting = false;
            Invalidate();
        }

        /// <summary>
        /// Will close and release all open tearaway windows currentl being maintained by the grid
        /// </summary>
        /// <remarks></remarks>
        public void KillAllTearAwayColumnWindows()
        {
            if (TearAways.Count == 0)
                return;

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                ta.Winform.KillMe(ta.ColID);
            }
        }

        /// <summary>
        /// Will kill any tearaway windows being maintained by the grid for the indicated column it colid
        /// </summary>
        /// <param name="colid"></param>
        /// <remarks></remarks>
        public void KillTearAwayColumnWindow(int colid)
        {
            if (colid == -1 | TearAways.Count == 0)
                return;

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                if (ta.ColID == colid)
                    // ta.KillTearAway()
                    TearAways.RemoveAt(t);
            }
        }

        /// <summary>
        /// Will fire the GridHoverleave event if the grid is not maintaining any tearaway windows at the moment
        /// </summary>
        /// <remarks></remarks>
        public void LowerGridHoverEvents()
        {
            if (!_TearAwayWork)
                GridHoverleave?.Invoke(this);
        }

        // New stuff as of May 31 2005

        /// <summary>
        /// Will populate the grid with the results of a call to the System.Management classes with a properly
        /// formatted wql query (Windows Query Language) call.
        /// 
        /// Example:
        /// <c>Select * from Win32_Printer</c>
        /// 
        /// </summary>
        /// <param name="wql"></param>
        /// <remarks></remarks>
        public void PopulateFromWQL(string wql)
        {
            StartedDatabasePopulateOperation?.Invoke(this);

            try
            {
                System.Management.ManagementObjectCollection moReturn;

                System.Management.ManagementObjectSearcher moSearch;

                //System.Management.ManagementObject mo;

                System.Management.PropertyData prop;

                bool HeaderDone = false;

                // moSearch = New Management.ManagementObjectSearcher("Select * from Win32_Printer")

                moSearch = new System.Management.ManagementObjectSearcher(wql);

                moReturn = moSearch.Get();

                if (_ShowProgressBar)
                {
                    pBar.Maximum = moReturn.Count;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }

                int x = 0;

                foreach (var mo in moReturn)
                {
                    int y = 0;

                    if (!HeaderDone)
                    {
                        InitializeTheGrid(moReturn.Count, mo.Properties.Count);

                        foreach (var prop1 in mo.Properties)
                        {
                            _GridHeader[y] = prop1.Name;
                            y += 1;
                        }

                        HeaderDone = true;
                    }

                    y = 0;

                    foreach (var prop1 in mo.Properties)
                    {
                        _grid[x, y] = Convert.ToString(prop1.Value);
                        y += 1;
                    }

                    if (_ShowProgressBar)
                    {
                        pBar.Increment(1);
                        pBar.Refresh();
                    }

                    x += 1;
                }

                AllCellsUseThisFont(_DefaultCellFont);
                AllCellsUseThisForeColor(_DefaultForeColor);

                AutoSizeCellsToContents = true;

                Refresh();

                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        #region PopulateGridFromArray Calls

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, employing the supplied parameters
        /// to control
        /// <list type="Bullet">
        /// <item><c>gridfont</c> will employ this font for the cell contents</item>
        /// <item><c>col</c> will be used as the color for the displays cell items</item>
        /// <item><c>FirstRowHeader</c> if true will treat the first row in the array as the names for each column header</item>
        /// <item><c>AutoHeader</c> if true will automatically name each column COLUMN - {ordinal} as it populates the grid</item>
        /// <item><c>hdr</c> an array of strings that will be used as the column labels if the other column options are False</item>
        /// </list>
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="FirstRowHeader"></param>
        /// <param name="AutoHeader"></param>
        /// <param name="hdr"></param>
        /// <remarks></remarks>
        private void PopulateGridFromArray(string[,] arr, Font gridfont, Color col, bool FirstRowHeader, bool AutoHeader, string[] hdr)
        {
            int x, y;
            int r, c;

            r = arr.GetUpperBound(0) + 1;
            c = arr.GetUpperBound(1) + 1;

            if (FirstRowHeader)
            {
                InitializeTheGrid(r - 1, c);
                var loopTo = c - 1;
                for (y = 0; y <= loopTo; y++)
                    _GridHeader[y] = arr[0, y];
                var loopTo1 = r - 1;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = c - 1;
                    for (y = 0; y <= loopTo2; y++)
                        _grid[x - 1, y] = arr[x, y];
                }
            }
            else
            {
                InitializeTheGrid(r, c);

                if (AutoHeader)
                {
                    var loopTo3 = c - 1;
                    for (y = 0; y <= loopTo3; y++)
                        _GridHeader[y] = "Column - " + y.ToString();
                }
                else
                    _GridHeader = hdr;
                var loopTo4 = r - 1;
                for (x = 0; x <= loopTo4; x++)
                {
                    var loopTo5 = c - 1;
                    for (y = 0; y <= loopTo5; y++)
                        _grid[x, y] = arr[x, y];
                }
            }

            AllCellsUseThisFont(gridfont);
            AllCellsUseThisForeColor(col);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, employing the supplied parameters
        /// to control
        /// <list type="Bullet">
        /// <item><c>gridfont</c> will employ this font for the cell contents</item>
        /// <item><c>col</c> will be used as the color for the displays cell items</item>
        /// <item><c>FirstRowHeader</c> if true will treat the first row in the array as the names for each column header
        /// if its false the columns will be automatically named</item>
        /// </list>
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="FirstRowHeader"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(string[,] arr, Font gridfont, Color col, bool FirstRowHeader)
        {
            PopulateGridFromArray(arr, gridfont, col, FirstRowHeader, true, _GridHeader);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
        /// to named each column
        /// </summary>
        /// <param name="arr"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(string[,] arr)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
        /// to name each column each cell will be displayed the the supplied <c>cellfont</c>
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="CellFont"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(string[,] arr, Font CellFont)
        {
            PopulateGridFromArray(arr, CellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
        /// to name each column each cell will be displayed the the supplied <c>Forecolor</c>
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Forecolor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(string[,] arr, Color Forecolor)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
        /// to name each column each cell will be displayed the the supplied <c>cellfont</c> and <c>Forecolor</c>
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Cellfont"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(string[,] arr, Font Cellfont, Color ForeColor)
        {
            PopulateGridFromArray(arr, Cellfont, ForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of stings <c>arr</c>, using the first row of values
        /// to name each column each cell will be displayed the the supplied <c>gridfont</c> and <c>col</c> color and
        /// if <c>FirstRowHeader</c> is true will use the first row to label each column, if not, then the first row will be auto
        /// labled with Column - {ordinal}
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="FirstRowHeader"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(int[,] arr, Font gridfont, Color col, bool FirstRowHeader)
        {
            int x, y;
            int r, c;

            r = arr.GetUpperBound(0) + 1;
            c = arr.GetUpperBound(1) + 1;

            if (FirstRowHeader)
            {
                InitializeTheGrid(r - 1, c);
                var loopTo = c - 1;
                for (y = 0; y <= loopTo; y++)
                    _GridHeader[y] = arr[0, y].ToString();
                var loopTo1 = r - 1;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = c - 1;
                    for (y = 0; y <= loopTo2; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }
            else
            {
                InitializeTheGrid(r, c);
                var loopTo3 = c - 1;
                for (y = 0; y <= loopTo3; y++)
                    _GridHeader[y] = "Column - " + y.ToString();
                var loopTo4 = r - 1;
                for (x = 0; x <= loopTo4; x++)
                {
                    var loopTo5 = c - 1;
                    for (y = 0; y <= loopTo5; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }

            AllCellsUseThisFont(gridfont);
            AllCellsUseThisForeColor(col);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
        /// to name each column
        /// </summary>
        /// <param name="arr"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(int[,] arr)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Cellfont</c> will be used as the font for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="CellFont"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(int[,] arr, Font CellFont)
        {
            PopulateGridFromArray(arr, CellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Forecolor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(int[,] arr, Color Forecolor)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of integers <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
        /// for each new cells font
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Cellfont"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(int[,] arr, Font Cellfont, Color ForeColor)
        {
            PopulateGridFromArray(arr, Cellfont, ForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>col</c> will be used as the foreground color for each new cell and <c>gridfont</c> will be used if
        /// <c>FirstRowHeader</c> is true the first row of data in the array will be used to name each column otherwise the columns will be
        /// named Column - {ordinal}
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="FirstRowHeader"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(long[,] arr, Font gridfont, Color col, bool FirstRowHeader)
        {
            int x, y;
            int r, c;

            r = arr.GetUpperBound(0) + 1;
            c = arr.GetUpperBound(1) + 1;

            if (FirstRowHeader)
            {
                InitializeTheGrid(r - 1, c);
                var loopTo = c - 1;
                for (y = 0; y <= loopTo; y++)
                    _GridHeader[y] = arr[0, y].ToString();
                var loopTo1 = r - 1;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = c - 1;
                    for (y = 0; y <= loopTo2; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }
            else
            {
                InitializeTheGrid(r, c);
                var loopTo3 = c - 1;
                for (y = 0; y <= loopTo3; y++)
                    _GridHeader[y] = "Column - " + y.ToString();
                var loopTo4 = r - 1;
                for (x = 0; x <= loopTo4; x++)
                {
                    var loopTo5 = c - 1;
                    for (y = 0; y <= loopTo5; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }

            AllCellsUseThisFont(gridfont);
            AllCellsUseThisForeColor(col);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
        /// to name each column
        /// </summary>
        /// <param name="arr"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(long[,] arr)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Cellfont</c> will be used as the font for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="CellFont"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(long[,] arr, Font CellFont)
        {
            PopulateGridFromArray(arr, CellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Forecolor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(long[,] arr, Color Forecolor)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of longs <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
        /// for each new cells font
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Cellfont"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(long[,] arr, Font Cellfont, Color ForeColor)
        {
            PopulateGridFromArray(arr, Cellfont, ForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of doubles <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>col</c> will be used as the foreground color for each new cell and <c>gridfont</c> will be used if
        /// <c>FirstRowHeader</c> is true the first row of data in the array will be used to name each column otherwise the columns will be
        /// named Column - {ordinal}
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="FirstRowHeader"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(double[,] arr, Font gridfont, Color col, bool FirstRowHeader)
        {
            int x, y;
            int r, c;

            r = arr.GetUpperBound(0) + 1;
            c = arr.GetUpperBound(1) + 1;

            if (FirstRowHeader)
            {
                InitializeTheGrid(r - 1, c);
                var loopTo = c - 1;
                for (y = 0; y <= loopTo; y++)
                    _GridHeader[y] = arr[0, y].ToString();
                var loopTo1 = r - 1;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = c - 1;
                    for (y = 0; y <= loopTo2; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }
            else
            {
                InitializeTheGrid(r, c);
                var loopTo3 = c - 1;
                for (y = 0; y <= loopTo3; y++)
                    _GridHeader[y] = "Column - " + y.ToString();
                var loopTo4 = r - 1;
                for (x = 0; x <= loopTo4; x++)
                {
                    var loopTo5 = c - 1;
                    for (y = 0; y <= loopTo5; y++)
                        _grid[x, y] = arr[x, y].ToString();
                }
            }

            AllCellsUseThisFont(gridfont);
            AllCellsUseThisForeColor(col);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
        /// to name each column
        /// </summary>
        /// <param name="arr"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(double[,] arr)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Cellfont</c> will be used as the font for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="CellFont"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(double[,] arr, Font CellFont)
        {
            PopulateGridFromArray(arr, CellFont, _DefaultForeColor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Forecolor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(double[,] arr, Color Forecolor)
        {
            PopulateGridFromArray(arr, _DefaultCellFont, Forecolor, true);
        }

        /// <summary>
        /// Will populate the grids contents from an 2 dimensional array of Doubles <c>arr</c> converted to strings, using the first row of values
        /// to name each column. <c>Forecolor</c> will be used as the foreground color for each new cell and <c>Cellfont</c> will be used
        /// for each new cells font
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="Cellfont"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void PopulateGridFromArray(double[,] arr, Font Cellfont, Color ForeColor)
        {
            PopulateGridFromArray(arr, Cellfont, ForeColor, true);
        }

        #endregion

        #region PopulateGridWithDataAt Calls

        /// <summary>
        /// Will allow a database populate of a grid within an already populated grid of data.
        /// The effect will be to insert data from a carefully crafted query into a rectangular region of an
        /// existing grid of data.
        /// <list type="Bullet">
        /// <item><c>ConnectionString</c> the database connection to be employed</item>
        /// <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
        /// <item><c>AtRow</c> the integer offset row to start populating the data at</item>
        /// <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
        /// <item><c>newheadercolor</c> the color to use for the header that will be created from the queried data</item>
        /// <item><c>ColOffset</c> the column offset from the edge to start populating</item>
        /// </list>
        /// </summary>
        /// <param name="ConnectionString">A legal connection string representing the database. </param>
        /// <param name="Sql">The actual SQL code to execute against the database represented in <c>ConnectionString</c>.</param>
        /// <param name="Atrow">The Row from wich you want the insertion to occur.</param>
        /// <param name="newbackcolor">The <c>System.Drawing.Color</c> to set the newly inserted rows to use as a background color. </param>
        /// <param name="newheadercolor">Newly inserted rows will have a new header applied at the first row inserted, this will be used as the background color for this new header</param>
        /// <param name="ColOffSet">New rows inserted will be placed starting the this column offset. Allowing insertion of data into rectangular regions of and existing grid</param>
        /// <remarks></remarks>
        public void PopulateGridWithDataAt(string ConnectionString, string Sql, int Atrow, Color newbackcolor, Color newheadercolor, int ColOffSet)
        {
            var cn = new System.Data.SqlClient.SqlConnection(ConnectionString);
            System.Data.SqlClient.SqlCommand dbc;
            System.Data.SqlClient.SqlCommand dbc2;
            System.Data.SqlClient.SqlDataReader dbr;
            System.Data.SqlClient.SqlDataReader dbr2;

            string sql2;
            long t;
            int x, y, yy, xx;

            StartedDatabasePopulateOperation?.Invoke(this);

            // _LastConnectionString = ConnectionString
            // _LastSQLString = Sql

            try
            {
                cn.Open();

                sql2 = Sql;
                dbc2 = new System.Data.SqlClient.SqlCommand(sql2, cn);
                dbc2.CommandTimeout = _dataBaseTimeOut;
                dbr2 = dbc2.ExecuteReader();
                y = 0;
                yy = 0;

                while (dbr2.Read())
                {
                    y = y + 1;
                    if (y > MaxRowsSelected & MaxRowsSelected > 0)
                    {
                        y = MaxRowsSelected;
                        yy = -1;
                        break;
                    }
                }

                dbr2.Close();
                dbc2.Dispose();
                cn.Close();

                InsertRowsIntoGridAt(Atrow, y + 1);

                cn.Open();
                dbc = new System.Data.SqlClient.SqlCommand(Sql, cn);
                dbc.CommandTimeout = _dataBaseTimeOut;

                dbr = dbc.ExecuteReader();

                // Me.Cols = dbr.FieldCount

                // InitializeTheGrid(y, dbr.FieldCount)

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }


                // AllCellsUseThisFont(Gridfont)
                // AllCellsUseThisForeColor(col)

                if (dbr.FieldCount + ColOffSet < Cols)
                    xx = dbr.FieldCount;
                else
                    xx = Cols - ColOffSet;
                var loopTo = xx - 1;
                for (x = 0; x <= loopTo; x++)
                {
                    _grid[Atrow, x + ColOffSet] = dbr.GetName(x);
                    set_CellBackColor(Atrow, x + ColOffSet, new SolidBrush(newheadercolor));
                }

                // For x = 0 To Me.Cols - 1
                // _GridHeader(x) = dbr.GetName(x)
                // Next

                t = Atrow + 1;
                while (dbr.Read())
                {
                    var loopTo1 = xx - 1;
                    // Me.Rows = t + 1
                    for (x = 0; x <= loopTo1; x++)
                    {
                        if (Information.IsDBNull(dbr[x]))
                        {
                            if (_omitNulls)
                                _grid[Conversions.ToInteger(t), x + ColOffSet] = "";
                            else
                                _grid[Conversions.ToInteger(t), x + ColOffSet] = "{NULL}";
                        }
                        else
                        // here we need to do some work on items of certain types
                        if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                            _grid[Conversions.ToInteger(t), x + ColOffSet] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                        else
                            _grid[Conversions.ToInteger(t), x + ColOffSet] = Conversions.ToString(dbr[x]);
                        set_CellBackColor(Conversions.ToInteger(t), x + ColOffSet, new SolidBrush(newbackcolor));
                    }
                    t = t + 1;

                    if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                        break;

                    if (_ShowProgressBar)
                    {
                        pBar.Increment(1);
                        pBar.Refresh();
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                pBar.Visible = false;
                gb1.Visible = false;

                AutoSizeCellsToContents = true;

                dbr.Close();

                dbc.Dispose();

                cn.Close();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// Will allow a database populate of a grid within an already populated grid of data.
        /// The effect will be to insert data from a carefully crafted query into a rectangular region of an
        /// existing grid of data.
        /// <list type="Bullet">
        /// <item><c>ConnectionString</c> the database connection to be employed</item>
        /// <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
        /// <item><c>AtRow</c> the integer offset row to start populating the data at</item>
        /// <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
        /// <item><c>newheadercolor</c> the color to use for the header that will be created from the queried data</item>
        /// </list>
        /// </summary>
        /// <param name="ConnectionString">A legal connection string representing the database. </param>
        /// <param name="Sql">The actual SQL code to execute against the database represented in <c>ConnectionString</c>.</param>
        /// <param name="Atrow">The Row from wich you want the insertion to occur.</param>
        /// <param name="newbackcolor">The <c>System.Drawing.Color</c> to set the newly inserted rows to use as a background color. </param>
        /// <param name="newheadercolor">Newly inserted rows will have a new header applied at the first row inserted, this will be used as the background color for this new header</param>
        /// <remarks></remarks>
        public void PopulateGridWithDataAt(string ConnectionString, string Sql, int Atrow, Color newbackcolor, Color newheadercolor)
        {
            var cn = new System.Data.SqlClient.SqlConnection(ConnectionString);
            System.Data.SqlClient.SqlCommand dbc;
            System.Data.SqlClient.SqlCommand dbc2;
            System.Data.SqlClient.SqlDataReader dbr;
            System.Data.SqlClient.SqlDataReader dbr2;

            string sql2;
            long t;
            int x, y, yy, xx;

            StartedDatabasePopulateOperation?.Invoke(this);

            // _LastConnectionString = ConnectionString
            // _LastSQLString = Sql

            try
            {
                cn.Open();

                sql2 = Sql;
                dbc2 = new System.Data.SqlClient.SqlCommand(sql2, cn);
                dbc2.CommandTimeout = _dataBaseTimeOut;
                dbr2 = dbc2.ExecuteReader();
                y = 0;
                yy = 0;

                while (dbr2.Read())
                {
                    y = y + 1;
                    if (y > MaxRowsSelected & MaxRowsSelected > 0)
                    {
                        y = MaxRowsSelected;
                        yy = -1;
                        break;
                    }
                }

                dbr2.Close();
                dbc2.Dispose();
                cn.Close();

                InsertRowsIntoGridAt(Atrow, y + 1);

                cn.Open();
                dbc = new System.Data.SqlClient.SqlCommand(Sql, cn);
                dbc.CommandTimeout = _dataBaseTimeOut;

                dbr = dbc.ExecuteReader();

                // Me.Cols = dbr.FieldCount

                // InitializeTheGrid(y, dbr.FieldCount)

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }


                // AllCellsUseThisFont(Gridfont)
                // AllCellsUseThisForeColor(col)

                if (dbr.FieldCount < Cols)
                    xx = dbr.FieldCount;
                else
                    xx = Cols;
                var loopTo = xx - 1;
                for (x = 0; x <= loopTo; x++)
                {
                    _grid[Atrow, x] = dbr.GetName(x);
                    set_CellBackColor(Atrow, x, new SolidBrush(newheadercolor));
                }

                // For x = 0 To Me.Cols - 1
                // _GridHeader(x) = dbr.GetName(x)
                // Next

                t = Atrow + 1;
                while (dbr.Read())
                {
                    var loopTo1 = xx - 1;
                    // Me.Rows = t + 1
                    for (x = 0; x <= loopTo1; x++)
                    {
                        if (Information.IsDBNull(dbr[x]))
                        {
                            if (_omitNulls)
                                _grid[Conversions.ToInteger(t), x] = "";
                            else
                                _grid[Conversions.ToInteger(t), x] = "{NULL}";
                        }
                        else
                        // here we need to do some work on items of certain types
                        if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                            _grid[Conversions.ToInteger(t), x] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                        else
                            _grid[Conversions.ToInteger(t), x] = Conversions.ToString(dbr[x]);
                        set_CellBackColor(Conversions.ToInteger(t), x, new SolidBrush(newbackcolor));
                    }
                    t = t + 1;

                    if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                        break;

                    if (_ShowProgressBar)
                    {
                        pBar.Increment(1);
                        pBar.Refresh();
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                pBar.Visible = false;
                gb1.Visible = false;

                AutoSizeCellsToContents = true;

                dbr.Close();

                dbc.Dispose();

                cn.Close();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// Will allow a database populate of a grid within an already populated grid of data.
        /// The effect will be to insert data from a carefully crafted query into a rectangular region of an
        /// existing grid of data. this call will omit the header for the inserted result...
        /// <list type="Bullet">
        /// <item><c>ConnectionString</c> the database connection to be employed</item>
        /// <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
        /// <item><c>AtRow</c> the integer offset row to start populating the data at</item>
        /// <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
        /// </list>
        /// </summary>
        /// <param name="ConnectionString">A legal connection string representing the database. </param>
        /// <param name="Sql">The actual SQL code to execute against the database represented in <c>ConnectionString</c>.</param>
        /// <param name="Atrow">The Row from wich you want the insertion to occur.</param>
        /// <param name="newbackcolor">The <c>System.Drawing.Color</c> to set the newly inserted rows to use as a background color. </param>
        /// <remarks></remarks>
        public void PopulateGridWithDataAt(string ConnectionString, string Sql, int Atrow, Color newbackcolor)
        {
            PopulateGridWithDataAt(ConnectionString, Sql, Atrow, newbackcolor, true);
        }

        /// <summary>
        /// Will allow a database populate of a grid within an already populated grid of data.
        /// The effect will be to insert data from a carefully crafted query into a rectangular region of an
        /// existing grid of data. this call will omit the header for the inserted result...
        /// <list type="Bullet">
        /// <item><c>ConnectionString</c> the database connection to be employed</item>
        /// <item><c>Sql</c> the sql code to be used to retrieve the data to be inserted</item>
        /// <item><c>AtRow</c> the integer offset row to start populating the data at</item>
        /// <item><c>newbackcolor</c> the color to be used to setup the background of the cells for the new data</item>
        /// <item><c>allowDups></c> Will not insert any rows that already exist in the grid if set to false</item>
        /// </list>
        /// </summary>
        /// <param name="ConnectionString">A legal connection string representing the database. </param>
        /// <param name="Sql">The actual SQL code to execute against the database represented in <c>ConnectionString</c>.</param>
        /// <param name="Atrow">The Row from wich you want the insertion to occur.</param>
        /// <param name="newbackcolor">The <c>System.Drawing.Color</c> to set the newly inserted rows to use as a background color. </param>
        /// <param name="allowDups"> Will not insert any rows that already exist in the grid if set to false</param>
        /// <remarks></remarks>
        public void PopulateGridWithDataAt(string ConnectionString, string Sql, int Atrow, Color newbackcolor, bool allowDups)
        {
            var cn = new System.Data.SqlClient.SqlConnection(ConnectionString);
            System.Data.SqlClient.SqlCommand dbc;
            System.Data.SqlClient.SqlCommand dbc2;
            System.Data.SqlClient.SqlDataReader dbr;
            System.Data.SqlClient.SqlDataReader dbr2;

            string sql2;
            long t, tt;
            int x, y, yy, xx;
            bool fnd = false;
            string hst = "";
            string hst2 = "";

            StartedDatabasePopulateOperation?.Invoke(this);

            // _LastConnectionString = ConnectionString
            // _LastSQLString = Sql

            try
            {

                // ' here we want to get whats in the grid already as a set of hashes for dup checking

                var ga = new List<string>();

                var sb = new StringBuilder();
                var loopTo = (long)(Rows - 1);
                for (t = 0; t <= loopTo; t++)
                {
                    sb = new StringBuilder();
                    var loopTo1 = Cols - 1;
                    for (x = 0; x <= loopTo1; x++)
                        sb.Append(x.ToString() + _grid[Conversions.ToInteger(t), x].ToUpper() + "|");
                    ga.Add(sb.ToString());
                }

                cn.Open();

                sql2 = Sql;
                dbc2 = new System.Data.SqlClient.SqlCommand(sql2, cn);
                dbc2.CommandTimeout = _dataBaseTimeOut;
                dbr2 = dbc2.ExecuteReader();
                y = 0;
                yy = 0;

                if (dbr2.FieldCount < Cols)
                    xx = dbr2.FieldCount;
                else
                    xx = Cols;

                var _ggrid = new string[xx + 1];

                while (dbr2.Read())
                {
                    hst = "";

                    if (!allowDups)
                    {
                        var loopTo2 = xx - 1;
                        for (x = 0; x <= loopTo2; x++)
                        {
                            if (Information.IsDBNull(dbr2[x]))
                            {
                                if (_omitNulls)
                                    _ggrid[x] = "";
                                else
                                    _ggrid[x] = "{NULL}";
                            }
                            else
                            // here we need to do some work on items of certain types
                            if ((dbr2[x].ToString() ?? "") == "System.Byte[]")
                                _ggrid[x] = ReturnByteArrayAsHexString((byte[])dbr2[x]);
                            else
                                _ggrid[x] = Conversions.ToString(dbr2[x]);
                        }

                        var loopTo3 = xx - 1;
                        for (x = 0; x <= loopTo3; x++)
                            hst += x.ToString() + _ggrid[x].ToUpper() + "|";

                        if (ga.Contains(hst))
                            fnd = true;
                        else
                            fnd = false;


                        if (!fnd)
                        {
                            y = y + 1;
                            if (y > MaxRowsSelected & MaxRowsSelected > 0)
                            {
                                y = MaxRowsSelected;
                                yy = -1;
                                break;
                            }
                        }
                        else
                            Console.WriteLine("Dupe");
                    }
                    else
                    {
                        y = y + 1;
                        if (y > MaxRowsSelected & MaxRowsSelected > 0)
                        {
                            y = MaxRowsSelected;
                            yy = -1;
                            break;
                        }
                    }
                }

                dbr2.Close();
                dbc2.Dispose();
                cn.Close();

                InsertRowsIntoGridAt(Atrow, y);

                cn.Open();
                dbc = new System.Data.SqlClient.SqlCommand(Sql, cn);
                dbc.CommandTimeout = _dataBaseTimeOut;

                dbr = dbc.ExecuteReader();

                // Me.Cols = dbr.FieldCount

                // InitializeTheGrid(y, dbr.FieldCount)

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }


                // AllCellsUseThisFont(Gridfont)
                // AllCellsUseThisForeColor(col)

                if (dbr.FieldCount < Cols)
                    xx = dbr.FieldCount;
                else
                    xx = Cols;

                t = Atrow;
                if (allowDups)
                {
                    while (dbr.Read())
                    {
                        var loopTo4 = xx - 1;
                        // Me.Rows = t + 1
                        for (x = 0; x <= loopTo4; x++)
                        {
                            if (Information.IsDBNull(dbr[x]))
                            {
                                if (_omitNulls)
                                    _grid[Conversions.ToInteger(t), x] = "";
                                else
                                    _grid[Conversions.ToInteger(t), x] = "{NULL}";
                            }
                            else
                            // here we need to do some work on items of certain types
                            if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                                _grid[Conversions.ToInteger(t), x] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                            else
                                _grid[Conversions.ToInteger(t), x] = Conversions.ToString(dbr[x]);
                            set_CellBackColor(Conversions.ToInteger(t), x, new SolidBrush(newbackcolor));
                        }
                        t = t + 1;

                        if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                            break;

                        if (_ShowProgressBar)
                        {
                            pBar.Increment(1);
                            pBar.Refresh();
                        }
                    }
                }
                else
                {
                    // ' here we are gonna not import any duplicate rows

                    tt = 0;

                    // 'Dim _ggrid(xx) As String

                    while (dbr.Read())
                    {
                        var loopTo5 = xx - 1;
                        // Me.Rows = t + 1
                        for (x = 0; x <= loopTo5; x++)
                        {
                            if (Information.IsDBNull(dbr[x]))
                            {
                                if (_omitNulls)
                                    _ggrid[x] = "";
                                else
                                    _ggrid[x] = "{NULL}";
                            }
                            else
                            // here we need to do some work on items of certain types
                            if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                                _ggrid[x] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                            else
                                _ggrid[x] = Conversions.ToString(dbr[x]);
                        }

                        // ' here we want to scan the current contents of the grid to see if these values are already in the thing

                        // ' first we will build a giant hash string of what we are looking for

                        hst = "";
                        hst2 = "";
                        var loopTo6 = xx - 1;
                        for (x = 0; x <= loopTo6; x++)
                            hst += x.ToString() + _ggrid[x].ToUpper() + "|";

                        if (ga.Contains(hst))
                            fnd = true;
                        else
                            fnd = false;

                        if (!fnd)
                        {
                            var loopTo7 = xx - 1;
                            for (x = 0; x <= loopTo7; x++)
                            {
                                _grid[Conversions.ToInteger(t), x] = _ggrid[x];
                                set_CellBackColor(Conversions.ToInteger(t), x, new SolidBrush(newbackcolor));
                            }

                            t += 1;
                        }

                        if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                            break;

                        if (_ShowProgressBar)
                        {
                            pBar.Increment(1);
                            pBar.Refresh();
                        }
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                pBar.Visible = false;
                gb1.Visible = false;

                AutoSizeCellsToContents = true;

                dbr.Close();

                dbc.Dispose();

                cn.Close();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);
            }
        }

        #endregion

        #region PopulateGridWithData Calls

        /// <summary>
        /// Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
        /// <c>col</c> for the foreground color and <c>gridfont</c> for the cells font style
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="col"></param>
        /// <param name="gridfont"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Color col, Font gridfont)
        {
            int x, numrows;

            try
            {
                StartedDatabasePopulateOperation?.Invoke(this);

                numrows = 0;

                InitializeTheGrid(0, SQLDR.FieldCount);
                var loopTo = Cols - 1;
                for (x = 0; x <= loopTo; x++)
                    _GridHeader[x] = SQLDR.GetName(x);

                while (SQLDR.Read())
                {
                    numrows += 1;

                    if (MaxRowsSelected > 0 & numrows > MaxRowsSelected)
                    {
                        PartialSelection?.Invoke(this);
                        break;
                    }


                    Rows = numrows;
                    var loopTo1 = Cols - 1;
                    // For x = 0 To _Cols - 1
                    for (x = 0; x <= loopTo1; x++)
                    {
                        if (Information.IsDBNull(SQLDR[x]))
                        {
                            if (_omitNulls)
                                _grid[numrows - 1, x] = "";
                            else
                                _grid[numrows - 1, x] = "{NULL}";
                        }
                        else
                        // here we need to do some work on items of certain types
                        if ((SQLDR[x].ToString() ?? "") == "System.Byte[]")
                            _grid[numrows - 1, x] = ReturnByteArrayAsHexString((byte[])SQLDR[x]);
                        else
                        // Console.WriteLine(SQLDR.Item(x).GetType.ToString())
                        if ((SQLDR[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                        {
                            if (_ShowDatesWithTime)
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(SQLDR[x]));

                                _grid[numrows - 1, x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                            }
                            else
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(SQLDR[x]));

                                _grid[numrows - 1, x] = _dt.ToShortDateString();
                            }
                        }
                        else if ((SQLDR[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.GUID")
                            _grid[numrows - 1, x] = "This is a GUID";
                        else
                            _grid[numrows - 1, x] = Conversions.ToString(SQLDR[x]);
                    }
                }

                AllCellsUseThisForeColor(col);

                AllCellsUseThisFont(gridfont);

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
        /// the grids default coloring and fonts for the cells content (settable using the propertries of the grid itself
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR)
        {
            PopulateGridWithData(ref SQLDR, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
        /// <c>ForeColor</c> for the foreground color and <c>gridfont</c> for the cells font style
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Color ForeColor)
        {
            PopulateGridWithData(ref SQLDR, ForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// Will take the supplied SQLDataReader <c>SQLDR</c> and will automatically populate the grid with its contents using
        /// <c>GridFont</c> for the cells font style
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="GridFont"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Font GridFont)
        {
            PopulateGridWithData(ref SQLDR, _DefaultForeColor, GridFont);
        }

        /// <summary>
        /// Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid
        /// with those results. <c>GridFont</c> and <c>col</c> be used to generate the font and the foreground color for the cell contents
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Gridfont"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(string ConnectionString, string Sql, Font Gridfont, Color col)
        {
            System.Data.SqlClient.SqlCommand dbc;
            System.Data.SqlClient.SqlCommand dbc2;
            string sql2;
            long t;
            int y, yy;

            StartedDatabasePopulateOperation?.Invoke(this);

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                y = 0;
                yy = 0;

                // Gets a count of the number of records returned by the query
                using (var cn = new System.Data.SqlClient.SqlConnection())
                {
                    cn.ConnectionString = ConnectionString;
                    cn.Open();
                    sql2 = Sql;
                    dbc2 = new System.Data.SqlClient.SqlCommand(sql2, cn);
                    dbc2.CommandTimeout = _dataBaseTimeOut;

                    using (var dbr2 = dbc2.ExecuteReader())
                    {
                        while (dbr2.Read())
                        {
                            y = y + 1;
                            if (y > MaxRowsSelected & MaxRowsSelected > 0)
                            {
                                y = MaxRowsSelected;
                                yy = -1;
                                break;
                            }
                        }
                    }
                }

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    pBar.Style = ProgressBarStyle.Continuous;
                    gb1.Visible = true;
                    pBar.Step = 1;
                    pBar.Refresh();
                    gb1.Refresh();
                }

                using (var cn2 = new System.Data.SqlClient.SqlConnection())
                {
                    cn2.ConnectionString = ConnectionString;
                    cn2.Open();
                    dbc = new System.Data.SqlClient.SqlCommand(Sql, cn2);
                    dbc.CommandTimeout = _dataBaseTimeOut;

                    using (var dbr = dbc.ExecuteReader())
                    {
                        InitializeTheGrid(y, dbr.FieldCount);
                        AllCellsUseThisFont(Gridfont);
                        AllCellsUseThisForeColor(col);

                        for (int x = 0, loopTo = Cols - 1; x <= loopTo; x++)
                            _GridHeader[x] = dbr.GetName(x);

                        t = 0;

                        while (dbr.Read())
                        {
                            var dbrRow = new List<object>();
                            int x = 0;

                            for (int i = 0, loopTo1 = Cols - 1; i <= loopTo1; i++)
                                dbrRow.Add(dbr[i]);

                            // Process the row item from the data reader
                            foreach (object o in dbrRow)
                            {
                                if (o.Equals(DBNull.Value))
                                {
                                    if (_omitNulls)
                                        _grid[Conversions.ToInteger(t), x] = "";
                                    else
                                        _grid[Conversions.ToInteger(t), x] = "{NULL}";
                                }
                                else if ((o.ToString() ?? "") == "System.Byte[]")
                                    _grid[Conversions.ToInteger(t), x] = ReturnByteArrayAsHexString((byte[])o);
                                else if ((o.GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                                {
                                    var _dt = Convert.ToDateTime(o); // DateTime.Parse(o)
                                    if (_ShowDatesWithTime)
                                        _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                                    else
                                        _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString();
                                }
                                else if ((o.GetType().ToString().ToUpper() ?? "") == "SYSTEM.GUID")
                                {
                                    string s = o.ToString();
                                    _grid[Conversions.ToInteger(t), x] = s;
                                }
                                else
                                    _grid[Conversions.ToInteger(t), x] = Conversions.ToString(o);
                                // increment column index
                                x += 1;
                            }

                            // increment the row index
                            t = t + 1;

                            if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                                break;

                            if (_ShowProgressBar)
                                pBar.PerformStep();
                        }
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                if (_ShowProgressBar)
                {
                    pBar.PerformStep();
                    pBar.Visible = false;
                    gb1.Visible = false;
                }

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }

            finally
            {

                // Set the cursor back to the default cursor
                Cursor.Current = Cursors.Default;
            }
        }

        /// <summary>
        /// Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid
        /// the grids defauls will be used for the cells fonts and coloring characteristics
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(string ConnectionString, string Sql)
        {
            PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor);
        }

        /// <summary>
        /// Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid
        /// the <c>col</c> parameter wwill be used for the cell foreground coloring
        /// the grids defauls will be used for the cells fonts and other coloring characteristics
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Col"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(string ConnectionString, string Sql, Color Col)
        {
            PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col);
        }

        /// <summary>
        /// Will take the supplied <c>ConnectionString</c> and <c>Sql</c> code and query the database gathering the results and populaating the grid
        /// the <c>fnt</c> parameter wwill be used for the cell fonts
        /// the grids defauls will be used for the cells other coloring characteristics
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PopulateGridWithData(string ConnectionString, string Sql, Font fnt)
        {
            PopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor);
        }

        #endregion

        #region SQLPopulateGridWithData Calls

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="col"></param>
        /// <param name="gridfont"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Color col, Font gridfont)
        {
            PopulateGridWithData(ref SQLDR, col, gridfont);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR)
        {
            PopulateGridWithData(ref SQLDR, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Color ForeColor)
        {
            PopulateGridWithData(ref SQLDR, ForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="SQLDR"></param>
        /// <param name="GridFont"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(ref System.Data.SqlClient.SqlDataReader SQLDR, Font GridFont)
        {
            PopulateGridWithData(ref SQLDR, _DefaultForeColor, GridFont);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Gridfont"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(string ConnectionString, string Sql, Font Gridfont, Color col)
        {
            PopulateGridWithData(ConnectionString, Sql, Gridfont, col);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(string ConnectionString, string Sql)
        {
            PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Col"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(string ConnectionString, string Sql, Color Col)
        {
            PopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col);
        }

        /// <summary>
        /// A synonym for the PopulateGridWithData method of the same signature
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void SQLPopulateGridWithData(string ConnectionString, string Sql, Font fnt)
        {
            PopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor);
        }

        #endregion

        #region OLEPopulateGridWithData Calls
        // OLE Populate Data Calls

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
        /// </summary>
        /// <param name="OLEDR"></param>
        /// <param name="col"></param>
        /// <param name="gridfont"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(ref System.Data.OleDb.OleDbDataReader OLEDR, Color col, Font gridfont)
        {
            int x, numrows;

            try
            {
                StartedDatabasePopulateOperation?.Invoke(this);

                numrows = 0;

                InitializeTheGrid(1, OLEDR.FieldCount);
                var loopTo = Cols - 1;
                for (x = 0; x <= loopTo; x++)
                    _GridHeader[x] = OLEDR.GetName(x);

                while (OLEDR.Read())
                {
                    numrows += 1;
                    if (MaxRowsSelected > 0 & numrows < MaxRowsSelected)
                    {
                        Rows = numrows;
                        var loopTo1 = _cols;
                        for (x = 0; x <= loopTo1; x++)
                        {
                            if (Information.IsDBNull(OLEDR[x]))
                            {
                                if (_omitNulls)
                                    _grid[numrows - 1, x] = "";
                                else
                                    _grid[numrows - 1, x] = "{NULL}";
                            }
                            else
                            // here we need to do some work on items of certain types
                            if ((OLEDR[x].ToString() ?? "") == "System.Byte[]")
                                _grid[numrows - 1, x] = ReturnByteArrayAsHexString((byte[])OLEDR[x]);
                            else if ((OLEDR[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                            {
                                if (_ShowDatesWithTime)
                                {
                                    var _dt = DateTime.Parse(Conversions.ToString(OLEDR[x]));

                                    _grid[numrows - 1, x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                                }
                                else
                                {
                                    var _dt = DateTime.Parse(Conversions.ToString(OLEDR[x]));

                                    _grid[numrows - 1, x] = _dt.ToShortDateString();
                                }
                            }
                            else
                                _grid[numrows - 1, x] = Conversions.ToString(OLEDR[x]);
                        }
                    }
                    else
                    {
                        PartialSelection?.Invoke(this);
                        break;
                    }
                }

                AllCellsUseThisForeColor(col);

                AllCellsUseThisFont(gridfont);

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
        /// </summary>
        /// <param name="OLEDR"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(ref System.Data.OleDb.OleDbDataReader OLEDR)
        {
            OLEPopulateGridWithData(ref OLEDR, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
        /// </summary>
        /// <param name="OLEDR"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(ref System.Data.OleDb.OleDbDataReader OLEDR, Color ForeColor)
        {
            OLEPopulateGridWithData(ref OLEDR, ForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OleDbDataReader <c>OLEDR</c> instead
        /// </summary>
        /// <param name="OLEDR"></param>
        /// <param name="GridFont"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(ref System.Data.OleDb.OleDbDataReader OLEDR, Font GridFont)
        {
            OLEPopulateGridWithData(ref OLEDR, _DefaultForeColor, GridFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for OLE data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Gridfont"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(string ConnectionString, string Sql, Font Gridfont, Color col)
        {
            var cn = new System.Data.OleDb.OleDbConnection(ConnectionString);
            System.Data.OleDb.OleDbCommand dbc;
            System.Data.OleDb.OleDbCommand dbc2;
            System.Data.OleDb.OleDbDataReader dbr;
            System.Data.OleDb.OleDbDataReader dbr2;

            string sql2;
            long t;
            int x, y, yy;

            StartedDatabasePopulateOperation?.Invoke(this);

            // _LastConnectionString = ConnectionString
            // _LastSQLString = Sql


            try
            {
                cn.Open();

                sql2 = Sql;
                dbc2 = new System.Data.OleDb.OleDbCommand(sql2, cn);
                dbc2.CommandTimeout = _dataBaseTimeOut;
                dbr2 = dbc2.ExecuteReader();
                y = 0;
                yy = 0;

                while (dbr2.Read())
                {
                    y = y + 1;
                    if (y > MaxRowsSelected & MaxRowsSelected > 0)
                    {
                        y = MaxRowsSelected;
                        yy = -1;
                        break;
                    }
                }

                dbr2.Close();
                dbc2.Dispose();
                cn.Close();

                cn.Open();
                dbc = new System.Data.OleDb.OleDbCommand(Sql, cn);
                dbc.CommandTimeout = _dataBaseTimeOut;

                dbr = dbc.ExecuteReader();

                // Me.Cols = dbr.FieldCount

                InitializeTheGrid(y, dbr.FieldCount);

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }


                AllCellsUseThisFont(Gridfont);
                AllCellsUseThisForeColor(col);
                var loopTo = Cols - 1;
                for (x = 0; x <= loopTo; x++)
                    _GridHeader[x] = dbr.GetName(x);

                t = 0;
                while (dbr.Read())
                {
                    var loopTo1 = Cols - 1;
                    // Me.Rows = t + 1
                    for (x = 0; x <= loopTo1; x++)
                    {
                        if (Information.IsDBNull(dbr[x]))
                        {
                            if (_omitNulls)
                                _grid[Conversions.ToInteger(t), x] = "";
                            else
                                _grid[Conversions.ToInteger(t), x] = "{NULL}";
                        }
                        else
                        // here we need to do some work on items of certain types
                        if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                            _grid[Conversions.ToInteger(t), x] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                        else if ((dbr[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                        {
                            if (_ShowDatesWithTime)
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(dbr[x]));

                                _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                            }
                            else
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(dbr[x]));

                                _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString();
                            }
                        }
                        else if ((dbr[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.GUID")
                            _grid[Conversions.ToInteger(t), x] = dbr[x].ToString();
                        else
                            _grid[Conversions.ToInteger(t), x] = Conversions.ToString(dbr[x]);
                    }
                    t = t + 1;

                    if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                        break;

                    if (_ShowProgressBar)
                    {
                        pBar.Increment(1);
                        pBar.Refresh();
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                pBar.Visible = false;
                gb1.Visible = false;

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                dbr.Close();

                dbc.Dispose();

                cn.Close();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for OLE data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(string ConnectionString, string Sql)
        {
            OLEPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for OLE data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Col"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(string ConnectionString, string Sql, Color Col)
        {
            OLEPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for OLE data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void OLEPopulateGridWithData(string ConnectionString, string Sql, Font fnt)
        {
            OLEPopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor);
        }

        #endregion

        #region ODBCPopulateGridWithData Calls

        // ODBC Populate Data Calls

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
        /// </summary>
        /// <param name="OdbcDR"></param>
        /// <param name="col"></param>
        /// <param name="gridfont"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(ref System.Data.Odbc.OdbcDataReader OdbcDR, Color col, Font gridfont)
        {
            int x, numrows;

            try
            {
                StartedDatabasePopulateOperation?.Invoke(this);

                numrows = 0;

                InitializeTheGrid(1, OdbcDR.FieldCount);
                var loopTo = Cols - 1;
                for (x = 0; x <= loopTo; x++)
                    _GridHeader[x] = OdbcDR.GetName(x);

                while (OdbcDR.Read())
                {
                    numrows += 1;
                    if (MaxRowsSelected > 0 & numrows < MaxRowsSelected)
                    {
                        Rows = numrows;
                        var loopTo1 = _cols;
                        for (x = 0; x <= loopTo1; x++)
                        {
                            if (Information.IsDBNull(OdbcDR[x]))
                            {
                                if (_omitNulls)
                                    _grid[numrows - 1, x] = "";
                                else
                                    _grid[numrows - 1, x] = "{NULL}";
                            }
                            else
                            // here we need to do some work on items of certain types
                            if ((OdbcDR[x].ToString() ?? "") == "System.Byte[]")
                                _grid[numrows - 1, x] = ReturnByteArrayAsHexString((byte[])OdbcDR[x]);
                            else if ((OdbcDR[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                            {
                                if (_ShowDatesWithTime)
                                {
                                    var _dt = DateTime.Parse(Conversions.ToString(OdbcDR[x]));

                                    _grid[numrows - 1, x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                                }
                                else
                                {
                                    var _dt = DateTime.Parse(Conversions.ToString(OdbcDR[x]));

                                    _grid[numrows - 1, x] = _dt.ToShortDateString();
                                }
                            }
                            else
                                _grid[numrows - 1, x] = Conversions.ToString(OdbcDR[x]);
                        }
                    }
                    else
                    {
                        PartialSelection?.Invoke(this);
                        break;
                    }
                }

                AllCellsUseThisForeColor(col);

                AllCellsUseThisFont(gridfont);

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
        /// </summary>
        /// <param name="OdbcDR"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(ref System.Data.Odbc.OdbcDataReader OdbcDR)
        {
            ODBCPopulateGridWithData(ref OdbcDR, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
        /// </summary>
        /// <param name="OdbcDR"></param>
        /// <param name="ForeColor"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(ref System.Data.Odbc.OdbcDataReader OdbcDR, Color ForeColor)
        {
            ODBCPopulateGridWithData(ref OdbcDR, ForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but uses an OdbcDataReader <c>OdbcDR</c> instead
        /// </summary>
        /// <param name="OdbcDR"></param>
        /// <param name="GridFont"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(ref System.Data.Odbc.OdbcDataReader OdbcDR, Font GridFont)
        {
            ODBCPopulateGridWithData(ref OdbcDR, _DefaultForeColor, GridFont);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for ODBC data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Gridfont"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(string ConnectionString, string Sql, Font Gridfont, Color col)
        {
            var cn = new System.Data.Odbc.OdbcConnection(ConnectionString);
            System.Data.Odbc.OdbcCommand dbc;
            System.Data.Odbc.OdbcCommand dbc2;
            System.Data.Odbc.OdbcDataReader dbr;
            System.Data.Odbc.OdbcDataReader dbr2;

            string sql2;
            long t;
            int x, y, yy;

            StartedDatabasePopulateOperation?.Invoke(this);

            // _LastConnectionString = ConnectionString
            // _LastSQLString = Sql


            try
            {
                cn.Open();

                sql2 = Sql;
                dbc2 = new System.Data.Odbc.OdbcCommand(sql2, cn);
                dbc2.CommandTimeout = _dataBaseTimeOut;
                dbr2 = dbc2.ExecuteReader();
                y = 0;
                yy = 0;

                while (dbr2.Read())
                {
                    y = y + 1;
                    if (y > MaxRowsSelected & MaxRowsSelected > 0)
                    {
                        y = MaxRowsSelected;
                        yy = -1;
                        break;
                    }
                }

                dbr2.Close();
                dbc2.Dispose();
                cn.Close();

                cn.Open();
                dbc = new System.Data.Odbc.OdbcCommand(Sql, cn);
                dbc.CommandTimeout = _dataBaseTimeOut;

                dbr = dbc.ExecuteReader();

                // Me.Cols = dbr.FieldCount

                InitializeTheGrid(y, dbr.FieldCount);

                if (_ShowProgressBar)
                {
                    pBar.Maximum = y;
                    pBar.Minimum = 0;
                    pBar.Value = 0;
                    pBar.Visible = true;
                    gb1.Visible = true;
                    pBar.Refresh();
                    gb1.Refresh();
                }

                AllCellsUseThisFont(Gridfont);
                AllCellsUseThisForeColor(col);
                var loopTo = Cols - 1;
                for (x = 0; x <= loopTo; x++)
                    _GridHeader[x] = dbr.GetName(x);

                t = 0;
                while (dbr.Read())
                {
                    var loopTo1 = Cols - 1;
                    // Me.Rows = t + 1
                    for (x = 0; x <= loopTo1; x++)
                    {
                        if (Information.IsDBNull(dbr[x]))
                        {
                            if (_omitNulls)
                                _grid[Conversions.ToInteger(t), x] = "";
                            else
                                _grid[Conversions.ToInteger(t), x] = "{NULL}";
                        }
                        else
                        // here we need to do some work on items of certain types
                        if ((dbr[x].ToString() ?? "") == "System.Byte[]")
                            _grid[Conversions.ToInteger(t), x] = ReturnByteArrayAsHexString((byte[])dbr[x]);
                        else if ((dbr[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.DATETIME")
                        {
                            if (_ShowDatesWithTime)
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(dbr[x]));

                                _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString() + " " + _dt.ToShortTimeString();
                            }
                            else
                            {
                                var _dt = DateTime.Parse(Conversions.ToString(dbr[x]));

                                _grid[Conversions.ToInteger(t), x] = _dt.ToShortDateString();
                            }
                        }
                        else if ((dbr[x].GetType().ToString().ToUpper() ?? "") == "SYSTEM.GUID")
                            _grid[Conversions.ToInteger(t), x] = dbr[x].ToString();
                        else
                            _grid[Conversions.ToInteger(t), x] = Conversions.ToString(dbr[x]);
                    }
                    t = t + 1;

                    if (MaxRowsSelected > 0 & t >= MaxRowsSelected)
                        break;

                    if (_ShowProgressBar)
                    {
                        pBar.Increment(1);
                        pBar.Refresh();
                    }
                }

                if (yy == -1)
                    PartialSelection?.Invoke(this);

                pBar.Visible = false;
                gb1.Visible = false;

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                dbr.Close();

                dbc.Dispose();

                cn.Close();

                FinishedDatabasePopulateOperation?.Invoke(this);

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                pBar.Visible = false;
                gb1.Visible = false;

                FinishedDatabasePopulateOperation?.Invoke(this);

                Interaction.MsgBox(ex.Message);
            }
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for ODBC data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(string ConnectionString, string Sql)
        {
            ODBCPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, _DefaultForeColor);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for ODBC data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="Col"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(string ConnectionString, string Sql, Color Col)
        {
            ODBCPopulateGridWithData(ConnectionString, Sql, _DefaultCellFont, Col);
        }

        /// <summary>
        /// As the PopulateGridWithData method of the same signature but <c>ConnectionString</c> parameter must be in the correct
        /// syntax for ODBC data access
        /// </summary>
        /// <param name="ConnectionString"></param>
        /// <param name="Sql"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void ODBCPopulateGridWithData(string ConnectionString, string Sql, Font fnt)
        {
            ODBCPopulateGridWithData(ConnectionString, Sql, fnt, _DefaultForeColor);
        }

        #endregion

        #region PopulateViaWebServiceString Calls

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults)
        {
            PopulateViaWebServiceString(WebServiceResults, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, Color col)
        {
            PopulateViaWebServiceString(WebServiceResults, col, _DefaultCellFont);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, Font fnt)
        {
            PopulateViaWebServiceString(WebServiceResults, _DefaultForeColor, fnt);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="col"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, Color col, Font fnt)
        {
            var argarray = WebServiceResults.Split('|');
            int x, y;
            int r, c;
            r = Conversions.ToInteger(Conversion.Val(argarray[argarray.GetUpperBound(0)]));
            c = Conversions.ToInteger(Conversion.Val(Conversions.ToDouble(argarray[argarray.GetUpperBound(0) - 1]) + 1) - 1);

            // Me.Rows = Val(argarray(argarray.GetUpperBound(0))) + 1
            // Me.Cols = Val(argarray(argarray.GetUpperBound(0) - 1) + 1) - 1

            InitializeTheGrid(r, c);
            var loopTo = Cols - 1;
            for (y = 0; y <= loopTo; y++)
                _GridHeader[y] = argarray[y];

            if (OmitNulls)
            {
                var loopTo1 = _rows;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = _cols - 1;
                    for (y = 0; y <= loopTo2; y++)
                    {
                        if ((Strings.UCase(argarray[x * _cols + y]) ?? "") == "{NULL}")
                            argarray[x * _cols + y] = "";
                    }
                }
            }

            var loopTo3 = _rows;
            for (x = 1; x <= loopTo3; x++)
            {
                var loopTo4 = _cols - 1;
                for (y = 0; y <= loopTo4; y++)
                    _grid[x - 1, y] = argarray[x * _cols + y];
            }

            AllCellsUseThisForeColor(col);
            AllCellsUseThisFont(fnt);
            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="delimiter"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, string delimiter)
        {
            PopulateViaWebServiceString(WebServiceResults, delimiter, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="delimiter"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, string delimiter, Color col)
        {
            PopulateViaWebServiceString(WebServiceResults, delimiter, col, _DefaultCellFont);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="delimiter"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, string delimiter, Font fnt)
        {
            PopulateViaWebServiceString(WebServiceResults, delimiter, _DefaultForeColor, fnt);
        }

        /// <summary>
        /// The Webservice set of populators actually are forms of populators from formatted string similar to populat from text files
        /// The difference being that defines webservices were returning these formatted strings of text. Back before the days of being
        /// able to send complex data types across the HTTP wire. Carefully crafted webservices could spit out their results as streams
        /// of delimited text. These methods would then parse that text and rehydrate the results into the grid for viuual presentation.
        /// the TAIClient Time Tracking tool employes this technique. The various parameters are self explanitory
        /// </summary>
        /// <param name="WebServiceResults"></param>
        /// <param name="Delimiter"></param>
        /// <param name="col"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PopulateViaWebServiceString(string WebServiceResults, string Delimiter, Color col, Font fnt)
        {
            // parse the rows and colmuns off the string
            int rows = Conversions.ToInteger(WebServiceResults.Substring(WebServiceResults.LastIndexOf(Delimiter) + Delimiter.Length));
            WebServiceResults = WebServiceResults.Substring(0, WebServiceResults.LastIndexOf(Delimiter));
            int cols = Conversions.ToInteger(WebServiceResults.Substring(WebServiceResults.LastIndexOf(Delimiter) + Delimiter.Length));
            WebServiceResults = WebServiceResults.Substring(0, WebServiceResults.LastIndexOf(Delimiter));

            //string[,] argarray = ReturnDelimitedStringAsArray(WebServiceResults, cols, rows, Delimiter);

            string[,] argarray = (string[,])ReturnDelimitedStringAsArray(WebServiceResults, cols, rows, Delimiter);

            InitializeTheGrid(rows - 1, cols);
            int x, y;
            var loopTo = cols - 1;
            for (y = 0; y <= loopTo; y++)
                _GridHeader[y] = argarray[0, y];

            if (OmitNulls)
            {
                var loopTo1 = _rows;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = _cols - 1;
                    for (y = 0; y <= loopTo2; y++)
                    {
                        if ((Strings.UCase(argarray[x, y]) ?? "") == "{NULL}")
                            argarray[x, y] = "";
                    }
                }
            }

            var loopTo3 = _rows - 1;
            for (x = 1; x <= loopTo3; x++)
            {
                var loopTo4 = _cols - 1;
                for (y = 0; y <= loopTo4; y++)
                    _grid[x, y] = argarray[x, y];
            }

            AllCellsUseThisFont(fnt);
            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();
            AllCellsUseThisForeColor(col);

            Refresh();

            NormalizeTearaways();
        }

        #endregion

        #region PivotPopulate Calls

        /// <summary>
        /// The pivot populate calls simulate pivot table functionality in excel.
        /// The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
        /// The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
        /// then for each unique value set in each of the two columns the
        /// </summary>
        /// <param name="sgrid"></param>
        /// <param name="xcol"></param>
        /// <param name="ycol"></param>
        /// <param name="scol"></param>
        /// <param name="FormatSpec"></param>
        /// <remarks></remarks>
        public void PivotPopulate(TAIGridControl sgrid, int xcol, int ycol, int scol, string FormatSpec)
        {
            PivotPopulate(sgrid, xcol, ycol, scol, FormatSpec, _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// The pivot populate calls simulate pivot table functionality in excel.
        /// The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
        /// The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
        /// then for each unique value set in each of the two columns the
        /// </summary>
        /// <param name="sgrid"></param>
        /// <param name="xcol"></param>
        /// <param name="ycol"></param>
        /// <param name="scol"></param>
        /// <remarks></remarks>
        public void PivotPopulate(TAIGridControl sgrid, int xcol, int ycol, int scol)
        {
            PivotPopulate(sgrid, xcol, ycol, scol, "0.0000", _DefaultForeColor, _DefaultCellFont);
        }

        /// <summary>
        /// The pivot populate calls simulate pivot table functionality in excel.
        /// The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
        /// The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
        /// then for each unique value set in each of the two columns the
        /// </summary>
        /// <param name="sgrid"></param>
        /// <param name="xcol"></param>
        /// <param name="ycol"></param>
        /// <param name="scol"></param>
        /// <param name="formatspec"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PivotPopulate(TAIGridControl sgrid, int xcol, int ycol, int scol, string formatspec, Color col)
        {
            PivotPopulate(sgrid, xcol, ycol, scol, formatspec, col, _DefaultCellFont);
        }

        /// <summary>
        /// The pivot populate calls simulate pivot table functionality in excel.
        /// The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
        /// The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
        /// then for each unique value set in each of the two columns the
        /// </summary>
        /// <param name="sgrid"></param>
        /// <param name="xcol"></param>
        /// <param name="ycol"></param>
        /// <param name="scol"></param>
        /// <param name="formatspec"></param>
        /// <param name="col"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PivotPopulate(TAIGridControl sgrid, int xcol, int ycol, int scol, string formatspec, Color col, Font fnt)
        {
            int x, y, xxx;
            string a, b, c;
            double aa;
            int sx = sgrid.Cols - 1;
            int sy = sgrid.Rows - 1;
            var uniquerows = new string[sy + 1];
            var uniquecols = new string[sx + 1];
            // Dim formatspec As String = "0.0000"

            var u = new ArrayList();
            var uu = new ArrayList();

            u.Clear();
            uu.Clear();
            var loopTo = sy;


            // how many unique vals do we have in the Xcol

            for (x = 0; x <= loopTo; x++)
            {
                a = sgrid.get_item(x, xcol);
                if (!u.Contains(a))
                    u.Add(a);
            }

            Cols = u.Count + 1;
            var loopTo1 = sy;

            // how many unique vals do we have in the Ycol

            for (x = 0; x <= loopTo1; x++)
            {
                a = sgrid.get_item(x, ycol);
                if (!uu.Contains(a))
                    uu.Add(a);
            }

            Rows = uu.Count;
            var loopTo2 = u.Count;

            // here we will populate the header and the y column with the values being rolled up
            for (x = 1; x <= loopTo2; x++)
                this.set_HeaderLabel(x, u[x - 1].ToString());

            set_HeaderLabel(0, sgrid.get_HeaderLabel(ycol));
            var loopTo3 = uu.Count - 1;
            for (y = 0; y <= loopTo3; y++)
                this.set_item(y, 0, uu[y].ToString());
            var loopTo4 = u.Count - 1;

            // here we will actually populate the values

            for (x = 0; x <= loopTo4; x++)
            {
                b = Conversions.ToString(u[x]);
                var loopTo5 = uu.Count - 1;
                for (y = 0; y <= loopTo5; y++)
                {
                    c = Conversions.ToString(uu[y]);
                    aa = 0;
                    var loopTo6 = sy;
                    for (xxx = 0; xxx <= loopTo6; xxx++)
                    {
                        if ((sgrid.get_item(xxx, xcol) ?? "") == (b ?? "") & (sgrid.get_item(xxx, ycol) ?? "") == (c ?? ""))
                            aa = aa + Conversion.Val(sgrid.get_item(xxx, scol));
                    }

                    set_item(y, x + 1, Strings.Format(aa, formatspec));
                }
            }

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();
            AllCellsUseThisForeColor(col);
            AllCellsUseThisFont(fnt);
            Refresh();

            NormalizeTearaways();
        }

        /// <summary>
        /// The pivot populate calls simulate pivot table functionality in excel.
        /// The instance grid that the pethod is called on will be populated with data from a source grid <c>sgrid</c>
        /// The defined <c>xcol</c> and <c>ycol</c> parameters will be used to search the source grid for unique values
        /// then for each unique value set in each of the two columns the
        /// </summary>
        /// <param name="sgrid"></param>
        /// <param name="xcol"></param>
        /// <param name="ycol"></param>
        /// <param name="scol"></param>
        /// <param name="formatspec"></param>
        /// <param name="fnt"></param>
        /// <remarks></remarks>
        public void PivotPopulate(TAIGridControl sgrid, int xcol, int ycol, int scol, string formatspec, Font fnt)
        {
            PivotPopulate(sgrid, xcol, ycol, scol, formatspec, _DefaultForeColor, fnt);
        }

        #endregion

        #region FrequencyDistribution Calls
        public void FrequencyDistribution(TAIGridControl sgrid, int ColForFrequency)
        {
            var codes = new ArrayList();

            int t;
            int tt;
            var loopTo = sgrid.Rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                string cd = sgrid.get_item(t, ColForFrequency);
                if (!codes.Contains(cd))
                    codes.Add(cd);
            }

            Rows = codes.Count;
            Cols = 2;
            set_HeaderLabel(0, sgrid.get_HeaderLabel(ColForFrequency));
            set_HeaderLabel(1, "Frequency");
            var loopTo1 = codes.Count - 1;
            for (t = 0; t <= loopTo1; t++)
            {
                this.set_item(t, 0, codes[t].ToString());

                int result = 0;
                var loopTo2 = sgrid.Rows - 1;
                for (tt = 0; tt <= loopTo2; tt++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(sgrid.get_item(tt, ColForFrequency), codes[t], false)))
                        result += 1;
                }

                set_item(t, 1, result.ToString());
            }

            AutoSizeCellsToContents = true;
            Refresh();

            NormalizeTearaways();
        }

        public void FrequencyDistribution(TAIGridControl sgrid, int ColForFrequency, bool SortDescending)
        {
            FrequencyDistribution(sgrid, ColForFrequency);

            SortGridOnColumnNumeric(1, SortDescending);

            AutoSizeCellsToContents = true;
            Refresh();
        }

        #endregion

        #region PopulateGridWithDataTable Calls

        /// <summary>
        /// Will take the supplied dataSet and extract the first table from that dataset and populate the grid with the
        /// contents of that datatable
        /// </summary>
        /// <param name="dset"></param>
        /// <remarks></remarks>
        public void PopulateGridWithADataTable(DataSet dset)
        {
            PopulateGridWithADataTable(dset.Tables[0]);
        }

        /// <summary>
        /// Will take thge supplied datatable and populate the grid with the contents oif that datatable
        /// </summary>
        /// <param name="dt"></param>
        /// <remarks></remarks>
        public void PopulateGridWithADataTable(DataTable dt)
        {
            int t = 0;
            int x = 0;
            int y = 0;
            string typ;

            InitializeTheGrid(dt.Rows.Count, dt.Columns.Count);
            var loopTo = dt.Columns.Count - 1;
            for (t = 0; t <= loopTo; t++)
                set_HeaderLabel(t, dt.Columns[t].ColumnName);
            var loopTo1 = dt.Rows.Count - 1;
            for (x = 0; x <= loopTo1; x++)
            {
                var loopTo2 = dt.Columns.Count - 1;
                for (y = 0; y <= loopTo2; y++)
                {
                    typ = dt.Rows[x][y].GetType().ToString().ToUpper();

                    if ((typ ?? "") == "SYSTEM.STRING")
                        _grid[x, y] = Convert.ToString(dt.Rows[x][y]);
                    else if ((typ ?? "") == "SYSTEM.DBNULL")
                    {
                        if (_omitNulls)
                            _grid[x, y] = "";
                        else
                            _grid[x, y] = "{NULL}";
                    }
                    else if ((typ ?? "") == "SYSTEM.DATETIME")
                    {
                        if (_ShowDatesWithTime)
                            _grid[x, y] = Convert.ToDateTime(dt.Rows[x][y]).ToString();
                        else
                            _grid[x, y] = Convert.ToDateTime(dt.Rows[x][y]).ToShortDateString();
                    }
                    else if ((typ ?? "") == "SYSTEM.SINGLE")
                        _grid[x, y] = Convert.ToSingle(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.INT32")
                        _grid[x, y] = Convert.ToInt32(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.INT16")
                        _grid[x, y] = Convert.ToInt16(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.INT64")
                        _grid[x, y] = Convert.ToInt64(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.BOOLEAN")
                    {
                        if (Convert.ToBoolean(dt.Rows[x][y]))
                            _grid[x, y] = "TRUE";
                        else
                            _grid[x, y] = "FALSE";
                    }
                    else if ((typ ?? "") == "SYSTEM.DECIMAL")
                        _grid[x, y] = Convert.ToDecimal(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.DOUBLE")
                        _grid[x, y] = Convert.ToDouble(dt.Rows[x][y]).ToString();
                    else if ((typ ?? "") == "SYSTEM.RUNTIMETYPE")
                        _grid[x, y] = Convert.ToString(dt.Rows[x][y]);
                    else
                        _grid[x, y] = dt.Rows[x][y].GetType().ToString();
                }
            }

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();

            NormalizeTearaways();
        }

        #endregion

        #region PopulateFrimADirectory Calls

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
        /// The grid will then be cleared and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname)
        {
            PopulateFromADirectory(Dirname, _DefaultCellFont, _DefaultForeColor, "*");
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
        /// The grid will then be cleared and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>gridfont</c> font will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="Gridfont"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Font Gridfont)
        {
            PopulateFromADirectory(Dirname, Gridfont, _DefaultForeColor, "*");
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
        /// The grid will then be cleared and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>col</c> color will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Color col)
        {
            PopulateFromADirectory(Dirname, _DefaultCellFont, col, "*");
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents.
        /// The grid will then be cleared and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>col</c> color and <c>gridfont</c> font will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="Gridfont"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Font Gridfont, Color col)
        {
            PopulateFromADirectory(Dirname, Gridfont, col, "*");
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
        /// pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="Pattern"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, string Pattern)
        {
            PopulateFromADirectory(Dirname, _DefaultCellFont, _DefaultForeColor, Pattern);
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
        /// pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>gridfont</c> font will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="Gridfont"></param>
        /// <param name="Pattern"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Font Gridfont, string Pattern)
        {
            PopulateFromADirectory(Dirname, Gridfont, _DefaultForeColor, Pattern);
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
        /// pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>col</c> color will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="col"></param>
        /// <param name="Pattern"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Color col, string Pattern)
        {
            PopulateFromADirectory(Dirname, _DefaultCellFont, col, Pattern);
        }

        /// <summary>
        /// Will open the directory specified by <c>Dirname</c> and will enumerate its contents via the supplied
        /// pattern <c>Pattern</c>. The grid will then be clears and the results enumerated in the grids
        /// contents showing the FileName, Last Update Time, and physical size. The supplied
        /// <c>col</c> color and <c>gridfont</c> font will be used to show the content generated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="gridfont"></param>
        /// <param name="col"></param>
        /// <param name="Pattern"></param>
        /// <remarks></remarks>
        public void PopulateFromADirectory(string Dirname, Font gridfont, Color col, string Pattern)
        {
            try
            {
                var dinf = new System.IO.DirectoryInfo(Dirname);

                var finf = dinf.GetFiles(Pattern);
                int y;
                int r, c;

                r = finf.GetUpperBound(0);
                c = 3;

                InitializeTheGrid(r + 1, c);

                _GridTitle = "Files in " + Dirname;

                _GridHeader[0] = "File Name";
                _GridHeader[1] = "File Time";
                _GridHeader[2] = "File Size";
                var loopTo = r;
                for (y = 0; y <= loopTo; y++)
                {
                    _grid[y, 0] = finf[y].FullName;
                    _grid[y, 1] = Conversions.ToString(finf[y].LastAccessTime);
                    _grid[y, 2] = finf[y].Length.ToString();
                }

                AllCellsUseThisFont(gridfont);
                AllCellsUseThisForeColor(col);

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();

                Refresh();

                NormalizeTearaways();
            }
            catch (Exception ex)
            {
                InitializeTheGrid(1, 3);
                _GridTitle = "Files in " + "We got a problem...";

                _GridHeader[0] = "File Name";
                _GridHeader[1] = "File Time";
                _GridHeader[2] = "File Size";

                Refresh();

                NormalizeTearaways();
            }
        }

        /// <summary>
        /// Attempts to open the directory specfied by <c>Dirname</c> and enumerate the entire contents
        /// The results are the appended to the current grids contents
        /// The FileName, Last Update Time, and physical size are enumerated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <remarks></remarks>
        public void AppendPopulate(string Dirname)
        {
            AppendPopulate(Dirname, "*");
        }

        /// <summary>
        /// Attempts to open the directory specfied by <c>Dirname</c> and enumerate the contents via the supplied <c>Pattern</c>
        /// The results are the appended to the current grids contents
        /// The FileName, Last Update Time, and physical size are enumerated.
        /// </summary>
        /// <param name="Dirname"></param>
        /// <param name="Pattern"></param>
        /// <remarks></remarks>
        public void AppendPopulate(string Dirname, string Pattern)
        {
            var dinf = new System.IO.DirectoryInfo(Dirname);

            var finf = dinf.GetFiles(Pattern);
            int y;
            int r;

            r = finf.GetUpperBound(0);

            _GridTitle = "Files in...";

            if (_cols != 3)
            {
                _cols = 3;

                _GridHeader[0] = "File Name";
                _GridHeader[1] = "File Time";
                _GridHeader[2] = "File Size";
            }

            int oldrows = _rows - 1;

            Rows += r + 1;
            var loopTo = r;
            for (y = 0; y <= loopTo; y++)
            {
                _grid[y + oldrows, 0] = finf[y].FullName;
                _grid[y + oldrows, 1] = Conversions.ToString(finf[y].LastAccessTime);
                _grid[y + oldrows, 2] = finf[y].Length.ToString();
            }

            AllCellsUseThisFont(_DefaultCellFont);
            AllCellsUseThisForeColor(_DefaultForeColor);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();
        }

        #endregion

        /// <summary>
        /// Will fire the CellClicked event from the outside world
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void RaiseCellClickedEvent(int Row, int col)
        {
            CellClicked?.Invoke(this, Row, col);
        }

        /// <summary>
        /// Will fire the CellDoubleClicked event from the outside world
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void RaiseCellDoubleClickedEvent(int row, int col)
        {
            CellDoubleClicked?.Invoke(this, row, col);
        }

        /// <summary>
        /// Will fire the GridHover event from the outside world
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="rowid"></param>
        /// <param name="colid"></param>
        /// <param name="textvalue"></param>
        /// <remarks></remarks>
        public void RaiseGridHoverEvents(object sender, int rowid, int colid, string textvalue)
        {
            if (!_TearAwayWork)
                GridHover?.Invoke(sender, rowid, colid, textvalue);
        }

        /// <summary>
        /// Will remove a specified row of data from the grids contents
        /// If rowid is greater than rows in the grid nothing will be removed
        /// </summary>
        /// <param name="rowid"></param>
        /// <remarks></remarks>
        public void RemoveRowFromGrid(int rowid)
        {
            var hdr = _GridHeader;

            if (rowid < 0 | rowid > _rows - 1)
                return;

            _Painting = true;    // stop the painting operations

            var tmpgrid = _grid;    // get the current grid
            var newtmpgrid = new string[_rows - 2 + 1, _cols];  // a temporary storage area
            int x = 0;
            int xx = 0;
            int y = 0;
            var loopTo = _rows - 1;
            for (x = 0; x <= loopTo; x++)
            {
                if (x != rowid)
                {
                    var loopTo1 = _cols - 1;
                    for (y = 0; y <= loopTo1; y++)
                        newtmpgrid[xx, y] = tmpgrid[x, y];
                    xx += 1;
                }
            }


            // For x = 0 To rowid - 1  ' loop through all the rows up to the one we want to get rid of
            // For y = 0 To _cols - 1
            // newtmpgrid(x, y) = tmpgrid(x, y)    ' all the colums
            // Next
            // Next

            // For x = rowid + 1 To _rows - 1  ' loop 
            // For y = 0 To _cols - 1
            // newtmpgrid(x - 1, y) = tmpgrid(x, y)
            // Next
            // Next

            // Me.PopulateGridFromArray(newtmpgrid)

            PopulateGridFromArray(newtmpgrid, _DefaultCellFont, _DefaultForeColor, false, false, hdr);

            _Painting = false;

            Refresh();
        }

        /// <summary>
        /// Will attempt to remove specific rowws from the grid contained in the supplied
        /// arraylist of integers
        /// </summary>
        /// <param name="ListOfRows"></param>
        /// <remarks></remarks>
        public void RemoveRowsFromGrid(ArrayList ListOfRows)
        {
            // will take arraylist rows and purge them from the _grid(x,y) array

            if (ListOfRows.Count == 0)
                // I see nothing... I Do Nothing...
                return;

            var hdr = _GridHeader;
            int x, y, t;

            _Painting = true;    // stop the painting operations

            // calculate final row count

            x = 0;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (!ListOfRows.Contains(t))
                    x += 1;
            }

            var finalgrid = new string[x, _cols];

            x = 0;
            var loopTo1 = _rows - 1;
            for (t = 0; t <= loopTo1; t++)
            {
                if (!ListOfRows.Contains(t))
                {
                    var loopTo2 = _cols - 1;
                    // we have a row to go
                    for (y = 0; y <= loopTo2; y++)
                        finalgrid[x, y] = _grid[t, y];
                    x += 1;
                }
            }

            // finalgrid should have what we need now

            _SelectedRow = -1;
            _SelectedRows.Clear();

            var colp = _colPasswords;
            var colmaxchars = _colMaxCharacters;
            var coledit = _colEditable;
            var rowedit = _rowEditable;
            var colhid = _colhidden;

            PopulateGridFromArray(finalgrid, _DefaultCellFont, _DefaultForeColor, false, false, hdr);

            _colPasswords = colp;
            _colMaxCharacters = colmaxchars;
            _colEditable = coledit;
            _colhidden = colhid;

            _rowEditable = rowedit;

            _Painting = false;

            Refresh();
            NormalizeTearaways();
        }

        /// <summary>
        /// Will attempt to remove a specific column from the grids contents
        /// If colid is greater than the number of columns in the grid nothing will
        /// be removed
        /// </summary>
        /// <param name="colid"></param>
        /// <remarks></remarks>
        public void RemoveColFromGrid(int colid)
        {
            if (colid < 0 | colid > _cols - 1)
                return;

            _Painting = true;    // stop the painting operations

            var tmpgrid = GetGridAsArray();    // get the current grid
            var newtmpgrid = new string[_rows, _cols - 2 + 1];  // a temporary storage area
            int x = 0;
            int y = 0;

            if (colid != 0)
            {
                var loopTo = colid - 1;
                // we are not skipping the first row
                for (y = 0; y <= loopTo; y++)
                {
                    var loopTo1 = _rows - 1;
                    for (x = 0; x <= loopTo1; x++)
                        newtmpgrid[x, y] = tmpgrid[x, y];
                }
            }

            if (colid < _cols - 1)
            {
                var loopTo2 = _cols - 1;
                // we are not skipping the laast row
                for (y = colid + 1; y <= loopTo2; y++)
                {
                    var loopTo3 = _rows - 1;
                    for (x = 0; x <= loopTo3; x++)
                        newtmpgrid[x, y - 1] = tmpgrid[x, y];
                }
            }

            PopulateGridFromArray(newtmpgrid);
            // PopulateGridFromArray(arr, _DefaultCellFont, _DefaultForeColor, True)

            CheckGridTearAways(colid);

            _Painting = false;

            Refresh();
        }

        /// <summary>
        /// Will attempt to walk the contents of a column for integers 1 through 12
        /// on finding a 1 through 12 it will replace the integers with the name of that
        /// month number IE 1 = January, 2 = February...
        /// </summary>
        /// <param name="columnid"></param>
        /// <remarks></remarks>
        public void ReplaceColMonthNumericWithMonthName(int columnid)
        {
            int y;
            string a;
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                if (_grid[y, columnid] == null)
                {
                }
                else
                {
                    a = _grid[y, columnid];
                    if (Information.IsNumeric(a))
                    {
                        // its at least a number
                        switch (Conversions.ToInteger(a))
                        {
                            case 1:
                                {
                                    a = "January";
                                    break;
                                }

                            case 2:
                                {
                                    a = "February";
                                    break;
                                }

                            case 3:
                                {
                                    a = "March";
                                    break;
                                }

                            case 4:
                                {
                                    a = "April";
                                    break;
                                }

                            case 5:
                                {
                                    a = "May";
                                    break;
                                }

                            case 6:
                                {
                                    a = "June";
                                    break;
                                }

                            case 7:
                                {
                                    a = "July";
                                    break;
                                }

                            case 8:
                                {
                                    a = "August";
                                    break;
                                }

                            case 9:
                                {
                                    a = "September";
                                    break;
                                }

                            case 10:
                                {
                                    a = "October";
                                    break;
                                }

                            case 11:
                                {
                                    a = "November";
                                    break;
                                }

                            case 12:
                                {
                                    a = "December";
                                    break;
                                }

                            default:
                                {
                                    break;
                                }
                        }

                        _grid[y, columnid] = a;
                    }
                }
            }

            Invalidate();
        }

        /// <summary>
        /// Will walk the list of columns tornaway and will set the windows to be siz width
        /// </summary>
        /// <param name="siz"></param>
        /// <remarks></remarks>
        public void ResizeTearawayColumnsHorizontally(int siz)
        {
            if (TearAways.Count == 0)
                // we ain't got any stinking tearaways so lets bail
                return;

            int t;

            TearAwayWindowEntry tear;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.Width = siz;
            }

            ArrangeTearAwayWindows();
        }

        /// <summary>
        /// Will walk the list of torn away columns and will set each one to siz height
        /// </summary>
        /// <param name="siz"></param>
        /// <remarks></remarks>
        public void ResizeTearawayColumnsVertically(int siz)
        {
            if (TearAways.Count == 0)
                // we ain't got any stinking tearaways so lets bail
                return;

            if (_TearAwayWork)
                return;
            _TearAwayWork = true;

            int t;

            TearAwayWindowEntry tear;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.SuspendLayout();
                tear.Winform.Height = siz;
            }

            ArrangeTearAwayWindows();
            var loopTo1 = TearAways.Count - 1;
            for (t = 0; t <= loopTo1; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.ResumeLayout();
            }

            _TearAwayWork = false;
        }

        /// <summary>
        /// Will walk the list of torn away columns and will set each to sizx width andf sizy height
        /// </summary>
        /// <param name="sizx"></param>
        /// <param name="sizy"></param>
        /// <remarks></remarks>
        public void ResizeTearawayColumnsVerticallyAndHorizontally(int sizx, int sizy)
        {
            if (TearAways.Count == 0)
                // we ain't got any stinking tearaways so lets bail
                return;

            int t;

            TearAwayWindowEntry tear;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.Height = sizy;
                tear.Winform.Width = sizx;
            }

            ArrangeTearAwayWindows();
        }

        /// <summary>
        /// Will set the column at colid edit restrictions to the list contained in the CaretDelimitedString
        /// </summary>
        /// <param name="colid"></param>
        /// <param name="CaretDelimitedString"></param>
        /// <remarks></remarks>
        public void RestrictColumnEditsTo(int colid, string CaretDelimitedString)
        {
            var reslist = new EditColumnRestrictor();

            reslist.ColumnID = colid;
            reslist.RestrictedList = CaretDelimitedString;

            foreach (EditColumnRestrictor it in _colEditRestrictions)
            {
                if (it.ColumnID == colid)
                    _colEditRestrictions.Remove(it);
            }

            _colEditRestrictions.Add(reslist);
        }

        /// <summary>
        /// Will set the column at colid edit restrictions list to the supplied ArrayListOfStrings
        /// </summary>
        /// <param name="colid"></param>
        /// <param name="ArrayListOfStrings"></param>
        /// <remarks></remarks>
        public void RestrictColumnEditsTo(int colid, ArrayList ArrayListOfStrings)
        {
            var reslist = new EditColumnRestrictor();

            string CaretDelimitedString = "";

            foreach (string ar in ArrayListOfStrings)
            {
                //ar = ar.Replace("^", "+");
                ar.Replace("^", "+");
                CaretDelimitedString += ar + "^";
            }

            if (CaretDelimitedString.EndsWith("^"))
                CaretDelimitedString = CaretDelimitedString.Substring(0, CaretDelimitedString.Length - 1);

            reslist.ColumnID = colid;
            reslist.RestrictedList = CaretDelimitedString;

            foreach (EditColumnRestrictor it in _colEditRestrictions)
            {
                if (it.ColumnID == colid)
                    _colEditRestrictions.Remove(it);
            }

            _colEditRestrictions.Add(reslist);
        }
        
        public Array ReturnDelimitedStringAsArray(string StringToParse, int Columns, int rows, string Delimiter)
        {
            // add a delimiter to the begning and the end of the string
            if (!StringToParse.StartsWith(Delimiter))
                StringToParse = Delimiter + StringToParse;
            if (!StringToParse.EndsWith(Delimiter))
                StringToParse = StringToParse + Delimiter;

            var mc = System.Text.RegularExpressions.Regex.Matches(StringToParse, Delimiter);
            int RegExCounter = 0;
            int rowCounter = 0;
            var argarray = new string[rows + 1, Columns];

            while (RegExCounter < mc.Count - 1)
            {
                int intCol = 0;
                // parse string
                argarray[rowCounter, intCol] = StringToParse.Substring(mc[RegExCounter].Index + Delimiter.Length, mc[RegExCounter + 1].Index - mc[RegExCounter].Index - Delimiter.Length);
                RegExCounter += 1;
                intCol += 1;
                while (intCol != Columns)
                {
                    if (mc[RegExCounter].Index + Delimiter.Length != StringToParse.Length)
                        argarray[rowCounter, intCol] = StringToParse.Substring(mc[RegExCounter].Index + Delimiter.Length, mc[RegExCounter + 1].Index - mc[RegExCounter].Index - Delimiter.Length);
                    intCol += 1;
                    RegExCounter += 1;
                }
                rowCounter += 1;
            }
            // clear out some memory
            mc = null;
            return argarray;
        }

        /// <summary>
        /// Will attempt to render the grids surface onto the supplied graphics context GR, The grid will be rendered into the rectangle denoted by
        /// xloc,yloc and width and height
        /// </summary>
        /// <param name="gr"></param>
        /// <param name="xloc"></param>
        /// <param name="yloc"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <remarks></remarks>
        public void PlaceGridOnGraphicsContext(Graphics gr, int xloc, int yloc, int width, int height)
        {
            var cr = new Rectangle(xloc, yloc, width, height);

            RenderGridToGraphicsContext(gr, cr);
        }

        /// <summary>
        /// Walks the list of open tear away columns and sets the to be on top of all windows
        /// </summary>
        /// <remarks></remarks>
        public void PullAllTearAwaysToTheFront()
        {
            if (TearAways.Count == 0)
                return;

            if (_TearAwayWork)
                return;

            _TearAwayWork = true;

            int t;
            TearAwayWindowEntry tear;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.TopMost = true;
                tear.Winform.BringToFront();
            }

            _TearAwayWork = false;
        }

        /// <summary>
        /// Walks the list of tear away columns and sets them to be behind all open windows.
        /// </summary>
        /// <remarks></remarks>
        public void PushAllTearAwaysToTheBack()
        {
            if (TearAways.Count == 0)
                return;


            if (_TearAwayWork)
                return;

            _TearAwayWork = true;


            int t;
            TearAwayWindowEntry tear;
            var loopTo = TearAways.Count - 1;
            for (t = 0; t <= loopTo; t++)
            {
                tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                tear.Winform.TopMost = false;
                tear.Winform.SendToBack();
            }

            _TearAwayWork = false;
        }

        /// <summary>
        /// Selects all rows in the current grid
        /// </summary>
        /// <remarks></remarks>
        public void SelectAllRows()
        {
            var aList = new ArrayList();
            int i;
            var loopTo = Rows - 1;
            for (i = 0; i <= loopTo; i++)
                aList.Add(i);

            SelectedRows = aList;
        }

        /// <summary>
        /// Selects rows in the current grid from an arraylist of row IDs
        /// </summary>
        /// <remarks></remarks>
        public void SelectRows(ArrayList rowIDs)
        {
            try
            {
                SelectedRows = rowIDs;
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.SelectRows Error...");
            }
        }

        /// <summary>
        /// Sets the background color for all cells in the grid to be the supplied color
        /// </summary>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetAllCellBackcolors(Color color)
        {
            int x, y;

            int colentry = GetGridBackColorListEntry(new SolidBrush(color));
            var loopTo = _cols - 1;
            for (x = 0; x <= loopTo; x++)
            {
                var loopTo1 = _rows - 1;
                for (y = 0; y <= loopTo1; y++)
                    _gridBackColor[y, x] = colentry;
            }

            Invalidate();
        }

        /// <summary>
        /// Sets the foreground color for all cells in the grid to be the supplied color
        /// </summary>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetAllCellForecolors(Color color)
        {
            int x, y;

            int colentry = GetGridForeColorListEntry(new Pen(color));
            var loopTo = _cols - 1;
            for (x = 0; x <= loopTo; x++)
            {
                var loopTo1 = _rows - 1;
                for (y = 0; y <= loopTo1; y++)
                    _gridForeColor[y, x] = colentry;
            }

            Invalidate();
        }

        /// <summary>
        /// Sets all the cells in the specific column Col to be the the supplied color in the background
        /// </summary>
        /// <param name="Col"></param>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetColBackColor(int Col, Color color)
        {
            try
            {
                if (Col >= 0 & Col < _cols)
                {
                    int iCol;
                    var loopTo = _rows - 1;
                    for (iCol = 0; iCol <= loopTo; iCol++)
                        _gridBackColor[iCol, Col] = GetGridBackColorListEntry(new SolidBrush(color));
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.SetColBackColor Error...");
            }
        }

        /// <summary>
        /// Sets all the cells in the specific column Col to be the the supplied color in the foreground
        /// </summary>
        /// <param name="Col"></param>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetColForeColor(int Col, Color color)
        {
            try
            {
                if (Col >= 0 & Col < _cols)
                {
                    int iCol;
                    var loopTo = _rows - 1;
                    for (iCol = 0; iCol <= loopTo; iCol++)
                        _gridForeColor[iCol, Col] = GetGridForeColorListEntry(new Pen(color));
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.SetColForeColor Error...");
            }
        }

        /// <summary>
        /// Applied the specified enumerated colorscheme to the grids contents
        /// </summary>
        /// <param name="Scheme"></param>
        /// <remarks></remarks>
        public void SetColorScheme(TaiGridColorSchemes Scheme)
        {
            switch (Scheme)
            {
                case TaiGridColorSchemes._Default:
                    {
                        _GridTitleBackcolor = Color.Blue;
                        _GridTitleForeColor = Color.White;
                        _GridHeaderBackcolor = Color.LightBlue;
                        _GridHeaderForecolor = Color.Black;
                        _CellOutlineColor = Color.Black;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.AntiqueWhite;
                        DefaultForegroundColor = Color.Black;
                        _RowHighLiteBackColor = Color.Blue;
                        _RowHighLiteForeColor = Color.White;
                        _ColHighliteBackColor = Color.MediumSlateBlue;
                        _ColHighliteForeColor = Color.LightGray;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }

                case TaiGridColorSchemes._Technical:
                    {
                        _GridTitleBackcolor = Color.DarkBlue;
                        _GridTitleForeColor = Color.GhostWhite;
                        _GridHeaderBackcolor = Color.LightBlue;
                        _GridHeaderForecolor = Color.Black;
                        _CellOutlineColor = Color.Black;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.LightYellow;
                        DefaultForegroundColor = Color.Black;
                        _RowHighLiteBackColor = Color.LightSlateGray;
                        _RowHighLiteForeColor = Color.Black;
                        _ColHighliteBackColor = Color.MediumSpringGreen;
                        _ColHighliteForeColor = Color.Black;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }

                case TaiGridColorSchemes._Colorful1:
                    {
                        _GridTitleBackcolor = Color.Blue;
                        _GridTitleForeColor = Color.Yellow;
                        _GridHeaderBackcolor = Color.Violet;
                        _GridHeaderForecolor = Color.Yellow;
                        _CellOutlineColor = Color.White;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.MediumPurple;
                        DefaultForegroundColor = Color.Yellow;
                        _RowHighLiteBackColor = Color.Blue;
                        _RowHighLiteForeColor = Color.White;
                        _ColHighliteBackColor = Color.MediumSlateBlue;
                        _ColHighliteForeColor = Color.LightGray;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }

                case TaiGridColorSchemes._Colorful2:
                    {
                        _GridTitleBackcolor = Color.Violet;
                        _GridTitleForeColor = Color.White;
                        _GridHeaderBackcolor = Color.Blue;
                        _GridHeaderForecolor = Color.White;
                        _CellOutlineColor = Color.Black;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.AntiqueWhite;
                        DefaultForegroundColor = Color.Black;
                        _RowHighLiteBackColor = Color.Blue;
                        _RowHighLiteForeColor = Color.White;
                        _ColHighliteBackColor = Color.MediumSlateBlue;
                        _ColHighliteForeColor = Color.LightGray;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }

                case TaiGridColorSchemes._Fancy:
                    {
                        _GridTitleBackcolor = Color.Blue;
                        _GridTitleForeColor = Color.White;
                        _GridHeaderBackcolor = Color.LightBlue;
                        _GridHeaderForecolor = Color.Black;
                        _CellOutlineColor = Color.Black;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.AntiqueWhite;
                        DefaultForegroundColor = Color.Black;
                        _RowHighLiteBackColor = Color.Blue;
                        _RowHighLiteForeColor = Color.White;
                        _ColHighliteBackColor = Color.MediumSlateBlue;
                        _ColHighliteForeColor = Color.LightGray;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }

                default:
                    {
                        _GridTitleBackcolor = Color.Blue;
                        _GridTitleForeColor = Color.White;
                        _GridHeaderBackcolor = Color.LightBlue;
                        _GridHeaderForecolor = Color.Black;
                        _CellOutlineColor = Color.Black;
                        _alternateColorationALTColor = Color.MediumSpringGreen;
                        _alternateColorationBaseColor = Color.AntiqueWhite;
                        DefaultBackgroundColor = Color.AntiqueWhite;
                        DefaultForegroundColor = Color.Black;
                        _RowHighLiteBackColor = Color.Blue;
                        _RowHighLiteForeColor = Color.White;
                        _ColHighliteBackColor = Color.MediumSlateBlue;
                        _ColHighliteForeColor = Color.LightGray;
                        _BorderColor = Color.Black;
                        _excelAlternateRowColor = Color.FromArgb(204, 255, 204);
                        Refresh();
                        break;
                    }
            }
        }

        /// <summary>
        /// Will apply the supplied <c>ItemToSet</c> to the cell currently being edited in the grid
        /// </summary>
        /// <param name="ItemToSet"></param>
        /// <remarks></remarks>
        public void SetEditItemText(string ItemToSet)
        {
            if (_EditMode)
            {
                if (txtInput.Visible)
                    txtInput.Text = ItemToSet;
                else
                    cmboInput.Text = ItemToSet;
            }
        }

        /// <summary>
        /// Attempts to set the cell at <c>row</c> and <c>col</c> to be in edit mode
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <remarks></remarks>
        public void SetEditItem(int row, int col)
        {
            // lets do some sanity checking here
            if (row > -1 & row < _rows & _rowEditable[row])
            {
                // the rows are in range
                if (col > -1 & col < _cols)
                {
                    // the cols are in range
                    // is that column editable
                    if (_colEditable[col])
                    {
                        // aye it be editable

                        Focus();

                        int xoff, yoff, r, c;

                        _RowClicked = row;
                        _ColClicked = col;

                        if (_RowClicked > -1 & _RowClicked < _rows)
                        {
                            if (_ColClicked > -1 & _ColClicked < _cols & _colEditable[_ColClicked] & _AllowInGridEdits)
                            {
                                if (IsColumnRestricted(_ColClicked))
                                {
                                    var it = GetColumnRestriction(_ColClicked);

                                    cmboInput.Items.Clear();

                                    var s = it.RestrictedList.Split("^".ToCharArray());
                                    foreach (string ss in s)
                                        cmboInput.Items.Add(ss);

                                    // we have selected a row and col lets move the txtinput there and bring it to the front
                                    xoff = 0;
                                    yoff = 0;

                                    if (_RowClicked > 0)
                                    {
                                        var loopTo = _RowClicked - 1;
                                        for (r = 0; r <= loopTo; r++)
                                            yoff = yoff + get_RowHeight(r);
                                    }

                                    if (GridheaderVisible)
                                        yoff = yoff + _GridHeaderHeight;

                                    if (_GridTitleVisible)
                                        yoff = yoff + _GridTitleHeight;

                                    if (_ColClicked > 0)
                                    {
                                        var loopTo1 = _ColClicked - 1;
                                        for (c = 0; c <= loopTo1; c++)
                                            xoff = xoff + get_ColWidth(c);
                                    }

                                    if (vs.Visible & vs.Value > 0)
                                        yoff = yoff - GimmeYOffset(vs.Value);

                                    if (hs.Visible & hs.Value > 0)
                                        xoff = xoff - GimmeXOffset(hs.Value);

                                    if (_CellOutlines)
                                    {
                                        cmboInput.Top = yoff + 1;
                                        cmboInput.Left = xoff + 1;
                                        cmboInput.Width = get_ColWidth(_ColClicked) - 1;
                                        cmboInput.Height = get_RowHeight(_RowClicked) - 2;
                                        cmboInput.BackColor = _colEditableTextBackColor;
                                    }
                                    else
                                    {
                                        cmboInput.Top = yoff;
                                        cmboInput.Left = xoff;
                                        cmboInput.Width = get_ColWidth(_ColClicked);
                                        cmboInput.Height = get_RowHeight(_RowClicked);
                                        cmboInput.BackColor = _colEditableTextBackColor;
                                    }

                                    cmboInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                                    cmboInput.Text = _grid[_RowClicked, _ColClicked];

                                    cmboInput.Visible = true;
                                    cmboInput.BringToFront();
                                    cmboInput.DroppedDown = true;
                                    _EditModeCol = _ColClicked;
                                    _EditModeRow = _RowClicked;
                                    _EditMode = true;

                                    cmboInput.Focus();
                                }
                                else
                                {
                                    // we have selected a row and col lets move the txtinput there and bring it to the front
                                    xoff = 0;
                                    yoff = 0;

                                    if (_RowClicked > 0)
                                    {
                                        var loopTo2 = _RowClicked - 1;
                                        for (r = 0; r <= loopTo2; r++)
                                            yoff = yoff + get_RowHeight(r);
                                    }

                                    if (GridheaderVisible)
                                        yoff = yoff + _GridHeaderHeight;

                                    if (_GridTitleVisible)
                                        yoff = yoff + _GridTitleHeight;

                                    if (_ColClicked > 0)
                                    {
                                        var loopTo3 = _ColClicked - 1;
                                        for (c = 0; c <= loopTo3; c++)
                                            xoff = xoff + get_ColWidth(c);
                                    }

                                    if (vs.Visible & vs.Value > 0)
                                        yoff = yoff - GimmeYOffset(vs.Value);

                                    if (hs.Visible & hs.Value > 0)
                                        xoff = xoff - GimmeXOffset(hs.Value);

                                    if (_CellOutlines)
                                    {
                                        txtInput.Top = yoff + 1;
                                        txtInput.Left = xoff + 1;
                                        txtInput.Width = get_ColWidth(_ColClicked) - 1;
                                        txtInput.Height = get_RowHeight(_RowClicked) - 2;
                                        txtInput.BackColor = _colEditableTextBackColor;
                                    }
                                    else
                                    {
                                        txtInput.Top = yoff;
                                        txtInput.Left = xoff;
                                        txtInput.Width = get_ColWidth(_ColClicked);
                                        txtInput.Height = get_RowHeight(_RowClicked);
                                        txtInput.BackColor = _colEditableTextBackColor;
                                    }

                                    txtInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                                    txtInput.Text = _grid[_RowClicked, _ColClicked];

                                    txtInput.Visible = true;
                                    txtInput.BringToFront();
                                    _EditModeCol = _ColClicked;
                                    _EditModeRow = _RowClicked;
                                    _EditMode = true;

                                    txtInput.Focus();
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Sets the row at <c>row</c> to be the corresponging <c>color</c> background color
        /// </summary>
        /// <param name="row"></param>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetRowBackColor(int row, Color color)
        {
            try
            {
                if (row >= 0 & row < _rows)
                {
                    int iCol;
                    var loopTo = Cols;
                    for (iCol = 0; iCol <= loopTo; iCol++)
                        _gridBackColor[row, iCol] = GetGridBackColorListEntry(new SolidBrush(color));
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.SetRowBackColor Error...");
            }
        }

        /// <summary>
        /// Sets the row at <c>row</c> to have the corresponding <c>color</c> foreground color
        /// </summary>
        /// <param name="row"></param>
        /// <param name="color"></param>
        /// <remarks></remarks>
        public void SetRowForeColor(int row, Color color)
        {
            try
            {
                if (row >= 0 & row < _rows)
                {
                    int iCol;
                    var loopTo = Cols;
                    for (iCol = 0; iCol <= loopTo; iCol++)
                        _gridForeColor[row, iCol] = GetGridForeColorListEntry(new Pen(color));
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.SetRowForeColor Error...");
            }
        }

        /// <summary>
        /// Will attempt to sort the contents of the current grid on <c>col</c>
        /// If <c>Descending</c> is true or false will dictate the order of the sort
        /// </summary>
        /// <param name="col"></param>
        /// <param name="Descending"></param>
        /// <remarks></remarks>
        public void SortGridOnColumn(int col, bool Descending)
        {
            if (col < 0 | col > _cols - 1)
                return;

            int oldcol = _ColOverOnMenuButton;

            _ColOverOnMenuButton = col;

            if (Descending)
                miSortDescending_Click(this, new EventArgs());
            else
                miSortAscending_Click(this, new EventArgs());

            _ColOverOnMenuButton = oldcol;
        }

        /// <summary>
        /// Will atempt to sort the grids contents on <c>col</c> treating the column contents as dates.
        /// The <c>Descending</c> parameter will distate the order of the sort
        /// </summary>
        /// <param name="col"></param>
        /// <param name="Descending"></param>
        /// <remarks></remarks>
        public void SortGridOnColumnDate(int col, bool Descending)
        {
            if (col < 0 | col > _cols - 1)
                return;

            int oldcol = _ColOverOnMenuButton;

            _ColOverOnMenuButton = col;

            if (Descending)
                miDateDesc_Click(this, new EventArgs());
            else
                miDateAsc_Click(this, new EventArgs());

            _ColOverOnMenuButton = oldcol;
        }

        /// <summary>
        /// Will atempt to sort the grids contents on <c>col</c> treating the column contents as numbers.
        /// The <c>Descending</c> parameter will distate the order of the sort
        /// </summary>
        /// <param name="col"></param>
        /// <param name="Descending"></param>
        /// <remarks></remarks>
        public void SortGridOnColumnNumeric(int col, bool Descending)
        {
            if (col < 0 | col > _cols - 1)
                return;

            int oldcol = _ColOverOnMenuButton;

            _ColOverOnMenuButton = col;

            if (Descending)
                miSortNumericDesc_Click(this, new EventArgs());
            else
                miSortNumericAsc_Click(this, new EventArgs());

            _ColOverOnMenuButton = oldcol;
        }

        /// <summary>
        /// Will attempt to add all the values in a column denoted by <c>colnum</c> and return
        /// the sum as a double
        /// </summary>
        /// <param name="colnum"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double SumUpColumn(int colnum)
        {
            int t;
            string a;
            double result = 0.0;

            if (colnum >= _cols | _rows < 1)
                return result;

            try
            {
                var loopTo = _rows - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    a = _grid[t, colnum].Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

                    result = result + Conversion.Val(a);
                }
            }
            catch (Exception ex)
            {
            }

            return result;
        }

        /// <summary>
        /// Manually sets the position of the verticle scrollbar of the grid if the contents are larger
        /// than the physical grid window.
        /// </summary>
        /// <param name="sb"></param>
        /// <remarks></remarks>
        public void SetVertScrollbarPosition(int sb)
        {
            if (vs.Visible)
            {
                if (sb <= vs.Maximum & sb >= vs.Minimum)
                    vs.Value = sb;
            }
        }

        /// <summary>
        /// Will attempt to wrap the text data in a specified <c>col</c> at <c>wraplen</c> length.
        /// The wrap is smat in that it tries to wrap on whitespace boundaries
        /// </summary>
        /// <param name="col">The specific column that you want to set the <c>wraplen</c> on</param>
        /// <param name="wraplen">sets the column <c>col</c> to workwrap on <c>wraplen</c> characters</param>
        /// <remarks></remarks>
        public void WordWrapColumn(int col, int wraplen)
        {
            int t;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
                _grid[t, col] = SplitLongString(_grid[t, col], wraplen);

            AutoSizeCellsToContents = true;

            Invalidate();
        }

        /// <summary>
        /// Will return the computed STDEV of all the numbers contained in the column denoted by <c>colid</c>
        /// </summary>
        /// <param name="colid"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double GetColumnSTDEV(int colid)
        {
            var arl = GetColAsCleanedArrayList(colid);
            double avg = 0;
            double res = 0;

            int t;

            if (arl.Count == 0)
            {
                return res;
                return default(double);
            }

            var loopTo = arl.Count - 1;
            for (t = 0; t <= loopTo; t++)
                avg += Convert.ToDouble(arl[t]);

            // now we can get the average

            avg = avg / arl.Count;
            var loopTo1 = arl.Count - 1;

            // now to subtract the aaverage from each element in the array and square it
            // giving us our squared deviations

            for (t = 0; t <= loopTo1; t++)
                arl[t] = Math.Pow(Convert.ToDouble(Convert.ToDouble(arl[t]) - avg), 2);

            // now lets get the sum of the squared deviations

            double sum = 0;
            var loopTo2 = arl.Count - 1;
            for (t = 0; t <= loopTo2; t++)
                sum += Convert.ToDouble(arl[t]);

            // Finally lets get the square root of the (sum / number in the set-1)

            res = Math.Sqrt(sum / (arl.Count - 1));

            return res;
        }

        /// <summary>
        /// Will return the computed STDEVP of all the numbers contained in the column denoted by <c>colid</c>
        /// </summary>
        /// <param name="colid"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double GetColumnSTDEVP(int colid)
        {
            var arl = GetColAsCleanedArrayList(colid);
            double avg = 0;
            double res = 0;

            int t;

            if (arl.Count == 0)
            {
                return res;
                return default(double);
            }

            var loopTo = arl.Count - 1;
            for (t = 0; t <= loopTo; t++)
                avg += Convert.ToDouble(arl[t]);

            // now we can get the average

            avg = avg / arl.Count;
            var loopTo1 = arl.Count - 1;

            // now to subtract the aaverage from each element in the array and square it
            // giving us our squared deviations

            for (t = 0; t <= loopTo1; t++)
                arl[t] = Math.Pow(Convert.ToDouble(Convert.ToDouble(arl[t]) - avg), 2);

            // now lets get the sum of the squared deviations

            double sum = 0;
            var loopTo2 = arl.Count - 1;
            for (t = 0; t <= loopTo2; t++)
                sum += Convert.ToDouble(arl[t]);

            // Finally lets get the square root of the (sum / number in the set)

            res = Math.Sqrt(sum / arl.Count);

            return res;
        }

        /// <summary>
        /// Will calculate the fuzzy membership of the values at <c>colid</c> beyond <c>targetval</c> from the direction of <c>outlier</c>
        /// Values will be between 0 and 1 where beyond <c>targetval</c> is 1 and between <c>outlier</c> and target are some portion
        /// of 0 to 1. Will use a Liner function between <c>outlier</c> and <c>targetval</c>
        /// </summary>
        /// <param name="colid"></param>
        /// <param name="targetval"></param>
        /// <param name="outlier"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public ArrayList FuzzyColumnMembership(int colid, double targetval, double outlier)
        {
            var arl = new ArrayList();
            int t, y;
            double tv, tv2;
            string tstr;


            bool lessthan = true;

            if (outlier <= targetval)
                lessthan = true;
            else
                lessthan = false;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
                arl.Add(Convert.ToDouble("0.0"));
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                if (!Information.IsNothing(_grid[y, colid]))
                {
                    // get the string contained at that grids cell coordinates stripped of some money crap
                    tstr = _grid[y, colid].Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "");
                    if (Information.IsNumeric(tstr))
                    {
                        // is it still a number? Yes it is 
                        tv = Convert.ToDouble(tstr);

                        // Ok we got the number whats the direction of the check

                        if (lessthan)
                        {
                            // we are checking up
                            if (tv >= targetval)
                                arl[y] = Convert.ToDouble("1");
                            else if (tv >= outlier)
                            {
                                // here we compute the weight

                                tv2 = targetval - outlier; // range of values

                                arl[y] = 1 / (tv2 / (tv - outlier));
                            }
                        }
                        else
                            // we are checking down
                            if (tv <= targetval)
                            arl[y] = Convert.ToDouble("1");
                        else if (tv <= outlier)
                        {
                            // here we compute the weight

                            tv2 = outlier - targetval; // range of values

                            arl[y] = 1 / (tv2 / (outlier - tv));
                        }
                    }
                }
            }

            return arl;
        }

        public ArrayList FuzzyColumnCombine(ArrayList colids)
        {
            var arl = new ArrayList();
            bool bail = false;
            int t, x;
            double sum;
            string s;
            var loopTo = _rows - 1;

            // Initialize our resultset
            for (t = 0; t <= loopTo; t++)
                arl.Add(Convert.ToDouble("0"));
            var loopTo1 = colids.Count - 1;

            // ensure all the ccols we are asking for membership values in are actually in the grid
            for (t = 0; t <= loopTo1; t++)
            {
                x = Convert.ToInt32(colids[t]);
                if (x < 0 | x > _cols - 1)
                {
                    // tyhis one is not there so lets setup to bail out of this 
                    bail = true;
                    break;
                }
            }

            if (bail)
            {
                // time to bail out just return the all 0 membership set we crafted at he start
                return arl;
                return null;
            }

            var loopTo2 = _rows - 1;

            // conditions are right for a test so lets check it out

            for (t = 0; t <= loopTo2; t++)
            {
                sum = 0;
                var loopTo3 = colids.Count - 1;
                for (x = 0; x <= loopTo3; x++)
                {
                    s = _grid[t, Convert.ToInt32(colids[x])];

                    s = s.Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "");

                    sum += Convert.ToDouble(s);
                }

                arl[t] = sum / colids.Count;
            }

            return arl;
        }

        public void InsertFuzzyMathResultSet(ArrayList arl, string ColName)
        {
            if (arl.Count != _rows)
                return;

            Cols += 1;

            set_HeaderLabel(_cols - 1, ColName);

            int y;
            var loopTo = arl.Count - 1;
            for (y = 0; y <= loopTo; y++)
                set_item(y, _cols - 1, Convert.ToDouble(arl[y]).ToString());

            AutoSizeCellsToContents = true;
            Refresh();
        }

        /// <summary>
        /// Will take the values at the specified <c>row</c> and starting at the specified <c>col</c> to the
        /// last column in the existing grid and add them up returning the result as a double.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RollupColumn(int row, int col)
        {
            double result = 0.0;
            int t;
            string a;

            // do some bounds checking
            if (row > _rows - 1)
            {
                return result;
                return default(double);
            }

            if (col > _cols - 1)
            {
                return result;
                return default(double);
            }

            var loopTo = _cols - 1;
            for (t = col; t <= loopTo; t++)
            {
                a = get_item(row, t).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                result = result + Conversion.Val(a);
            }

            return result;
        }

        /// <summary>
        /// will take the values at the specified <c>row</c> and <c>col</c> continuing to the edge of the grid
        /// and will add them up returning the result as a double
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RollupCube(int row, int col)
        {
            double result = 0.0;
            int t, tt;
            string a;

            // do some bounds checking
            if (row > _rows - 1)
            {
                return result;
                return default(double);
            }

            if (col > _cols - 1)
            {
                return result;
                return default(double);
            }

            var loopTo = _cols - 1;
            for (t = col; t <= loopTo; t++)
            {
                var loopTo1 = _rows - 1;
                for (tt = row; tt <= loopTo1; tt++)
                {
                    a = get_item(tt, t).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                    result = result + Conversion.Val(a);
                }
            }

            return result;
        }

        /// <summary>
        /// will take the values in a specified <c>col</c> and will take all the valies from the specfied <c>row</c>
        /// until the last row in the grid and will add them up returning the result as a double
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RollupRow(int row, int col)
        {
            double result = 0.0;
            int t;
            string a;

            // do some bounds checking
            if (row > _rows - 1)
            {
                return result;
                return default(double);
            }

            if (col > _cols - 1)
            {
                return result;
                return default(double);
            }

            var loopTo = _rows - 1;
            for (t = row; t <= loopTo; t++)
            {
                a = get_item(t, col).Replace("$", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                result = result + Conversion.Val(a);
            }

            return result;
        }

        /// <summary>
        /// Will populate the grid from the supplied <c>sFilename</c>
        /// The call will assume that the caller wants the first set of items from the supplied xml file
        /// </summary>
        /// <param name="sFilename"></param>
        /// <remarks></remarks>
        public void ImportFromXML(string sFilename)
        {
            try
            {
                var _ds = new DataSet();

                _ds.ReadXml(sFilename, XmlReadMode.Auto);

                // determine how many rows and columns
                int iRows = _ds.Tables[0].Rows.Count;
                int iCols = _ds.Tables[0].Columns.Count;

                Rows = iRows;
                Cols = iCols;

                // fill in the column names
                DataRow row;
                int iCol = 0;
                int iRow = 0;
                var loopTo = iCols - 1;
                for (iCol = 0; iCol <= loopTo; iCol++)
                    _GridHeader[iCol] = _ds.Tables[0].Columns[iCol].ColumnName;
                var loopTo1 = iRows - 1;
                for (iRow = 0; iRow <= loopTo1; iRow++)
                {
                    row = _ds.Tables[0].Rows[iRow];
                    var loopTo2 = iCols - 1;
                    for (iCol = 0; iCol <= loopTo2; iCol++)
                    {
                        if (!Information.IsDBNull(row[iCol]))
                            _grid[iRow, iCol] = Conversions.ToString(row[iCol]);
                        else
                            _grid[iRow, iCol] = "{NULL}";
                    }
                }

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ImportFromXML Error...");
            }
        }

        /// <summary>
        /// Will populate the grid from the supplied <c>sFilename</c>
        /// The call will attempt to get the data from the designated <c>tblnum</c> in the supplied xml file
        /// </summary>
        /// <param name="sFilename"></param>
        /// <param name="tblnum"></param>
        /// <remarks></remarks>
        public void ImportFromXML(string sFilename, int tblnum)
        {
            try
            {
                var _ds = new DataSet();

                _ds.ReadXml(sFilename, XmlReadMode.Auto);

                if (_ds.Tables.Count - 1 < tblnum)
                    tblnum = _ds.Tables.Count - 1;

                // determine how many rows and columns
                int iRows = _ds.Tables[tblnum].Rows.Count;
                int iCols = _ds.Tables[tblnum].Columns.Count;

                Rows = iRows;
                Cols = iCols;

                // fill in the column names
                DataRow row;
                int iCol = 0;
                int iRow = 0;
                var loopTo = iCols - 1;
                for (iCol = 0; iCol <= loopTo; iCol++)
                    _GridHeader[iCol] = _ds.Tables[tblnum].Columns[iCol].ColumnName;
                var loopTo1 = iRows - 1;
                for (iRow = 0; iRow <= loopTo1; iRow++)
                {
                    row = _ds.Tables[tblnum].Rows[iRow];
                    var loopTo2 = iCols - 1;
                    for (iCol = 0; iCol <= loopTo2; iCol++)
                    {
                        if (!Information.IsDBNull(row[iCol]))
                            _grid[iRow, iCol] = Conversions.ToString(row[iCol]);
                        else
                            _grid[iRow, iCol] = "{NULL}";
                    }
                }

                AutoSizeCellsToContents = true;
                _colEditRestrictions.Clear();
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ImportFromXML Error...");
            }
        }

        /// <summary>
        /// Will attempt to print the contents of the grid to the default printer in the system
        /// The grids own properties for Outlining printed cells, Printing page numbers,
        /// Previewing the output first, Page orientation will be employed in the resulting process
        /// </summary>
        /// <remarks></remarks>
        public void PrintTheGrid()
        {
            PrintTheGrid("", _gridReportMatchColors, _gridReportOutlineCells, _gridReportNumberPages, _gridReportPreviewFirst, _gridReportOrientLandscape);
        }

        /// <summary>
        /// Will attempt to print the contents of the grid to the default printer in the system
        /// The grids own properties for Outlining printed cells, Printing page numbers,
        /// Previewing the output first, Page orientation will be employed in the resulting process
        /// The supplied <c>Title</c> will be used to lable the pages of output
        /// </summary>
        /// <param name="Title"></param>
        /// <remarks></remarks>
        public void PrintTheGrid(string Title)
        {
            PrintTheGrid(Title, _gridReportMatchColors, _gridReportOutlineCells, _gridReportNumberPages, _gridReportPreviewFirst, _gridReportOrientLandscape);
        }

        /// <summary>
        /// Will attempt to print the contents of the grid to the default printer in the system
        /// The grids own properties for Outlining printed cells, Previewing the output first,
        /// The supplied <c>Title</c> will be used to lable the pages of output as well as the
        /// supplied values for <c>NumberPages</c> and <c>Landscapemode</c> will override those setup
        /// in the grid properties
        /// </summary>
        /// <param name="Title"></param>
        /// <param name="NumberPages"></param>
        /// <param name="Landscapemode"></param>
        /// <remarks></remarks>
        public void PrintTheGrid(string Title, bool NumberPages, bool Landscapemode)
        {
            PrintTheGrid(Title, _gridReportMatchColors, _gridReportOutlineCells, NumberPages, _gridReportPreviewFirst, Landscapemode);
        }

        /// <summary>
        /// Will attempt to print the contents of the grid to the default printer in the system
        /// using supplied values for
        /// <list type="Bullet">
        /// <item> <c>Title</c> will use thee supplied strin g to title the resulting output</item>
        /// <item> <c>MatchColors</c> attempting to match the colors on the grid with printed output</item>
        /// <item> <c>OutlineCells</c> will draw an outline around each cell of output on the printed page</item>
        /// <item> <c>NumberPages</c> will number each page as its printed</item>
        /// <item> <c>PreviewFirst</c> will show the print preview windows forst before sending the results to the printer</item>
        /// <item> <c>Landscapemode</c> will dictate that the resulting output be in landscape mode</item>
        /// </list>
        /// </summary>
        /// <param name="Title"></param>
        /// <param name="MatchColors"></param>
        /// <param name="OutlineCells"></param>
        /// <param name="NumberPages"></param>
        /// <param name="PreviewFirst"></param>
        /// <param name="LandscapeMode"></param>
        /// <remarks></remarks>
        public void PrintTheGrid(string Title, bool MatchColors, bool OutlineCells, bool NumberPages, bool PreviewFirst, bool LandscapeMode)
        {
            if (_psets == null)
                _psets = new System.Drawing.Printing.PageSettings();

            _gridReportMatchColors = MatchColors;
            _gridReportNumberPages = NumberPages;
            _gridReportOutlineCells = OutlineCells;
            _gridReportPreviewFirst = PreviewFirst;
            _gridReportOrientLandscape = LandscapeMode;
            _gridReportTitle = Title;

            try
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(_psets.PrinterSettings.PrintRange, System.Drawing.Printing.PrintRange.AllPages, false)))
                {
                    _gridReportPageNumbers = 1;
                    _gridReportCurrentrow = 0;
                    _gridReportCurrentColumn = 0;
                    _gridReportPrintedOn = DateAndTime.Now;
                }
                else
                {
                    CalculatePageRange();
                    _gridReportPageNumbers = _gridStartPage;
                    _gridReportCurrentrow = _gridStartPageRow;
                    _gridReportCurrentColumn = 0;
                    _gridReportPrintedOn = DateAndTime.Now;
                }

                if (LandscapeMode)
                    _psets.Landscape = true;
                else
                    _psets.Landscape = false;

                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(_psets.PrinterSettings.PrinterName, _OriginalPrinterName, false)))
                {
                    // We changed the printer invoke the Set default printer via the System.Management class

                    System.Management.ManagementObjectCollection moReturn;

                    System.Management.ManagementObjectSearcher moSearch;

                    //System.Management.ManagementObject mo;

                    moSearch = new System.Management.ManagementObjectSearcher("Select * from Win32_Printer");

                    moReturn = moSearch.Get();

                    foreach (System.Management.ManagementObject mo in moReturn)
                    {
                        object[] objReturn = new object[10];
                        Console.WriteLine(mo["Name"]);
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(mo["Name"], _psets.PrinterSettings.PrinterName, false)))
                            mo.InvokeMethod("SetDefaultPrinter", objReturn);
                    }
                }


                // pdoc.DefaultPageSettings.Landscape = LandscapeMode
                pdoc.DefaultPageSettings = _psets;

                if (_gridReportPreviewFirst)
                {
                    var pview = new PrintPreviewDialog();
                    pview.Document = pdoc;
                    pview.WindowState = FormWindowState.Maximized;
                    pview.ShowDialog();
                }
                else
                    pdoc.Print();
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.Message, MsgBoxStyle.Information, "Print Error");
            }
            finally
            {
                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(_psets.PrinterSettings.PrinterName, _OriginalPrinterName, false)))
                {
                    // We changed the printer invoke the Set default printer via the System.Management class

                    System.Management.ManagementObjectCollection moReturn;

                    System.Management.ManagementObjectSearcher moSearch;

                    //System.Management.ManagementObject mo;

                    moSearch = new System.Management.ManagementObjectSearcher("Select * from Win32_Printer");

                    moReturn = moSearch.Get();

                    foreach (System.Management.ManagementObject mo in moReturn)
                    {
                        object[] objReturn = new object[10];
                        // Console.WriteLine(mo("Name"))
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(mo["Name"], _OriginalPrinterName, false)))
                            mo.InvokeMethod("SetDefaultPrinter", objReturn);
                    }
                }
            }
        }

        /// <summary>
        /// Instructs the grid to stop is continuous redrawing
        /// Can be used to speed up population oiperations that are being performed manually
        /// </summary>
        /// <remarks></remarks>
        public void SuspendGridPaintOperations()
        {
            _Painting = true;
        }

        /// <summary>
        /// Will resume the grid automatic redrawing operations
        /// </summary>
        /// <remarks></remarks>
        public void ResumeGridPaintOperations()
        {
            _Painting = false;
            Refresh();
        }

        #region Private Methods
        private int AllColWidths()
        {
            int t;
            int res = 0;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
                res = res + _colwidths[t];

            return res + 1;
        }

        private int AllRowHeights()
        {
            int t;
            int res = 0;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
                res = res + _rowheights[t];

            if (TitleVisible)
                res = res + TitleFont.Height;

            return res + 1;
        }

        private int CalculatePageRange()
        {

            // Dim psets As New System.Drawing.printing.PageSettings

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(_psets.PrinterSettings.PrintRange, System.Drawing.Printing.PrintRange.SomePages, false)))
            {
                _gridPrintingAllPages = false;
                _gridStartPage = Conversions.ToInteger(_psets.PrinterSettings.FromPage);
                _gridEndPage = Conversions.ToInteger(_psets.PrinterSettings.ToPage);

                var ppea = new System.Drawing.Printing.PrintPageEventArgs(CreateGraphics(), new Rectangle(Conversions.ToInteger(_psets.Margins.Left), Conversions.ToInteger(_psets.Margins.Top), Conversions.ToInteger(_psets.Margins.Right - _psets.Margins.Left), Conversions.ToInteger(_psets.Margins.Bottom - _psets.Margins.Top)), (Rectangle)_psets.Bounds, _psets);

                _gridReportPageNumbers = 1;
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;

                Fake_PrintPage(this, ppea);

                while (ppea.HasMorePages)
                    Fake_PrintPage(this, ppea);

                int maxpage = _gridReportPageNumbers;

                _gridReportPageNumbers = 1;
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;
                return maxpage;
            }
            else
            {
                _gridPrintingAllPages = true;
                _gridStartPage = 1;
                _gridStartPageRow = -1;

                var ppea = new System.Drawing.Printing.PrintPageEventArgs(CreateGraphics(), new Rectangle(Conversions.ToInteger(_psets.Margins.Left), Conversions.ToInteger(_psets.Margins.Top), Conversions.ToInteger(_psets.Margins.Right - _psets.Margins.Left), Conversions.ToInteger(_psets.Margins.Bottom - _psets.Margins.Top)), (Rectangle)_psets.Bounds, _psets);

                _gridReportPageNumbers = 1;
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;

                Fake_PrintPage(this, ppea);

                while (ppea.HasMorePages)
                    Fake_PrintPage(this, ppea);

                int maxpage = _gridReportPageNumbers;

                _gridReportPageNumbers = 1;
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;

                _gridEndPage = maxpage;

                return maxpage;
            }
        }

        private void CheckGridTearAways(int colid)
        {
            // 
            // to be use in methods that affect a specific grid column
            // like RemoveColFromGrid
            // 

            // dont bother unless we actualy have some to act on 
            if (TearAways.Count == 0)
                return;

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];

                if (ta.ColID == colid)
                {
                    // we are showing the column that needs to get the boot
                    ta.KillTearAway();
                    TearAways.RemoveAt(t);
                }
            }

            if (TearAways.Count > 0)
            {
                var loopTo = TearAways.Count - 1;
                // we still have some so lets look to see if any colids were greater than 
                // the intended colid for deletion if so we need to decrement them by one
                for (t = 0; t <= loopTo; t++)
                {
                    if (((TearAwayWindowEntry)TearAways[t]).ColID > colid)
                        ((TearAwayWindowEntry)TearAways[t]).ColID -= 1;
                }
            }
        }

        private string CleanMoneyString(string s)
        {
            return s.Replace("$", "").Replace("(", "").Replace(")", "").Replace(",", "");
        }

        private void ClearToBackgroundColor()
        {
            var gr = CreateGraphics();
            gr.FillRectangle(new SolidBrush(BackColor), gr.ClipBounds);
        }

        private void ClearToBackgroundColor(Graphics gr)
        {
            gr.FillRectangle(new SolidBrush(BackColor), gr.ClipBounds);
        }

        private void DoAutoSizeCheck(Graphics gr)
        {
            int r;
            int c;
            int rr = 0;
            int cc = 0;
            string t;

            int rrr = 0;
            var sz = new SizeF(0, 0);

            if (!_AutoSizeSemaphore | _AutoSizeAlreadyCalculated)
                return;

            if (_AutosizeCellsToContents)
            {
                _AutoSizeSemaphore = false;
                var loopTo = _rows - 1;
                for (r = 0; r <= loopTo; r++)
                    _rowheights[r] = 0;
                var loopTo1 = _cols - 1;
                for (c = 0; c <= loopTo1; c++)
                {
                    t = " " + _GridHeader[c] + " ";
                    cc = Conversions.ToInteger(gr.MeasureString(t, _GridHeaderFont).Width);
                    if (cc > rr)
                        rr = cc;
                    var loopTo2 = _rows - 1;
                    for (r = 0; r <= loopTo2; r++)
                    {
                        if (!string.IsNullOrEmpty(_colPasswords[c]))
                            t = " " + _colPasswords[c] + " ";
                        else if (_grid[r, c] == null)
                            t = "  ";
                        else
                            t = " " + _grid[r, c] + " ";

                        if (_colMaxCharacters[c] != 0)
                        {
                            if (t.Length > _colMaxCharacters[c])
                                t = t.Substring(0, _colMaxCharacters[c]);
                        }

                        if (!_AllowWhiteSpaceInCells)
                        {
                            t = t.Replace(Constants.vbCr, " ").Replace(Constants.vbLf, " ").Replace(Constants.vbTab, " ").Replace(Constants.vbFormFeed, " ");

                            while ((t.Replace("  ", " ") ?? "") != (t ?? ""))
                                t = t.Replace("  ", " ");
                        }

                        sz = gr.MeasureString(t, _gridCellFontsList[_gridCellFonts[r, c]]);

                        cc = Conversions.ToInteger(sz.Width);

                        if (cc > rr)
                            rr = cc;

                        if (_rowheights[r] < sz.Height)
                            _rowheights[r] = Conversions.ToInteger(sz.Height);
                    }

                    _colwidths[c] = rr;
                    rr = 0;
                }

                cc = Conversions.ToInteger(gr.MeasureString("Yy", _GridHeaderFont).Height);

                _GridHeaderHeight = cc;

                cc = Conversions.ToInteger(gr.MeasureString("Yy", _GridTitleFont).Height);

                _GridTitleHeight = cc;

                // For r = 0 To _rows - 1
                // For c = 0 To _cols - 1
                // t = " " & _grid(r, c) & " "
                // cc = gr.MeasureString(t, _gridCellFontsList(_gridCellFonts(r, c))).Height
                // If cc > rr Then
                // rr = cc
                // End If
                // Next

                // _rowheights(r) = rr
                // rr = 0

                // Next

                _AutoSizeSemaphore = true;

                _AutoSizeAlreadyCalculated = true;
            }
            else
            {
            }
        }

        private void Fake_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int x, y, xx, r, c;
            var fnt = new Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Regular, GraphicsUnit.Pixel);
            var fnt2 = new Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Bold, GraphicsUnit.Pixel);

            float m;

            Font ft;

            var greypen = new Pen(Color.Gray);

            int pagewidth = e.PageSettings.Bounds.Size.Width;
            int pageheight = e.PageSettings.Bounds.Size.Height;

            int lrmargin = 40;
            int tbmargin = 70;

            bool colprintedonpage = false;

            if (AllColWidths() * _gridReportScaleFactor < pagewidth - 2 * lrmargin)
                xx = Conversions.ToInteger((pagewidth - 2 * lrmargin - AllColWidths() * _gridReportScaleFactor) / 2);
            else
                xx = 0;


            var rect = new RectangleF(0, 0, 1, 1);

            x = lrmargin;
            y = tbmargin;

            int coloffset = 0;
            bool morecols = true;
            int currow = _gridReportCurrentrow;

            if ((int)e.PageSettings.PrinterSettings.PrintRange == (int)System.Drawing.Printing.PrintRange.SomePages)
            {
                // we may be printing just a range so lets see if we can calculate the row to start printing on
                if (_gridStartPageRow <= 0)
                {
                    // we have not set this up yet so lets check the bounds
                    if (_gridReportPageNumbers >= _gridStartPage)
                        _gridStartPageRow = currow;
                }
            }

            ft = _GridHeaderFont;

            ft = new Font(_GridHeaderFont.FontFamily, (_GridHeaderFont.SizeInPoints - 1) * _gridReportScaleFactor, _GridHeaderFont.Style, _GridHeaderFont.Unit);

            // calculate size and place te printed on date on the page

            m = e.Graphics.MeasureString(_gridReportPrintedOn.ToLongDateString() + Constants.vbCrLf
                                        + _gridReportPrintedOn.ToLongTimeString(), fnt).Width;

            if (_gridReportNumberPages)
                // we want to number the pages here

                m = e.Graphics.MeasureString("Page " + _gridReportPageNumbers.ToString(), fnt).Height;
            var loopTo = Cols - 1;

            // print the grid header

            for (c = _gridReportCurrentColumn; c <= loopTo; c++)
            {
                if (x + _colwidths[c] + xx > pagewidth - lrmargin & colprintedonpage)
                    break;

                colprintedonpage = true;

                rect.X = Convert.ToSingle(x + xx);
                rect.Y = Convert.ToSingle(y);
                rect.Width = Convert.ToSingle(_colwidths[c]);
                rect.Height = Convert.ToSingle(_GridHeaderHeight);

                x = x + _colwidths[c];
            }


            y += _GridHeaderHeight;
            x = lrmargin;
            var loopTo1 = Rows - 1;
            for (r = _gridReportCurrentrow; r <= loopTo1; r++)
            {
                var loopTo2 = Cols - 1;
                for (c = _gridReportCurrentColumn; c <= loopTo2; c++)
                {
                    if (x + _colwidths[c] + xx > pagewidth - lrmargin & colprintedonpage)
                    {
                        coloffset = c;
                        morecols = true;
                        break;
                    }
                    else
                        morecols = false;

                    colprintedonpage = true;

                    rect.X = Convert.ToSingle(x + xx);
                    rect.Y = Convert.ToSingle(y);
                    rect.Width = Convert.ToSingle(_colwidths[c]);
                    rect.Height = Convert.ToSingle(_rowheights[r]);

                    // ft = New Font(_gridCellFontsList(_gridCellFonts(r, c)).FontFamily, _
                    // _gridCellFontsList(_gridCellFonts(r, c)).SizeInPoints - 1, _
                    // _gridCellFontsList(_gridCellFonts(r, c)).Style, _
                    // _gridCellFontsList(_gridCellFonts(r, c)).Unit)

                    // e.Graphics.DrawString(_grid(r, c), ft, _
                    // Brushes.Black, rect, _gridCellAlignmentList(_gridCellAlignment(r, c)))

                    x = x + _colwidths[c];
                }
                x = lrmargin;
                y += _rowheights[r];
                _gridReportCurrentrow += 1;

                // do we need to skip to next page here
                if (y >= pageheight - tbmargin)
                    break;
                else
                {
                }

                Application.DoEvents();
            }

            if (_gridReportCurrentrow >= Rows - 1 & !morecols)
            {
                e.HasMorePages = false;
                // _gridReportPageNumbers = 1
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;
            }
            else
            {
                if (morecols)
                {
                    _gridReportCurrentColumn = coloffset;
                    _gridReportCurrentrow = currow;
                }
                else
                    _gridReportCurrentColumn = 0;
                e.HasMorePages = true;
                _gridReportPageNumbers += 1;
            }
        }

        private int GetGridBackColorListEntry(Brush bcol)
        {
            int t;
            int flag = -1;

            SolidBrush bbcol;
            SolidBrush aacol;
            var loopTo = _gridBackColorList.GetUpperBound(0);
            for (t = 0; t <= loopTo; t++)
            {
                if (_gridBackColorList[t] == null)
                {
                }
                else
                {
                    bbcol = (SolidBrush)_gridBackColorList[t];
                    aacol = (SolidBrush)bcol;

                    if (aacol.Color.A == bbcol.Color.A)
                    {
                        if (aacol.Color.R == bbcol.Color.R)
                        {
                            if (aacol.Color.G == bbcol.Color.G)
                            {
                                if (aacol.Color.B == bbcol.Color.B)
                                {
                                    flag = t;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (flag == -1)
            {
                // we dont have that fnt2find in the list so we need to add it
                t = _gridBackColorList.GetUpperBound(0);
                t += 1;
                var old_gridBackColorList = _gridBackColorList;
                _gridBackColorList = new Brush[t + 1 + 1];
                if (old_gridBackColorList != null)
                    Array.Copy(old_gridBackColorList, _gridBackColorList, Math.Min(t + 1 + 1, old_gridBackColorList.Length));
                _gridBackColorList[t] = bcol;
                flag = t;
            }

            return flag;
        }

        private int GetGridCellAlignmentListEntry(StringFormat sfmt)
        {
            int t;
            int flag = -1;
            var loopTo = _gridCellAlignmentList.GetUpperBound(0);
            for (t = 0; t <= loopTo; t++)
            {
                if (sfmt.Equals(_gridCellAlignmentList[t]))
                {
                    flag = t;
                    break;
                }
            }

            if (flag == -1)
            {
                // we dont have that fnt2find in the list so we need to add it
                t = _gridCellAlignmentList.GetUpperBound(0);
                t += 1;
                var old_gridCellAlignmentList = _gridCellAlignmentList;
                _gridCellAlignmentList = new StringFormat[t + 1 + 1];
                if (old_gridCellAlignmentList != null)
                    Array.Copy(old_gridCellAlignmentList, _gridCellAlignmentList, Math.Min(t + 1 + 1, old_gridCellAlignmentList.Length));
                _gridCellAlignmentList[t] = sfmt;
                flag = t;
            }

            return flag;
        }

        private int GetGridCellFontListEntry(Font fnt2find)
        {
            int t;
            int flag = -1;
            var loopTo = _gridCellFontsList.GetUpperBound(0);
            for (t = 0; t <= loopTo; t++)
            {
                if (fnt2find.Equals(_gridCellFontsList[t]))
                {
                    flag = t;
                    break;
                }
            }

            if (flag == -1)
            {
                // we dont have that fnt2find in the list so we need to add it
                t = _gridCellFontsList.GetUpperBound(0);
                t += 1;
                var old_gridCellFontsList = _gridCellFontsList;
                _gridCellFontsList = new Font[t + 1 + 1];
                if (old_gridCellFontsList != null)
                    Array.Copy(old_gridCellFontsList, _gridCellFontsList, Math.Min(t + 1 + 1, old_gridCellFontsList.Length));
                _gridCellFontsList[t] = fnt2find;
                flag = t;
            }

            return flag;
        }

        private int GetGridForeColorListEntry(Pen fcol)
        {
            int t;
            int flag = -1;
            var loopTo = _gridForeColorList.GetUpperBound(0);
            for (t = 0; t <= loopTo; t++)
            {
                if (_gridForeColorList[t] == null)
                {
                }
                else if (fcol.Color.A == _gridForeColorList[t].Color.A)
                {
                    if (fcol.Color.R == _gridForeColorList[t].Color.R)
                    {
                        if (fcol.Color.G == _gridForeColorList[t].Color.G)
                        {
                            if (fcol.Color.B == _gridForeColorList[t].Color.B)
                            {
                                flag = t;
                                break;
                            }
                        }
                    }
                }
            }

            if (flag == -1)
            {
                // we dont have that fnt2find in the list so we need to add it
                t = _gridForeColorList.GetUpperBound(0);
                t += 1;
                var old_gridForeColorList = _gridForeColorList;
                _gridForeColorList = new Pen[t + 1 + 1];
                if (old_gridForeColorList != null)
                    Array.Copy(old_gridForeColorList, _gridForeColorList, Math.Min(t + 1 + 1, old_gridForeColorList.Length));
                _gridForeColorList[t] = fcol;
                flag = t;
            }

            return flag;
        }

        private string GetLetter(int iNumber)
        {
            try
            {
                switch (iNumber)
                {
                    case 0:
                        {
                            return "";
                        }

                    case 1:
                        {
                            return "A";
                        }

                    case 2:
                        {
                            return "B";
                        }

                    case 3:
                        {
                            return "C";
                        }

                    case 4:
                        {
                            return "D";
                        }

                    case 5:
                        {
                            return "E";
                        }

                    case 6:
                        {
                            return "F";
                        }

                    case 7:
                        {
                            return "G";
                        }

                    case 8:
                        {
                            return "H";
                        }

                    case 9:
                        {
                            return "I";
                        }

                    case 10:
                        {
                            return "J";
                        }

                    case 11:
                        {
                            return "K";
                        }

                    case 12:
                        {
                            return "L";
                        }

                    case 13:
                        {
                            return "M";
                        }

                    case 14:
                        {
                            return "N";
                        }

                    case 15:
                        {
                            return "O";
                        }

                    case 16:
                        {
                            return "P";
                        }

                    case 17:
                        {
                            return "Q";
                        }

                    case 18:
                        {
                            return "R";
                        }

                    case 19:
                        {
                            return "S";
                        }

                    case 20:
                        {
                            return "T";
                        }

                    case 21:
                        {
                            return "U";
                        }

                    case 22:
                        {
                            return "V";
                        }

                    case 23:
                        {
                            return "W";
                        }

                    case 24:
                        {
                            return "X";
                        }

                    case 25:
                        {
                            return "Y";
                        }

                    case 26:
                        {
                            return "Z";
                        }

                    default:
                        {
                            return "";
                        }
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.GetUpperColumn Error...");
                return "";
            }
        }

        private string GetUpperColumn(int iCols)
        {
            try
            {
                int iMajor = Conversions.ToInteger(iCols / (double)26);
                int iMinor = iCols % 26;

                return GetLetter(iMajor) + GetLetter(iMinor);
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.GetLetter Error...");
                return GetLetter(1) + GetLetter(1);
            }
        }

        private Point GimmeGridSize()
        {
            var pnt = default(Point);
            int x, y;
            int siz = 0;
            var loopTo = _cols - 1;
            for (x = 0; x <= loopTo; x++)
                siz = siz + _colwidths[x];

            pnt.X = siz;
            siz = 0;
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
                siz = siz + _rowheights[y];

            if (_GridTitleVisible)
                siz = siz + _GridTitleHeight;

            if (_GridHeaderVisible)
                siz = siz + _GridHeaderHeight;

            pnt.Y = siz;

            return pnt;
        }

        private int GimmeXOffset(int col)
        {
            int t;
            int ret = 0;

            if (col == 0)
                ret = 0;
            else
            {
                var loopTo = col - 1;
                for (t = 0; t <= loopTo; t++)
                    ret = ret + _colwidths[t];
            }
            return ret;
        }

        private int GimmeYOffset(int row)
        {
            int t;
            int ret = 0;

            if (row == 0)
                ret = 0;
            else
            {
                var loopTo = row - 1;
                for (t = 0; t <= loopTo; t++)
                    ret = ret + _rowheights[t];
            }
            return ret;
        }

        private void InitializeTheGrid()
        {
            int r, c;

            _grid = new string[3, 3];
            _gridBackColor = new int[3, 3];
            _gridForeColor = new int[3, 3];
            _gridCellFonts = new int[3, 3];
            _gridCellFontsList = new Font[2];
            _gridForeColorList = new Pen[2];
            _gridBackColorList = new Brush[2];
            _gridCellAlignment = new int[3, 3];
            _gridCellAlignmentList = new StringFormat[2];
            _colwidths = new int[3];
            _colhidden = new bool[3];
            _rowhidden = new bool[3];
            _colEditable = new bool[3];
            _rowEditable = new bool[3];
            _colboolean = new bool[3];
            _colPasswords = new string[3];
            _colMaxCharacters = new int[3];
            _rowheights = new int[3];
            _GridHeader = new string[3];
            _rows = 2;
            _cols = 2;

            _SelectedRow = -1;
            _SelectedColumn = -1;

            _gridCellFontsList[0] = _DefaultCellFont;
            _gridForeColorList[0] = new Pen(_DefaultForeColor);
            _gridCellAlignmentList[0] = _DefaultStringFormat;
            _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);


            hs.Visible = false;
            vs.Visible = false;
            hs.Value = 0;
            vs.Value = 0;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = _cols - 1;
                for (c = 0; c <= loopTo1; c++)
                {
                    _gridBackColor[r, c] = 0; // New SolidBrush(_DefaultBackColor.AntiqueWhite)
                    _gridForeColor[r, c] = 0; // New Pen(_DefaultForeColor.Blue)
                    _gridCellFonts[r, c] = 0; // _DefaultCellFont
                    _gridCellAlignment[r, c] = 0; // _DefaultStringFormat
                }
            }
        }

        private void InitializeTheGrid(int row, int col)
        {
            int r, c;

            _grid = new string[row + 1, col + 1];
            _gridBackColor = new int[row + 1, col + 1];
            _gridForeColor = new int[row + 1, col + 1];
            _gridCellFonts = new int[row + 1, col + 1];
            _gridCellFontsList = new Font[2];
            _gridForeColorList = new Pen[2];
            _gridBackColorList = new Brush[2];
            _gridCellAlignment = new int[row + 1, col + 1];
            _gridCellAlignmentList = new StringFormat[2];
            _colwidths = new int[col + 1];
            _colEditable = new bool[col + 1];
            _rowEditable = new bool[row + 1];
            _colhidden = new bool[col + 1];
            _colboolean = new bool[col + 1];
            _rowhidden = new bool[row + 1];
            _colPasswords = new string[col + 1];
            _colMaxCharacters = new int[col + 1];
            _rowheights = new int[row + 1];
            _GridHeader = new string[col + 1];

            _SelectedRows = new ArrayList();
            // '_SelectedRows.Clear()

            _rows = row;
            _cols = col;

            _SelectedRow = -1;
            _SelectedColumn = -1;

            _gridCellFontsList[0] = _DefaultCellFont;
            _gridForeColorList[0] = new Pen(_DefaultForeColor);
            _gridCellAlignmentList[0] = _DefaultStringFormat;
            _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);


            hs.Visible = false;
            vs.Visible = false;
            hs.Value = 0;
            vs.Value = 0;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                _rowEditable[r] = true;
                var loopTo1 = _cols - 1;
                for (c = 0; c <= loopTo1; c++)
                {
                    _gridBackColor[r, c] = 0; // New SolidBrush(_DefaultBackColor)
                    _gridForeColor[r, c] = 0; // New Pen(_DefaultForeColor)
                    _gridCellFonts[r, c] = 0; // _DefaultCellFont
                    _gridCellAlignment[r, c] = 0; // _DefaultStringFormat
                }
            }

            var loopTo2 = _cols - 1;
            for (c = 0; c <= loopTo2; c++)
            {
                _colPasswords[c] = "";
                _colEditable[c] = false;
                _colhidden[c] = false;
                _colboolean[c] = false;
            }
        }

        private void NormalizeTearaways()
        {
            if (TearAways.Count == 0)
                // we dont have any tearaways lets blow this pop stand
                return;

            // we have some so lets fix things here

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                if (ta.ColID >= _cols)
                    // we have a tearaway open on a column that no longer exists so lets close it
                    ta.Winform.KillMe(ta.ColID);
                else
                {
                    // we the column is still there so lets change its title and its contents
                    ta.Winform.Text = get_HeaderLabel(ta.ColID);
                    ta.Winform.ListItems = GetColAsArrayList(ta.ColID);
                    ta.SetTearAwayScrollParameters(vs.Minimum, vs.Maximum, vs.Visible);
                    ta.SetTearAwayScrollIndex(vs.Value);
                }
            }
        }

        private void OleRenderGrid(Graphics gr)
        {
            int w = AllColWidths();
            int h = AllRowHeights();
            var orig = default(Point);
            int t;
            int xof;
            int xxof, yyof;
            int r, c;
            int rh, rhy, rhx; // use for checkbox renderings
            int rowstart = -1;
            int rowend = -1;
            int colstart = -1;
            int colend = -1;
            int gyofset;
            string renderstring = "";

            if (_gridForeColorList[0] == null)
                _gridForeColorList[0] = new Pen(_DefaultForeColor);

            if (_gridBackColorList[0] == null)
                _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);

            if (_GridHeaderVisible)
                h += _GridHeaderHeight;

            if (_antialias)
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            }
            else
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault;
            }

            ClearToBackgroundColor(gr);

            // If we are disallowing selection of columns then make sure the Selected column variable is out of bounds
            if (!_AllowColumnSelection)
                _SelectedColumn = -1;

            if (_GridTitleVisible)
            {
                // we need to draw the title
                gr.FillRectangle(new SolidBrush(_GridTitleBackcolor), 0, 0, w, _GridTitleHeight);
                gr.DrawString(_GridTitle, _GridTitleFont, new SolidBrush(_GridTitleForeColor), 0, 0);
                orig.X = 0;
                orig.Y = _GridTitleHeight;
            }
            else
            {
                orig.X = 0;
                orig.Y = 0;
            }

            if (_cols != 0 & _GridHeaderVisible)
                orig.Y = orig.Y + _GridHeaderHeight;

            yyof = 0;
            xxof = 0;

            if (_rows == 0 & _cols == 0)
            {
            }
            else
            {
                rowstart = 0;
                rowend = _rows - 1;

                colstart = 0;
                colend = _cols - 1;
                var loopTo = rowend;

                // time to render the grid here
                for (r = rowstart; r <= loopTo; r++)
                {
                    gyofset = GimmeYOffset(r);
                    var loopTo1 = colend;
                    for (c = colstart; c <= loopTo1; c++)
                    {
                        xof = GimmeXOffset(c);
                        if (_colwidths[c] > 0)
                        {
                            if (_colPasswords[c] == null)
                                renderstring = _grid[r, c];
                            else if (string.IsNullOrEmpty(_colPasswords[c]))
                                renderstring = _grid[r, c];
                            else
                                renderstring = _colPasswords[c];

                            // handle the Max characters display here

                            if (_colMaxCharacters[c] != 0)
                            {
                                if (renderstring.Length > _colMaxCharacters[c])
                                    renderstring = renderstring.Substring(0, _colMaxCharacters[c]) + "...";
                            }

                            if (r == _SelectedRow | c == _SelectedColumn | _SelectedRows.Contains(r))
                            {
                                if (r == _SelectedRow | _SelectedRows.Contains(r))
                                {
                                    // we have a selected row override of selected column

                                    gr.FillRectangle(new SolidBrush(_RowHighLiteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        // we have to render the the checkbox

                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_RowHighLiteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);


                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                                }
                                else
                                {
                                    // we have a selected Col

                                    gr.FillRectangle(new SolidBrush(_ColHighliteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        // we have to render the the checkbox
                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_ColHighliteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);

                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                                }
                            }
                            else
                            {
                                gr.FillRectangle(_gridBackColorList[_gridBackColor[r, c]], xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                if (_colboolean[c])
                                {
                                    // we have to render the the checkbox
                                    rh = _rowheights[r] - 2;

                                    if (rh > 14)
                                        rh = 14;

                                    if (rh < 6)
                                        rh = 6;

                                    rhx = _colwidths[c] / 2 - rh / 2;

                                    if (rhx < 0)
                                        rhx = 0;

                                    rhy = _rowheights[r] / 2 - rh / 2;

                                    if (rhy < 0)
                                        rhy = 0;

                                    if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                    else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                    else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                }
                                else
                                    gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_gridForeColorList[_gridForeColor[r, c]].Color), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);
                                if (_CellOutlines)
                                    gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                            }
                        }
                    }
                }

                // recalc the top area so we can draw the header if its vivible
                if (_GridTitleVisible)
                {
                    orig.X = 0;
                    orig.Y = _GridTitleHeight;
                }
                else
                {
                    orig.X = 0;
                    orig.Y = 0;
                }

                gr.SetClip(new RectangleF(0, 0, w, h));

                if (_cols != 0 & _GridHeaderVisible)
                {
                    var loopTo2 = _cols - 1;
                    // we need to render the Header

                    for (t = 0; t <= loopTo2; t++)
                    {
                        xof = GimmeXOffset(t);
                        if (_colwidths[t] > 0)
                        {
                            gr.FillRectangle(new SolidBrush(_GridHeaderBackcolor), xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight);
                            gr.DrawString(_GridHeader[t], _GridHeaderFont, new SolidBrush(_GridHeaderForecolor), new RectangleF(xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight), _GridHeaderStringFormat);
                            if (_CellOutlines)
                                gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight));
                        }
                    }
                    orig.Y = orig.Y + _GridHeaderHeight;
                }

                // do we need to display the scrollbars

                // RecalcScrollBars()

                if ((int)_BorderStyle == (int)BorderStyle.Fixed3D | (int)_BorderStyle == (int)BorderStyle.FixedSingle)
                    gr.DrawRectangle(new Pen(_BorderColor, 1), 0, 0, w - 1, h - 1);

                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
                _Painting = false;
            }
        }

        private void PrivatePopulateGridFromArray(string[,] arr, Font gridfont, Color col, bool FirstRowHeader)
        {
            int x, y;
            int r, c;

            r = arr.GetUpperBound(0) + 1;
            c = arr.GetUpperBound(1) + 1;

            if (FirstRowHeader)
            {
                InitializeTheGrid(r - 1, c);
                var loopTo = c - 1;
                for (y = 0; y <= loopTo; y++)
                    _GridHeader[y] = arr[0, y];
                var loopTo1 = r - 1;
                for (x = 1; x <= loopTo1; x++)
                {
                    var loopTo2 = c - 1;
                    for (y = 0; y <= loopTo2; y++)
                        _grid[x, y] = arr[x, y];
                }
            }
            else
            {
                var loopTo3 = r - 1;
                // InitializeTheGrid(r, c)
                // For y = 0 To c - 1
                // _GridHeader(y) = "Column - " & y.ToString
                // Next
                for (x = 0; x <= loopTo3; x++)
                {
                    var loopTo4 = c - 1;
                    for (y = 0; y <= loopTo4; y++)
                        _grid[x, y] = arr[x, y];
                }
            }

            AllCellsUseThisFont(gridfont);
            AllCellsUseThisForeColor(col);

            AutoSizeCellsToContents = true;
            _colEditRestrictions.Clear();

            Refresh();
        }

        private void RecalcScrollBars()
        {
            // ##Recalculates the positions and the visibility for the scroll bars

            int ClientHeight;

            ClientHeight = Height;


            _GridSize = GimmeGridSize();

            if (_GridHeaderVisible)
                ClientHeight -= _GridHeaderHeight;

            if (_GridTitleVisible)
                ClientHeight -= _GridTitleHeight;

            if (_GridSize.X > Width)
            {
                hs.Visible = true;
                hs.Height = _ScrollBarWeight;

                hs.Maximum = _cols + 2;
                hs.LargeChange = 4;
                hs.SmallChange = 1;
                ClientHeight -= _ScrollBarWeight;
            }
            else
            {
                hs.Visible = false;
                hs.Maximum = 1;
                hs.Minimum = 0;
                hs.Value = 0;
            }

            if (_GridSize.Y > ClientHeight)
            {
                vs.Visible = true;
                vs.Width = _ScrollBarWeight;

                vs.Maximum = _rows + 10;
                vs.LargeChange = 10;
                vs.SmallChange = 1;
            }
            else
            {
                vs.Visible = false;
                vs.Maximum = 1;
                vs.Minimum = 0;
                vs.Value = 0;
            }
        }

        private void RedimTable()
        {
            var oldrowhidden = new bool[_rowhidden.GetUpperBound(0) + 1];
            var oldcolhidden = new bool[_colhidden.GetUpperBound(0) + 1];
            var oldcolboolean = new bool[_colboolean.GetUpperBound(0) + 1];
            var oldcoleditable = new bool[_colEditable.GetUpperBound(0) + 1];
            var oldroweditable = new bool[_rowEditable.GetUpperBound(0) + 1];
            var oldcolwidths = new int[_colwidths.GetUpperBound(0) + 1];
            var oldrowheights = new int[_rowheights.GetUpperBound(0) + 1];
            var oldgridheader = new string[_GridHeader.GetUpperBound(0) + 1];
            var oldgrid = new string[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridbcolor = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridfcolor = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridfonts = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            var oldgridcolpasswords = new string[_colPasswords.GetUpperBound(0) + 1];
            var oldcolmaxcharacters = new int[_colMaxCharacters.GetUpperBound(0) + 1];
            var oldgridcellalignment = new int[_grid.GetUpperBound(0) + 1, _grid.GetUpperBound(1) + 1];
            int r, c;
            int x, y;

            x = oldgrid.GetUpperBound(0);
            y = oldgrid.GetUpperBound(1);
            var loopTo = x;
            for (r = 0; r <= loopTo; r++)
            {
                var loopTo1 = y;
                for (c = 0; c <= loopTo1; c++)
                {
                    oldgrid[r, c] = _grid[r, c];
                    oldgridbcolor[r, c] = _gridBackColor[r, c];
                    oldgridfcolor[r, c] = _gridForeColor[r, c];
                    oldgridfonts[r, c] = _gridCellFonts[r, c];
                    oldgridcellalignment[r, c] = _gridCellAlignment[r, c];
                }
            }

            var loopTo2 = Math.Min(_GridHeader.GetUpperBound(0), _colwidths.GetUpperBound(0));
            for (c = 0; c <= loopTo2; c++)
            {
                oldgridheader[c] = _GridHeader[c];
                oldcolwidths[c] = _colwidths[c];
                oldgridcolpasswords[c] = _colPasswords[c];
                oldcolhidden[c] = _colhidden[c];
                oldcolboolean[c] = _colboolean[c];
                oldcoleditable[c] = _colEditable[c];
                oldcolmaxcharacters[c] = _colMaxCharacters[c];
            }

            var loopTo3 = _rowheights.GetUpperBound(0);
            for (r = 0; r <= loopTo3; r++)
                oldrowheights[r] = _rowheights[r];
            var loopTo4 = _rowhidden.GetUpperBound(0);
            for (r = 0; r <= loopTo4; r++)
                oldrowhidden[r] = _rowhidden[r];
            var loopTo5 = _rowEditable.GetUpperBound(0);
            for (r = 0; r <= loopTo5; r++)
                oldroweditable[r] = _rowEditable[r];

            _rowhidden = new bool[_rows + 1];
            _colhidden = new bool[_cols + 1];
            _colboolean = new bool[_cols + 1];
            _colEditable = new bool[_cols + 1];
            _rowEditable = new bool[_rows + 1];
            _rowheights = new int[_rows + 1];
            _colwidths = new int[_cols + 1];
            _GridHeader = new string[_cols + 1];
            _grid = new string[_rows + 1, _cols + 1];
            _gridBackColor = new int[_rows + 1, _cols + 1];
            _gridForeColor = new int[_rows + 1, _cols + 1];
            _gridCellFonts = new int[_rows + 1, _cols + 1];
            _gridCellAlignment = new int[_rows + 1, _cols + 1];
            _colPasswords = new string[_cols + 1];
            _colMaxCharacters = new int[_cols + 1];

            if (_rows < x)
                x = _rows;

            if (_cols < y)
                y = _cols;
            var loopTo6 = y;
            for (c = 0; c <= loopTo6; c++)
            {
                _colPasswords[c] = oldgridcolpasswords[c];
                _GridHeader[c] = oldgridheader[c];
                _colwidths[c] = oldcolwidths[c];
                _colhidden[c] = oldcolhidden[c];
                _colboolean[c] = oldcolboolean[c];
                _colEditable[c] = oldcoleditable[c];
                _colMaxCharacters[c] = oldcolmaxcharacters[c];
            }

            var loopTo7 = x;
            for (r = 0; r <= loopTo7; r++)
            {
                _rowheights[r] = oldrowheights[r];
                _rowhidden[r] = oldrowhidden[r];
                _rowEditable[r] = oldroweditable[r];
            }

            if (x == 0)
            {
                r = x;
                var loopTo8 = y;
                for (c = 0; c <= loopTo8; c++)
                {
                    _grid[r, c] = oldgrid[r, c];
                    _gridBackColor[r, c] = GetGridBackColorListEntry(new SolidBrush(_DefaultBackColor));
                    _gridForeColor[r, c] = GetGridForeColorListEntry(new Pen(_DefaultForeColor));
                    _gridCellFonts[r, c] = GetGridCellFontListEntry(_DefaultCellFont);
                    _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(_DefaultStringFormat);
                }
            }
            else
            {
                var loopTo9 = x;
                for (r = 0; r <= loopTo9; r++)
                {
                    var loopTo10 = y;
                    for (c = 0; c <= loopTo10; c++)
                    {
                        _grid[r, c] = oldgrid[r, c];
                        _gridBackColor[r, c] = oldgridbcolor[r, c];
                        _gridForeColor[r, c] = oldgridfcolor[r, c];
                        _gridCellFonts[r, c] = oldgridfonts[r, c];
                        _gridCellAlignment[r, c] = oldgridcellalignment[r, c];
                    }
                }
            }

            if (oldcolwidths.GetUpperBound(0) < _colwidths.GetUpperBound(0))
            {
                var loopTo11 = _colwidths.GetUpperBound(0);
                for (c = oldcolwidths.GetUpperBound(0) + 1; c <= loopTo11; c++)
                {
                    _colwidths[c] = _DefaultColWidth;
                    _colEditable[c] = false; // default all new columns to not editable
                    _colhidden[c] = false; // cols default to not hidden
                    _colboolean[c] = false; // cols default to not boolean
                }
            }

            if (oldrowheights.GetUpperBound(0) < _rowheights.GetUpperBound(0))
            {
                var loopTo12 = _rowheights.GetUpperBound(0);
                for (r = oldrowheights.GetUpperBound(0) + 1; r <= loopTo12; r++)
                    _rowheights[r] = _DefaultRowHeight;
            }

            var loopTo13 = _cols - 1;
            for (c = 0; c <= loopTo13; c++)
            {
                if (_colwidths[c] == 0 & !_colhidden[c])
                    _colwidths[c] = _DefaultColWidth;
            }

            var loopTo14 = _rows - 1;
            for (r = 0; r <= loopTo14; r++)
            {
                if (_rowheights[r] == 0 & !_rowhidden[r])
                    _rowheights[r] = _DefaultRowHeight;
            }
        }

        private void RenderGrid(Graphics grview)
        {
            int w = Conversions.ToInteger(grview.VisibleClipBounds.Width);
            int h = Conversions.ToInteger(grview.VisibleClipBounds.Height);
            var orig = default(Point);
            int t;
            int xof;
            int xxof, yyof;
            int r, c;
            int rh, rhy, rhx; // use for checkbox renderings
            int rowstart = -1;
            int rowend = -1;
            int colstart = -1;
            int colend = -1;
            int gyofset;
            string renderstring = "";

            // 
            // Here we want to just bail if the size is less than some small size
            // 

            if (w < 10 | h < 10)
                return;

            if (_Painting)
                return;
            else
                _Painting = true;

            if (_gridForeColorList[0] == null)
                _gridForeColorList[0] = new Pen(_DefaultForeColor);

            if (_gridBackColorList[0] == null)
                _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);

            Graphics gr;
            Bitmap bmp;

            bmp = new Bitmap(w, h, grview);
            gr = Graphics.FromImage(bmp);

            if (_antialias)
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            }
            else
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault;
            }

            DoAutoSizeCheck(gr);

            ClearToBackgroundColor(gr);

            RecalcScrollBars();

            // If we are disallowing selection of columns then make sure the Selectedcolumn variable is out of bounds
            if (!_AllowColumnSelection)
                _SelectedColumn = -1;

            if (_GridTitleVisible)
            {
                // we need to draw the title
                gr.FillRectangle(new SolidBrush(_GridTitleBackcolor), 0, 0, w, _GridTitleHeight);
                gr.DrawString(_GridTitle, _GridTitleFont, new SolidBrush(_GridTitleForeColor), 0, 0);
                orig.X = 0;
                orig.Y = _GridTitleHeight;
            }
            else
            {
                orig.X = 0;
                orig.Y = 0;
            }

            if (_cols != 0 & _GridHeaderVisible)
                orig.Y = orig.Y + _GridHeaderHeight;

            if (vs.Visible)
                yyof = GimmeYOffset(vs.Value);
            else
                yyof = 0;

            if (hs.Visible)
                xxof = GimmeXOffset(hs.Value);
            else
                xxof = 0;

            if (_rows == 0 & _cols == 0)
                // We have nothing else to draw so lets bail
                _Painting = false;
            else
            {
                // if we are possible needing to draw the background we had better do it here
                if (!(hs.Visible & vs.Visible))
                {
                    if (_GridTitleVisible)
                        gr.FillRectangle(new SolidBrush(BackColor), new RectangleF(0, _GridTitleHeight, w, h - _GridTitleHeight));
                    else
                        gr.FillRectangle(new SolidBrush(BackColor), gr.VisibleClipBounds);
                }

                // here we want to validate the starting and ending rows for the render process
                if (vs.Visible)
                {

                    // If _SelectedRow <> -1 Then
                    // vs.Value = _SelectedRow
                    // End If

                    rowstart = vs.Value;
                    rowend = _rows - 1;
                    var loopTo = _rows - 1;
                    for (r = rowstart; r <= loopTo; r++)
                    {
                        if (GimmeYOffset(r) - yyof >= h)
                        {
                            rowend = r;
                            break;
                        }
                    }
                }
                else
                {
                    rowstart = 0;
                    rowend = _rows - 1;
                }

                if (hs.Visible)
                {
                    colstart = hs.Value;
                    colend = _cols - 1;
                    var loopTo1 = _cols - 1;
                    for (c = colstart; c <= loopTo1; c++)
                    {
                        if (GimmeXOffset(c) - xxof >= w)
                        {
                            colend = c;
                            break;
                        }
                    }
                }
                else
                {
                    colstart = 0;
                    colend = _cols - 1;
                }

                // If _SelectedRow <> -1 And vs.Visible Then
                // If _SelectedRow < rowstart Then
                // vs.Value = vs.Value - (rowstart - _SelectedRow)
                // End If
                // If _SelectedRow > rowend Then
                // vs.Value = vs.Value + (_SelectedRow - rowend)
                // End If
                // End If

                // from now on all drawing ops occur below the grid title if its visible and the header if its visible

                gr.SetClip(new RectangleF(0, orig.Y, w, h - orig.Y));
                var loopTo2 = rowend;

                // Console.WriteLine(rowstart.ToString & " - " & rowend.ToString & " ------- " & colstart.ToString & " - " & colend)

                // time to render the grid here
                for (r = rowstart; r <= loopTo2; r++)
                {
                    gyofset = GimmeYOffset(r);
                    var loopTo3 = colend;
                    for (c = colstart; c <= loopTo3; c++)
                    {
                        xof = GimmeXOffset(c);
                        if (_colwidths[c] > 0)
                        {
                            if (_colPasswords[c] == null)
                                renderstring = _grid[r, c];
                            else if (string.IsNullOrEmpty(_colPasswords[c]))
                                renderstring = _grid[r, c];
                            else
                                renderstring = _colPasswords[c];

                            // handle the Max characters display here

                            if (_colMaxCharacters[c] != 0)
                            {
                                if (renderstring.Length > _colMaxCharacters[c])
                                    renderstring = renderstring.Substring(0, _colMaxCharacters[c]) + "...";
                            }

                            if (r == _SelectedRow | c == _SelectedColumn | _SelectedRows.Contains(r))
                            {
                                if (r == _SelectedRow | _SelectedRows.Contains(r))
                                {
                                    // we have a selected row override of selected column

                                    gr.FillRectangle(new SolidBrush(_RowHighLiteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        // we have to render the the checkbox

                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_RowHighLiteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);


                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                                }
                                else
                                {
                                    // we have a selected Col

                                    gr.FillRectangle(new SolidBrush(_ColHighliteBackColor), xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        // we have to render the the checkbox
                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_ColHighliteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);

                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                                }
                            }
                            else
                            {
                                gr.FillRectangle(_gridBackColorList[_gridBackColor[r, c]], xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]);

                                if (_colboolean[c])
                                {
                                    // we have to render the the checkbox
                                    rh = _rowheights[r] - 2;

                                    if (rh > 14)
                                        rh = 14;

                                    if (rh < 6)
                                        rh = 6;

                                    rhx = _colwidths[c] / 2 - rh / 2;

                                    if (rhx < 0)
                                        rhx = 0;

                                    rhy = _rowheights[r] / 2 - rh / 2;

                                    if (rhy < 0)
                                        rhy = 0;

                                    if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                    else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                    else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                }
                                else
                                    gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_gridForeColorList[_gridForeColor[r, c]].Color), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);
                                if (_CellOutlines)
                                    gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]));
                            }
                        }
                    }
                }

                // recalc the top area so we can draw the header if its vivible
                if (_GridTitleVisible)
                {
                    orig.X = 0;
                    orig.Y = _GridTitleHeight;
                }
                else
                {
                    orig.X = 0;
                    orig.Y = 0;
                }

                gr.SetClip(new RectangleF(0, 0, w, h));

                if (_cols != 0 & _GridHeaderVisible)
                {
                    var loopTo4 = _cols - 1;
                    // we need to render the Header

                    for (t = 0; t <= loopTo4; t++)
                    {
                        xof = GimmeXOffset(t);
                        if (_colwidths[t] > 0)
                        {
                            gr.FillRectangle(new SolidBrush(_GridHeaderBackcolor), xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight);
                            gr.DrawString(_GridHeader[t], _GridHeaderFont, new SolidBrush(_GridHeaderForecolor), new RectangleF(xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight), _GridHeaderStringFormat);
                            if (_CellOutlines)
                                gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof, orig.Y, _colwidths[t], _GridHeaderHeight));
                        }
                    }
                    orig.Y = orig.Y + _GridHeaderHeight;
                }

                // do we need to display the scrollbars

                RecalcScrollBars();

                if ((int)_BorderStyle == (int)BorderStyle.Fixed3D | (int)_BorderStyle == (int)BorderStyle.FixedSingle)
                    gr.DrawRectangle(new Pen(_BorderColor, 1), 0, 0, w - 1, h - 1);

                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
                _Painting = false;
            }

            grview.DrawImageUnscaled(bmp, 0, 0);
            gr.Dispose();
            bmp.Dispose();
            bmp = null;
        }

        private void RenderGridToGraphicsContext(Graphics gr, Rectangle Cliprect)
        {
            int w = AllColWidths();
            int h = AllRowHeights();
            var orig = default(Point);
            int t;
            int xof;
            int xxof, yyof, ofx, ofy;
            int r, c;
            int rh, rhy, rhx; // use for checkbox renderings
            int rowstart = -1;
            int rowend = -1;
            int colstart = -1;
            int colend = -1;
            int gyofset;
            string renderstring = "";

            gr.SetClip(Cliprect);
            ofx = Cliprect.X;
            ofy = Cliprect.Y;

            if (_gridForeColorList[0] == null)
                _gridForeColorList[0] = new Pen(_DefaultForeColor);

            if (_gridBackColorList[0] == null)
                _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);

            if (_GridHeaderVisible)
                h += _GridHeaderHeight;

            if (_antialias)
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            }
            else
            {
                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
                gr.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault;
            }

            ClearToBackgroundColor(gr);

            // If we are disallowing selection of columns then make sure the Selected column variable is out of bounds
            if (!_AllowColumnSelection)
                _SelectedColumn = -1;

            if (_GridTitleVisible)
            {
                // we need to draw the title
                gr.FillRectangle(new SolidBrush(_GridTitleBackcolor), 0 + ofx, 0 + ofy, w, _GridTitleHeight);
                gr.DrawString(_GridTitle, _GridTitleFont, new SolidBrush(_GridTitleForeColor), 0 + ofx, 0 + ofy);
                orig.X = 0;
                orig.Y = _GridTitleHeight;
            }
            else
            {
                orig.X = 0;
                orig.Y = 0;
            }

            if (_cols != 0 & _GridHeaderVisible)
                orig.Y = orig.Y + _GridHeaderHeight;

            yyof = 0;
            xxof = 0;

            if (_rows == 0 & _cols == 0)
            {
            }
            else
            {
                rowstart = 0;
                rowend = _rows - 1;

                colstart = 0;
                colend = _cols - 1;
                var loopTo = rowend;

                // time to render the grid here
                for (r = rowstart; r <= loopTo; r++)
                {
                    gyofset = GimmeYOffset(r);
                    var loopTo1 = colend;
                    for (c = colstart; c <= loopTo1; c++)
                    {
                        xof = GimmeXOffset(c);
                        if (_colwidths[c] > 0)
                        {
                            if (_colPasswords[c] == null)
                                renderstring = _grid[r, c];
                            else if (string.IsNullOrEmpty(_colPasswords[c]))
                                renderstring = _grid[r, c];
                            else
                                renderstring = _colPasswords[c];

                            // handle the Max characters display here

                            if (_colMaxCharacters[c] != 0)
                            {
                                if (renderstring.Length > _colMaxCharacters[c])
                                    renderstring = renderstring.Substring(0, _colMaxCharacters[c]) + "...";
                            }

                            if (r == _SelectedRow | c == _SelectedColumn | _SelectedRows.Contains(r))
                            {
                                if (r == _SelectedRow | _SelectedRows.Contains(r))
                                {
                                    // we have a selected row override of selected column

                                    gr.FillRectangle(new SolidBrush(_RowHighLiteBackColor), xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_RowHighLiteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);

                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]));
                                }
                                else
                                {
                                    // we have a selected Col

                                    gr.FillRectangle(new SolidBrush(_ColHighliteBackColor), xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]);

                                    if (_colboolean[c])
                                    {
                                        // we have to render the the checkbox
                                        rh = _rowheights[r] - 2;

                                        if (rh > 14)
                                            rh = 14;

                                        if (rh < 6)
                                            rh = 6;

                                        rhx = _colwidths[c] / 2 - rh / 2;

                                        if (rhx < 0)
                                            rhx = 0;

                                        rhy = _rowheights[r] / 2 - rh / 2;

                                        if (rhy < 0)
                                            rhy = 0;

                                        if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                        else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                        else
                                            ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                    }
                                    else
                                        gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_RowHighLiteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);

                                    if (_CellOutlines)
                                        gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]));
                                }
                            }
                            else
                            {
                                gr.FillRectangle(_gridBackColorList[_gridBackColor[r, c]], xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]);

                                if (_colboolean[c])
                                {
                                    // we have to render the the checkbox
                                    rh = _rowheights[r] - 2;

                                    if (rh > 14)
                                        rh = 14;

                                    if (rh < 6)
                                        rh = 6;

                                    rhx = _colwidths[c] / 2 - rh / 2;

                                    if (rhx < 0)
                                        rhx = 0;

                                    rhy = _rowheights[r] / 2 - rh / 2;

                                    if (rhy < 0)
                                        rhy = 0;

                                    if ((Strings.UCase(renderstring) ?? "") == "TRUE" | (Strings.UCase(renderstring) ?? "") == "YES" | (Strings.UCase(renderstring) ?? "") == "Y" | (Strings.UCase(renderstring) ?? "") == "1")
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Checked);
                                    else if (string.IsNullOrEmpty(Strings.UCase(renderstring)))
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Inactive);
                                    else
                                        ControlPaint.DrawCheckBox(gr, xof - xxof + rhx, orig.Y + gyofset - yyof + rhy, rh, rh, ButtonState.Normal);
                                }
                                else
                                    gr.DrawString(renderstring, _gridCellFontsList[_gridCellFonts[r, c]], new SolidBrush(_RowHighLiteForeColor), new RectangleF(xof - xxof, orig.Y + gyofset - yyof, _colwidths[c], _rowheights[r]), _gridCellAlignmentList[_gridCellAlignment[r, c]]);
                                if (_CellOutlines)
                                    gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof + ofx, orig.Y + gyofset - yyof + ofy, _colwidths[c], _rowheights[r]));
                            }
                        }
                    }
                }

                // recalc the top area so we can draw the header if its vivible
                if (_GridTitleVisible)
                {
                    orig.X = 0;
                    orig.Y = _GridTitleHeight;
                }
                else
                {
                    orig.X = 0;
                    orig.Y = 0;
                }

                gr.SetClip(new RectangleF(0 + ofx, 0 + ofy, Cliprect.Width, Cliprect.Height));

                if (_cols != 0 & _GridHeaderVisible)
                {
                    var loopTo2 = _cols - 1;
                    // we need to render the Header

                    for (t = 0; t <= loopTo2; t++)
                    {
                        xof = GimmeXOffset(t);
                        if (_colwidths[t] > 0)
                        {
                            gr.FillRectangle(new SolidBrush(_GridHeaderBackcolor), xof - xxof + ofx, orig.Y + ofy, _colwidths[t], _GridHeaderHeight);
                            gr.DrawString(_GridHeader[t], _GridHeaderFont, new SolidBrush(_GridHeaderForecolor), new RectangleF(xof - xxof + ofx, orig.Y + ofy, _colwidths[t], _GridHeaderHeight), _GridHeaderStringFormat);
                            if (_CellOutlines)
                                gr.DrawRectangle(new Pen(_CellOutlineColor), new Rectangle(xof - xxof + ofx, orig.Y + ofy, _colwidths[t], _GridHeaderHeight));
                        }
                    }
                    orig.Y = orig.Y + _GridHeaderHeight;
                }

                // do we need to display the scrollbars

                // RecalcScrollBars()

                if ((int)_BorderStyle == (int)BorderStyle.Fixed3D | (int)_BorderStyle == (int)BorderStyle.FixedSingle)
                    gr.DrawRectangle(new Pen(_BorderColor, 1), 0 + ofx, 0 + ofy, Cliprect.Width - 1, Cliprect.Height - 1);

                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped EndIfDirectiveTrivia */
                _Painting = false;
            }
        }

        private string ReturnExcelColumn(int intColumn)
        {
            try
            {
                var arrAlphabet = new ArrayList();
                arrAlphabet.Add("A");
                arrAlphabet.Add("B");
                arrAlphabet.Add("C");
                arrAlphabet.Add("D");
                arrAlphabet.Add("E");
                arrAlphabet.Add("F");
                arrAlphabet.Add("G");
                arrAlphabet.Add("H");
                arrAlphabet.Add("I");
                arrAlphabet.Add("J");
                arrAlphabet.Add("K");
                arrAlphabet.Add("L");
                arrAlphabet.Add("M");
                arrAlphabet.Add("N");
                arrAlphabet.Add("O");
                arrAlphabet.Add("P");
                arrAlphabet.Add("Q");
                arrAlphabet.Add("R");
                arrAlphabet.Add("S");
                arrAlphabet.Add("T");
                arrAlphabet.Add("U");
                arrAlphabet.Add("V");
                arrAlphabet.Add("W");
                arrAlphabet.Add("X");
                arrAlphabet.Add("Y");
                arrAlphabet.Add("Z");

                if (intColumn <= 25)
                    return arrAlphabet[intColumn].ToString();
                else
                {
                    int idx = intColumn / 26;
                    if (idx == 0)
                        idx += 1;
                    if (idx >= 1)
                        // If (intColumn - 1) - (idx * 26) < 0 Then
                        // Return arrAlphabet.Item(idx - 1) + arrAlphabet.Item((intColumn) - (idx * 26))
                        // Else
                        return arrAlphabet[idx - 1].ToString() + arrAlphabet[intColumn - idx * 26].ToString();
                    else
                        return "A" + arrAlphabet[intColumn - idx * 26];
                }
            }
            catch (Exception ex)
            {
                Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "TAIGRIDControl.ExportToExcel.ReturnExcelColumn Error...");
                return "";
            }
        }

        private string ReturnByteArrayAsHexString(byte[] Bytes)
        {
            int t;
            string result = "";
            string a;

            try
            {
                var loopTo = Bytes.GetLength(0) - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    a = Strings.Right("00" + Conversion.Hex(Bytes[t]), 2);
                    result = result + a;
                }
            }
            catch (Exception ex)
            {
            }

            return result;
        }

        private string ReturnHTMLColor(Color col)
        {
            string result = Conversions.ToString((char)34) + "#";

            result += Strings.Right("0" + Conversion.Hex(col.R), 2);
            result += Strings.Right("0" + Conversion.Hex(col.G), 2);
            result += Strings.Right("0" + Conversion.Hex(col.B), 2) + Conversions.ToString((char)34);

            return result;
        }

        private void SetCols(int newColVal)
        {
            int t;

            hs.Value = 0;

            if (_cols == 0)
            {
                // we have no columns now so lets just set them
                _colwidths = new int[newColVal + 1];
                var loopTo = newColVal - 1;
                for (t = 0; t <= loopTo; t++)
                    _colwidths[t] = _DefaultColWidth;
                _cols = newColVal;
            }
            else
                _cols = newColVal;

            RedimTable();
        }

        private void SetRows(int newRowVal)
        {
            int t;

            vs.Value = 0;

            if (_rows == 0)
            {
                // we have no rows now so lets start things off
                _rowheights = new int[newRowVal + 1];
                var loopTo = newRowVal - 1;
                for (t = 0; t <= loopTo; t++)
                    _rowheights[t] = _DefaultRowHeight;
                _rows = newRowVal;
            }
            else
                _rows = newRowVal;

            RedimTable();
        }

        private string SplitLongString(string input, int breaklen)
        {
            var splitstringarray = input.Split(" ".ToCharArray());

            string ret = "";
            string subret = "";

            int t = 0;
            var loopTo = splitstringarray.GetUpperBound(0);
            for (t = 0; t <= loopTo; t++)
            {
                if ((splitstringarray[t].Trim() ?? "") == (Environment.NewLine ?? ""))
                    splitstringarray[t] = "";

                subret += " " + splitstringarray[t].Trim();

                if (subret.Length >= breaklen)
                {
                    ret += subret + Environment.NewLine;
                    subret = "";
                }
            }


            ret += subret;

            ret = ret.Trim();

            if (ret.EndsWith(Environment.NewLine))
                ret = ret.Substring(1, ret.Length - Environment.NewLine.Length);


            return ret;
        }

        private void TearAwayColumID(int id)
        {
            if (TearAways.Count > 0)
            {
                int t;
                var loopTo = TearAways.Count - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                    if (ta.ColID == id)
                    {
                        // we already got one of these
                        ta.Winform.BringToFront();
                        ta.Winform.Focus();
                        return;
                    }
                }
            }

            var tear = new TearAwayWindowEntry();
            var TearItem = new frmColumnTearAway(get_HeaderLabel(id));
            TearItem.Show();

            TearItem.ListItems = GetColAsArrayList(id);
            TearItem.GridParent = this;
            TearItem.Colid = id;
            TearItem.DefaultSelectionColor = _RowHighLiteBackColor;
            TearItem.GridDefaultBackColor = _DefaultBackColor;
            TearItem.GridDefaultForeColor = _DefaultForeColor;
            TearItem.SelectedRow = _SelectedRow;

            tear.Winform = TearItem;

            tear.ColID = id;
            tear.SetTearAwayScrollParameters(vs.Minimum, vs.Maximum, vs.Visible);

            // tear.ShowTearAway()
            TearAways.Add(tear);
        }

        private void TAIGRIDv2_Paint(object sender, PaintEventArgs e)
        {
            RenderGrid(e.Graphics);

            if (TearAways.Count != 0)
            {
                int t;
                var loopTo = TearAways.Count - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    TearAwayWindowEntry tear = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                    tear.Winform.SelectedRow = _SelectedRow;
                }
            }
        }

        private void TAIGRIDv2_SizeChanged(object sender, EventArgs e)
        {
            ClearToBackgroundColor();
            Invalidate();
        }

        private void hs_ValueChanged(object sender, EventArgs e)
        {
            Invalidate();

            if (hs.Visible & _EditMode)
            {
                int xxoff = GimmeXOffset(hs.Value);
                int xxxoff = GimmeXOffset(_EditModeCol);
                int xoff = xxxoff - xxoff;

                cmboInput.Left = xoff;
                txtInput.Left = xoff;
            }
        }

        private void vs_ValueChanged(object sender, EventArgs e)
        {
            int t;

            Invalidate();

            if (TearAways.Count > 0)
            {
                var loopTo = TearAways.Count - 1;
                // we have some tear sway windows open so lets set their verticle scrollers
                for (t = 0; t <= loopTo; t++)
                {
                    TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                    ta.SetTearAwayScrollIndex(vs.Value);
                }
            }

            if (vs.Visible & _EditMode)
            {
                int yyoff = GimmeYOffset(vs.Value);

                int yyyoff = GimmeYOffset(_EditModeRow);

                int yoff = yyyoff - yyoff;

                if (_GridTitleVisible & _GridHeaderVisible)
                {
                    yoff += _GridTitleHeight + _GridHeaderHeight;

                    if (yoff < _GridTitleHeight + _GridHeaderHeight)
                        yoff = -100;
                }
                else if (_GridTitleVisible)
                {
                    yoff += _GridTitleHeight;

                    if (yoff < _GridTitleHeight)
                        yoff = -100;
                }
                else if (GridheaderVisible)
                {
                    yoff += GridHeaderHeight;

                    if (yoff < _GridHeaderHeight)
                        yoff = -100;
                }

                cmboInput.Top = yoff;
                txtInput.Top = yoff;
            }
        }

        private void MouseEnterHandler(object sender, EventArgs e)
        {
            if (_AutoFocus)
                Focus();
        }

        private void MouseWheelHandler(object sender, MouseEventArgs e)
        {
            int v;
            int del = e.Delta;

            if (del < 0)
                del = -_MouseWheelScrollAmount;
            else if (del > 0)
                del = _MouseWheelScrollAmount;

            if (vs.Visible)
            {
                v = vs.Value - del;

                if (v < 0)
                    v = 0;
                if (v > _rows)
                    v = _rows;

                // _SelectedRow = -1
                vs.Value = v;
            }
            else if (hs.Visible)
            {
                v = hs.Value - del;

                if (v < 0)
                    v = 0;
                if (v > _cols)
                    v = _cols;

                hs.Value = v;
            }
        }

        private void MouseUpHandler(object sender, MouseEventArgs e)
        {
            int xoff, yoff, r, c, t;
            string ss;

            // Console.WriteLine("MouseUP")

            if (_DoubleClickSemaphore)
            {
                _DoubleClickSemaphore = false;
                return;
            }

            if (!(_OldContextMenu == null) & ContextMenu == null)
                ContextMenu = _OldContextMenu;


            if (_MouseDownOnHeader)
            {
                // we were in column resize mode so lets clear all than and blow this pop stand
                // that should prevent the unwanted selection of the top visible row for folks with
                // shakey mouse control like yours truely...

                _LastMouseX = -1;
                _LastMouseY = -1;
                _MouseDownOnHeader = false;
                _ColOverOnMouseDown = -1;
                _RowOverOnMouseDown = -1;
                Cursor = Cursors.Default;
                return;
            }

            if ((int)e.Button == (int)MouseButtons.Right)
                // bail on a right mousebutton
                return;

            txtHandler.Focus();

            yoff = 0;

            if (_GridTitleVisible)
                yoff = yoff + _GridTitleHeight;

            if (_GridHeaderVisible)
                yoff = yoff + _GridHeaderHeight;

            if (e.Y < yoff)
            {
                // we have clicked on the header or the title

                if (_GridHeaderVisible)
                {
                    // have we clicked on the header
                    if (_GridTitleVisible)
                    {
                        if (hs.Visible)
                            xoff = GimmeXOffset(hs.Value) + e.X;
                        else
                            xoff = e.X;

                        if (e.Y > _GridTitleHeight)
                        {
                            r = 0;
                            c = 0;
                            var loopTo = _cols - 1;
                            for (c = 0; c <= loopTo; c++)
                            {
                                r = r + get_ColWidth(c);
                                if (r > xoff)
                                {
                                    // we got the column
                                    if (_SelectedColumn == c)
                                    {
                                        ColumnDeSelected?.Invoke(this, _SelectedColumn);
                                        _SelectedColumn = -1;
                                        Invalidate();
                                    }
                                    else
                                    {
                                        _SelectedColumn = c;
                                        ColumnSelected?.Invoke(this, _SelectedColumn);
                                        Invalidate();
                                    }
                                    break;
                                }
                            }
                        }
                        else
                        {
                            r = 0;
                            c = 0;
                            var loopTo1 = _cols - 1;
                            for (c = 0; c <= loopTo1; c++)
                            {
                                r = r + get_ColWidth(c);
                                if (r > xoff)
                                {
                                    // we got the column
                                    if (_SelectedColumn == c)
                                    {
                                        ColumnDeSelected?.Invoke(this, _SelectedColumn);
                                        _SelectedColumn = -1;
                                        Invalidate();
                                    }
                                    else
                                    {
                                        _SelectedColumn = c;
                                        ColumnSelected?.Invoke(this, _SelectedColumn);
                                        Invalidate();
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                return;
            }

            if (vs.Visible)
                yoff = GimmeYOffset(vs.Value) + e.Y;
            else
                yoff = e.Y;

            if (_GridTitleVisible)
                yoff = yoff - _GridTitleHeight;

            if (_GridHeaderVisible)
                yoff = yoff - _GridHeaderHeight;

            if (hs.Visible)
                xoff = GimmeXOffset(hs.Value) + e.X;
            else
                xoff = e.X;

            _RowClicked = -1;
            _ColClicked = -1;
            r = 0;
            c = 0;
            if (yoff < 0 | !_AllowRowSelection)
            {
            }
            else
            {
                var loopTo2 = _rows - 1;
                for (r = 0; r <= loopTo2; r++)
                {
                    c = c + _rowheights[r];
                    if (c > yoff)
                    {
                        // we got the row
                        _RowClicked = r;

                        if (!_AllowMultipleRowSelections)
                        {
                            // handle like a regular selection 

                            if (_SelectedRow > -1 & _SelectedRow != r)
                                RowDeSelected?.Invoke(this, _SelectedRow);

                            // If _SelectedRow = r Then
                            // '_SelectedRow = -1
                            // '_SelectedRows.Clear()
                            // '_SelectedRows.Add(_SelectedRow)
                            // RaiseEvent RowDeSelected(Me, _SelectedRow)
                            // Me.Invalidate()
                            // Exit For
                            // Else
                            _SelectedRow = r;
                            _SelectedRows.Clear();
                            _SelectedRows.Add(_SelectedRow);
                            // _SelectedRows.Add(_SelectedRow)
                            RowSelected?.Invoke(this, _SelectedRow);
                            Invalidate();
                            break;
                        }
                        else if ((int)ModifierKeys == (int)Keys.Control)
                        {
                            if (_SelectedRows.Contains(r) & (int)ModifierKeys == (int)Keys.Control)
                            {
                                // we need to de select that row here
                                if (_SelectedRow == r)
                                    _SelectedRow = -1;
                                _SelectedRows.Remove(r);
                                RowDeSelected?.Invoke(this, r);
                                Invalidate();
                                break;
                            }
                            else if ((int)ModifierKeys == (int)Keys.Control)
                            {
                                _SelectedRow = r;
                                _SelectedRows.Add(r);
                                RowSelected?.Invoke(this, _SelectedRow);
                                Invalidate();
                                break;
                            }
                            else if (_SelectedRow == r)
                            {
                                RowDeSelected?.Invoke(this, _SelectedRow);
                                _SelectedRow = -1;
                                _SelectedRows.Clear();
                                Invalidate();
                                break;
                            }
                            else
                            {
                                _SelectedRow = r;
                                _SelectedRows.Clear();
                                _SelectedRows.Add(_SelectedRow);
                                RowSelected?.Invoke(this, _SelectedRow);
                                Invalidate();
                                break;
                            }
                        }
                        else if ((int)ModifierKeys == (int)Keys.Shift)
                        {
                            if (_ShiftMultiSelectSelectedRowCrap > -1)
                            {
                                // we have a row selected already

                                _SelectedRows.Clear();

                                if (_ShiftMultiSelectSelectedRowCrap > r)
                                {
                                    var loopTo3 = _ShiftMultiSelectSelectedRowCrap;
                                    for (t = r; t <= loopTo3; t++)
                                        _SelectedRows.Add(t);

                                    _SelectedRow = r;
                                    RowSelected?.Invoke(this, _SelectedRow);
                                    Invalidate();
                                    break;
                                }
                                else
                                {
                                    var loopTo4 = r;
                                    for (t = _ShiftMultiSelectSelectedRowCrap; t <= loopTo4; t++)
                                        _SelectedRows.Add(t);

                                    _SelectedRow = r;
                                    RowSelected?.Invoke(this, _SelectedRow);
                                    Invalidate();
                                    break;
                                }
                            }
                            else
                            {
                                // we dont have a selectedrow already so lets haandle this like a ragular selection
                                if (_SelectedRow > -1 & _SelectedRow != r)
                                    RowDeSelected?.Invoke(this, _SelectedRow);
                                else if (SelectedRow == -1)
                                    _ShiftMultiSelectSelectedRowCrap = r;

                                if (_SelectedRow == r)
                                {
                                    _SelectedRow = -1;
                                    _ShiftMultiSelectSelectedRowCrap = -1;
                                    _SelectedRows.Clear();
                                    // _SelectedRows.Add(_SelectedRow)
                                    RowDeSelected?.Invoke(this, _SelectedRow);
                                    Invalidate();
                                    break;
                                }
                                else
                                {
                                    _SelectedRow = r;
                                    _ShiftMultiSelectSelectedRowCrap = r;
                                    _SelectedRows.Clear();
                                    _SelectedRows.Add(_SelectedRow);
                                    RowSelected?.Invoke(this, _SelectedRow);
                                    Invalidate();
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // handle like a regular selection
                            if (_SelectedRow > -1 & _SelectedRow != r)
                                RowDeSelected?.Invoke(this, _SelectedRow);
                            else if (SelectedRow == -1)
                                _ShiftMultiSelectSelectedRowCrap = r;

                            // If _SelectedRow = r Then
                            // _SelectedRow = -1
                            // _ShiftMultiSelectSelectedRowCrap = -1
                            // _SelectedRows.Clear()
                            // '_SelectedRows.Add(_SelectedRow)
                            // RaiseEvent RowDeSelected(Me, _SelectedRow)
                            // Me.Invalidate()
                            // Exit For
                            // Else
                            _SelectedRow = r;
                            _ShiftMultiSelectSelectedRowCrap = r;
                            _SelectedRows.Clear();
                            _SelectedRows.Add(_SelectedRow);
                            RowSelected?.Invoke(this, _SelectedRow);
                            Invalidate();
                            break;
                        }
                    }
                }
            }

            if (_RowClicked == -1)
                // we did not click on a row so we should bail
                return;

            r = 0;
            c = 0;
            var loopTo5 = _cols - 1;
            for (c = 0; c <= loopTo5; c++)
            {
                r = r + get_ColWidth(c);
                if (r > xoff)
                {
                    // we got the column
                    _ColClicked = c;
                    break;
                }
            }

            if (_ColClicked == -1)
                // we did not click on a a column so lets bail
                return;

            if (Visible)
            {
                CellClicked?.Invoke(this, _RowClicked, _ColClicked);
                if (_RowClicked > -1 & _RowClicked < _rows)
                {
                    if (_ColClicked > -1 & _ColClicked < _cols & _colEditable[_ColClicked] & _AllowInGridEdits)
                    {
                        if (IsColumnRestricted(_ColClicked))
                        {
                            var it = GetColumnRestriction(_ColClicked);

                            cmboInput.Items.Clear();

                            var s = it.RestrictedList.Split("^".ToCharArray());
                            foreach (string ss1 in s)
                                cmboInput.Items.Add(ss1);

                            // we have selected a row and col lets move the txtinput there and bring it to the front
                            xoff = 0;
                            yoff = 0;

                            if (_RowClicked > 0)
                            {
                                var loopTo6 = _RowClicked - 1;
                                for (r = 0; r <= loopTo6; r++)
                                    yoff = yoff + get_RowHeight(r);
                            }

                            if (GridheaderVisible)
                                yoff = yoff + _GridHeaderHeight;

                            if (_GridTitleVisible)
                                yoff = yoff + _GridTitleHeight;

                            if (_ColClicked > 0)
                            {
                                var loopTo7 = _ColClicked - 1;
                                for (c = 0; c <= loopTo7; c++)
                                    xoff = xoff + get_ColWidth(c);
                            }

                            if (vs.Visible & vs.Value > 0)
                                yoff = yoff - GimmeYOffset(vs.Value);

                            if (hs.Visible & hs.Value > 0)
                                xoff = xoff - GimmeXOffset(hs.Value);

                            if (_CellOutlines)
                            {
                                cmboInput.Top = yoff + 1;
                                cmboInput.Left = xoff + 1;
                                cmboInput.Width = get_ColWidth(_ColClicked) - 1;
                                cmboInput.Height = get_RowHeight(_RowClicked) - 2;
                                cmboInput.BackColor = _colEditableTextBackColor;
                            }
                            else
                            {
                                cmboInput.Top = yoff;
                                cmboInput.Left = xoff;
                                cmboInput.Width = get_ColWidth(_ColClicked);
                                cmboInput.Height = get_RowHeight(_RowClicked);
                                cmboInput.BackColor = _colEditableTextBackColor;
                            }

                            cmboInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                            cmboInput.Text = _grid[_RowClicked, _ColClicked];

                            cmboInput.Visible = true;
                            cmboInput.BringToFront();
                            cmboInput.DroppedDown = true;
                            _EditModeCol = _ColClicked;
                            _EditModeRow = _RowClicked;
                            _EditMode = true;

                            cmboInput.Focus();
                        }
                        else if (_colboolean[_ColClicked])
                        {
                            // we have clicked on a boolean editable cell lets flip those bits baby

                            ss = Strings.Trim(Strings.UCase(_grid[_RowClicked, _ColClicked]));

                            switch (ss)
                            {
                                case "TRUE":
                                    {
                                        ss = "FALSE";
                                        break;
                                    }

                                case "FALSE":
                                    {
                                        ss = "TRUE";
                                        break;
                                    }

                                case "YES":
                                    {
                                        ss = "NO";
                                        break;
                                    }

                                case "NO":
                                    {
                                        ss = "YES";
                                        break;
                                    }

                                case "1":
                                    {
                                        ss = "0";
                                        break;
                                    }

                                case "0":
                                    {
                                        ss = "1";
                                        break;
                                    }

                                case "Y":
                                    {
                                        ss = "N";
                                        break;
                                    }

                                case "N":
                                    {
                                        ss = "Y";
                                        break;
                                    }

                                default:
                                    {
                                        ss = "";
                                        break;
                                    }
                            }

                            _grid[_RowClicked, _ColClicked] = ss;

                            Refresh();
                        }
                        else
                        {
                            // we have selected a row and col lets move the txtinput there and bring it to the front
                            xoff = 0;
                            yoff = 0;

                            if (_RowClicked > 0)
                            {
                                var loopTo8 = _RowClicked - 1;
                                for (r = 0; r <= loopTo8; r++)
                                    yoff = yoff + get_RowHeight(r);
                            }

                            if (GridheaderVisible)
                                yoff = yoff + _GridHeaderHeight;

                            if (_GridTitleVisible)
                                yoff = yoff + _GridTitleHeight;

                            if (_ColClicked > 0)
                            {
                                var loopTo9 = _ColClicked - 1;
                                for (c = 0; c <= loopTo9; c++)
                                    xoff = xoff + get_ColWidth(c);
                            }

                            if (vs.Visible & vs.Value > 0)
                                yoff = yoff - GimmeYOffset(vs.Value);

                            if (hs.Visible & hs.Value > 0)
                                xoff = xoff - GimmeXOffset(hs.Value);

                            if (_CellOutlines)
                            {
                                txtInput.Top = yoff + 1;
                                txtInput.Left = xoff + 1;
                                txtInput.Width = get_ColWidth(_ColClicked) - 1;
                                txtInput.Height = get_RowHeight(_RowClicked) - 2;
                                txtInput.BackColor = _colEditableTextBackColor;
                            }
                            else
                            {
                                txtInput.Top = yoff;
                                txtInput.Left = xoff;
                                txtInput.Width = get_ColWidth(_ColClicked);
                                txtInput.Height = get_RowHeight(_RowClicked);
                                txtInput.BackColor = _colEditableTextBackColor;
                            }

                            txtInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                            txtInput.Text = _grid[_RowClicked, _ColClicked];

                            txtInput.Visible = true;
                            txtInput.BringToFront();
                            _EditModeCol = _ColClicked;
                            _EditModeRow = _RowClicked;
                            _EditMode = true;

                            txtInput.Focus();
                        }
                    }
                }
            }
        }

        private bool IsColumnRestricted(int colid)
        {
            bool ret = false;

            foreach (EditColumnRestrictor it in _colEditRestrictions)
            {
                if (it.ColumnID == colid)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }

        private EditColumnRestrictor GetColumnRestriction(int colid)
        {
            var ret = new EditColumnRestrictor();

            foreach (EditColumnRestrictor it in _colEditRestrictions)
            {
                if (it.ColumnID == colid)
                {
                    ret = it;
                    break;
                }
            }

            return ret;
        }

        private void DoubleClickHandler(object sender, EventArgs e)
        {

            // Console.WriteLine("MouseDoubleClick")
            if (_RowClicked != -1 & _ColClicked != -1 & Visible)
            {
                CellDoubleClicked?.Invoke(this, _RowClicked, _ColClicked);
                _DoubleClickSemaphore = true;
            }
        }

        private void txtHandler_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void txtHandler_KeyDown(object sender, KeyEventArgs e)
        {

            // we are NOT in editmode so lets handle this like any other

            KeyPressedInGrid?.Invoke(this, e.KeyCode);

            if ((int)e.KeyCode == (int)Keys.Return | (int)e.KeyCode == (int)Keys.Enter)
            {
                if (_SelectedRow != -1 & Visible)
                    CellDoubleClicked?.Invoke(this, _SelectedRow, 0);
            }

            if ((int)e.KeyCode == (int)Keys.Left)
            {
                if (hs.Visible)
                {
                    int v;
                    v = hs.Value;
                    v -= 1;
                    if (v < 0)
                        v = 0;
                    hs.Value = v;
                }
            }

            if ((int)e.KeyCode == (int)Keys.Right)
            {
                if (hs.Visible)
                {
                    int v;
                    v = hs.Value;
                    v += 1;
                    if (v >= hs.Maximum)
                        v = hs.Maximum - 1;
                    hs.Value = v;
                }
            }

            if ((int)e.KeyCode == (int)Keys.PageDown)
            {
                RowDeSelected?.Invoke(this, _SelectedRow);

                _SelectedRow = _SelectedRow + 10;

                if (_SelectedRow >= _rows)
                    _SelectedRow = _rows - 1;

                _SelectedRows.Clear();
                _SelectedRows.Add(_SelectedRow);

                if (vs.Visible)
                {
                    bool flag = true;
                    int x, xx;
                    while (flag)
                    {
                        x = GimmeYOffset(_SelectedRow);
                        x = x - GimmeYOffset(vs.Value);

                        if (_GridTitleVisible)
                            xx = _GridTitleHeight;
                        else
                            xx = 0;

                        if (_GridHeaderVisible)
                            xx = xx + _GridHeaderHeight;

                        if (hs.Visible)
                            xx = xx + hs.Height;

                        xx = xx + _rowheights[_SelectedRow];

                        if (x < Height - xx)
                            flag = false;
                        else
                            vs.Value = vs.Value + 1;
                    }
                }

                Invalidate();

                RowSelected?.Invoke(this, _SelectedRow);
            }


            if ((int)e.KeyCode == (int)Keys.Down)
            {
                RowDeSelected?.Invoke(this, _SelectedRow);

                _SelectedRow = _SelectedRow + 1;

                if (_SelectedRow >= _rows)
                    _SelectedRow = _rows - 1;

                _SelectedRows.Clear();
                _SelectedRows.Add(_SelectedRow);

                if (vs.Visible)
                {
                    bool flag = true;
                    int x, xx;
                    while (flag)
                    {
                        x = GimmeYOffset(_SelectedRow);
                        x = x - GimmeYOffset(vs.Value);

                        if (_GridTitleVisible)
                            xx = _GridTitleHeight;
                        else
                            xx = 0;

                        if (_GridHeaderVisible)
                            xx = xx + _GridHeaderHeight;

                        if (hs.Visible)
                            xx = xx + hs.Height;

                        xx = xx + _rowheights[_SelectedRow];

                        if (x < Height - xx)
                            flag = false;
                        else
                            vs.Value = vs.Value + 1;
                    }
                }

                Invalidate();

                RowSelected?.Invoke(this, _SelectedRow);
            }

            if ((int)e.KeyCode == (int)Keys.PageUp)
            {
                RowDeSelected?.Invoke(this, _SelectedRow);

                _SelectedRow = _SelectedRow - 10;
                if (_SelectedRow < 0)
                    _SelectedRow = 0;

                _SelectedRows.Clear();
                _SelectedRows.Add(_SelectedRow);

                if (vs.Visible)
                {
                    bool flag = true;
                    while (flag)
                    {
                        if (_SelectedRow >= vs.Value)
                            flag = false;
                        else
                            vs.Value = vs.Value - 1;
                    }
                }

                Invalidate();

                RowSelected?.Invoke(this, _SelectedRow);
            }

            if ((int)e.KeyCode == (int)Keys.Up)
            {
                RowDeSelected?.Invoke(this, _SelectedRow);

                _SelectedRow = _SelectedRow - 1;

                if (_SelectedRow < 0)
                    _SelectedRow = 0;

                _SelectedRows.Clear();
                _SelectedRows.Add(_SelectedRow);

                if (vs.Visible)
                {
                    bool flag = true;
                    while (flag)
                    {
                        if (_SelectedRow >= vs.Value)
                            flag = false;
                        else
                            vs.Value = vs.Value - 1;
                    }
                }

                Invalidate();

                RowSelected?.Invoke(this, _SelectedRow);
            }

            if ((int)e.Modifiers == (int)Keys.Control)
            {
                if ((int)e.KeyCode == (int)Keys.F)
                {
                    if (!string.IsNullOrEmpty(_LastSearchText) & _LastSearchColumn != -1 & _SelectedRow != -1)
                    {
                        int t;

                        if (_SelectedRow + 1 >= _rows)
                            // we are on the last row already flip to the first row
                            _SelectedRow = 0;

                        bool found = false;
                        var loopTo = _rows - 1;
                        for (t = _SelectedRow + 1; t <= loopTo; t++)
                        {
                            if (Strings.InStr(Strings.UCase(_grid[t, _LastSearchColumn]), Strings.UCase(_LastSearchText), CompareMethod.Text) != 0)
                            {
                                // we have a match
                                if (vs.Visible)
                                {
                                    vs.Value = t;
                                    found = true;
                                    _SelectedRow = t;
                                    Invalidate();
                                    e.Handled = true;
                                    break;
                                }
                                else
                                {
                                    found = true;
                                    _SelectedRow = t;
                                    Invalidate();
                                    e.Handled = true;
                                    break;
                                }
                            }
                        }
                        if (!found)
                        {
                            e.Handled = true;
                            Interaction.MsgBox("Search found nothing further");
                        }
                    }
                }
            }
        }

        private void vs_Scroll(object sender, ScrollEventArgs e)
        {
        }

        private void MouseDownHandler(object sender, MouseEventArgs e)
        {
            Point p;
            int x, y, xx, yy, r, c = default(int), rr, cc, yoff;


            // Console.WriteLine("MouseDOWN")

            xx = -1;
            yy = -1;

            // handler for leftmousebuttons on column resizing
            if ((int)e.Button == (int)MouseButtons.Left & _AllowUserColumnResizing & _GridHeaderVisible)
            {
                p = PointToClient(MousePosition);
                x = p.X; // + Math.Abs(Me.Left)
                y = p.Y; // + Math.Abs(Me.Top)

                // is the title visible
                if (_GridTitleVisible)
                    rr = _GridTitleHeight;
                else
                    rr = 0;

                // adjust for the hs scroll bar
                if (hs.Visible)
                    x = x + GimmeXOffset(hs.Value);

                // first of all are we actually in the header

                if (y >= rr & y <= rr + _GridHeaderHeight)
                {
                    // yes we are lets setup for resizing

                    cc = 0;
                    var loopTo = _cols - 1;
                    for (c = 0; c <= loopTo; c++)
                    {
                        if (cc <= x & cc + _colwidths[c] >= x)
                        {
                            // we have the column
                            xx = c;
                            break;
                        }
                        else
                            cc = cc + _colwidths[c];
                    }

                    _ColOverOnMenuButton = -1;
                    _ColOverOnMouseDown = xx;

                    _MouseDownOnHeader = true;

                    _LastMouseX = MousePosition.X;
                    _LastMouseY = MousePosition.Y;
                    _AutosizeCellsToContents = false;
                    Cursor = Cursors.SizeWE;
                }
                else
                {
                    // No we aren't so lets NOT resize

                    _MouseDownOnHeader = false;
                    _LastMouseX = -1;
                    _LastMouseY = -1;
                    _ColOverOnMenuButton = -1;
                    _ColOverOnMouseDown = -1;
                }
            }

            // handler for rightmousebuttons and popupmenus and allowing / disallowing ctrl key menus

            if ((int)e.Button == (int)MouseButtons.Right & (_AllowPopupMenu | (int)ModifierKeys == (int)Keys.Control & _AllowControlKeyMenuPopup))
            {
                if (!(ContextMenu == null))
                    _OldContextMenu = ContextMenu;

                p = PointToClient(MousePosition);
                x = p.X + Math.Abs(Left);
                y = p.Y + Math.Abs(Top);

                if (vs.Visible)
                    yoff = GimmeYOffset(vs.Value) + e.Y;
                else
                    yoff = e.Y;

                if (_GridTitleVisible)
                    yoff = yoff - _GridTitleHeight;

                if (_GridHeaderVisible)
                    yoff = yoff - _GridHeaderHeight;

                if (yoff < 0)
                    // we have clicked on the header or the title area so we should skip the row section
                    _RowOverOnMenuButton = -1;
                else
                {
                    var loopTo1 = _rows - 1;
                    for (r = 0; r <= loopTo1; r++)
                    {
                        c = c + _rowheights[r];
                        if (c > yoff)
                        {
                            // we got the row
                            _RowOverOnMenuButton = r;
                            break;
                        }
                    }
                }

                // adjust for the hs scroll bar
                if (hs.Visible)
                    x = x + GimmeXOffset(hs.Value);

                // adjust for the vs scroll bar
                if (vs.Visible)
                    y = y + GimmeYOffset(vs.Value);

                cc = Math.Abs(Left);
                var loopTo2 = _cols - 1;
                for (c = 0; c <= loopTo2; c++)
                {
                    if (cc <= x & cc + _colwidths[c] >= x)
                    {
                        // we have the column
                        xx = c;
                        break;
                    }
                    else
                        cc = cc + _colwidths[c];
                }

                // If _GridTitleVisible Then
                // If _GridHeaderVisible Then
                // ' we have the header and the titlevisible

                // End If
                // End If

                _ColOverOnMenuButton = xx;

                if (ContextMenu == null)
                {
                    miStats.Text = "Rows = " + _rows.ToString() + " : Cols = " + _cols.ToString();

                    menu.Show(this, p);
                }
                else
                    ContextMenu.Show(this, p);
                // menu.Show(Me, p)
                return;
            }
            else if ((int)e.Button == (int)MouseButtons.Right)
            {
                RightMouseButtonInGrid?.Invoke(this);
                return;
            }
        }

        private void MouseMoveHandler(object sender, MouseEventArgs e)
        {
            int x, delta;

            if (_LMouseX == e.X & _LMouseY == e.Y)
                return;
            else
            {
                _LMouseX = e.X;
                _LMouseY = e.Y;
            }

            if (_MouseDownOnHeader & _ColOverOnMouseDown > -1 & _AllowUserColumnResizing & _ColOverOnMouseDown < _cols)
            {
                x = MousePosition.X;

                // calculate Deltas
                if (x >= _LastMouseX)
                    delta = x - _LastMouseX;
                else
                    delta = -(_LastMouseX - x);
                _LastMouseX = x;

                _colwidths[_ColOverOnMouseDown] = _colwidths[_ColOverOnMouseDown] + delta;

                if (_colwidths[_ColOverOnMouseDown] < _UserColResizeMinimum)
                    _colwidths[_ColOverOnMouseDown] = _UserColResizeMinimum;

                ColumnResized?.Invoke(this, _ColOverOnMouseDown);

                _AutoSizeAlreadyCalculated = false;

                Invalidate();
            }
            else
                MouseHoverHandler(sender, new EventArgs());
        }

        private void MouseHoverHandler(object sender, EventArgs e)
        {
            var mp = PointToClient(MousePosition);

            int xoff, yoff, r, c, row, col;

            if (_GridTitleVisible & mp.Y <= _GridTitleHeight)
            {
                // we are hovering over the grid title
                GridHoverleave?.Invoke(this);
                return;
            }

            if (_GridTitleVisible & _GridHeaderVisible & mp.Y <= _GridTitleHeight + _GridHeaderHeight)
            {
                // we are hovering over the Header or the title so lets bail
                GridHoverleave?.Invoke(this);
                return;
            }

            if (vs.Visible)
                yoff = GimmeYOffset(vs.Value) + mp.Y;
            else
                yoff = mp.Y;

            if (_GridTitleVisible)
                yoff = yoff - _GridTitleHeight;

            if (_GridHeaderVisible)
                yoff = yoff - _GridHeaderHeight;

            if (hs.Visible)
                xoff = GimmeXOffset(hs.Value) + mp.X;
            else
                xoff = mp.X;

            // here xoff and yoff are converted to real grid coordinates if the exist

            r = 0;
            c = 0;
            row = -1;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                c = c + _rowheights[r];
                if (c > yoff)
                {
                    // we got the row
                    row = r;
                    break;
                }
            }

            r = 0;
            c = 0;
            col = -1;
            var loopTo1 = _cols - 1;
            for (c = 0; c <= loopTo1; c++)
            {
                r = r + get_ColWidth(c);
                if (r > xoff)
                {
                    // we got the column
                    col = c;
                    break;
                }
            }

            if (row > -1 & col > -1)
                // we gots a winner

                GridHover?.Invoke(this, row, col, _grid[row, col]);
            else
                // we gots a loser
                GridHoverleave?.Invoke(this);
        }

        private void TAIGRIDControl_Load(object sender, EventArgs e)
        {
            _gridCellFontsList[0] = _DefaultCellFont;
            _gridForeColorList[0] = new Pen(_DefaultForeColor);
            _gridCellAlignmentList[0] = _DefaultStringFormat;
            _gridBackColorList[0] = new SolidBrush(_DefaultBackColor);
        }

        private void pdoc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int x, y, xx, r, c;
            var fnt = new Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Regular, GraphicsUnit.Pixel);
            var fnt2 = new Font("Courier New", 10 * _gridReportScaleFactor, FontStyle.Bold, GraphicsUnit.Pixel);

            var tfnt = new Font("Courier New", 8, FontStyle.Regular, GraphicsUnit.Pixel);


            float m;

            Font ft;

            var greypen = new Pen(Color.Gray);

            int pagewidth = e.PageSettings.Bounds.Size.Width;
            int pageheight = e.PageSettings.Bounds.Size.Height;

            int lrmargin = 40;
            int tbmargin = 70;

            bool colprintedonpage = false;

            if (AllColWidths() * _gridReportScaleFactor < pagewidth - 2 * lrmargin)
                xx = Conversions.ToInteger((pagewidth - 2 * lrmargin - AllColWidths() * _gridReportScaleFactor) / 2);
            else
                xx = 0;


            var rect = new RectangleF(0, 0, 1, 1);

            x = lrmargin;
            y = tbmargin;

            int coloffset = 0;
            bool morecols = true;
            int currow = _gridReportCurrentrow;


            ft = _GridHeaderFont;

            ft = new Font(_GridHeaderFont.FontFamily, (_GridHeaderFont.SizeInPoints - 1) * _gridReportScaleFactor, _GridHeaderFont.Style, _GridHeaderFont.Unit);

            // calculate size and place the printed on date on the page

            m = e.Graphics.MeasureString(_gridReportPrintedOn.ToLongDateString() + Constants.vbCrLf
                                        + _gridReportPrintedOn.ToLongTimeString(), fnt).Width;

            e.Graphics.DrawString(_gridReportPrintedOn.ToLongDateString() + Constants.vbCrLf + _gridReportPrintedOn.ToLongTimeString(), fnt, Brushes.Black, pagewidth - m - lrmargin, Conversions.ToSingle(tbmargin / (double)2));

            if (!string.IsNullOrEmpty(_gridReportTitle))
            {
                // we want to title each page here

                string ttit = _gridReportTitle;

                if (ttit.Length > 98)
                {
                    var ttitarray = ttit.Split(" ".ToCharArray());

                    int ttitidx, curlen;
                    ttit = "";
                    curlen = 0;
                    var loopTo = ttitarray.GetUpperBound(0);
                    for (ttitidx = 0; ttitidx <= loopTo; ttitidx++)
                    {
                        ttit += ttitarray[ttitidx] + " ";

                        curlen += ttitarray[ttitidx].Length;

                        if (curlen > 98)
                        {
                            ttit += Constants.vbCrLf;
                            curlen = 0;
                        }
                    }
                }

                e.Graphics.DrawString(ttit, tfnt, Brushes.Black, lrmargin, Conversions.ToSingle(tbmargin / (double)2));
            }
            else
            {
                string ttit = _GridTitle;

                if (ttit.Length > 98)
                {
                    var ttitarray = ttit.Split(" ".ToCharArray());

                    int ttitidx, curlen;
                    ttit = "";
                    curlen = 0;
                    var loopTo1 = ttitarray.GetUpperBound(0);
                    for (ttitidx = 0; ttitidx <= loopTo1; ttitidx++)
                    {
                        ttit += ttitarray[ttitidx] + " ";

                        curlen += ttitarray[ttitidx].Length;

                        if (curlen > 98)
                        {
                            ttit += Constants.vbCrLf;
                            curlen = 0;
                        }
                    }
                }

                e.Graphics.DrawString(ttit, tfnt, Brushes.Black, lrmargin, Conversions.ToSingle(tbmargin / (double)2));
            }

            if (_gridReportNumberPages)
            {
                // we want to number the pages here

                m = e.Graphics.MeasureString("Page " + _gridReportPageNumbers.ToString(), fnt).Height;


                e.Graphics.DrawString("Page " + _gridReportPageNumbers.ToString(), fnt, Brushes.Black, lrmargin, pageheight - tbmargin + m + 2);
            }

            var loopTo2 = Cols - 1;

            // print the grid header

            for (c = _gridReportCurrentColumn; c <= loopTo2; c++)
            {
                if (x + _colwidths[c] + xx > pagewidth - lrmargin & colprintedonpage)
                    break;

                colprintedonpage = true;

                if (_gridReportMatchColors)
                    e.Graphics.FillRectangle(new SolidBrush(_GridHeaderBackcolor), x + xx, y, _colwidths[c], _GridHeaderHeight);

                rect.X = Convert.ToSingle(x + xx);
                rect.Y = Convert.ToSingle(y);
                rect.Width = Convert.ToSingle(_colwidths[c]);
                rect.Height = Convert.ToSingle(_GridHeaderHeight);



                e.Graphics.DrawString(_GridHeader[c], ft, Brushes.Black, rect, _GridHeaderStringFormat);

                if (_gridReportOutlineCells)
                    e.Graphics.DrawRectangle(greypen, x + xx, y, _colwidths[c], _GridHeaderHeight);

                x = x + _colwidths[c];
            }


            y += _GridHeaderHeight;
            x = lrmargin;
            var loopTo3 = Rows - 1;
            for (r = _gridReportCurrentrow; r <= loopTo3; r++)
            {
                var loopTo4 = Cols - 1;
                for (c = _gridReportCurrentColumn; c <= loopTo4; c++)
                {
                    if (x + _colwidths[c] + xx > pagewidth - lrmargin & colprintedonpage)
                    {
                        coloffset = c;
                        morecols = true;
                        break;
                    }
                    else
                        morecols = false;

                    colprintedonpage = true;

                    if (_gridReportMatchColors)
                        e.Graphics.FillRectangle(_gridBackColorList[_gridBackColor[r, c]], x + xx, y, _colwidths[c], _rowheights[r]);

                    rect.X = Convert.ToSingle(x + xx);
                    rect.Y = Convert.ToSingle(y);
                    rect.Width = Convert.ToSingle(_colwidths[c]);
                    rect.Height = Convert.ToSingle(_rowheights[r]);

                    ft = new Font(_gridCellFontsList[_gridCellFonts[r, c]].FontFamily, _gridCellFontsList[_gridCellFonts[r, c]].SizeInPoints - 1, _gridCellFontsList[_gridCellFonts[r, c]].Style, _gridCellFontsList[_gridCellFonts[r, c]].Unit);

                    e.Graphics.DrawString(_grid[r, c], ft, Brushes.Black, rect, _gridCellAlignmentList[_gridCellAlignment[r, c]]);

                    if (_gridReportOutlineCells)
                        e.Graphics.DrawRectangle(greypen, x + xx, y, _colwidths[c], _rowheights[r]);

                    x = x + _colwidths[c];
                }
                x = lrmargin;
                y += _rowheights[r];
                _gridReportCurrentrow += 1;

                // do we need to skip to next page here
                if (y >= pageheight - tbmargin)
                    break;
                else
                {
                }
            }


            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(_psets.PrinterSettings.PrintRange, System.Drawing.Printing.PrintRange.SomePages, false)))
            {
                if (_gridReportCurrentrow >= Rows - 1 & !morecols | _gridReportPageNumbers >= _gridEndPage)
                {
                    e.HasMorePages = false;
                    _gridReportPageNumbers = 1;
                    _gridReportCurrentrow = 0;
                    _gridReportCurrentColumn = 0;
                }
                else
                {
                    if (morecols)
                    {
                        _gridReportCurrentColumn = coloffset;
                        _gridReportCurrentrow = currow;
                    }
                    else
                        _gridReportCurrentColumn = 0;
                    e.HasMorePages = true;
                    _gridReportPageNumbers += 1;
                }
            }
            else if (_gridReportCurrentrow >= Rows - 1 & !morecols)
            {
                e.HasMorePages = false;
                _gridReportPageNumbers = 1;
                _gridReportCurrentrow = 0;
                _gridReportCurrentColumn = 0;
            }
            else
            {
                if (morecols)
                {
                    _gridReportCurrentColumn = coloffset;
                    _gridReportCurrentrow = currow;
                }
                else
                    _gridReportCurrentColumn = 0;
                e.HasMorePages = true;
                _gridReportPageNumbers += 1;
            }
        }

        private void PageOrientationChange(bool lsorientation)
        {
            bool oldorientation = Conversions.ToBoolean(_psets.Landscape);

            _psets.Landscape = lsorientation;


            _PageSetupForm.MaxPage = CalculatePageRange();

            _psets.Landscape = oldorientation;
        }

        private void PageSetupChange(System.Drawing.Printing.PaperSize psiz)
        {
            try
            {
                System.Drawing.Printing.PaperSize ps = _psets.PaperSize;

                _psets.PaperSize = psiz;

                _PageSetupForm.MaxPage = CalculatePageRange();

                _psets.PaperSize = ps;
            }
            catch (Exception ex)
            {
            }
        }

        private void PageMetricsChange(System.Drawing.Printing.PaperSize psiz, bool lsorientation)
        {
            try
            {
                LogThis("Inside the event handler for Page Meterics changing...");

                System.Drawing.Printing.PaperSize ps = _psets.PaperSize;


                _psets.PaperSize = psiz;
                _psets.Landscape = lsorientation;
                _PageSetupForm.MaxPage = CalculatePageRange();
                _psets.PaperSize = ps;
            }
            catch (Exception ex)
            {
            }
        }

        private void txtInput_Leave(object sender, EventArgs e)
        {
            if ((int)_GridEditMode == (int)GridEditModes.LostFocus)
            {
                if ((_grid[_RowClicked, _ColClicked] ?? "") != (txtInput.Text ?? ""))
                {
                    string oldval = _grid[_RowClicked, _ColClicked];
                    string newval = txtInput.Text;
                    _grid[_RowClicked, _ColClicked] = txtInput.Text;
                    CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                }

                txtInput.SendToBack();
                txtInput.Visible = false;

                Invalidate();
            }
            else
            {
                txtInput.Visible = false;
                _EditModeCol = -1;
                _EditModeRow = -1;
                _EditMode = false;
            }
        }

        private void txtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if ((int)e.KeyCode == (int)Keys.Return & _GridEditMode == (int)GridEditModes.KeyReturn)
            {
                if ((_grid[_RowClicked, _ColClicked] ?? "") != (txtInput.Text ?? ""))
                {
                    string oldval = _grid[_RowClicked, _ColClicked];
                    string newval = txtInput.Text;
                    _grid[_RowClicked, _ColClicked] = txtInput.Text;
                    CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                }

                txtInput.SendToBack();
                txtInput.Visible = false;

                Invalidate();
                e.Handled = false;
            }

            if ((int)e.KeyCode == (int)Keys.Tab)
            {
                Console.WriteLine("Tab Key Pressed");
                e.Handled = true;
            }
        }

        private void cmboInput_Leave(object sender, EventArgs e)
        {
            if (GridEditMode == (int)GridEditModes.KeyReturn)
            {
                if ((_grid[_RowClicked, _ColClicked] ?? "") != (cmboInput.Text ?? ""))
                {
                    string oldval = _grid[_RowClicked, _ColClicked];
                    string newval = cmboInput.Text;
                    _grid[_RowClicked, _ColClicked] = cmboInput.Text;
                    CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                }

                cmboInput.SendToBack();
                cmboInput.Visible = false;

                Invalidate();
            }
            else
            {
                cmboInput.Visible = false;
                _EditModeCol = -1;
                _EditModeRow = -1;
                _EditMode = false;
            }
        }

        private void cmboInput_keyDown(object sender, KeyEventArgs e)
        {
            if ((int)e.KeyCode == (int)Keys.Return & _GridEditMode == (int)GridEditModes.KeyReturn)
            {
                if ((_grid[_RowClicked, _ColClicked] ?? "") != (cmboInput.Text ?? ""))
                {
                    string oldval = _grid[_RowClicked, _ColClicked];
                    string newval = cmboInput.Text;
                    _grid[_RowClicked, _ColClicked] = cmboInput.Text;
                    CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                }

                cmboInput.SendToBack();
                cmboInput.Visible = false;

                Invalidate();
                e.Handled = false;
            }
        }

        private void cmboInput_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((_grid[_RowClicked, _ColClicked] ?? "") != (cmboInput.Text ?? ""))
            {
                string oldval = _grid[_RowClicked, _ColClicked];
                string newval = cmboInput.Text;
                _grid[_RowClicked, _ColClicked] = cmboInput.Text;
                CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
            }

            cmboInput.SendToBack();
            cmboInput.Visible = false;

            Invalidate();
        }

        private void TAIGridControl_HandleDestroyed(object sender, EventArgs e)
        {
            // implemented to to destroy all tearaways if the parent grid gets destroyed

            KillAllTearAwayColumnWindows();
        }

        private void menu_Popup(object sender, EventArgs e)
        {
            if (Antialias)
                miSmoothing.Checked = true;
            else
                miSmoothing.Checked = false;

            if (ExcelAutoFitColumn)
                miAutoFitCols.Checked = true;
            else
                miAutoFitCols.Checked = false;

            if (ExcelAutoFitRow)
                miAutoFitRows.Checked = true;
            else
                miAutoFitRows.Checked = false;

            if (ExcelUseAlternateRowColor)
                miALternateRowColors.Checked = true;
            else
                miALternateRowColors.Checked = false;

            if (ExcelOutlineCells)
                miOutlineExportedCells.Checked = true;
            else
                miOutlineExportedCells.Checked = false;

            if (ExcelMatchGridColorScheme)
                miMatchGridColors.Checked = true;
            else
                miMatchGridColors.Checked = false;

            if (_ColOverOnMenuButton != -1)
                miSearchInColumn.Enabled = true;
            else
                miSearchInColumn.Enabled = false;

            if (_AllowUserColumnResizing)
                miAllowUserColumnResizing.Checked = true;
            else
                miAllowUserColumnResizing.Checked = false;


            miExportToExcelMenu.Enabled = _AllowExcelFunctionality;
            miTearColumnAway.Enabled = _AllowTearAwayFuncionality;
            miArrangeTearAways.Enabled = _AllowTearAwayFuncionality;
            miHideAllTearAwayColumns.Enabled = _AllowTearAwayFuncionality;
            miHideColumnTearAway.Enabled = _AllowTearAwayFuncionality;
            miMultipleColumnTearAway.Enabled = _AllowTearAwayFuncionality;

            miExportToTextFile.Enabled = _AllowTextFunctionality;
            miExportToHTMLTable.Enabled = _AllowHTMLFunctionality;
            miExportToSQLScript.Enabled = _AllowSQLScriptFunctionality;

            MenuItem5.Enabled = _AllowMathFunctionality;
            miFormatStuff.Enabled = _AllowFormatFunctionality;
            MenuItem2.Enabled = _AllowSettingsFunctionality;
            MenuItem3.Enabled = _AllowSortFunctionality;
            MenuItem4.Enabled = _AllowRowAndColumnFunctionality;
        }

        private void miSmoothing_Click(object sender, EventArgs e)
        {
            Antialias = !miSmoothing.Checked;
        }

        private void miFontsLarger_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = DefaultCellFont;

            if (fnt.Size > 72)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit);

            AllCellsUseThisFont(fnt2);
        }

        private void miFontsSmaller_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = DefaultCellFont;

            if (fnt.Size < 4)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit);

            AllCellsUseThisFont(fnt2);
        }

        private void miFormatAsMoney_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat();
            int c = _ColOverOnMenuButton;
            string a;

            if (c >= _cols | _rows < 2 | c < 0)
                return;

            // sf.LineAlignment = StringAlignment.Far

            sf.LineAlignment = StringAlignment.Near;
            sf.Alignment = StringAlignment.Far;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                if (Information.IsNumeric(_grid[r, c]))
                {
                    a = _grid[r, c];
                    if (a.StartsWith("$"))
                        a = a.Substring(1);
                    _grid[r, c] = Strings.Format(Conversion.Val(a), "C");
                    _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);
                }
            }

            Refresh();
        }

        private void miFormatAsDecimal_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat();
            int c = _ColOverOnMenuButton;
            string a;

            if (c >= _cols | _rows < 2 | c < 0)
                return;

            // sf.LineAlignment = StringAlignment.Far

            sf.LineAlignment = StringAlignment.Near;
            sf.Alignment = StringAlignment.Far;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)
            {
                if (Information.IsNumeric(_grid[r, c]))
                {
                    a = _grid[r, c];
                    if (a.StartsWith("$"))
                        a = a.Substring(1);
                    _grid[r, c] = Strings.Format(Conversion.Val(a), "G");
                    _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);
                }
            }

            Refresh();
        }

        private void miFormatAsText_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat(StringFormatFlags.FitBlackBox);
            int c = _ColOverOnMenuButton;


            if (c >= _cols | _rows < 2 | c < 0)
                return;

            sf.LineAlignment = StringAlignment.Far;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)

                // _grid(r, c) = Format(Val(_grid(r, c)), "C")
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);

            Refresh();
        }

        private void miCenter_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat();
            int c = _ColOverOnMenuButton;


            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;

            if (c >= _cols | _rows < 2 | c < 0)
                return;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)

                // _grid(r, c) = Format(Val(_grid(r, c)), "C")
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);

            Refresh();
        }

        private void miLeft_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat();
            int c = _ColOverOnMenuButton;


            sf.LineAlignment = StringAlignment.Near;
            sf.Alignment = StringAlignment.Near;

            if (c >= _cols | _rows < 2 | c < 0)
                return;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)

                // _grid(r, c) = Format(Val(_grid(r, c)), "C")
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);

            Refresh();
        }

        private void miRight_Click(object sender, EventArgs e)
        {
            int r;
            var sf = new StringFormat();
            int c = _ColOverOnMenuButton;


            sf.LineAlignment = StringAlignment.Far;
            sf.Alignment = StringAlignment.Far;

            if (c >= _cols | _rows < 2 | c < 0)
                return;
            var loopTo = _rows - 1;
            for (r = 0; r <= loopTo; r++)

                // _grid(r, c) = Format(Val(_grid(r, c)), "C")
                _gridCellAlignment[r, c] = GetGridCellAlignmentListEntry(sf);

            Refresh();
        }

        private void miExportToExcel_Click_1(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void miAutoFitCols_Click(object sender, EventArgs e)
        {
            ExcelAutoFitColumn = !ExcelAutoFitColumn;
        }

        private void miAutoFitRows_Click(object sender, EventArgs e)
        {
            ExcelAutoFitRow = !ExcelAutoFitRow;
        }

        private void miALternateRowColors_Click(object sender, EventArgs e)
        {
            ExcelUseAlternateRowColor = !ExcelUseAlternateRowColor;
        }

        private void miMatchGridColors_Click(object sender, EventArgs e)
        {
            ExcelMatchGridColorScheme = !ExcelMatchGridColorScheme;
        }

        private void miOutlineExportedCells_Click(object sender, EventArgs e)
        {
            ExcelOutlineCells = !ExcelOutlineCells;
        }

        private void miExportToTextFile_Click(object sender, EventArgs e)
        {
            ExportToText();
        }

        private void miHeaderFontSmaller_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = _GridHeaderFont;

            if (fnt.Size < 4)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit);

            _GridHeaderFont = fnt2;

            Invalidate();
        }

        private void miHeaderFontLarger_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = _GridHeaderFont;

            if (fnt.Size > 60)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit);

            _GridHeaderFont = fnt2;

            Invalidate();
        }

        private void miTitleFontSmaller_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = _GridTitleFont;

            if (fnt.Size < 4)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size - 1, fnt.Style, fnt.Unit);

            _GridTitleFont = fnt2;

            Invalidate();
        }

        private void miTitleFontLarger_Click(object sender, EventArgs e)
        {
            Font fnt;

            fnt = _GridTitleFont;

            if (fnt.Size > 60)
                return;

            var fnt2 = new Font(fnt.FontFamily, fnt.Size + 1, fnt.Style, fnt.Unit);

            _GridTitleFont = fnt2;

            Invalidate();
        }

        private void miSearchInColumn_Click(object sender, EventArgs e)
        {
            var frm = new frmSearchInColumn(PointToScreen(Location));

            frm.ColumnName = _GridHeader[_ColOverOnMenuButton];

            frm.ShowDialog();

            if (!frm.Canceled)
            {
                // we wanna search
                string srch = frm.SearchText;
                int t;

                if (string.IsNullOrEmpty(srch))
                    // we have nothing to look for so lets bail
                    return;

                _LastSearchText = srch;
                _LastSearchColumn = _ColOverOnMenuButton;
                var loopTo = _rows - 1;
                for (t = 0; t <= loopTo; t++)
                {
                    if (Strings.InStr(Strings.UCase(_grid[t, _ColOverOnMenuButton]), Strings.UCase(srch), CompareMethod.Text) != 0)
                    {
                        // we have a match
                        if (vs.Visible)
                        {
                            vs.Value = t;
                            _SelectedRow = t;
                            Invalidate();
                            break;
                        }
                        else
                        {
                            _SelectedRow = t;
                            Invalidate();
                            break;
                        }
                    }
                }
            }
        }

        private void miAutoSizeToContents_Click(object sender, EventArgs e)
        {
            AutoSizeCellsToContents = true;
        }

        private void miAllowUserColumnResizing_Click(object sender, EventArgs e)
        {
            _AllowUserColumnResizing = !_AllowUserColumnResizing;
        }

        private void miSortAscending_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;

            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);

            Refresh();

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = _grid[y, _ColOverOnMenuButton];
                sitem.Iord = y;
                lb.Items.Add(sitem);
                lb.DisplayMember = "Ivis";
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
            }
            lb.Sorted = true;
            var loopTo1 = lb.Items.Count - 1;
            for (rr = 0; rr <= loopTo1; rr++)
            {
                var loopTo2 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo2; cc++)
                {

                    var ssi = new SortItem();
                    ssi = (SortItem)lb.Items[rr];

                    newgrid[rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
            }


            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);

            _SelectedRow = -1;
            _SelectedRows.Clear();
            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            Invalidate();


            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miSortDescending_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;
            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);



            Refresh();

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = _grid[y, _ColOverOnMenuButton];
                sitem.Iord = y;
                lb.Items.Add(sitem);
                lb.DisplayMember = "Ivis";
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
            }
            lb.Sorted = true;

            for (rr = lb.Items.Count - 1; rr >= 0; rr += -1)
            {
                var loopTo1 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo1; cc++)
                {
                    SortItem ssi = new SortItem();
                    ssi = (SortItem)lb.Items[rr];
                    newgrid[lb.Items.Count - 1 - rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
            }

            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            Invalidate();

            _SelectedRow = -1;
            _SelectedRows.Clear();
            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miDateAsc_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;

            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);



            Refresh();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                if (!Information.IsDate(_grid[y, _ColOverOnMenuButton]) & string.IsNullOrEmpty(_grid[y, _ColOverOnMenuButton].Trim()))
                    _grid[y, _ColOverOnMenuButton] = DateTime.MinValue.ToString("MM/dd/yyyy");
                else if (!Information.IsDate(_grid[y, _ColOverOnMenuButton] + ""))
                {
                    Interaction.MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a date", MsgBoxStyle.Critical, "Sort date descending message");
                    return;
                }
            }

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = Strings.Format(Conversions.ToDate(_grid[y, _ColOverOnMenuButton]), "yyyyMMdd");
                sitem.Iord = y;
                lb.Items.Add(sitem);
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }
            lb.DisplayMember = "Ivis";
            lb.Sorted = true;
            var loopTo2 = lb.Items.Count - 1;
            for (rr = 0; rr <= loopTo2; rr++)
            {
                var loopTo3 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo3; cc++)
                {
                    SortItem ssi = new SortItem();

                    ssi = (SortItem)lb.Items[rr];

                    newgrid[rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }

            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);
            var loopTo4 = _rows - 1;
            for (y = 0; y <= loopTo4; y++)
            {
                if ((_grid[y, _ColOverOnMenuButton].Trim() ?? "") == "01/01/0001")
                    _grid[y, _ColOverOnMenuButton] = "";
            }

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            Refresh();

            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miDateDesc_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;

            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);


            Refresh();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                if (!Information.IsDate(_grid[y, _ColOverOnMenuButton]) & string.IsNullOrEmpty(_grid[y, _ColOverOnMenuButton].Trim()))
                    _grid[y, _ColOverOnMenuButton] = DateTime.MinValue.ToString("MM/dd/yyyy");
                else if (!Information.IsDate(_grid[y, _ColOverOnMenuButton] + ""))
                {
                    Interaction.MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a date", MsgBoxStyle.Critical, "Sort date descending message");
                    return;
                }
            }

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = Strings.Format(Conversions.ToDate(_grid[y, _ColOverOnMenuButton]), "yyyyMMdd");
                sitem.Iord = y;
                lb.Items.Add(sitem);
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }
            lb.DisplayMember = "Ivis";
            lb.Sorted = true;

            for (rr = lb.Items.Count - 1; rr >= 0; rr += -1)
            {
                var loopTo2 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo2; cc++)
                {
                    SortItem ssi = new SortItem();

                    ssi = (SortItem)lb.Items[rr];

                    newgrid[lb.Items.Count - 1 - rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                //newgrid[lb.Items.Count - 1 - rr, cc] = _grid[Conversions.ToInteger(lb.Items[rr].iord), cc];
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }

            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);
            var loopTo3 = _rows - 1;
            for (y = 0; y <= loopTo3; y++)
            {
                if ((_grid[y, _ColOverOnMenuButton].Trim() ?? "") == "01/01/0001")
                    _grid[y, _ColOverOnMenuButton] = "";
            }

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            Refresh();

            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miSortNumericAsc_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;

            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);



            Refresh();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                if (!Information.IsNumeric(_grid[y, _ColOverOnMenuButton]))
                {
                    Interaction.MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a Number", MsgBoxStyle.Critical, "Sort date ascending message");
                    return;
                }
            }

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = Strings.Right("000000000000000" + _grid[y, _ColOverOnMenuButton], 15);
                sitem.Iord = y;
                lb.Items.Add(sitem);
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }
            lb.DisplayMember = "Ivis";
            lb.Sorted = true;
            var loopTo2 = lb.Items.Count - 1;
            for (rr = 0; rr <= loopTo2; rr++)
            {
                var loopTo3 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo3; cc++)
                {
                    SortItem ssi = new SortItem();

                    ssi = (SortItem)lb.Items[rr];

                    newgrid[rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                //newgrid[rr, cc] = _grid[Conversions.ToInteger(lb.Items[rr].iord), cc];
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }

            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miSortNumericDesc_Click(object sender, EventArgs e)
        {
            var newgrid = new string[_rows + 1, _cols + 1];
            var lb = new ListBox();
            int y, rr, cc;

            var presortcolwidths = new int[_cols + 1];
            bool oldautosizecells = _AutosizeCellsToContents;

            Array.Copy(_colwidths, presortcolwidths, _cols);

            Refresh();
            var loopTo = _rows - 1;
            for (y = 0; y <= loopTo; y++)
            {
                if (!Information.IsNumeric(_grid[y, _ColOverOnMenuButton]))
                {
                    Interaction.MsgBox("Cannot sort this column as a date because some value in the column cannot be converted to a Number", MsgBoxStyle.Critical, "Sort date ascending message");
                    return;
                }
            }

            if (_ShowProgressBar)
            {
                pBar.Maximum = _rows * 2;
                pBar.Minimum = 0;
                pBar.Value = 0;
                pBar.Visible = true;
                gb1.Visible = true;
                pBar.Refresh();
                gb1.Refresh();
            }

            lb.Items.Clear();
            var loopTo1 = _rows - 1;
            for (y = 0; y <= loopTo1; y++)
            {
                var sitem = new SortItem();
                sitem.Ivis = Strings.Right("000000000000000" + _grid[y, _ColOverOnMenuButton], 15);
                sitem.Iord = y;
                lb.Items.Add(sitem);
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }
            lb.DisplayMember = "Ivis";
            lb.Sorted = true;


            for (rr = lb.Items.Count - 1; rr >= 0; rr += -1)
            {
                var loopTo2 = _cols - 1;
                // loop through the current listbox
                for (cc = 0; cc <= loopTo2; cc++)
                {
                    SortItem ssi = new SortItem();

                    ssi = (SortItem)lb.Items[rr];

                    newgrid[lb.Items.Count - 1 - rr, cc] = _grid[Conversions.ToInteger(ssi.Iord), cc];
                }
                //newgrid[lb.Items.Count - 1 - rr, cc] = _grid[Conversions.ToInteger(lb.Items[rr].iord), cc];
                if (_ShowProgressBar)
                {
                    pBar.Increment(1);
                    pBar.Refresh();
                }
                Application.DoEvents();
            }

            PrivatePopulateGridFromArray(newgrid, _DefaultCellFont, _DefaultForeColor, false);

            Array.Copy(presortcolwidths, _colwidths, _cols);

            _AutosizeCellsToContents = oldautosizecells;

            GridResorted?.Invoke(this, _ColOverOnMenuButton);

            pBar.Visible = false;
            gb1.Visible = false;
        }

        private void miHideRow_Click(object sender, EventArgs e)
        {
            if (_RowOverOnMenuButton != -1 & _RowOverOnMenuButton < _rows)
            {
                _rowheights[_RowOverOnMenuButton] = 0;
                _AutosizeCellsToContents = false;
                Invalidate();
            }
        }

        private void miHideColumn_Click(object sender, EventArgs e)
        {
            if (_ColOverOnMenuButton != -1 & _ColOverOnMenuButton < _cols)
            {
                _colwidths[_ColOverOnMenuButton] = 0;
                _AutosizeCellsToContents = false;
                Invalidate();
            }
        }

        private void miShowAllRowsAndColumns_Click(object sender, EventArgs e)
        {
            AutoSizeCellsToContents = true;
            Invalidate();
        }

        private void miSetRowColor_Click(object sender, EventArgs e)
        {
            int r;
            int ccol;
            if ((int)clrdlg.ShowDialog() == (int)DialogResult.OK)
            {
                ccol = GetGridBackColorListEntry(new SolidBrush(clrdlg.Color));
                var loopTo = _cols - 1;
                for (r = 0; r <= loopTo; r++)
                    _gridBackColor[_RowOverOnMenuButton, r] = ccol;
                Invalidate();
            }
        }

        private void miSetColumnColor_Click(object sender, EventArgs e)
        {
            int r;
            int ccol;
            if ((int)clrdlg.ShowDialog() == (int)DialogResult.OK)
            {
                ccol = GetGridBackColorListEntry(new SolidBrush(clrdlg.Color));
                var loopTo = _rows - 1;
                for (r = 0; r <= loopTo; r++)
                    _gridBackColor[r, _ColOverOnMenuButton] = ccol;
                Invalidate();
            }
        }

        private void miSetCellColor_Click(object sender, EventArgs e)
        {
            if ((int)clrdlg.ShowDialog() == (int)DialogResult.OK)
            {
                _gridBackColor[_RowOverOnMenuButton, _ColOverOnMenuButton] = GetGridBackColorListEntry(new SolidBrush(clrdlg.Color));
                Invalidate();
            }
        }

        private void miSumColumn_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[t, _ColOverOnMenuButton]))
                {
                    v += Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                    flag = true;
                }
            }

            if (flag)
                Interaction.MsgBox("The sum of the column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "SUM COLUMN");
            else
                Interaction.MsgBox("The column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "SUM COLUMN");
        }

        private void miSumRow_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[_RowOverOnMenuButton, t]))
                {
                    v += Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                    flag = true;
                }
            }

            if (flag)
                Interaction.MsgBox("The sum of the row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "SUM ROW");
            else
                Interaction.MsgBox("The row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "SUM ROW");
        }

        private void miMaxCol_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[t, _ColOverOnMenuButton]))
                {
                    if (!flag)
                    {
                        v = Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                        flag = true;
                    }
                    else if (Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]) > v)
                        v = Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                }
            }

            if (flag)
                Interaction.MsgBox("The max value in the column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "MAX IN COLUMN");
            else
                Interaction.MsgBox("The column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "MAX IN COLUMN");
        }

        private void miMaxRow_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[_RowOverOnMenuButton, t]))
                {
                    if (!flag)
                    {
                        v = Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                        flag = true;
                    }
                    else if (Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]) > v)
                        v = Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                }
            }

            if (flag)
                Interaction.MsgBox("The max value in the row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "MAX IN ROW");
            else
                Interaction.MsgBox("The row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "MAX IN ROW");
        }

        private void miMinCol_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[t, _ColOverOnMenuButton]))
                {
                    if (!flag)
                    {
                        v = Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                        flag = true;
                    }
                    else if (Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]) < v)
                        v = Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                }
            }

            if (flag)
                Interaction.MsgBox("The min value in the column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "MIN IN COLUMN");
            else
                Interaction.MsgBox("The column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "MIN IN COLUMN");
        }

        private void miMinRow_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[_RowOverOnMenuButton, t]))
                {
                    if (!flag)
                    {
                        v = Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                        flag = true;
                    }
                    else if (Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]) < v)
                        v = Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                }
            }

            if (flag)
                Interaction.MsgBox("The min value in the row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "is " + v.ToString(), MsgBoxStyle.Information, "MIN IN ROW");
            else
                Interaction.MsgBox("The row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "MIN IN ROW");
        }

        private void miColAverage_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _rows - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[t, _ColOverOnMenuButton]))
                {
                    v += Conversions.ToDouble(_grid[t, _ColOverOnMenuButton]);
                    flag = true;
                }
            }

            if (flag)
            {
                v = v / _rows;
                Interaction.MsgBox("The average value in the column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "is " + Strings.Format(v, "##,##0.00"), MsgBoxStyle.Information, "AVERAGE IN COLUMN");
            }
            else
                Interaction.MsgBox("The column named " + _GridHeader[_ColOverOnMenuButton] + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "AVERAGE IN COLUMN");
        }

        private void miRowAverage_Click(object sender, EventArgs e)
        {
            int t;
            double v = 0;
            bool flag = false;
            var loopTo = _cols - 1;
            for (t = 0; t <= loopTo; t++)
            {
                if (Information.IsNumeric(_grid[_RowOverOnMenuButton, t]))
                {
                    v += Conversions.ToDouble(_grid[_RowOverOnMenuButton, t]);
                    flag = true;
                }
            }

            if (flag)
            {
                v = v / _cols;
                Interaction.MsgBox("The average value in the row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "is " + Strings.Format(v, "##,##0.00"), MsgBoxStyle.Information, "AVERAGE IN ROW");
            }
            else
                Interaction.MsgBox("The row numbered " + Conversions.ToString(_RowOverOnMenuButton) + Constants.vbCrLf + "Contains no numeric data...", MsgBoxStyle.Information, "AVERAGE IN ROW");
        }

        private void miCopyCellToClipboard_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(_grid[_RowOverOnMenuButton, _ColOverOnMenuButton], true);
        }

        private void miPrintTheGrid_Click(object sender, EventArgs e)
        {
            PrintTheGrid(_gridReportTitle, true, true, true, false, false);
        }

        private void miPreviewTheGrid_Click(object sender, EventArgs e)
        {
            PrintTheGrid(_gridReportTitle, true, true, true, true, false);
        }

        private void LogThis(string str)
        {
        }

        private void PurgeLog()
        {
        }

        private void miPageSetup_Click(object sender, EventArgs e)
        {
            if (_rows == 0 | _cols == 0)
                // If we got nothing to print dont bring up the page setup 
                return;

            try
            {
                _psets = new System.Drawing.Printing.PageSettings();

                // MsgBox(_psets.ToString())

                _OriginalPrinterName = Conversions.ToString(_psets.PrinterSettings.PrinterName);

                _PageSetupForm = new frmPageSetup(_psets);
            }
            catch (Exception ex)
            {
                miPreviewTheGrid.Enabled = false;
                miPrintTheGrid.Enabled = false;
                miPageSetup.Enabled = false;
                return;
            }

            PurgeLog();

            LogThis("Dim ps As System.Drawing.Printing.PageSettings = _psets");

            System.Drawing.Printing.PageSettings ps = _psets;

            LogThis("PageSetupForm = Nothing");

            _PageSetupForm = null;

            LogThis("_PageSetupForm = New frmPageSetup(_psets)");

            _PageSetupForm = new frmPageSetup(_psets);

            LogThis("PageSetupForm.MaxPage = CalculatePageRange()");

            _PageSetupForm.MaxPage = CalculatePageRange();

            LogThis(" _PageSetupForm.ShowDialog()");

            _PageSetupForm.ShowDialog();

            if (_PageSetupForm.Canceled)
                _psets = ps;
            else
            {
                _psets = _PageSetupForm.Psets;

                if (_PageSetupForm.Print)
                    PrintTheGrid(_gridReportTitle, true, true, true, false, Conversions.ToBoolean(_psets.Landscape));
                else if (_PageSetupForm.Preview)
                    PrintTheGrid(_gridReportTitle, true, true, true, true, Conversions.ToBoolean(_psets.Landscape));
            }

            Refresh();
        }

        private void miExportToSQLScript_Click(object sender, EventArgs e)
        {
            var frm = new frmScriptToSQL(this);

            frm.Location = PointToScreen(Location);

            frm.ShowDialog();
        }

        private void miExportToHTMLTable_Click(object sender, EventArgs e)
        {
            var frm = new frmScriptToHTML(this);

            frm.Location = PointToScreen(Location);

            frm.ShowDialog();
        }

        private void miDisplayFrequencyDistribution_Click(object sender, EventArgs e)
        {
            var frm = new frmFreqDist(this, _ColOverOnMenuButton);

            frm.ShowDialog();
        }

        private void miProperties_Click(object sender, EventArgs e)
        {
            var frm = new frmGridProperties(this);

            frm.Location = PointToScreen(Location);

            frm.ShowDialog();

            BringToFront();
        }

        private void miTearColumnAway_Click(object sender, EventArgs e)
        {
            if (_ColOverOnMenuButton == -1)
                return;

            _TearAwayWork = true;

            TearAwayColumID(_ColOverOnMenuButton-1);

            _TearAwayWork = false;
        }

        private void miHideColumnTearAway_Click(object sender, EventArgs e)
        {
            if (TearAways.Count == 0)
                return;

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                if (ta.ColID == _ColOverOnMenuButton-1)
                    // call into the child form to start the death spiral a happening
                    ta.Winform.KillMe(_ColOverOnMenuButton-1);
            }
        }

        private void miHideAllTearAwayColumns_Click(object sender, EventArgs e)
        {
            if (TearAways.Count == 0)
                return;

            int t;

            for (t = TearAways.Count - 1; t >= 0; t += -1)
            {
                TearAwayWindowEntry ta = (TAIGridControl2.TAIGridControl.TearAwayWindowEntry)TearAways[t];
                // call into the child form to start the death spiral a happening (one at a time as we
                // be a killing em all
                ta.Winform.KillMe(ta.ColID);
            }
        }

        private void miMultipleColumnTearAway_Click(object sender, EventArgs e)
        {
            if (_cols == 0)
                // there be no columns to tear away  
                return;

            _TearAwayWork = true;

            var frm = new frmMultipleColumnTearAway(_GridHeader);

            frm.Location = PointToScreen(Location);

            frm.ShowDialog();


            if (!frm.Canceled)
            {
                // here we need to tear away the columns if they are selected

                if (frm.SelectedIndices.Count > 0)
                {
                    int t;
                    var loopTo = frm.SelectedIndices.Count - 1;
                    for (t = 0; t <= loopTo; t++)
                        TearAwayColumID(frm.SelectedIndices[t]);
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                _TearAwayWork = false;

                ArrangeTearAwayWindows();
            }

            _TearAwayWork = false;
        }

        private void miArrangeTearAways_Click(object sender, EventArgs e)
        {
            ArrangeTearAwayWindows();
        }

        #endregion

        #region OverRides

        protected override void OnPaintBackground(PaintEventArgs pevent)
        {
        }

        protected override bool ProcessDialogKey(Keys kd)
        {
            if (_EditMode & (int)kd == (int)Keys.Tab & _AllowInGridEdits)
            {
                // we may need to bounce to the next col for edits or tab off the grid in the case of being at the end

                int x, y, xoff, yoff, r, c, nrc, ncc;

                bool flag = false;

                nrc = _RowClicked;
                ncc = _ColClicked;

                if (ncc < _cols - 1)
                    x = ncc + 1;
                else
                {
                    x = 0;

                    nrc += 1;
                }

                y = nrc;

                if (y > _rows - 1)
                {
                    flag = true;
                    x = -1;
                    y = -1;
                }
                else
                    while (!flag)
                    {
                        if (_colEditable[x] & !_colboolean[x])
                            flag = true;
                        else
                        {
                            x += 1;
                            if (x > _cols - 1)
                            {
                                x = 0;
                                y += 1;

                                if (y > _rows - 1)
                                {
                                    flag = true;
                                    x = -1;
                                    y = -1;
                                }
                            }
                        }
                    }

                // if we get here and not flag  and x > -1 and y > -1 then we are at the end
                // otherwise x = newcolumn for edit y = new row for edit
                // we need to clean up the existing edit and jump to the new one

                // who has focus

                if (flag & x > -1 & y > -1)
                {
                    if (txtInput.Visible)
                    {
                        // the txtinput has it

                        if ((_grid[_RowClicked, _ColClicked] ?? "") != (txtInput.Text ?? ""))
                        {
                            string oldval = _grid[_RowClicked, _ColClicked];
                            string newval = txtInput.Text;
                            _grid[_RowClicked, _ColClicked] = txtInput.Text;
                            CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                        }

                        txtInput.SendToBack();
                        txtInput.Visible = false;
                    }
                    else
                    {
                        // the cmboinput does

                        if ((_grid[_RowClicked, _ColClicked] ?? "") != (cmboInput.Text ?? "") & !string.IsNullOrEmpty(cmboInput.Text.Trim()))
                        {
                            string oldval = _grid[_RowClicked, _ColClicked];
                            string newval = cmboInput.Text;
                            _grid[_RowClicked, _ColClicked] = cmboInput.Text;
                            CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                        }

                        cmboInput.SendToBack();
                        cmboInput.Visible = false;
                    }

                    // we now need to ensure that the selectedrows collection is clear

                    _SelectedRows.Clear();

                    // now lets setup for the new edit

                    if (_SelectedRow != y)
                        SelectedRow = y;

                    _RowClicked = y;
                    _ColClicked = x;

                    if (_RowClicked > -1 & _RowClicked < _rows & _rowEditable[_RowClicked])
                    {
                        if (_ColClicked > -1 & _ColClicked < _cols & _colEditable[_ColClicked] & _AllowInGridEdits)
                        {
                            if (IsColumnRestricted(_ColClicked))
                            {
                                var it = GetColumnRestriction(_ColClicked);

                                cmboInput.Items.Clear();

                                var s = it.RestrictedList.Split("^".ToCharArray());
                                foreach (string ss in s)
                                    cmboInput.Items.Add(ss);

                                // we have selected a row and col lets move the txtinput there and bring it to the front
                                xoff = 0;
                                yoff = 0;

                                if (_RowClicked > 0)
                                {
                                    var loopTo = _RowClicked - 1;
                                    for (r = 0; r <= loopTo; r++)
                                        yoff = yoff + get_RowHeight(r);
                                }

                                if (GridheaderVisible)
                                    yoff = yoff + _GridHeaderHeight;

                                if (_GridTitleVisible)
                                    yoff = yoff + _GridTitleHeight;

                                if (_ColClicked > 0)
                                {
                                    var loopTo1 = _ColClicked - 1;
                                    for (c = 0; c <= loopTo1; c++)
                                        xoff = xoff + get_ColWidth(c);
                                }

                                if (vs.Visible & vs.Value > 0)
                                    yoff = yoff - GimmeYOffset(vs.Value);

                                if (hs.Visible & hs.Value > 0)
                                    xoff = xoff - GimmeXOffset(hs.Value);

                                if (_CellOutlines)
                                {
                                    cmboInput.Top = yoff + 1;
                                    cmboInput.Left = xoff + 1;
                                    cmboInput.Width = get_ColWidth(_ColClicked) - 1;
                                    cmboInput.Height = get_RowHeight(_RowClicked) - 2;
                                    cmboInput.BackColor = _colEditableTextBackColor;
                                }
                                else
                                {
                                    cmboInput.Top = yoff;
                                    cmboInput.Left = xoff;
                                    cmboInput.Width = get_ColWidth(_ColClicked);
                                    cmboInput.Height = get_RowHeight(_RowClicked);
                                    cmboInput.BackColor = _colEditableTextBackColor;
                                }

                                cmboInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                                cmboInput.Text = _grid[_RowClicked, _ColClicked];

                                cmboInput.Visible = true;
                                cmboInput.BringToFront();
                                cmboInput.DroppedDown = true;
                                _EditModeCol = _ColClicked;
                                _EditModeRow = _RowClicked;
                                _EditMode = true;

                                cmboInput.Focus();
                            }
                            else
                            {
                                // we have selected a row and col lets move the txtinput there and bring it to the front
                                xoff = 0;
                                yoff = 0;

                                if (_RowClicked > 0)
                                {
                                    var loopTo2 = _RowClicked - 1;
                                    for (r = 0; r <= loopTo2; r++)
                                        yoff = yoff + get_RowHeight(r);
                                }

                                if (GridheaderVisible)
                                    yoff = yoff + _GridHeaderHeight;

                                if (_GridTitleVisible)
                                    yoff = yoff + _GridTitleHeight;

                                if (_ColClicked > 0)
                                {
                                    var loopTo3 = _ColClicked - 1;
                                    for (c = 0; c <= loopTo3; c++)
                                        xoff = xoff + get_ColWidth(c);
                                }

                                if (vs.Visible & vs.Value > 0)
                                    yoff = yoff - GimmeYOffset(vs.Value);

                                if (hs.Visible & hs.Value > 0)
                                    xoff = xoff - GimmeXOffset(hs.Value);

                                if (_CellOutlines)
                                {
                                    txtInput.Top = yoff + 1;
                                    txtInput.Left = xoff + 1;
                                    txtInput.Width = get_ColWidth(_ColClicked) - 1;
                                    txtInput.Height = get_RowHeight(_RowClicked) - 2;
                                    txtInput.BackColor = _colEditableTextBackColor;
                                }
                                else
                                {
                                    txtInput.Top = yoff;
                                    txtInput.Left = xoff;
                                    txtInput.Width = get_ColWidth(_ColClicked);
                                    txtInput.Height = get_RowHeight(_RowClicked);
                                    txtInput.BackColor = _colEditableTextBackColor;
                                }

                                txtInput.Font = _gridCellFontsList[_gridCellFonts[_RowClicked, _ColClicked]];

                                txtInput.Text = _grid[_RowClicked, _ColClicked];

                                txtInput.Visible = true;
                                txtInput.BringToFront();
                                _EditModeCol = _ColClicked;
                                _EditModeRow = _RowClicked;
                                _EditMode = true;

                                txtInput.Focus();
                            }
                        }
                    }

                    return true;
                }
                else
                {
                    if (txtInput.Visible)
                    {
                        // the txtinput has it

                        if ((_grid[_RowClicked, _ColClicked] ?? "") != (txtInput.Text ?? ""))
                        {
                            string oldval = _grid[_RowClicked, _ColClicked];
                            string newval = txtInput.Text;
                            _grid[_RowClicked, _ColClicked] = txtInput.Text;
                            CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                        }

                        txtInput.SendToBack();
                        txtInput.Visible = false;
                    }


                    if (cmboInput.Visible)
                    {
                        // the cmboinput does

                        if ((_grid[_RowClicked, _ColClicked] ?? "") != (cmboInput.Text ?? "") & !string.IsNullOrEmpty(cmboInput.Text.Trim()))
                        {
                            string oldval = _grid[_RowClicked, _ColClicked];
                            string newval = cmboInput.Text;
                            _grid[_RowClicked, _ColClicked] = cmboInput.Text;
                            CellEdited?.Invoke(this, _RowClicked, _ColClicked, oldval, newval);
                        }

                        cmboInput.SendToBack();
                        cmboInput.Visible = false;
                    }

                    base.ProcessDialogKey(kd);

                    return false;
                } // If flag And x > -1 And y > -1 Then
            }
            else
            {

                // If _EditMode And _AllowInGridEdits Then
                // If txtInput.Visible Then
                // ' the txtinput has it

                // If _grid(_RowClicked, _ColClicked) <> txtInput.Text Then
                // Dim oldval As String = _grid(_RowClicked, _ColClicked)
                // Dim newval As String = txtInput.Text
                // _grid(_RowClicked, _ColClicked) = txtInput.Text
                // RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                // End If

                // txtInput.SendToBack()
                // txtInput.Visible = False

                // End If


                // If cmboInput.Visible Then
                // ' the cmboinput does

                // If _grid(_RowClicked, _ColClicked) <> cmboInput.Text And cmboInput.Text.Trim() <> "" Then
                // Dim oldval As String = _grid(_RowClicked, _ColClicked)
                // Dim newval As String = cmboInput.Text
                // _grid(_RowClicked, _ColClicked) = cmboInput.Text
                // RaiseEvent CellEdited(Me, _RowClicked, _ColClicked, oldval, newval)
                // End If

                // cmboInput.SendToBack()
                // cmboInput.Visible = False

                // End If

                // End If

                base.ProcessDialogKey(kd);
                return false;
            } // If _EditMode And kd = Keys.Tab And _AllowInGridEdits
        }

        #endregion

        #region Private Classes

        private class frmExportToText : Form
        {
            public frmExportToText() : base()
            {
                _PageSetupForm = new frmPageSetup();

                // This call is required by the Windows Form Designer.
                InitializeComponent();
            }

            // Form overrides dispose to clean up the component list.
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    if (!(components == null))
                        components.Dispose();
                }
                base.Dispose(disposing);
            }

            // Required by the Windows Form Designer
            private IContainer components;

            // NOTE: The following procedure is required by the Windows Form Designer
            // It can be modified using the Windows Form Designer.  
            // Do not modify it using the code editor.
            private Button _cmdOK;

            internal Button cmdOK
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _cmdOK;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_cmdOK != null)
                    {
                        _cmdOK.Click -= cmdOK_Click;
                    }

                    _cmdOK = value;
                    if (_cmdOK != null)
                    {
                        _cmdOK.Click += cmdOK_Click;
                    }
                }
            }

            private Button _cmdCancel;

            internal Button cmdCancel
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _cmdCancel;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_cmdCancel != null)
                    {
                    }

                    _cmdCancel = value;
                    if (_cmdCancel != null)
                    {
                    }
                }
            }

            private GroupBox _GroupBox1;

            internal GroupBox GroupBox1
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _GroupBox1;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_GroupBox1 != null)
                    {
                    }

                    _GroupBox1 = value;
                    if (_GroupBox1 != null)
                    {
                    }
                }
            }

            private RadioButton _rbTab;

            internal RadioButton rbTab
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _rbTab;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_rbTab != null)
                    {
                    }

                    _rbTab = value;
                    if (_rbTab != null)
                    {
                    }
                }
            }

            private RadioButton _rbSemicolon;

            internal RadioButton rbSemicolon
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _rbSemicolon;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_rbSemicolon != null)
                    {
                    }

                    _rbSemicolon = value;
                    if (_rbSemicolon != null)
                    {
                    }
                }
            }

            private RadioButton _rbComma;

            internal RadioButton rbComma
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _rbComma;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_rbComma != null)
                    {
                    }

                    _rbComma = value;
                    if (_rbComma != null)
                    {
                    }
                }
            }

            private RadioButton _rbSpace;

            internal RadioButton rbSpace
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _rbSpace;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_rbSpace != null)
                    {
                    }

                    _rbSpace = value;
                    if (_rbSpace != null)
                    {
                    }
                }
            }

            private RadioButton _rbOther;

            internal RadioButton rbOther
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _rbOther;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_rbOther != null)
                    {
                    }

                    _rbOther = value;
                    if (_rbOther != null)
                    {
                    }
                }
            }

            private TextBox _txtOther;

            internal TextBox txtOther
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _txtOther;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_txtOther != null)
                    {
                    }

                    _txtOther = value;
                    if (_txtOther != null)
                    {
                    }
                }
            }

            private CheckBox _chkIncludeFieldNames;

            internal CheckBox chkIncludeFieldNames
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _chkIncludeFieldNames;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_chkIncludeFieldNames != null)
                    {
                    }

                    _chkIncludeFieldNames = value;
                    if (_chkIncludeFieldNames != null)
                    {
                    }
                }
            }

            private Label _Label1;

            internal Label Label1
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _Label1;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_Label1 != null)
                    {
                    }

                    _Label1 = value;
                    if (_Label1 != null)
                    {
                    }
                }
            }

            private TextBox _txtExportFile;

            internal TextBox txtExportFile
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _txtExportFile;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_txtExportFile != null)
                    {
                    }

                    _txtExportFile = value;
                    if (_txtExportFile != null)
                    {
                    }
                }
            }

            private Button _cmdBrowse;

            internal Button cmdBrowse
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _cmdBrowse;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_cmdBrowse != null)
                    {
                        _cmdBrowse.Click -= cmdBrowse_Click;
                    }

                    _cmdBrowse = value;
                    if (_cmdBrowse != null)
                    {
                        _cmdBrowse.Click += cmdBrowse_Click;
                    }
                }
            }

            private CheckBox _chkIncludeLineTerminator;

            internal CheckBox chkIncludeLineTerminator
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _chkIncludeLineTerminator;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_chkIncludeLineTerminator != null)
                    {
                    }

                    _chkIncludeLineTerminator = value;
                    if (_chkIncludeLineTerminator != null)
                    {
                    }
                }
            }

            [DebuggerStepThrough()]
            private void InitializeComponent()
            {
                _cmdOK = new Button();
                _cmdOK.Click += cmdOK_Click;
                _cmdCancel = new Button();
                _GroupBox1 = new GroupBox();
                _txtOther = new TextBox();
                _rbOther = new RadioButton();
                _rbSpace = new RadioButton();
                _rbComma = new RadioButton();
                _rbSemicolon = new RadioButton();
                _rbTab = new RadioButton();
                _chkIncludeFieldNames = new CheckBox();
                _Label1 = new Label();
                _txtExportFile = new TextBox();
                _cmdBrowse = new Button();
                _cmdBrowse.Click += cmdBrowse_Click;
                _chkIncludeLineTerminator = new CheckBox();
                _GroupBox1.SuspendLayout();
                SuspendLayout();
                // 
                // cmdOK
                // 
                _cmdOK.DialogResult = DialogResult.OK;
                _cmdOK.Location = new Point(472, 8);
                _cmdOK.Name = "cmdOK";
                _cmdOK.Size = new Size(104, 24);
                _cmdOK.TabIndex = 0;
                _cmdOK.Text = "OK";
                // 
                // cmdCancel
                // 
                _cmdCancel.DialogResult = DialogResult.Cancel;
                _cmdCancel.Location = new Point(472, 40);
                _cmdCancel.Name = "cmdCancel";
                _cmdCancel.Size = new Size(104, 24);
                _cmdCancel.TabIndex = 1;
                _cmdCancel.Text = "Cancel";
                // 
                // GroupBox1
                // 
                _GroupBox1.Controls.Add(_txtOther);
                _GroupBox1.Controls.Add(_rbOther);
                _GroupBox1.Controls.Add(_rbSpace);
                _GroupBox1.Controls.Add(_rbComma);
                _GroupBox1.Controls.Add(_rbSemicolon);
                _GroupBox1.Controls.Add(_rbTab);
                _GroupBox1.Location = new Point(8, 8);
                _GroupBox1.Name = "GroupBox1";
                _GroupBox1.Size = new Size(456, 56);
                _GroupBox1.TabIndex = 2;
                _GroupBox1.TabStop = false;
                _GroupBox1.Text = "Choose the delimiter that separates your fields";
                // 
                // txtOther
                // 
                _txtOther.Location = new Point(376, 22);
                _txtOther.Name = "txtOther";
                _txtOther.Size = new Size(32, 20);
                _txtOther.TabIndex = 5;
                _txtOther.Text = "";
                // 
                // rbOther
                // 
                _rbOther.Location = new Point(312, 24);
                _rbOther.Name = "rbOther";
                _rbOther.Size = new Size(56, 16);
                _rbOther.TabIndex = 4;
                _rbOther.Text = "&Other";
                // 
                // rbSpace
                // 
                _rbSpace.Location = new Point(248, 24);
                _rbSpace.Name = "rbSpace";
                _rbSpace.Size = new Size(56, 16);
                _rbSpace.TabIndex = 3;
                _rbSpace.Text = "S&pace";
                // 
                // rbComma
                // 
                _rbComma.Checked = true;
                _rbComma.Location = new Point(168, 24);
                _rbComma.Name = "rbComma";
                _rbComma.Size = new Size(72, 16);
                _rbComma.TabIndex = 2;
                _rbComma.TabStop = true;
                _rbComma.Text = "&Comma";
                // 
                // rbSemicolon
                // 
                _rbSemicolon.Location = new Point(80, 24);
                _rbSemicolon.Name = "rbSemicolon";
                _rbSemicolon.Size = new Size(80, 16);
                _rbSemicolon.TabIndex = 1;
                _rbSemicolon.Text = "&Semicolon";
                // 
                // rbTab
                // 
                _rbTab.Location = new Point(24, 24);
                _rbTab.Name = "rbTab";
                _rbTab.Size = new Size(48, 16);
                _rbTab.TabIndex = 0;
                _rbTab.Text = "&Tab";
                // 
                // chkIncludeFieldNames
                // 
                _chkIncludeFieldNames.Checked = true;
                _chkIncludeFieldNames.CheckState = CheckState.Checked;
                _chkIncludeFieldNames.Location = new Point(32, 72);
                _chkIncludeFieldNames.Name = "chkIncludeFieldNames";
                _chkIncludeFieldNames.Size = new Size(200, 24);
                _chkIncludeFieldNames.TabIndex = 3;
                _chkIncludeFieldNames.Text = "Include Field Names on First Row";
                // 
                // Label1
                // 
                _Label1.Location = new Point(16, 104);
                _Label1.Name = "Label1";
                _Label1.Size = new Size(128, 24);
                _Label1.TabIndex = 4;
                _Label1.Text = "Export To File:";
                _Label1.TextAlign = ContentAlignment.MiddleLeft;
                // 
                // txtExportFile
                // 
                _txtExportFile.Location = new Point(16, 128);
                _txtExportFile.Name = "txtExportFile";
                _txtExportFile.Size = new Size(448, 20);
                _txtExportFile.TabIndex = 5;
                _txtExportFile.Text = "";
                // 
                // cmdBrowse
                // 
                _cmdBrowse.Location = new Point(472, 128);
                _cmdBrowse.Name = "cmdBrowse";
                _cmdBrowse.Size = new Size(104, 24);
                _cmdBrowse.TabIndex = 6;
                _cmdBrowse.Text = "Browse";
                // 
                // chkIncludeLineTerminator
                // 
                _chkIncludeLineTerminator.Checked = true;
                _chkIncludeLineTerminator.CheckState = CheckState.Checked;
                _chkIncludeLineTerminator.Location = new Point(256, 72);
                _chkIncludeLineTerminator.Name = "chkIncludeLineTerminator";
                _chkIncludeLineTerminator.Size = new Size(200, 24);
                _chkIncludeLineTerminator.TabIndex = 7;
                _chkIncludeLineTerminator.Text = "Include Line Terminator";
                // 
                // frmExportToText
                // 
                AutoScaleBaseSize = new Size(5, 13);
                ClientSize = new Size(584, 176);
                ControlBox = false;
                Controls.Add(_chkIncludeLineTerminator);
                Controls.Add(_cmdBrowse);
                Controls.Add(_txtExportFile);
                Controls.Add(_Label1);
                Controls.Add(_chkIncludeFieldNames);
                Controls.Add(_GroupBox1);
                Controls.Add(_cmdCancel);
                Controls.Add(_cmdOK);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MinimumSize = new Size(590, 200);
                Name = "frmExportToText";
                StartPosition = FormStartPosition.CenterParent;
                Text = "Export Grid Data To Text File...";
                _GroupBox1.ResumeLayout(false);
                ResumeLayout(false);
            }

            private string _delimiter = ",";

            private string _filename;

            private bool _includeFieldNames = true;
            
            private bool _includeLineTerminator = true;
            
            private frmPageSetup _PageSetupForm;

            private void cmdBrowse_Click(object sender, EventArgs e)
            {
                try
                {
                    var openFile = new SaveFileDialog();

                    openFile.InitialDirectory = Environment.CurrentDirectory;
                    openFile.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";

                    openFile.DefaultExt = "txt";

                    if ((int)openFile.ShowDialog() == (int)DialogResult.OK)
                        txtExportFile.Text = openFile.FileName;
                }
                catch (Exception ex)
                {
                    Interaction.MsgBox(ex.ToString(), (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "frmExportToText.cmdBrowse_Click Error...");
                }
            }

            private void cmdOK_Click(object sender, EventArgs e)
            {
                if (rbOther.Checked)
                {
                    if (string.IsNullOrEmpty(txtOther.Text.Trim()))
                    {
                        Interaction.MsgBox("You have selected Other as your delimiter but you did not specify what the delimiter should be! " + "Please correct this before proceeding!", (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "Export To Text Error...");
                        DialogResult = DialogResult.None;
                        return;
                    }
                }

                if (string.IsNullOrEmpty(txtExportFile.Text))
                {
                    Interaction.MsgBox("You must select the file to export the data to! " + "Please correct this before proceeding!", (MsgBoxStyle)((int)MsgBoxStyle.Information + (int)MsgBoxStyle.OkOnly), "Export To Text Error...");
                    DialogResult = DialogResult.None;
                    return;
                }

                if (rbTab.Checked)
                    _delimiter = Constants.vbTab;
                else if (rbSemicolon.Checked)
                    _delimiter = ";";
                else if (rbComma.Checked)
                    _delimiter = ",";
                else if (rbSpace.Checked)
                    _delimiter = " ";
                else if (rbOther.Checked)
                    _delimiter = txtOther.Text;

                _filename = txtExportFile.Text;
                _includeFieldNames = Conversions.ToBoolean(chkIncludeFieldNames.CheckState);
                _includeLineTerminator = Conversions.ToBoolean(chkIncludeLineTerminator.CheckState);
            }

            public string Delimiter
            {
                get
                {
                    return _delimiter;
                }
                set
                {
                    _delimiter = value;
                }
            }

            public string Filename
            {
                get
                {
                    return _filename;
                }
                set
                {
                    _filename = value;
                }
            }

            public bool IncludeFieldNames
            {
                get
                {
                    return _includeFieldNames;
                }
                set
                {
                    _includeFieldNames = value;
                }
            }

            public bool IncludeLineTerminator
            {
                get
                {
                    return _includeLineTerminator;
                }
                set
                {
                    _includeLineTerminator = value;
                }
            }
        }

        private class frmExportingToExcelWorking : Form
        {
            public frmExportingToExcelWorking() : base()
            {

                // This call is required by the Windows Form Designer.
                InitializeComponent();
            }

            // Form overrides dispose to clean up the component list.
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    if (!(components == null))
                        components.Dispose();
                }
                base.Dispose(disposing);
            }

            // Required by the Windows Form Designer
            private IContainer components;

            // NOTE: The following procedure is required by the Windows Form Designer
            // It can be modified using the Windows Form Designer.  
            // Do not modify it using the code editor.
            private Label _Label1;

            internal Label Label1
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _Label1;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_Label1 != null)
                    {
                    }

                    _Label1 = value;
                    if (_Label1 != null)
                    {
                    }
                }
            }

            private Label _Label2;

            internal Label Label2
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _Label2;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_Label2 != null)
                    {
                    }

                    _Label2 = value;
                    if (_Label2 != null)
                    {
                    }
                }
            }

            [DebuggerStepThrough()]
            private void InitializeComponent()
            {
                _Label1 = new Label();
                _Label2 = new Label();
                SuspendLayout();
                // 
                // Label1
                // 
                _Label1.Font = new Font("Microsoft Sans Serif", 15.75F, FontStyle.Regular, GraphicsUnit.Point, Conversions.ToByte(0));
                _Label1.Location = new Point(8, 8);
                _Label1.Name = "Label1";
                _Label1.Size = new Size(288, 32);
                _Label1.TabIndex = 0;
                _Label1.Text = "Sending data to Excel.";
                // 
                // Label2
                // 
                _Label2.Font = new Font("Microsoft Sans Serif", 15.75F, FontStyle.Regular, GraphicsUnit.Point, Conversions.ToByte(0));
                _Label2.Location = new Point(80, 60);
                _Label2.Name = "Label2";
                _Label2.Size = new Size(376, 32);
                _Label2.TabIndex = 1;
                _Label2.Text = "This may take a few moments...";
                _Label2.TextAlign = ContentAlignment.MiddleCenter;
                // 
                // frmExportingToExcelWorking
                // 
                AutoScaleBaseSize = new Size(5, 13);
                BackColor = Color.AntiqueWhite;
                ClientSize = new Size(516, 157);
                ControlBox = false;
                Controls.Add(_Label2);
                Controls.Add(_Label1);
                Name = "frmExportingToExcelWorking";
                StartPosition = FormStartPosition.CenterScreen;
                Text = "Exporting.....";
                TopMost = true;
                ResumeLayout(false);
            }

            public void UpdateDisplay(string msg)
            {
                Label2.Text = msg;
                Label2.Refresh();
                Application.DoEvents();
            }
        }

        private class frmSearchInColumn : Form
        {
            public frmSearchInColumn(Point Loc) : base()
            {
                _PageSetupForm = new frmPageSetup();

                // This call is required by the Windows Form Designer.
                InitializeComponent();

                // Add any initialization after the InitializeComponent() call

                StartPosition = FormStartPosition.Manual;
                Location = Loc;
            }


            public frmSearchInColumn() : base()
            {
                _PageSetupForm = new frmPageSetup();

                // This call is required by the Windows Form Designer.
                InitializeComponent();
            }

            // Form overrides dispose to clean up the component list.
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    if (!(components == null))
                        components.Dispose();
                }
                base.Dispose(disposing);
            }

            // Required by the Windows Form Designer
            private IContainer components;

            // NOTE: The following procedure is required by the Windows Form Designer
            // It can be modified using the Windows Form Designer.  
            // Do not modify it using the code editor.
            private Button _btnCancel;

            internal Button btnCancel
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _btnCancel;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_btnCancel != null)
                    {
                        _btnCancel.Click -= btnCancel_Click;
                    }

                    _btnCancel = value;
                    if (_btnCancel != null)
                    {
                        _btnCancel.Click += btnCancel_Click;
                    }
                }
            }

            private Button _btnSearch;

            internal Button btnSearch
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _btnSearch;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_btnSearch != null)
                    {
                        _btnSearch.Click -= btnSearch_Click;
                    }

                    _btnSearch = value;
                    if (_btnSearch != null)
                    {
                        _btnSearch.Click += btnSearch_Click;
                    }
                }
            }

            private TextBox _txtSearchItem;

            internal TextBox txtSearchItem
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _txtSearchItem;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_txtSearchItem != null)
                    {
                    }

                    _txtSearchItem = value;
                    if (_txtSearchItem != null)
                    {
                    }
                }
            }

            private Label _Label1;

            internal Label Label1
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _Label1;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_Label1 != null)
                    {
                    }

                    _Label1 = value;
                    if (_Label1 != null)
                    {
                    }
                }
            }

            [DebuggerStepThrough()]
            private void InitializeComponent()
            {
                _btnCancel = new Button();
                _btnCancel.Click += btnCancel_Click;
                _btnSearch = new Button();
                _btnSearch.Click += btnSearch_Click;
                _txtSearchItem = new TextBox();
                _Label1 = new Label();
                SuspendLayout();
                // 
                // btnCancel
                // 
                _btnCancel.DialogResult = DialogResult.Cancel;
                _btnCancel.Location = new Point(452, 12);
                _btnCancel.Name = "btnCancel";
                _btnCancel.Size = new Size(88, 24);
                _btnCancel.TabIndex = 1;
                _btnCancel.Text = "Cancel";
                // 
                // btnSearch
                // 
                _btnSearch.DialogResult = DialogResult.Cancel;
                _btnSearch.Location = new Point(452, 44);
                _btnSearch.Name = "btnSearch";
                _btnSearch.Size = new Size(88, 24);
                _btnSearch.TabIndex = 2;
                _btnSearch.Text = "Search";
                // 
                // txtSearchItem
                // 
                _txtSearchItem.Location = new Point(24, 28);
                _txtSearchItem.Name = "txtSearchItem";
                _txtSearchItem.Size = new Size(360, 20);
                _txtSearchItem.TabIndex = 0;
                _txtSearchItem.Text = "";
                // 
                // Label1
                // 
                _Label1.Location = new Point(28, 52);
                _Label1.Name = "Label1";
                _Label1.Size = new Size(284, 16);
                _Label1.TabIndex = 3;
                _Label1.Text = "Enter the text you wish to search for...";
                // 
                // frmSearchInColumn
                // 
                AutoScaleBaseSize = new Size(5, 13);
                BackColor = Color.AntiqueWhite;
                ClientSize = new Size(552, 90);
                Controls.Add(_Label1);
                Controls.Add(_txtSearchItem);
                Controls.Add(_btnSearch);
                Controls.Add(_btnCancel);
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                Name = "frmSearchInColumn";
                StartPosition = FormStartPosition.WindowsDefaultLocation;
                Text = "Search for something in this column named...";
                ResumeLayout(false);
            }


            private bool _Canceled = true;
            private string _SearchText = "";
            private string _ColumnName = "";
            private frmPageSetup _PageSetupForm;

            public bool Canceled
            {
                get
                {
                    return _Canceled;
                }
                set
                {
                    _Canceled = value;
                }
            }

            public string SearchText
            {
                get
                {
                    return _SearchText;
                }
                set
                {
                    _SearchText = value;
                    txtSearchItem.Text = value;
                }
            }

            public string ColumnName
            {
                get
                {
                    return _ColumnName;
                }
                set
                {
                    _ColumnName = value;
                    Text = "Search for something in this column named..." + value;
                }
            }

            private void btnCancel_Click(object sender, EventArgs e)
            {
                _Canceled = true;
                Hide();
            }

            private void btnSearch_Click(object sender, EventArgs e)
            {
                _Canceled = false;
                _SearchText = txtSearchItem.Text;
                Hide();
            }
        }

        private class SortItem
        {
            private string _ItemVisual;
            private int _ItemOrdinal;

            public string Ivis
            {
                get
                {
                    return _ItemVisual;
                }
                set
                {
                    _ItemVisual = value;
                }
            }

            public int Iord
            {
                get
                {
                    return _ItemOrdinal;
                }
                set
                {
                    _ItemOrdinal = value;
                }
            }
        }

        private class EditColumnRestrictor
        {
            private int _ColumnID;
            private string _RestrictorList;

            public int ColumnID
            {
                get
                {
                    return _ColumnID;
                }
                set
                {
                    _ColumnID = value;
                }
            }

            public string RestrictedList
            {
                get
                {
                    return _RestrictorList;
                }
                set
                {
                    _RestrictorList = value;
                }
            }

            public override string ToString()
            {
                return _ColumnID.ToString();
            }
        }

        private class TearAwayWindowEntry
        {
            private int _columnID;
            private frmColumnTearAway _Winform;

            public int ColID
            {
                get
                {
                    return _columnID;
                }
                set
                {
                    _columnID = value;
                }
            }

            public frmColumnTearAway Winform
            {
                get
                {
                    return _Winform;
                }
                set
                {
                    _Winform = value;
                }
            }

            public void KillTearAway()
            {
                if (_Winform == null)
                    return;

                _Winform.Close();
            }

            public void HideTearAway()
            {
                if (_Winform == null)
                    return;

                _Winform.Hide();
            }

            public void ShowTearAway()
            {
                if (_Winform == null)
                    return;

                _Winform.Show();
            }

            public void SetTearAwayScrollParameters(int min, int max, bool visible)
            {
                if (_Winform == null)
                    return;

                _Winform.vscroller.Visible = visible;
                _Winform.vscroller.Minimum = min;
                _Winform.vscroller.Maximum = max;
            }

            public void SetTearAwayScrollIndex(int index)
            {
                if (_Winform == null)
                    return;

                if (index >= _Winform.vscroller.Minimum & index <= _Winform.vscroller.Maximum)
                    _Winform.vscroller.Value = index;
            }
        }

        #endregion
    }
}
