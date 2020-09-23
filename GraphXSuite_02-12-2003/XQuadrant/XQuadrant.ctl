VERSION 5.00
Begin VB.UserControl XQuadrant 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   ToolboxBitmap   =   "XQuadrant.ctx":0000
   Begin VB.PictureBox picToPrinterLegend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   570
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   3090
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picToPrinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   570
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   13
      Top             =   2460
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picInfoQuadrant 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   900
      ScaleHeight     =   405
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   1260
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picSplitter 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   3300
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5415
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox picCommands 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   60
      ScaleHeight     =   330
      ScaleWidth      =   1935
      TabIndex        =   4
      Top             =   60
      Width           =   1935
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   5
         Left            =   1605
         Picture         =   "XQuadrant.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "XQuadrant.ctx":089C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   3
         Left            =   975
         Picture         =   "XQuadrant.ctx":0E26
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   1
         Left            =   330
         Picture         =   "XQuadrant.ctx":13B0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   4
         Left            =   1290
         Picture         =   "XQuadrant.ctx":193A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   2
         Left            =   660
         Picture         =   "XQuadrant.ctx":1EC4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   315
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   4
         Left            =   1470
         Picture         =   "XQuadrant.ctx":224E
         Top             =   585
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "XQuadrant.ctx":2398
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   2
         Left            =   930
         Picture         =   "XQuadrant.ctx":24E2
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   1
         Left            =   660
         Picture         =   "XQuadrant.ctx":262C
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "XQuadrant.ctx":2776
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F5F5&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFF0F0&
      ForeColor       =   &H00FF7040&
      Height          =   5430
      Left            =   3360
      ScaleHeight     =   5430
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2130
      Begin VB.VScrollBar vsbContainer 
         Height          =   5445
         LargeChange     =   5
         Left            =   1875
         Max             =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F0F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   150
         ScaleHeight     =   5205
         ScaleWidth      =   1665
         TabIndex        =   2
         Top             =   0
         Width           =   1665
         Begin VB.PictureBox picDescription 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   300
            ScaleHeight     =   195
            ScaleWidth      =   765
            TabIndex        =   15
            Top             =   150
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.PictureBox picBox 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   60
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   14
            Top             =   150
            Visible         =   0   'False
            Width           =   195
         End
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   480
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMainSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuMainPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuMainCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainSelectionInfo 
         Caption         =   "Selection information"
      End
      Begin VB.Menu mnuMainQuadrantInfo 
         Caption         =   "Quadrant information"
      End
      Begin VB.Menu mnuMainViewLegend 
         Caption         =   "Display Legend"
      End
      Begin VB.Menu mnuMainCustomItemsSeparator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "2"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "3"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "4"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "5"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "6"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "7"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainCustomItems 
         Caption         =   "8"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuLegend 
      Caption         =   "&Legend"
      Begin VB.Menu mnuLegendHide 
         Caption         =   "Hide"
      End
   End
End
Attribute VB_Name = "XQuadrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Private uColumnsHeight()        As Double   'array of column height values
Private uColumnsBase()          As Double   'array of column bases values
                                            'used to determine hittest feature.

Public Enum LegendPrintConstants            'the enumerated for legend printing
    legPrintNone = 0
    legPrintGraph
    legPrintText
End Enum

Private uLegendPrintMode As LegendPrintConstants

Public Enum PrinterFitConstants             'the enumerated for printing
    prtFitCentered = 0
    prtFitStretched
    prtFitTopLeft
    prtFitTopRight
    prtFitBottomLeft
    prtFitBottomRight
End Enum

Private uPrinterFit As PrinterFitConstants
Private uPrinterOrientation As PrinterObjectConstants
Private uDataFormat       As String       'the data format for numeric values
Private uPicture          As StdPicture   'the background picture
Private uPictureTile      As Boolean      'marker indicating if the background picture must be tiled
                                          '(TRUE) or stretched (FALSE)
Private uAutoRedraw       As Boolean      'indicates if the chart is auto-redrawn upon every property change
Private uRangeY           As Integer      'the absolute range between Y-axis min. ad max. values
Private uDataTypeY        As Integer      'indicates the data distribution in the Y axis
Private uRangeX           As Integer      'the absolute range between X-axis min. ad max. values
Private uDataTypeX        As Integer      'indicates the data distribution in the X axis
Private Const DT_BOTH = 0                 ' 0 = range(-Y0, +Y1)
Private Const DT_NEG = 1                  ' 1 = range(-Y0, -Y1)
Private Const DT_POS = 2                  ' 2 = range(+Y0, +Y1)

Private uQuadrantXValue   As Double       'maximum quadrant-x value
Private uQuadrantYValue   As Double       'maximum quadrant-y value
Private uMaxYValue        As Double       'maximum y value
Private uMaxXValue        As Double       'maximum x value
Private uMinYValue        As Double       'minimum y value
Private uMinXValue        As Double       'minimum x value
Private uMarkerColor      As Long         'the color of the symbol
Private uMinorGridColor   As Long         'the minor intersect grid color
Private uMajorGridColor   As Long         'the major intersect grid color
Private uMinorGridOn      As Boolean      'marker indicating display of minor grid
Private uMajorGridOn      As Boolean      'marker indicating display of major grid
Private uLegendBackColor  As Long         'the legend background color
Private uLegendForeColor  As Long         'the legend foreground color
Private uInfoBackColor    As Long         'the information picBox background color
Private uInfoForeColor    As Long         'the information picBox foreground color
Private uXAxisLabelColor  As Long         'the X axis label color
Private uYAxisLabelColor  As Long         'the Y axis label color
Private uXAxisItemsColor  As Long         'the X axis items color
Private uYAxisItemsColor  As Long         'the Y axis items color
Private uChartTitleColor  As Long         'the chart title color
Private uChartSubTitleColor As Long       'the chart subtitle color
Private uSelectedColumn   As Integer      'marker indicating the selected column
Private uSelectable       As Boolean      'marker indicating whether user can select a column
Private uSelectedColor    As Long         'the selected marker forecolor
Private uInnerColor       As Long         'the inner background color
Private uSaveAsCaption    As String       'the SaveAs dialog picBox caption
Private uOldSelection     As Long

Private uInfoItems        As String       'the information items (to be displayed in the info picBox)
Private Const INFO_ITEMS = "Value XY|Description"

Private uInfoQuadrantBackColor    As Long         'the quadrant information picBox background color
Private uInfoQuadrantForeColor    As Long         'the quadrant information picBox foreground color
Private uInfoQuadrantItems        As String       'the information items (to be displayed in the info picBox)
Private Const INFO_QUADRANT_ITEMS = "Quadrant 1|Quadrant 2|Quadrant 3|Quadrant 4"

Public Enum ChartMenuConstants             'the enumerated for menu type
    xcPopUpMenu = 0
    xcButtonMenu
End Enum

Private uMenuType         As ChartMenuConstants 'the menu type.
Private uMenuItems        As String       'the menu's items.
Private Const MENU_ITEMS = "&Save as...|&Print|&Copy|Selection &information|&Quadrant information|&Legend|&Hide"

Private uCustomMenuItems  As String       'the custom menu's items.
Private Const CUSTOM_MENU_ITEMS = Empty

Private uLegendCaption    As String       'the legend's tooltip string
Private Const LEGEND_CAPTION = "Display legend"

Private Const IDX_SAVE = 0                'the command buttons' indexs
Private Const IDX_PRINT = 1
Private Const IDX_COPY = 2
Private Const IDX_INFO = 3
Private Const IDX_QUAD_INFO = 4
Private Const IDX_LEGEND = 5

Private uColWidth         As Single       'the calculated width of each column
Private uRowHeightPortion As Single       'the minimum height of a column
Private uColWidthPortion  As Single       'the minimum width of a column
Private uTopMargin        As Single       '--------------------------------------
Private uBottomMargin     As Single       'margins used around the chart content
Private uLeftMargin       As Single       '
Private uRightMargin      As Single
Private uRightMarginOrg   As Single       '--------------------------------------
Private uContentBorder    As Boolean      'border around the chart content?
Private uDisplayDescript  As Boolean      'display description when selectable
Private uDisplayQuadrantDescript As Boolean  'display quadrant description
Private uChartTitle       As String       'chart title
Private uChartSubTitle    As String       'chart sub title
Private uChartAsQuadrant  As Boolean      'chart as quadrant (divide chart into 4 quadrants)
Private uAxisXOn          As Boolean      'marker indicating display of x axis
Private uAxisYOn          As Boolean      'marker indicating display of y axis
Private uIntersectMajorY  As Single       'major intersect value
Private uIntersectMinorY  As Single       'minor intersect value
Private uIntersectMajorX  As Single       'major intersect value
Private uIntersectMinorX  As Single       'minor intersect value
Private uXAxisLabel       As String       'label to be displayed below the X-Axis
Private uYAxisLabel       As String       'label to be displayed left of the Y-Axis
Private uHotTracking      As Boolean      'marker indicating use of hot tracking
Private cItems            As Collection   'collection of chart items

Private Const QUADRANT_COLORS = "255|16711680|65280|65535"   'red|blue|green|yellow
Private uQuadrantColor(3) As Long         'the colors of the quadrants
Private uQuadrantColors   As String       'the colors of the quadrants
Private uQuadrantDividerColor As Long     'the color of the quadrant divider
Private uQuadrantColorsOverridePicture As Boolean  'the colors of the quadrants override the back picture

Public Enum MarkerDirectionConstants      'the enumerated for marker direction
    xcMarkerUp = 0
    xcMarkerDown
    xcMarkerRight
    xcMarkerLeft
End Enum

Public Enum MarkerSymbolConstants         'the enumerated for marker symbol
    xcMarkerSymBox = 0
    xcMarkerSymCircle
    xcMarkerSymTriangle
    xcMarkerSymTrapezium
    xcMarkerSymRhombus
End Enum

Private uMarkerSymbol     As MarkerSymbolConstants  'the marker type to be displayed
Private uMarkerWidth      As Integer     'the marker width
Private uMarkerLabelAngle As Integer     'rotation degree of marker label
Private uMarkerLabelColor As Long        'the color of the marker
Private uMarkerLabelDirection As MarkerDirectionConstants

Private offsetX           As Long
Private offsetY           As Long

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean
Private bResizeLegend     As Boolean

Private bProcessingOver   As Boolean      'marker to speed up mouse over effects

Public Type ChartItem
    ItemID As String
    SelectedDescription As String
    LegendDescription As String
    Description As String
    XValue As Double
    YValue As Double
End Type

Public Event ItemClick(cItem As ChartItem)
Public Event MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'-----------------------------------------------
' for Ballon ToolTip
'-----------------------------------------------
Private ttpBalloon As New Tooltip

Public Property Let QuadrantX(dblVal As Double)
    If dblVal >= uMinXValue And dblVal <= uMaxXValue Then
        uQuadrantXValue = dblVal
        DrawChart
        PropertyChanged "QuadrantX"
    End If
End Property

Public Property Get QuadrantY() As Double
    QuadrantY = uQuadrantYValue
End Property

Public Property Let QuadrantY(dblVal As Double)
    If dblVal >= uMinYValue And dblVal <= uMaxYValue Then
        uQuadrantYValue = dblVal
        DrawChart
        PropertyChanged "QuadrantY"
    End If
End Property


Public Property Get QuadrantX() As Double
    QuadrantX = uQuadrantXValue
End Property

Private Function DrawMarkerpicBox(sngX As Single, sngY As Single, lngColor As Long) As PointAPI()
    
    Dim uaPts(1) As PointAPI
    Dim sngMarkYOff As Single
    Dim sngMarkXOff As Single
    
    sngMarkYOff = uMarkerWidth / 2 * Screen.TwipsPerPixelY
    sngMarkXOff = uMarkerWidth / 2 * Screen.TwipsPerPixelX
    uaPts(0).X = sngX - sngMarkXOff
    uaPts(1).X = sngX + sngMarkXOff
    uaPts(0).Y = sngY - sngMarkYOff
    uaPts(1).Y = sngY + sngMarkYOff
    UserControl.Line (uaPts(0).X, uaPts(0).Y)-(uaPts(1).X, uaPts(1).Y), lngColor, BF
    
    DrawMarkerpicBox = uaPts()

    'free the memory
    Erase uaPts

End Function

Private Function DrawMarkerCircle(sngX As Single, sngY As Single, lngColor As Long) As PointAPI()

    Dim lngFillColor As Long
    Dim lngFillStyle As Long
    Dim uaPts(1) As PointAPI
    Dim sngMarkYOff As Single
    Dim sngMarkXOff As Single
    
    sngMarkYOff = uMarkerWidth / 2 * Screen.TwipsPerPixelY
    sngMarkXOff = uMarkerWidth / 2 * Screen.TwipsPerPixelX
    With UserControl
        lngFillColor = .FillColor
        lngFillStyle = .FillStyle
        .FillColor = lngColor
        .FillStyle = vbFSSolid
        UserControl.Circle (sngX, sngY), sngMarkXOff, uMarkerColor
        .FillColor = lngFillColor
        .FillStyle = lngFillStyle
    End With
    uaPts(0).X = sngX - sngMarkXOff
    uaPts(1).X = sngX + sngMarkXOff
    uaPts(0).Y = sngY - sngMarkYOff
    uaPts(1).Y = sngY + sngMarkYOff

    DrawMarkerCircle = uaPts()

    'free the memory
    Erase uaPts

End Function

Private Function DrawMarkerTriangle(sngX As Single, sngY As Single, lngColor As Long) As PointAPI()

    'input parameters represent the center of the triangle
    
    On Error Resume Next
    
    Dim lRet As Long
    Dim lngFillColor As Long
    Dim lngFillStyle As Long
    Dim sngMarkYOff As Single
    Dim sngMarkXOff As Single
    Dim intScaleMode As Integer
    Dim uaPts(2) As PointAPI
    Dim uaPtspicBox(1) As PointAPI

    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 3 points of the triangle anti-clockwise
    '     (1)
    '    /   \
    '   /     \
    ' (0)-----(2)
    
    sngMarkYOff = uMarkerWidth * Screen.TwipsPerPixelY / 2
    sngMarkXOff = uMarkerWidth * Screen.TwipsPerPixelX / 2
    
    uaPts(0).X = (sngX - sngMarkXOff) / Screen.TwipsPerPixelX
    uaPts(1).X = sngX / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX + sngMarkXOff) / Screen.TwipsPerPixelX
    
    uaPts(0).Y = (sngY + sngMarkYOff) / Screen.TwipsPerPixelY
    uaPts(1).Y = sngY / Screen.TwipsPerPixelY
    uaPts(2).Y = (sngY + sngMarkYOff) / Screen.TwipsPerPixelY
    
    'draw the filled triangle
    lngFillColor = UserControl.FillColor
    lngFillStyle = UserControl.FillStyle
    UserControl.FillStyle = vbSolid
    UserControl.FillColor = lngColor
    lRet = Polygon(UserControl.hDC, uaPts(0), 3)
    UserControl.FillColor = lngFillColor
    UserControl.FillStyle = lngFillStyle
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'assign return values
    uaPtspicBox(0).X = uaPts(0).X * Screen.TwipsPerPixelX
    uaPtspicBox(0).Y = uaPts(1).Y * Screen.TwipsPerPixelY
    uaPtspicBox(1).X = uaPts(2).X * Screen.TwipsPerPixelX
    uaPtspicBox(1).Y = uaPts(2).Y * Screen.TwipsPerPixelY
    
    'free the memory
    Erase uaPts
    
    DrawMarkerTriangle = uaPtspicBox()
    
End Function


Private Function DrawMarkerRhombus(sngX As Single, sngY As Single, lngColor As Long) As PointAPI()

    'input parameters represent the center of the rhombus
    
    On Error Resume Next
    
    Dim lRet As Long
    Dim lngFillStyle As Long
    Dim lngFillColor As Long
    Dim sngMarkYOff As Single
    Dim sngMarkXOff As Single
    Dim intScaleMode As Integer
    Dim uaPts(3) As PointAPI
    Dim uaPtspicBox(1) As PointAPI
    
    sngMarkXOff = (uMarkerWidth * 2 * Screen.TwipsPerPixelX) / 4    'consider the 25% as X-offset
    sngMarkYOff = uMarkerWidth / 2 * Screen.TwipsPerPixelY          'consider the 50% as Y-offset
    
    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 4 points of the Rhombus anti-clockwise
    '     (1)
    '    /   \
    '   /     \
    ' (0)     (2)
    '   \     /
    '    \   /
    '     (3)
    
    uaPts(0).X = (sngX - sngMarkXOff) / Screen.TwipsPerPixelX
    uaPts(1).X = sngX / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX + sngMarkXOff) / Screen.TwipsPerPixelX
    uaPts(3).X = sngX / Screen.TwipsPerPixelX
    
    uaPts(0).Y = sngY / Screen.TwipsPerPixelY
    uaPts(1).Y = (sngY - sngMarkYOff) / Screen.TwipsPerPixelY
    uaPts(2).Y = sngY / Screen.TwipsPerPixelY
    uaPts(3).Y = (sngY + sngMarkYOff) / Screen.TwipsPerPixelY
    
    'draw the filled Rhombus
    lngFillColor = UserControl.FillColor
    lngFillStyle = UserControl.FillStyle
    UserControl.FillStyle = vbSolid
    UserControl.FillColor = lngColor
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    UserControl.FillColor = lngFillColor
    UserControl.FillStyle = lngFillStyle
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'assign return values
    uaPtspicBox(0).X = uaPts(0).X * Screen.TwipsPerPixelX
    uaPtspicBox(0).Y = uaPts(1).Y * Screen.TwipsPerPixelY
    uaPtspicBox(1).X = uaPts(2).X * Screen.TwipsPerPixelX
    uaPtspicBox(1).Y = uaPts(3).Y * Screen.TwipsPerPixelY
    
    'free the memory
    Erase uaPts
    
    DrawMarkerRhombus = uaPtspicBox()

End Function


Private Function DrawMarkerTrapezium(sngX As Single, sngY As Single, lngColor As Long) As PointAPI()

    'input parameters represent the center of the trapezium
    
    On Error Resume Next
    
    Dim lRet As Long
    Dim lngFillStyle As Long
    Dim lngFillColor As Long
    Dim sngMarkYOff As Single
    Dim sngMarkXOff As Single
    Dim intScaleMode As Integer
    Dim uaPts(3) As PointAPI
    Dim uaPtspicBox(1) As PointAPI
    
    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 4 points of the trapezium
    sngMarkYOff = (uMarkerWidth * Screen.TwipsPerPixelX) / 2 'consider the 50% as Y-offset
    sngMarkXOff = (uMarkerWidth * Screen.TwipsPerPixelX * 2) / 4 'consider the 25% as X-offset
    'set the points anti-clockwise
    '     (1)-----(2)
    '    /           \
    '   /             \
    ' (0)-------------(3)
    
    uaPts(0).X = (sngX - sngMarkXOff * 2) / Screen.TwipsPerPixelX
    uaPts(1).X = (sngX - sngMarkXOff) / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX + sngMarkXOff) / Screen.TwipsPerPixelX
    uaPts(3).X = (sngX + sngMarkXOff * 2) / Screen.TwipsPerPixelX
    
    uaPts(0).Y = (sngY + sngMarkYOff) / Screen.TwipsPerPixelY
    uaPts(1).Y = (sngY - sngMarkYOff) / Screen.TwipsPerPixelY
    uaPts(2).Y = (sngY - sngMarkYOff) / Screen.TwipsPerPixelY
    uaPts(3).Y = (sngY + sngMarkYOff) / Screen.TwipsPerPixelY
    
    'draw the filled trapezium
    lngFillColor = UserControl.FillColor
    lngFillStyle = UserControl.FillStyle
    UserControl.FillStyle = vbSolid
    UserControl.FillColor = lngColor
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    UserControl.FillColor = lngFillColor
    UserControl.FillStyle = lngFillStyle
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'assign return values
    uaPtspicBox(0).X = uaPts(0).X * Screen.TwipsPerPixelX
    uaPtspicBox(0).Y = uaPts(1).Y * Screen.TwipsPerPixelY
    uaPtspicBox(1).X = uaPts(3).X * Screen.TwipsPerPixelX
    uaPtspicBox(1).Y = uaPts(1).Y * Screen.TwipsPerPixelY
    
    'free the memory
    Erase uaPts
    
    DrawMarkerTrapezium = uaPtspicBox()

End Function

Public Property Let LegendPrintMode(val As LegendPrintConstants)
    uLegendPrintMode = val
    PropertyChanged "LegendPrintMode"
End Property

Public Property Get LegendPrintMode() As LegendPrintConstants
    LegendPrintMode = uLegendPrintMode
End Property

Public Property Let Selectable(blnVal As Boolean)
    If blnVal <> uSelectable Then
        uSelectable = blnVal
        DrawChart
        PropertyChanged "Selectable"
    End If
End Property

Public Property Get Selectable() As Boolean
    Selectable = uSelectable
End Property
Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = uSelectedColor
End Property
Public Property Let SelectedColor(lngVal As OLE_COLOR)
    If lngVal <> uSelectedColor Then
        uSelectedColor = lngVal
        PropertyChanged "SelectedColor"
    End If
End Property


Public Function AddItem(cItem As ChartItem) As Boolean
    
    cItems.Add cItem
    
End Function

Public Property Let AutoRedraw(blnVal As Boolean)
    If blnVal <> uAutoRedraw Then
        uAutoRedraw = blnVal
        DrawChart
        PropertyChanged "AutoRedraw"
    End If
End Property

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the option to force the chart redrawing upon each change."
    
    AutoRedraw = uAutoRedraw

End Property
Public Property Get MarkerLabelAngle() As Integer
Attribute MarkerLabelAngle.VB_Description = "Returns/sets a value that determines the rotation angle of the marker label."
    
    MarkerLabelAngle = uMarkerLabelAngle

End Property
Public Property Get MarkerLabelDirection() As MarkerDirectionConstants
Attribute MarkerLabelDirection.VB_Description = "Returns/sets a value that determines the direction of the marker label."
    
    MarkerLabelDirection = uMarkerLabelDirection

End Property



Public Property Get DataFormat() As String
Attribute DataFormat.VB_Description = "Determines the format which the Y-values are displayed with."
    DataFormat = uDataFormat
End Property

Public Property Get PrinterOrientation() As PrinterObjectConstants
Attribute PrinterOrientation.VB_Description = "Returns/sets a value that determines the orientation of the output sent to the printer."
    PrinterOrientation = uPrinterOrientation
End Property
Public Property Get PrinterFit() As PrinterFitConstants
    PrinterFit = uPrinterFit
End Property

Public Property Let DataFormat(stgVal As String)
    uDataFormat = stgVal
    PropertyChanged "DataFormat"
End Property

Public Property Let PrinterOrientation(intVal As PrinterObjectConstants)
    If intVal = vbPRORLandscape Or intVal = vbPRORPortrait Then
        uPrinterOrientation = intVal
        PropertyChanged "PrinterOrientation"
    End If
End Property
Public Property Let PrinterFit(intVal As PrinterFitConstants)
    uPrinterFit = intVal
    PropertyChanged "PrinterFit"
End Property
Private Sub DisplayInfo(intIdx As Integer)

    Dim sDescription    As String
    Dim varItems        As Variant
    
    'it's important to let the info label invisible at beginning to avoid flickering effect
    lblInfo.Visible = False
    If uDisplayDescript Then
        If intIdx > -1 Then
            With cItems.Item(intIdx + 1)
                'this kind of error trapping is useful in case the user
                'did not define any item in the menu items string, so the default is used
                On Error GoTo DrawChart_error
        
                If uInfoItems = Empty Then uInfoItems = INFO_ITEMS
                varItems = Split(uInfoItems, "|")
                sDescription = CStr(varItems(0)) & ": " & Format(.XValue, uDataFormat) & "/" & Format(.YValue, uDataFormat)
                If Len(.SelectedDescription) > 0 Then
                    sDescription = CStr(varItems(1)) & ": " & .SelectedDescription & vbCrLf & sDescription
                End If
            End With
        End If
        If sDescription <> Empty Then
            lblInfo.Caption = sDescription
            lblInfo.Width = UserControl.TextWidth(sDescription) + 5 * Screen.TwipsPerPixelX
            lblInfo.Height = UserControl.TextHeight(sDescription) * 1.2
            lblInfo.Visible = True
        End If
    End If
    Exit Sub

DrawChart_error:
    uInfoItems = INFO_ITEMS
    Resume Next

End Sub

Private Sub DisplayQuadrantInfo()

    Dim x1 As Single
    Dim y1 As Single
    Dim sngY As Single
    Dim intIdx As Integer
    Dim varItems As Variant
    Dim sDescription As String

    'it's important to let the info label invisible at beginning to avoid flickering effect
    picInfoQuadrant.Visible = False
    If uDisplayQuadrantDescript Then
        'this kind of error trapping is useful in case the user
        'did not define any item in the menu items string, so the default is used
        On Error GoTo DisplayQuadrantInfo_error
    
        If uInfoQuadrantItems = Empty Then uInfoQuadrantItems = INFO_QUADRANT_ITEMS
        varItems = Split(uInfoQuadrantItems, "|")
        For intIdx = 0 To UBound(varItems)
            sDescription = sDescription & CStr(varItems(intIdx)) & vbCrLf
        Next
        If sDescription <> Empty Then
            With picInfoQuadrant
                .BackColor = uInfoQuadrantBackColor
                .ForeColor = uInfoQuadrantForeColor
                .Cls
                .Width = .TextWidth(sDescription) + 15 * Screen.TwipsPerPixelX
                .Height = (UBound(varItems) + 1) * .TextHeight("A") + 5 * Screen.TwipsPerPixelY
                sngY = 2 * Screen.TwipsPerPixelY
                For intIdx = 0 To UBound(varItems)
                    x1 = 3 * Screen.TwipsPerPixelX
                    y1 = sngY + 4 * Screen.TwipsPerPixelY
                    .CurrentY = sngY
                    .CurrentX = 10 * Screen.TwipsPerPixelX
                    picInfoQuadrant.Print CStr(varItems(intIdx))
                    sngY = .CurrentY
                    picInfoQuadrant.Line (x1, y1)-(x1 + 3 * Screen.TwipsPerPixelX, y1 + 3 * Screen.TwipsPerPixelY), uQuadrantColor(intIdx), BF
                Next
                .Visible = True
            End With
        End If
    End If
    Exit Sub

DisplayQuadrantInfo_error:
    uInfoQuadrantItems = INFO_QUADRANT_ITEMS
    Resume Next

End Sub

Private Sub DrawPicture(sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, blnTile As Boolean, pic As StdPicture)

    On Error Resume Next
    
    Dim x1 As Single
    Dim x2 As Single
    Dim y1 As Single
    Dim y2 As Single
    Dim sngH As Single
    Dim sngW As Single
    Dim xTemp As Single
    Dim yTemp As Single
    
    If blnTile = True Then
        'I found the ratio of 1.75 to adjust size, but I really don't know why!!!
        sngH = Round(pic.Height / 1.75)
        sngW = Round(pic.Width / 1.75)
        If (sngH Mod Screen.TwipsPerPixelY) <> 0 Then
            sngH = Round(sngH / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        End If
        If (sngW Mod Screen.TwipsPerPixelX) <> 0 Then
            sngW = Round(sngW / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        End If
        y1 = sngY1
        y2 = sngY2
        x2 = sngX2
        Do While y1 < y2
            x1 = sngX1
            Do While x1 < x2
                If (x1 + sngW) > x2 Then
                    xTemp = (x2 - x1)
                Else
                    xTemp = sngW
                End If
                xTemp = IIf(xTemp < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, xTemp)
                If (y1 + sngH) > y2 Then
                    yTemp = (y2 - y1)
                Else
                    yTemp = sngH
                End If
                yTemp = IIf(yTemp < Screen.TwipsPerPixelY, Screen.TwipsPerPixelY, yTemp)
'If (yTemp Mod Screen.TwipsPerPixelY) <> 0 Then
'    yTemp = Round(yTemp / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
'End If
'If (xTemp Mod Screen.TwipsPerPixelX) <> 0 Then
'    xTemp = Round(xTemp / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
'End If
                UserControl.PaintPicture pic, _
                            x1, y1, _
                            xTemp, _
                            yTemp, _
                            0, 0, xTemp, yTemp
                x1 = (x1 + sngW)
            Loop
            y1 = (y1 + sngH)
        Loop
    Else
        'stretch the picture
        UserControl.PaintPicture pic, _
                            sngX1, sngY1, _
                            IIf((sngX2 - sngX1) < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, (sngX2 - sngX1)), _
                            IIf((sngY2 - sngY1) < Screen.TwipsPerPixelY, Screen.TwipsPerPixelY, (sngY2 - sngY1))
    End If

End Sub

Private Sub FixLegendCaption()
    uLegendCaption = IIf(uLegendCaption = Empty, LEGEND_CAPTION, uLegendCaption)
End Sub

Public Property Let LegendCaption(stgVal As String)
    uLegendCaption = stgVal
    FixLegendCaption
End Property

Public Property Let MarkerWidth(intVal As Integer)
    
    If intVal <> uMarkerWidth Then
        If intVal > 0 And intVal <= 16 Then
            uMarkerWidth = intVal
            DrawChart
            PropertyChanged "MarkerWidth"
        End If
    End If

End Property

Public Property Get MarkerWidth() As Integer
Attribute MarkerWidth.VB_Description = "Returns/sets the width of the line displayed in the chart."
    MarkerWidth = uMarkerWidth
End Property
Public Property Get MinorGridOn() As Boolean
Attribute MinorGridOn.VB_Description = "Returns/sets a value that determines if the minor grid is visible or hidden."
    MinorGridOn = uMinorGridOn
End Property

Public Property Get MajorGridOn() As Boolean
Attribute MajorGridOn.VB_Description = "Returns/sets a value that determines if the major grid is visible or hidden."
    MajorGridOn = uMajorGridOn
End Property

Public Property Let MinorGridOn(blnVal As Boolean)
    If blnVal <> uMinorGridOn Then
        uMinorGridOn = blnVal
        DrawChart
        PropertyChanged "MinorGridOn"
    End If
End Property

Public Property Let MajorGridOn(blnVal As Boolean)
    If blnVal <> uMajorGridOn Then
        uMajorGridOn = blnVal
        DrawChart
        PropertyChanged "MajorGridOn"
    End If
End Property

Public Property Set Picture(ByVal picVal As StdPicture)
    Set uPicture = picVal
    DrawChart
End Property


Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed as background of the chart."
    Set Picture = uPicture
End Property

Public Property Get MarkerSymbol() As MarkerSymbolConstants
Attribute MarkerSymbol.VB_Description = "Returns/sets the character to be displayed in place of the bar."
    MarkerSymbol = uMarkerSymbol
End Property

Public Property Let MarkerSymbol(intVal As MarkerSymbolConstants)
    If intVal <> uMarkerSymbol Then
        uMarkerSymbol = intVal
        DrawChart
        PropertyChanged "MarkerSymbol"
    End If
End Property

Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

Private Sub FixData()

    If uMinYValue < 0 And uMaxYValue < 0 Then
        uDataTypeY = DT_NEG
        uRangeY = (Abs(uMinYValue) - Abs(uMaxYValue))
    ElseIf uMinYValue >= 0 And uMaxYValue >= 0 Then
        uDataTypeY = DT_POS
        uRangeY = (Abs(uMaxYValue) - Abs(uMinYValue))
    Else
        uDataTypeY = DT_BOTH
        uRangeY = (Abs(uMaxYValue) + Abs(uMinYValue))
    End If

    If uRangeY = 0 Then uRangeY = 1
    If uIntersectMajorY = 0 Then uIntersectMajorY = uRangeY / 10
    If uIntersectMinorY = 0 Then uIntersectMinorY = uIntersectMajorY / 5
    
    If uMinXValue < 0 And uMaxXValue < 0 Then
        uDataTypeX = DT_NEG
        uRangeX = (Abs(uMinXValue) - Abs(uMaxXValue))
    ElseIf uMinXValue >= 0 And uMaxXValue >= 0 Then
        uDataTypeX = DT_POS
        uRangeX = (Abs(uMaxXValue) - Abs(uMinXValue))
    Else
        uDataTypeX = DT_BOTH
        uRangeX = (Abs(uMaxXValue) + Abs(uMinXValue))
    End If

    If uRangeX = 0 Then uRangeX = 1
    If uIntersectMajorX = 0 Then uIntersectMajorX = uRangeX / 10
    If uIntersectMinorX = 0 Then uIntersectMinorX = uIntersectMajorX / 5

End Sub

Private Sub FixMenu()
    
    'this kind of error trapping is useful in case the user
    'did not define any item in the menu items string, so the default is used
    On Error GoTo FixMenu_error
    
    Dim varItems As Variant
    
    If uMenuItems = Empty Then
        uMenuItems = MENU_ITEMS
    End If
    varItems = Split(uMenuItems, "|")
    
    If varItems(0) <> Empty Then
        mnuMainSaveAs.Caption = CStr(varItems(0))
    Else
        mnuMainSaveAs.Caption = "&Save as..."
    End If
    
    If varItems(1) <> Empty Then
        mnuMainPrint.Caption = CStr(varItems(1))
    Else
        mnuMainPrint.Caption = "&Print"
    End If
    
    If varItems(2) <> Empty Then
        mnuMainCopy.Caption = CStr(varItems(2))
    Else
        mnuMainCopy.Caption = "&Copy"
    End If
    
    If varItems(3) <> Empty Then
        mnuMainSelectionInfo.Caption = CStr(varItems(3))
    Else
        mnuMainSelectionInfo.Caption = "Selection &information"
    End If
    
    If varItems(4) <> Empty Then
        mnuMainQuadrantInfo.Caption = CStr(varItems(4))
    Else
        mnuMainQuadrantInfo.Caption = "&Quadrant information"
    End If
    
    If varItems(5) <> Empty Then
        mnuMainViewLegend.Caption = CStr(varItems(5))
    Else
        mnuMainViewLegend.Caption = "&Legend"
    End If
    
    If varItems(6) <> Empty Then
        mnuLegendHide.Caption = CStr(varItems(6))
    Else
        mnuLegendHide.Caption = "&Hide"
    End If

    If uMenuType = xcButtonMenu Then
        picCommands.Visible = True
        picCommands.BackColor = UserControl.BackColor
        picCommands.Move 60, 60
        If lblInfo.Visible = False Then
            lblInfo.Move picCommands.Left + picCommands.ScaleWidth + 60, 60
        End If
    Else
        picCommands.Visible = False
        If lblInfo.Visible = False Then
            lblInfo.Move 60, 60
        End If
    End If
    If picInfoQuadrant.Visible = False Then
        picInfoQuadrant.Move lblInfo.Left, lblInfo.Top + lblInfo.Height + 60
    End If
    Exit Sub
    
FixMenu_error:
    uMenuItems = MENU_ITEMS
    Resume Next

End Sub

Private Sub FixQuadrantColors()
    
    'this kind of error trapping is useful in case the user
    'did not define any item in the colors string, so the default is used
    On Error GoTo FixQuadrantColors_error
    
    Dim varItems As Variant
    
    If uQuadrantColors = Empty Then
        uQuadrantColors = QUADRANT_COLORS
    End If
    varItems = Split(uQuadrantColors, "|")
    
    If varItems(0) <> Empty Then
        uQuadrantColor(0) = CLng(varItems(0))
    Else
        uQuadrantColor(0) = vbRed
    End If
    
    If varItems(1) <> Empty Then
        uQuadrantColor(1) = CLng(varItems(1))
    Else
        uQuadrantColor(1) = vbBlue
    End If
    
    If varItems(2) <> Empty Then
        uQuadrantColor(2) = CLng(varItems(2))
    Else
        uQuadrantColor(2) = vbGreen
    End If
    
    If varItems(3) <> Empty Then
        uQuadrantColor(3) = CLng(varItems(3))
    Else
        uQuadrantColor(3) = vbYellow
    End If
    
    Exit Sub
    
FixQuadrantColors_error:
    uQuadrantColors = QUADRANT_COLORS
    Resume Next

End Sub

Private Sub FixCustomMenu()
    
    On Error Resume Next
    
    Dim ctl As Control
    Dim intIdx As Integer
    Dim stgItem As String
    Dim varItems As Variant
    Dim intItemCnt As Integer
    
    For Each ctl In mnuMainCustomItems
        ctl.Visible = False
    Next
    If Trim(uCustomMenuItems) <> Empty Then
        varItems = Split(uCustomMenuItems, "|")
        intItemCnt = 0
        For intIdx = 0 To UBound(varItems)
            stgItem = Trim(CStr(varItems(intIdx)))
            If stgItem <> Empty Then
                'eight items allowed in the custom menu
                If intItemCnt > 7 Then Exit For
                mnuMainCustomItems(intItemCnt).Caption = stgItem
                mnuMainCustomItems(intItemCnt).Visible = True
                intItemCnt = intItemCnt + 1
            End If
        Next
    End If
    'let the separator visible if at least one custom menu item is visible
    mnuMainCustomItemsSeparator.Visible = (mnuMainCustomItems(0).Visible)

End Sub


Public Property Let MarginTop(lMargin As Long)
    uTopMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
    PropertyChanged "MarginTop"
End Property

Public Property Get MarginTop() As Long
Attribute MarginTop.VB_Description = "Determines the distance between the top edge of the chart and the top edge of its container (in pixels)."
    MarginTop = uTopMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginBottom(lMargin As Long)
    uBottomMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
    PropertyChanged "MarginBottom"
End Property

Public Property Get MarginBottom() As Long
Attribute MarginBottom.VB_Description = "Determines the distance between the bottom edge of the chart and the bottom edge of its container (in pixels)."
    MarginBottom = uBottomMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginLeft(lMargin As Long)
    uLeftMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
    PropertyChanged "MarginLeft"
End Property

Public Property Get MarginLeft() As Long
Attribute MarginLeft.VB_Description = "Determines the distance between the left edge of the chart and the left edge of its container (in pixels)."
    MarginLeft = uLeftMargin / Screen.TwipsPerPixelX
End Property

Public Property Let MarginRight(lMargin As Long)
    uRightMargin = lMargin * Screen.TwipsPerPixelX
    uRightMarginOrg = uRightMargin
    DrawChart
    PropertyChanged "MarginRight"
End Property

Public Property Get MarginRight() As Long
Attribute MarginRight.VB_Description = "Determines the distance between the right edge of the chart and the right edge of its container (in pixels)."
    MarginRight = uRightMargin / Screen.TwipsPerPixelX
End Property

Public Property Let ContentBorder(blnVal As Boolean)
    If blnVal <> uContentBorder Then
        uContentBorder = blnVal
        DrawChart
        PropertyChanged "ContentBorder"
    End If
End Property

Public Property Get ContentBorder() As Boolean
Attribute ContentBorder.VB_Description = "Returns/sets a value that determines if the border of the chart must be drawn."
    ContentBorder = uContentBorder
End Property

Public Property Get MenuType() As ChartMenuConstants
Attribute MenuType.VB_Description = "Determines the type of the menu to be used."
    MenuType = uMenuType
End Property

Public Property Let MenuType(intVal As ChartMenuConstants)
    If intVal <> uMenuType Then
        uMenuType = intVal
        FixMenu
        PropertyChanged "MenuType"
    End If
End Property

Public Property Let PictureTile(blnVal As Boolean)
    If blnVal <> uPictureTile Then
        uPictureTile = blnVal
        DrawChart
        PropertyChanged "PictureTile"
    End If
End Property

Public Property Get PictureTile() As Boolean
Attribute PictureTile.VB_Description = "Determines if the picture used as the background of the chart must be tiled."
    PictureTile = uPictureTile
End Property

Public Sub PrintChart()
    
    On Error Resume Next
    
    Dim sngX As Single
    Dim sngY As Single
    Dim sngW As Single
    Dim sngH As Single
    Dim sngXBox As Single
    Dim sngWBox As Single
    Dim sngXDesc As Single
    Dim sngWDesc As Single
    Dim sngYoff As Single
    Dim sngXoff As Single
    Dim intIdx As Integer

    Screen.MousePointer = vbHourglass
    Printer.Orientation = uPrinterOrientation
    
    With picToPrinter
        .Cls
        sngW = IIf(bDisplayLegend = True, picSplitter.Left, UserControl.ScaleWidth)
        sngH = UserControl.ScaleHeight
        Select Case uPrinterFit
            Case prtFitStretched
                If (uLegendPrintMode = legPrintGraph) Then
                    .Width = Printer.ScaleWidth * Printer.ScaleX(picSplitter.Left, UserControl.ScaleMode, Printer.ScaleMode) / UserControl.ScaleWidth
                Else
                    .Width = Printer.ScaleWidth * Printer.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, Printer.ScaleMode) / UserControl.ScaleWidth
                End If
                .Height = Printer.ScaleHeight
                .PaintPicture UserControl.Image, 0, 0, .Width, .Height, 0, 0, sngW, sngH
            
            Case Else
                .Width = sngW
                .Height = sngH
                .PaintPicture UserControl.Image, 0, 0, sngW, sngH, 0, 0, sngW, sngH
            
        End Select
        sngW = .Width
        sngH = .Height
    End With
    
    If (bDisplayLegend = True) Then
        If (uLegendPrintMode = legPrintGraph) Then
            With picToPrinterLegend
                .Width = picLegend.Width
                .Height = picLegend.Height
                .Cls
                picToPrinterLegend.Line (0, 0)-(.Width, .Height), uLegendBackColor, BF
                picToPrinterLegend.Line (0, 0)-(2 * Screen.TwipsPerPixelX, .Height), picSplitter.BackColor, BF
                Set .Font = picDescription(0).Font
            
                sngXBox = picBox(0).Left
                sngWBox = picBox(0).Width
                sngXDesc = picDescription(0).Left
                sngWDesc = picDescription(0).ScaleWidth
                For intIdx = 0 To picBox.Count - 1
                    sngY = picBox(intIdx).Top
                    .ForeColor = uLegendForeColor
                    .CurrentX = sngXDesc
                    .CurrentY = sngY
                    picToPrinterLegend.Print picDescription(intIdx).Tag
                    picToPrinterLegend.Line (sngXBox, sngY)-(sngXBox + sngWBox, sngY + sngWBox), picBox(intIdx).BackColor, BF
                Next
                Select Case uPrinterFit
                    Case prtFitStretched
                        sngXoff = Printer.ScaleWidth * Printer.ScaleX(.ScaleWidth, .ScaleMode, Printer.ScaleMode) / UserControl.ScaleWidth
                    Case Else
                        sngXoff = .ScaleWidth
                End Select
                picToPrinter.Width = picToPrinter.Width + sngXoff
                picToPrinter.PaintPicture .Image, picToPrinter.Width - sngXoff, 0, sngXoff, sngH
            End With
        End If
    End If
    
    With picToPrinter
        Select Case uPrinterFit
            Case prtFitCentered
                sngY = ((Printer.ScaleHeight - .ScaleHeight) / 2)
                sngX = ((Printer.ScaleWidth - .ScaleWidth) / 2)
            
            Case prtFitStretched, prtFitTopLeft
                sngX = 0
                sngY = 0
        
            Case prtFitTopRight
                sngX = Printer.ScaleWidth - .ScaleWidth
                sngY = 0
        
            Case prtFitBottomLeft
                sngX = 0
                sngY = Printer.ScaleHeight - .ScaleHeight
            
            Case prtFitBottomRight
                sngX = Printer.ScaleWidth - .ScaleWidth
                sngY = Printer.ScaleHeight - .ScaleHeight
        
        End Select
    
        Printer.PaintPicture .Image, sngX, sngY, .ScaleWidth, .ScaleHeight
        Printer.EndDoc
    End With
    
    If (bDisplayLegend = True) And (uLegendPrintMode = legPrintText) Then
        Call PrintLegend
    End If
    
    Screen.MousePointer = vbDefault

End Sub


Public Sub PrintLegend()

    Dim stg As String
    Dim sngX As Single
    Dim intIdx As Integer
    Dim varItems As Variant
    Dim oChartItem As ChartItem
    
    If cItems.Count > 0 Then
        'this kind of error trapping is useful in case the user
        'did not define any item in the menu items string, so the default is used
        On Error GoTo Printlegend_error

        If uInfoItems = Empty Then uInfoItems = INFO_ITEMS
        varItems = Split(uInfoItems, "|")
        
        Printer.FontBold = True
        'dump chart title
        Printer.FontSize = UserControl.FontSize
        sngX = (Printer.ScaleWidth - Printer.TextWidth(uChartTitle)) / 2
        Printer.CurrentX = sngX
        Printer.Print uChartTitle
        
        'dump chart subtitle
        Printer.FontSize = Printer.FontSize - 2
        sngX = (Printer.ScaleWidth - Printer.TextWidth(uChartSubTitle)) / 2
        Printer.CurrentX = sngX
        Printer.Print uChartSubTitle
        Printer.FontSize = Printer.FontSize + 2
        Printer.Print
        
        Printer.FontBold = False
        For intIdx = 1 To cItems.Count
            With cItems(intIdx)
                stg = .LegendDescription & " (" & Format(.XValue, uDataFormat) & "/" & Format(.YValue, uDataFormat)
            End With
            Printer.Print stg
        Next
        
        Printer.EndDoc
    End If
    Exit Sub

Printlegend_error:
    uInfoItems = INFO_ITEMS
    Resume Next

End Sub


Public Property Get LegendCaption() As String
    LegendCaption = uLegendCaption
End Property

Public Property Let ChartTitle(stgVal As String)
    If stgVal <> uChartTitle Then
        uChartTitle = stgVal
        DrawChart
        PropertyChanged "ChartTitle"
    End If
End Property

Public Property Let ChartAsQuadrant(blnVal As Boolean)
    If blnVal <> uChartAsQuadrant Then
        uChartAsQuadrant = blnVal
        DrawChart
        PropertyChanged "ChartAsQuadrant"
    End If
End Property
Public Property Get ChartAsQuadrant() As Boolean
    ChartAsQuadrant = uChartAsQuadrant
End Property
Public Property Get QuadrantColorsOverridePicture() As Boolean
    QuadrantColorsOverridePicture = uQuadrantColorsOverridePicture
End Property



Public Property Get ChartTitle() As String
Attribute ChartTitle.VB_Description = "Determines the title of the chart."
    ChartTitle = uChartTitle
End Property

Public Property Let MenuItems(stgVal As String)
    uMenuItems = stgVal
    FixMenu
    PropertyChanged "MenuItems"
End Property

Public Property Let CustomMenuItems(stgVal As String)
    uCustomMenuItems = stgVal
    FixCustomMenu
    PropertyChanged "CustomMenuItems"
End Property


Public Property Let InfoItems(stgVal As String)
    uInfoItems = stgVal
    PropertyChanged "InfoItems"
End Property
Public Property Let InfoQuadrantItems(stgVal As String)
    uInfoQuadrantItems = stgVal
    PropertyChanged "InfoQuadrantItems"
End Property
Public Property Get InfoItems() As String
Attribute InfoItems.VB_Description = "Determines the string values displayed when selection information is enabled (separated by |)."
    InfoItems = uInfoItems
End Property
Public Property Get InfoQuadrantItems() As String
    InfoQuadrantItems = uInfoQuadrantItems
End Property


Public Property Get MenuItems() As String
Attribute MenuItems.VB_Description = "Determines the string values displayed when popup menu is enabled (separated by |)."
    MenuItems = uMenuItems
End Property
Public Property Get CustomMenuItems() As String
    CustomMenuItems = uCustomMenuItems
End Property


Public Property Let ChartSubTitle(stgVal As String)
    If stgVal <> uChartSubTitle Then
        uChartSubTitle = stgVal
        DrawChart
        PropertyChanged "ChartSubTitle"
    End If
End Property

Public Property Get ChartSubTitle() As String
Attribute ChartSubTitle.VB_Description = "Determines the subtitle of the chart."
    ChartSubTitle = uChartSubTitle
End Property

Public Property Let IntersectMajorY(sngVal As Single)
    If sngVal <> uIntersectMajorY Then
        uIntersectMajorY = sngVal
        DrawChart
        PropertyChanged "IntersectMajorY"
    End If
End Property

Public Property Let IntersectMajorX(sngVal As Single)
    If sngVal <> uIntersectMajorX Then
        uIntersectMajorX = sngVal
        DrawChart
        PropertyChanged "IntersectMajorX"
    End If
End Property


Public Property Get IntersectMajorY() As Single
Attribute IntersectMajorY.VB_Description = "Determines the value which the major intersection line is displayed for."
    IntersectMajorY = uIntersectMajorY
End Property

Public Property Get IntersectMajorX() As Single
    IntersectMajorX = uIntersectMajorX
End Property

Public Property Let IntersectMinorY(sngVal As Single)
    If sngVal <> uIntersectMinorY Then
        uIntersectMinorY = sngVal
        DrawChart
        PropertyChanged "IntersectMinorY"
    End If
End Property

Public Property Let IntersectMinorX(sngVal As Single)
    If sngVal <> uIntersectMinorX Then
        uIntersectMinorX = sngVal
        DrawChart
        PropertyChanged "IntersectMinorX"
    End If
End Property


Public Property Get IntersectMinorY() As Single
Attribute IntersectMinorY.VB_Description = "Determines the value which the minor intersection line is displayed for."
    IntersectMinorY = uIntersectMinorY
End Property

Public Property Get IntersectMinorX() As Single
    IntersectMinorX = uIntersectMinorX
End Property


Public Property Let AxisYOn(blnVal As Boolean)
Attribute AxisYOn.VB_Description = "Returns/sets the value that determines if the Y-axis items must be displayed."
    If blnVal <> uAxisYOn Then
        uAxisYOn = blnVal
        DrawChart
        PropertyChanged "AxisYOn"
    End If
End Property
Public Property Get AxisYOn() As Boolean
    AxisYOn = uAxisYOn
End Property

Public Property Let AxisXOn(blnVal As Boolean)
Attribute AxisXOn.VB_Description = "Returns/sets the value that determines if the X-axis items must be displayed."
    If blnVal <> uAxisXOn Then
        uAxisXOn = blnVal
        DrawChart
        PropertyChanged "AxisXOn"
    End If
End Property
Public Property Get AxisXOn() As Boolean
    AxisXOn = uAxisXOn
End Property

Public Property Let MaxY(dblMax As Double)
Attribute MaxY.VB_Description = "Returns/sets the maximum Y value."
    If dblMax > uMinYValue Then
        uMaxYValue = dblMax
        DrawChart
        PropertyChanged "MaxY"
    End If
End Property
Public Property Let MaxX(dblMax As Double)
    If dblMax > uMinXValue Then
        uMaxXValue = dblMax
        DrawChart
        PropertyChanged "MaxX"
    End If
End Property

Public Property Let MarkerLabelAngle(intValue As Integer)
    
    If intValue >= 0 And intValue <= 360 Then
        uMarkerLabelAngle = intValue
        DrawChart
        PropertyChanged "MarkerLabelAngle"
    End If

End Property
Public Property Let MarkerLabelDirection(intValue As MarkerDirectionConstants)
    
    uMarkerLabelDirection = intValue
    DrawChart
    PropertyChanged "MarkerLabelDirection"

End Property



Public Property Let MinY(dblMin As Double)
Attribute MinY.VB_Description = "Returns/sets the minimum Y value."
    If dblMin < uMaxYValue Then
        uMinYValue = dblMin
        DrawChart
        PropertyChanged "MinY"
    End If
End Property

Public Property Let MinX(dblMin As Double)
    If dblMin < uMaxXValue Then
        uMinXValue = dblMin
        DrawChart
        PropertyChanged "MinX"
    End If
End Property


Public Property Get MinY() As Double
    MinY = uMinYValue
End Property
Public Property Get MinX() As Double
    MinX = uMinXValue
End Property

Public Property Get MaxY() As Double
    MaxY = uMaxYValue
End Property
Public Property Get MaxX() As Double
    MaxX = uMaxXValue
End Property

Public Property Let QuadrantColors(stgVal As String)
    uQuadrantColors = stgVal
    DrawChart
    PropertyChanged "QuadrantColors"
End Property
Public Property Let QuadrantColorsOverridePicture(blnVal As Boolean)
    If blnVal <> QuadrantColorsOverridePicture Then
        uQuadrantColorsOverridePicture = blnVal
        DrawChart
        PropertyChanged "QuadrantColorsOverridePicture"
    End If
End Property

Public Property Get QuadrantColors() As String
    QuadrantColors = uQuadrantColors
End Property









Public Property Let SelectionInformation(blnVal As Boolean)
Attribute SelectionInformation.VB_Description = "Determines if the information box about the selected bar must be visible or hidden."
    If blnVal <> uDisplayDescript Then
        uDisplayDescript = blnVal
        DrawChart
        PropertyChanged "SelectionInformation"
    End If
End Property
Public Property Let QuadrantInformation(blnVal As Boolean)
    If blnVal <> uDisplayQuadrantDescript Then
        uDisplayQuadrantDescript = blnVal
        DrawChart
        PropertyChanged "QuadrantInformation"
    End If
End Property

Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
End Property

Public Property Get QuadrantInformation() As Boolean
    QuadrantInformation = uDisplayQuadrantDescript
End Property
Public Property Let AxisLabelY(stgCaption As String)
Attribute AxisLabelY.VB_Description = "Returns/sets the Y-axis label."
    If stgCaption <> uYAxisLabel Then
        uYAxisLabel = stgCaption
        DrawChart
        PropertyChanged "AxisLabelY"
    End If
End Property
Public Property Get AxisLabelY() As String
    AxisLabelY = uYAxisLabel
End Property

Public Property Let AxisLabelX(stgCaption As String)
Attribute AxisLabelX.VB_Description = "Returns/sets the  X-axis label."
    If stgCaption <> uXAxisLabel Then
        uXAxisLabel = stgCaption
        DrawChart
        PropertyChanged "AxisLabelX"
    End If
End Property
Public Property Let AxisLabelXColor(lngVal As OLE_COLOR)
Attribute AxisLabelXColor.VB_Description = "Returns/sets the color used to display the X-axis label."
    If lngVal <> uXAxisLabelColor Then
        uXAxisLabelColor = lngVal
        DrawChart
        PropertyChanged "AxisLabelXColor"
    End If
End Property

Public Property Let AxisLabelYColor(lngVal As OLE_COLOR)
Attribute AxisLabelYColor.VB_Description = "Returns/sets the color used to display the Y-axis label."
    If lngVal <> uYAxisLabelColor Then
        uYAxisLabelColor = lngVal
        DrawChart
        PropertyChanged "AxisLabelYColor"
    End If
End Property


Public Property Let AxisItemsYColor(lngVal As OLE_COLOR)
Attribute AxisItemsYColor.VB_Description = "Returns/sets the color used to display the Y-axis items."
    If lngVal <> uYAxisItemsColor Then
        uYAxisItemsColor = lngVal
        DrawChart
        PropertyChanged "AxisItemsYColor"
    End If
End Property
Public Property Let QuadrantDividerColor(lngVal As OLE_COLOR)
    If lngVal <> uQuadrantDividerColor Then
        uQuadrantDividerColor = lngVal
        DrawChart
        PropertyChanged "QuadrantDividerColor"
    End If
End Property
Public Property Let MarkerLabelColor(lngVal As OLE_COLOR)
    If lngVal <> uMarkerLabelColor Then
        uMarkerLabelColor = lngVal
        DrawChart
        PropertyChanged "MarkerLabelColor"
    End If
End Property

Public Property Let AxisItemsXColor(lngVal As OLE_COLOR)
Attribute AxisItemsXColor.VB_Description = "Returns/sets the color used to display the X-axis items."
    If lngVal <> uXAxisItemsColor Then
        uXAxisItemsColor = lngVal
        DrawChart
        PropertyChanged "AxisItemsXColor"
    End If
End Property
Public Property Get AxisItemsYColor() As OLE_COLOR
    AxisItemsYColor = uYAxisItemsColor
End Property
Public Property Get AxisItemsXColor() As OLE_COLOR
    AxisItemsXColor = uXAxisItemsColor
End Property
Public Property Get QuadrantDividerColor() As OLE_COLOR
    QuadrantDividerColor = uQuadrantDividerColor
End Property
Public Property Get MarkerLabelColor() As OLE_COLOR
    MarkerLabelColor = uMarkerLabelColor
End Property

Public Property Get AxisLabelYColor() As OLE_COLOR
    AxisLabelYColor = uYAxisLabelColor
End Property



Public Property Get AxisLabelXColor() As OLE_COLOR
    AxisLabelXColor = uXAxisLabelColor
End Property




Public Property Get AxisLabelX() As String
    AxisLabelX = uXAxisLabel
End Property

Public Property Let BackColor(lngVal As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns/sets the color of the chart background."
    If lngVal <> UserControl.BackColor Then
        UserControl.BackColor = lngVal
        DrawChart
        PropertyChanged "BackColor"
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get MajorGridColor() As OLE_COLOR
Attribute MajorGridColor.VB_Description = "Returns/sets the color of the major grid."
    MajorGridColor = uMajorGridColor
End Property

Public Property Get ChartTitleColor() As OLE_COLOR
Attribute ChartTitleColor.VB_Description = "Returns/sets the color used to display the chart title."
    ChartTitleColor = uChartTitleColor
End Property
Public Property Get InnerColor() As OLE_COLOR
    InnerColor = uInnerColor
End Property

Public Property Get SaveAsCaption() As String
Attribute SaveAsCaption.VB_Description = "Returns/sets the caption of the dialog box displayed when saving the picture."
    SaveAsCaption = uSaveAsCaption
End Property
Public Property Let SaveAsCaption(stgVal As String)
    uSaveAsCaption = stgVal
    PropertyChanged "SaveAsCaption"
End Property
Public Property Let ChartTitleColor(lngVal As OLE_COLOR)
    If lngVal <> uChartTitleColor Then
        uChartTitleColor = lngVal
        DrawChart
        PropertyChanged "ChartTitleColor"
    End If
End Property
Public Property Let ChartSubTitleColor(lngVal As OLE_COLOR)
Attribute ChartSubTitleColor.VB_Description = "Returns/sets the color used to display the chart subtitle."
    If lngVal <> uChartSubTitleColor Then
        uChartSubTitleColor = lngVal
        DrawChart
        PropertyChanged "ChartSubTitleColor"
    End If
End Property

Public Property Get ChartSubTitleColor() As OLE_COLOR
    ChartSubTitleColor = uChartSubTitleColor
End Property

Public Property Get MinorGridColor() As OLE_COLOR
Attribute MinorGridColor.VB_Description = "Returns/sets the color of the minor grid."
    MinorGridColor = uMinorGridColor
End Property

Public Property Let MinorGridColor(lngVal As OLE_COLOR)
    If lngVal <> uMinorGridColor Then
        uMinorGridColor = lngVal
        DrawChart
        PropertyChanged "MinorGridColor"
    End If
End Property


Public Property Let MajorGridColor(lngVal As OLE_COLOR)
    If lngVal <> uMajorGridColor Then
        uMajorGridColor = lngVal
        DrawChart
        PropertyChanged "MajorGridColor"
    End If
End Property



Public Property Get LegendBackColor() As OLE_COLOR
Attribute LegendBackColor.VB_Description = "Returns/sets the legend  background color."
    LegendBackColor = uLegendBackColor
End Property


Public Property Get LegendForeColor() As OLE_COLOR
Attribute LegendForeColor.VB_Description = "Returns/sets the legend foreground color."
    LegendForeColor = uLegendForeColor
End Property



Public Property Let LegendForeColor(lngVal As OLE_COLOR)
    If lngVal <> uLegendForeColor Then
        uLegendForeColor = lngVal
        DrawChart
        PropertyChanged "LegendForeColor"
    End If
End Property




Public Property Let InfoBackColor(lngVal As OLE_COLOR)
Attribute InfoBackColor.VB_Description = "Returns/sets the selection information background color."
    If lngVal <> uInfoBackColor Then
        uInfoBackColor = lngVal
        DrawChart
        PropertyChanged "InfoBackColor"
    End If
End Property
Public Property Let InfoQuadrantBackColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoQuadrantBackColor Then
        uInfoQuadrantBackColor = lngVal
        DrawChart
        PropertyChanged "InfoQuadrantBackColor"
    End If
End Property

Public Property Let InfoForeColor(lngVal As OLE_COLOR)
Attribute InfoForeColor.VB_Description = "Returns/sets the selection information foreground color."
    If lngVal <> uInfoForeColor Then
        uInfoForeColor = lngVal
        DrawChart
        PropertyChanged "InfoForeColor"
    End If
End Property
Public Property Let InnerColor(lngVal As OLE_COLOR)
    If lngVal <> uInnerColor Then
        uInnerColor = lngVal
        DrawChart
        PropertyChanged "InnerColor"
    End If
End Property


Public Property Let InfoQuadrantForeColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoQuadrantForeColor Then
        uInfoQuadrantForeColor = lngVal
        DrawChart
        PropertyChanged "InfoQuadrantForeColor"
    End If
End Property


Public Property Get InfoBackColor() As OLE_COLOR
    InfoBackColor = uInfoBackColor
End Property

Public Property Get InfoQuadrantBackColor() As OLE_COLOR
    InfoQuadrantBackColor = uInfoQuadrantBackColor
End Property


Public Property Get InfoForeColor() As OLE_COLOR
    InfoForeColor = uInfoForeColor
End Property

Public Property Get InfoQuadrantForeColor() As OLE_COLOR
    InfoQuadrantForeColor = uInfoQuadrantForeColor
End Property
Public Property Let LegendBackColor(lngVal As OLE_COLOR)
    If lngVal <> uLegendBackColor Then
        uLegendBackColor = lngVal
        DrawChart
        PropertyChanged "LegendBackColor"
    End If
End Property

Private Sub Swap(ByRef var1 As Variant, ByRef var2 As Variant)
    
    Dim varDummy As Variant
    
    varDummy = var1
    var1 = var2
    var2 = varDummy

End Sub

Private Sub cmdCmd_Click(Index As Integer)

    Select Case Index
        Case IDX_SAVE:      mnuMainSaveAs_Click
        Case IDX_PRINT:     mnuMainPrint_Click
        Case IDX_COPY:      mnuMainCopy_Click
        Case IDX_INFO:      mnuMainSelectionInfo_Click
        Case IDX_QUAD_INFO: mnuMainQuadrantInfo_Click
        Case IDX_LEGEND:    mnuMainViewLegend_Click
    End Select

End Sub

Private Sub cmdCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    Dim stgToolTipText As String

    With ttpBalloon
        If .ObjName <> ("cmdCmd" & Index) Then
            Select Case Index
                Case IDX_SAVE:      stgToolTipText = mnuMainSaveAs.Caption
                Case IDX_PRINT:     stgToolTipText = mnuMainPrint.Caption
                Case IDX_COPY:      stgToolTipText = mnuMainCopy.Caption
                Case IDX_INFO:      stgToolTipText = mnuMainSelectionInfo.Caption
                Case IDX_LEGEND:    stgToolTipText = mnuMainViewLegend.Caption
                Case IDX_QUAD_INFO: stgToolTipText = mnuMainQuadrantInfo.Caption
            End Select
            .Title = ""
            .TipText = Replace(stgToolTipText, "&", "")
            .Icon = TTIconInfo
            .Style = TTBalloon
            .Centered = False
            Set .ParentControl = cmdCmd(Index)
            .ObjName = "cmdCmd" & Index
            .Create
        End If
    End With

End Sub


Private Sub picDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lScrollvalue As Integer
    
    If Button = vbLeftButton Then
        If uSelectable Then
            uSelectedColumn = Index
            uOldSelection = uSelectedColumn
            lScrollvalue = vsbContainer.Value
            bLegendClicked = True
            DrawChart
            'display information
            Call DisplayInfo(Index)
            vsbContainer.Value = lScrollvalue
            bLegendClicked = False
        End If
    Else
        Call picContainer_MouseDown(Button, Shift, X, Y)
    End If
    
End Sub

Public Property Get SelectedColumn() As Long
    SelectedColumn = uSelectedColumn
End Property
Public Property Let SelectedColumn(lngColumn As Long)
    
    Dim oItem As ChartItem
    On Error Resume Next
    
    If lngColumn <> uSelectedColumn Then
        uSelectedColumn = lngColumn
        DrawChart
        PropertyChanged "SelectedColumn"
        
        If Err.Number Then
            uSelectedColumn = -1
        Else
            oItem = cItems(lngColumn + 1)
            RaiseEvent ItemClick(oItem)
        End If
    End If

End Property

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        lblInfo.Drag
    Else
        PopupMenu mnuMain
    End If
End Sub


Private Sub lblSplitter_Click()
        
    mnuMainViewLegend.Checked = Not mnuMainViewLegend.Checked
    bDisplayLegend = mnuMainViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart

End Sub

Private Sub mnuMainCopy_Click()
    Clipboard.SetData UserControl.Image
End Sub

Private Sub mnuLegendHide_Click()
    mnuMainViewLegend.Checked = Not mnuMainViewLegend.Checked
    bDisplayLegend = mnuMainViewLegend.Checked
    ShowLegend True
    DrawChart
End Sub



Private Sub mnuMainCustomItems_Click(Index As Integer)
    RaiseEvent MenuItemClick(Index, mnuMainCustomItems(Index).Caption)
End Sub

Private Sub mnuMainPrint_Click()
    
    Call PrintChart

End Sub

Private Sub mnuMainQuadrantInfo_Click()
    
    mnuMainQuadrantInfo.Checked = Not mnuMainQuadrantInfo.Checked
    uDisplayQuadrantDescript = mnuMainQuadrantInfo.Checked
    Call DisplayQuadrantInfo

End Sub

Private Sub mnuMainSaveAs_Click()
   
    Dim sFilters As String
    Dim OFN As OPENFILENAME
    Dim lRet As Long
    
    'used after call
    Dim buff As String
    Dim sLname As String
    Dim sSname As String
    Dim strBuffer As String
    Dim blnReturn As Boolean
    
    'create string of filters for the dialog
    sFilters = "Windows Bitmap" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
    If uSaveAsCaption = Empty Then
        uSaveAsCaption = "Save graph"
    End If
    
    With OFN
        .nStructSize = Len(OFN)
        .hWndOwner = UserControl.hWnd
        .sFilter = sFilters
        .nFilterIndex = 0
        .sFile = "XQuadrant.bmp" & Space$(1024) & vbNullChar & vbNullChar
        .nMaxFile = Len(.sFile)
        .sDefFileExt = "bmp" & vbNullChar & vbNullChar
        .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
        .nMaxTitle = Len(OFN.sFileTitle)
        .sInitialDir = strBuffer & vbNullChar & vbNullChar
        .sDialogTitle = uSaveAsCaption
        .flags = OFS_FILE_SAVE_FLAGS
    End With
   
    'call the API
    blnReturn = GetSaveFileName(OFN)
    
    If blnReturn Then
        SavePicture UserControl.Image, OFN.sFile
    End If

End Sub

Private Sub mnuMainSelectionInfo_Click()
    
    mnuMainSelectionInfo.Checked = Not mnuMainSelectionInfo.Checked
    uDisplayDescript = mnuMainSelectionInfo.Checked
    Call DisplayInfo(uSelectedColumn)
    
End Sub

Private Sub mnuMainViewLegend_Click()
    mnuMainViewLegend.Checked = Not mnuMainViewLegend.Checked
    bDisplayLegend = mnuMainViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub


Private Sub picBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call picDescription_MouseDown(Index, Button, Shift, X, Y)

End Sub

Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub picDescription_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static intIndex As Integer
    
    If intIndex <> Index Then
        If Right$(picDescription(Index).Tag, Len(TooltipNeeded)) = TooltipNeeded Then
            intIndex = Index
            ttpBalloon.Title = ""
            ttpBalloon.TipText = picDescription(Index).Tag
            ttpBalloon.Icon = TTIconInfo
            ttpBalloon.Style = TTBalloon
            ttpBalloon.Centered = False
            Set ttpBalloon.ParentControl = picDescription(Index)
            ttpBalloon.Create
        End If
    End If

End Sub

Private Sub picInfoQuadrant_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        picInfoQuadrant.Drag
    Else
        PopupMenu mnuMain
    End If
End Sub


Private Sub picLegend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picContainer_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        If Y <= 5 * Screen.TwipsPerPixelY Then
            mnuMainViewLegend.Checked = Not mnuMainViewLegend.Checked
            bDisplayLegend = mnuMainViewLegend.Checked
            ShowLegend Not (bDisplayLegend)
            DrawChart
        Else
            bResizeLegend = True
            picSplitter.BackColor = vbButtonShadow
        End If
    End If

End Sub


Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    
    Dim sngX As Single
    
    If bResizeLegend = True Then
        sngX = X + picSplitter.Left
        If sngX >= (UserControl.Width / 2) And sngX < (UserControl.Width - picSplitter.Width) Then
            picSplitter.Left = sngX
        End If
    Else
        If Y > 5 * Screen.TwipsPerPixelY Then
            picSplitter.MousePointer = 9
            Set ttpBalloon = Nothing
        Else
            picSplitter.MousePointer = 0
            With ttpBalloon
                If .ObjName <> "picSplitter" Then
                    .Title = ""
                    .TipText = uLegendCaption
                    .Icon = TTIconInfo
                    .Style = TTBalloon
                    .Centered = False
                    Set .ParentControl = picSplitter
                    .ObjName = "picSplitter"
                    .Create
                End If
            End With
        End If
    End If

End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next

    Dim lngW As Long

    If bResizeLegend = True Then
        picSplitter.BackColor = vbButtonFace
        lngW = UserControl.ScaleWidth - picSplitter.Left - picSplitter.Width
        If lngW < 0 Then lngW = 0
        picLegend.Width = lngW
        picContainer.Width = lngW - picContainer.Left - vsbContainer.Width
        vsbContainer.Left = lngW - vsbContainer.Width
    
        uRightMargin = uRightMarginOrg
        mnuMainViewLegend.Checked = True
        bDisplayLegend = mnuMainViewLegend.Checked
        ShowLegend Not (bDisplayLegend)
        DrawChart
        bResizeLegend = False
        picSplitter.MousePointer = 0
    End If
    
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Left = X - offsetX
    Source.Top = Y - offsetY
End Sub

Private Sub UserControl_Initialize()
    Set cItems = New Collection
End Sub

Private Sub UserControl_InitProperties()
    
    Dim X As Integer
    Dim oChartItem As ChartItem
    
    uTopMargin = 50 * Screen.TwipsPerPixelY
    uBottomMargin = 55 * Screen.TwipsPerPixelY
    uLeftMargin = 55 * Screen.TwipsPerPixelX
    uRightMargin = 55 * Screen.TwipsPerPixelX
    uRightMarginOrg = uRightMargin
    uSelectedColumn = -1
    uContentBorder = True
    uChartTitle = UserControl.Name
    uChartSubTitle = ""
    uAxisYOn = True
    uAxisXOn = True
    uIntersectMajorY = 10
    uIntersectMinorY = 2
    uIntersectMajorX = 10
    uIntersectMinorX = 2
    uMaxYValue = 100
    uMaxXValue = 100
    uMinYValue = 0
    uMinXValue = 0
    uQuadrantXValue = 50
    uQuadrantYValue = 50
    uMajorGridColor = vbWhite
    uMinorGridColor = vbBlack
    UserControl.BackColor = vbWindowBackground
    UserControl.ForeColor = vbWindowText
    uInnerColor = UserControl.BackColor
    uLegendBackColor = UserControl.BackColor
    uLegendForeColor = UserControl.ForeColor
    uInfoBackColor = vbInfoBackground
    uInfoForeColor = vbInfoText
    uInfoQuadrantBackColor = vbInfoBackground
    uInfoQuadrantForeColor = vbInfoText
    uXAxisLabelColor = UserControl.ForeColor
    uYAxisLabelColor = UserControl.ForeColor
    uXAxisItemsColor = UserControl.ForeColor
    uYAxisItemsColor = UserControl.ForeColor
    uChartTitleColor = UserControl.ForeColor
    uChartSubTitleColor = UserControl.ForeColor
    uMenuType = xcPopUpMenu
    uMarkerSymbol = xcMarkerSymBox
    uMenuItems = Empty
    uCustomMenuItems = Empty
    uInfoItems = Empty
    uSaveAsCaption = Empty
    uAutoRedraw = True
    Set uPicture = Nothing
    uPictureTile = False
    uMinorGridOn = True
    uMajorGridOn = True
    uMarkerWidth = 1
    uDataFormat = Empty
    uPrinterFit = prtFitCentered
    uPrinterOrientation = vbPRORLandscape
    uLegendCaption = LEGEND_CAPTION
    uMarkerLabelAngle = 45
    uMarkerLabelDirection = xcMarkerRight
    uChartAsQuadrant = False
    uQuadrantDividerColor = UserControl.ForeColor
    uQuadrantColorsOverridePicture = False
    uMarkerLabelColor = UserControl.ForeColor
    uSelectedColor = vbCyan
    uQuadrantColors = Empty
    uHotTracking = False
    uOldSelection = -1
    uLegendPrintMode = legPrintGraph

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim oItem As ChartItem
    Dim intSelectedCol As Integer
    
    If Button = vbLeftButton Then
        
        On Error GoTo TrackExit
        
        intSelectedCol = InColumn(X, Y)
        If intSelectedCol >= 0 Then
            If Not bProcessingOver Then
                bProcessingOver = True
                uSelectedColumn = intSelectedCol
                If Not uSelectedColumn = uOldSelection Then
                    DrawChart
                    uOldSelection = uSelectedColumn
                    oItem = cItems(uSelectedColumn + 1)
                    RaiseEvent ItemClick(oItem)
                End If
                bProcessingOver = False
             End If
        End If
    ElseIf Button = vbRightButton Then
        If uMenuType = xcPopUpMenu Then
            FixMenu
            FixCustomMenu
            mnuMainSelectionInfo.Visible = (uSelectable = True)
            PopupMenu mnuMain
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
TrackExit:
    Exit Sub

End Sub

Private Function InColumn(X As Single, Y As Single) As Integer

    Dim sngY1 As Single
    Dim sngY2 As Single
    Dim sngX1 As Single
    Dim sngX2 As Single
    Dim intCol As Integer
    Dim sngTemp As Single
    Dim intSelectedCol As Integer

    intSelectedCol = -1
    If (Y <= UserControl.ScaleHeight - uBottomMargin) And (Y >= uTopMargin) _
    And (uSelectable = True) Then
        For intCol = 1 To cItems.Count
            sngY1 = uColumnsHeight(intCol - 1, 0)
            sngY2 = uColumnsHeight(intCol - 1, 1)
            sngX1 = uColumnsBase(intCol - 1, 0)
            sngX2 = uColumnsBase(intCol - 1, 1)
            If sngY1 > sngY2 Then
                sngTemp = sngY1
                sngY1 = sngY2
                sngY2 = sngTemp
            End If
            If sngX1 > sngX2 Then
                sngTemp = sngX1
                sngX1 = sngX2
                sngX2 = sngTemp
            End If
            If (Y >= sngY1 And Y <= sngY2) Then
                If (X >= sngX1 And X <= sngX2) Then
                    intSelectedCol = intCol - 1
                    Exit For
                End If
            End If
        Next
    End If
    InColumn = intSelectedCol

End Function


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (uHotTracking = True) Or (Button = vbLeftButton) Then
        'either in case of hot tracking or not, simulate the mouse left button down
        Call UserControl_MouseDown(vbLeftButton, Shift, X, Y)
    End If
    
End Sub

Public Sub Refresh()
    DrawChart
End Sub

Public Sub Clear()
    
    Set cItems = Nothing
    Set cItems = New Collection
    ClearLegendItems
    'the following forces the drawing chart routine to not enhance the description
    'in the legend (if it is visible); the legend items were already deleted!
    DrawChart

End Sub

Public Sub DrawChart()
    
    On Error Resume Next

    Dim x1              As Single
    Dim x2              As Single
    Dim y1              As Single
    Dim y2              As Single
    Dim X               As Single
    Dim Y               As Single
    Dim xTemp           As Single
    Dim yTemp           As Single
    Dim xPrev           As Single
    Dim yPrev           As Single
    Dim sngRowHeight    As Single
    Dim sngColWidth     As Single
    Dim CurrentColor    As Integer
    Dim iCols           As Integer
    Dim xMiddle         As Single
    Dim yMiddle         As Single
    Dim xMiddleArea     As Single
    Dim sngMarkXOff     As Single
    Dim sngMarkYOff     As Single
    Dim lngColor        As Long
    ReDim uaPts(0) As PointAPI
    Dim oChartItem      As ChartItem
    
    'do not redraw the chart if not required
    If uAutoRedraw = False Then Exit Sub

    'calculate the data distribution in the y-axis
    FixData
    
    With lblInfo
        .ForeColor = uInfoForeColor
        .BackColor = uInfoBackColor
        .Visible = IIf((uDisplayDescript And uSelectedColumn > -1), True, False)
    End With
    mnuMainSelectionInfo.Checked = uDisplayDescript
    picDescription(0).ForeColor = uLegendForeColor
    
    iCols = cItems.Count
    
    If Not bResize Then ClearLegendItems

    With UserControl
        uRowHeightPortion = ((.ScaleHeight - (uTopMargin + uBottomMargin)) / uRangeY)
        uColWidthPortion = (.ScaleWidth - uRightMargin - uLeftMargin) / uRangeX
        If iCols Then
            uColWidth = ((.ScaleWidth - (uLeftMargin + uRightMargin)) / iCols)
        End If
        
        .Cls
        If uPicture Is Nothing Then
            UserControl.Line (uLeftMargin, uTopMargin)-(.ScaleWidth - uRightMargin, .ScaleHeight - uBottomMargin), uInnerColor, BF
        Else
            'paint the background image
            Call DrawPicture(uLeftMargin, .ScaleWidth - uRightMargin, _
                             uTopMargin, .ScaleHeight - uBottomMargin, _
                             uPictureTile, uPicture)
        End If
    
        If iCols Then
            ReDim uColumnsHeight(iCols - 1, 1)
            ReDim uColumnsBase(iCols - 1, 1)
        End If
    
        On Error Resume Next
        
        'dump chart title
        If bDisplayLegend Then
            xMiddleArea = (picSplitter.Left / 2)
        Else
            xMiddleArea = (.ScaleWidth / 2)
        End If
        .ForeColor = uChartTitleColor
        .CurrentX = xMiddleArea - (.TextWidth(uChartTitle) / 2)
        .CurrentY = 0
        .FontBold = True
        UserControl.Print uChartTitle
        .FontBold = False
        
        'dump chart subtitle
        .ForeColor = uChartSubTitleColor
        .FontSize = .FontSize - 2
        .CurrentX = xMiddleArea - (.TextWidth(uChartSubTitle) / 2)
        UserControl.Print uChartSubTitle
        .FontSize = .FontSize + 2
        
        'check if the chart must be divided into quadrants (draw quadrants)
        If uChartAsQuadrant And uQuadrantColorsOverridePicture Then
            If uQuadrantXValue >= uMinXValue And uQuadrantXValue <= uMaxXValue Then
                If uQuadrantYValue >= uMinYValue And uQuadrantYValue <= uMaxYValue Then
                    FixQuadrantColors
                    'calculate Y position of horizontal divider
                    If uDataTypeY = DT_POS Then
                        sngRowHeight = uRowHeightPortion * (uQuadrantYValue - uMinYValue)
                        y1 = .ScaleHeight - uBottomMargin
                        y2 = y1 - sngRowHeight
                    ElseIf uDataTypeY = DT_NEG Then
                        sngRowHeight = uRowHeightPortion * (Abs(CDbl(uQuadrantYValue)) - Abs(uMaxYValue))
                        y1 = uTopMargin
                        y2 = y1 + sngRowHeight
                    Else
                        sngRowHeight = (-CDbl(uQuadrantYValue) * uRowHeightPortion)
                        y1 = .ScaleHeight - uBottomMargin
                        y1 = y1 - uRowHeightPortion * Abs(uMinYValue)
                        y2 = y1 + sngRowHeight
                    End If
                    
                    'calculate X position of vertical divider
                    If uDataTypeX = DT_POS Then
                        sngColWidth = uColWidthPortion * (uQuadrantXValue - uMinXValue)
                        x1 = uLeftMargin
                        x2 = x1 + sngColWidth
                    ElseIf uDataTypeX = DT_NEG Then
                        sngColWidth = uColWidthPortion * (Abs(CDbl(uQuadrantXValue)) - Abs(uMaxXValue))
                        x1 = .ScaleWidth - uRightMargin
                        x2 = x1 - sngColWidth
                    Else
                        sngColWidth = (CDbl(uQuadrantXValue) * uColWidthPortion)
                        x1 = .ScaleWidth - uRightMargin
                        x1 = x1 - uColWidthPortion * Abs(uMaxXValue)
                        x2 = x1 + sngColWidth
                    End If
                    UserControl.Line (uLeftMargin, uTopMargin)-(x2, y2), uQuadrantColor(0), BF
                    UserControl.Line (x2, uTopMargin)-(.ScaleWidth - uRightMargin, y2), uQuadrantColor(1), BF
                    UserControl.Line (uLeftMargin, y2)-(x2, .ScaleHeight - uBottomMargin), uQuadrantColor(2), BF
                    UserControl.Line (x2, y2)-(.ScaleWidth - uRightMargin, .ScaleHeight - uBottomMargin), uQuadrantColor(3), BF
                End If
            End If
        End If
        
        'draw border
        If uContentBorder Then
            UserControl.Line (uLeftMargin, uTopMargin)-(.ScaleWidth - uRightMargin, .ScaleHeight - uBottomMargin), uMajorGridColor, B
        End If
        
        'draw Y axis
        If uAxisYOn Then
            .ForeColor = uYAxisItemsColor
            x1 = uLeftMargin + (2 * Screen.TwipsPerPixelX)
            x2 = .ScaleWidth - uRightMargin
            If uMinorGridOn Then
                For X = uMinYValue To uMaxYValue Step uIntersectMinorY
                    y1 = (.ScaleHeight - uBottomMargin)
                    If uDataTypeY = DT_NEG Then
                        y1 = y1 + ((Abs(X) - Abs(uMinYValue)) * uRowHeightPortion)
                    Else
                        y1 = y1 - ((X - uMinYValue) * uRowHeightPortion)
                    End If
                    UserControl.Line (x1, y1)-(x2, y1), uMinorGridColor
                Next
            End If
            For X = uMinYValue To uMaxYValue Step uIntersectMajorY
                y1 = (.ScaleHeight - uBottomMargin)
                If uDataTypeY = DT_NEG Then
                    y1 = y1 + ((Abs(X) - Abs(uMinYValue)) * uRowHeightPortion)
                Else
                    y1 = y1 - ((X - uMinYValue) * uRowHeightPortion)
                End If
                If uMajorGridOn Then
                    UserControl.Line (x1, y1)-(x2, y1), uMajorGridColor
                End If
                .FontSize = .FontSize - 2
                .CurrentX = uLeftMargin - .TextWidth(X) - (5 * Screen.TwipsPerPixelX)
                .CurrentY = y1 - (.TextHeight("0") / 2)
                UserControl.Print X
                .FontSize = .FontSize + 2
            Next X
        End If
    
        'draw X axis grid
        If uAxisXOn Then
            .ForeColor = uXAxisItemsColor
            y1 = (.ScaleHeight - uBottomMargin)
            y2 = uTopMargin
            If uMinorGridOn Then
                For Y = uMinXValue To uMaxXValue Step uIntersectMinorX
                    x1 = uLeftMargin
                    If uDataTypeX = DT_NEG Then
                        x1 = x1 - ((Abs(Y) - Abs(uMinXValue)) * uColWidthPortion)
                    Else
                        x1 = x1 + ((Y - uMinXValue) * uColWidthPortion)
                    End If
                    UserControl.Line (x1, y1)-(x1, y2), uMinorGridColor
                Next Y
            End If
            For Y = uMinXValue To uMaxXValue Step uIntersectMajorX
                x1 = uLeftMargin
                If uDataTypeX = DT_NEG Then
                    x1 = x1 - ((Abs(Y) - Abs(uMinXValue)) * uColWidthPortion)
                Else
                    x1 = x1 + ((Y - uMinXValue) * uColWidthPortion)
                End If
                If uMajorGridOn Then
                    UserControl.Line (x1, y1)-(x1, y2), uMajorGridColor
                End If
                .FontSize = .FontSize - 2
                .CurrentX = x1 - (.TextWidth(Y) / 2)
                .CurrentY = (.ScaleHeight - uBottomMargin) + .TextHeight("0") + (5 * Screen.TwipsPerPixelX)
                UserControl.Print Y
                .FontSize = .FontSize + 2
            Next Y
        End If
    
        'check if the chart must be divided into quadrants (draw divider)
        If uChartAsQuadrant Then
            If uQuadrantXValue >= uMinXValue And uQuadrantXValue <= uMaxXValue Then
                If uQuadrantYValue >= uMinYValue And uQuadrantYValue <= uMaxYValue Then
                    .ForeColor = uQuadrantDividerColor
                    'calculate Y position of horizontal divider
                    If uDataTypeY = DT_POS Then
                        sngRowHeight = uRowHeightPortion * (uQuadrantYValue - uMinYValue)
                        y1 = .ScaleHeight - uBottomMargin
                        y2 = y1 - sngRowHeight
                    ElseIf uDataTypeY = DT_NEG Then
                        sngRowHeight = uRowHeightPortion * (Abs(CDbl(uQuadrantYValue)) - Abs(uMaxYValue))
                        y1 = uTopMargin
                        y2 = y1 + sngRowHeight
                    Else
                        sngRowHeight = (-CDbl(uQuadrantYValue) * uRowHeightPortion)
                        y1 = .ScaleHeight - uBottomMargin
                        y1 = y1 - uRowHeightPortion * Abs(uMinYValue)
                        y2 = y1 + sngRowHeight
                    End If
                    UserControl.Line (uLeftMargin, y2)-(.ScaleWidth - uRightMargin, y2)
                    
                    'calculate X position of vertical divider
                    If uDataTypeX = DT_POS Then
                        sngColWidth = uColWidthPortion * (uQuadrantXValue - uMinXValue)
                        x1 = uLeftMargin
                        x2 = x1 + sngColWidth
                    ElseIf uDataTypeX = DT_NEG Then
                        sngColWidth = uColWidthPortion * (Abs(CDbl(uQuadrantXValue)) - Abs(uMaxXValue))
                        x1 = .ScaleWidth - uRightMargin
                        x2 = x1 - sngColWidth
                    Else
                        sngColWidth = (CDbl(uQuadrantXValue) * uColWidthPortion)
                        x1 = .ScaleWidth - uRightMargin
                        x1 = x1 - uColWidthPortion * Abs(uMaxXValue)
                        x2 = x1 + sngColWidth
                    End If
                    UserControl.Line (x2, uTopMargin)-(x2, .ScaleHeight - uBottomMargin)
                End If
            End If
        End If
        
        'draw markers
        For X = 0 To cItems.Count - 1
            oChartItem = cItems(X + 1)
            If uDataTypeY = DT_POS Then
                sngRowHeight = uRowHeightPortion * (oChartItem.YValue - uMinYValue)
                y1 = .ScaleHeight - uBottomMargin
                y2 = y1 - sngRowHeight
            ElseIf uDataTypeY = DT_NEG Then
                sngRowHeight = uRowHeightPortion * (Abs(CDbl(oChartItem.YValue)) - Abs(uMaxYValue))
                y1 = uTopMargin
                y2 = y1 + sngRowHeight
            Else
                sngRowHeight = (-CDbl(oChartItem.YValue) * uRowHeightPortion)
                y1 = .ScaleHeight - uBottomMargin
                y1 = y1 - uRowHeightPortion * Abs(uMinYValue)
                y2 = y1 + sngRowHeight
            End If
            
            If uDataTypeX = DT_POS Then
                sngColWidth = uColWidthPortion * (oChartItem.XValue - uMinXValue)
                x1 = uLeftMargin
                x2 = x1 + sngColWidth
            ElseIf uDataTypeX = DT_NEG Then
                sngColWidth = uColWidthPortion * (Abs(CDbl(oChartItem.XValue)) - Abs(uMaxXValue))
                x1 = .ScaleWidth - uRightMargin
                x2 = x1 - sngColWidth
            Else
                sngColWidth = (CDbl(oChartItem.XValue) * uColWidthPortion)
                x1 = .ScaleWidth - uRightMargin
                x1 = x1 - uColWidthPortion * Abs(uMinXValue)
                x2 = x1 + sngColWidth
            End If
            
            lngColor = IIf((X = uSelectedColumn And uSelectable), uSelectedColor, uMarkerColor)
            Select Case uMarkerSymbol
                Case xcMarkerSymBox
                    uaPts() = DrawMarkerpicBox(x2, y2, lngColor)
                
                Case xcMarkerSymCircle
                    uaPts() = DrawMarkerCircle(x2, y2, lngColor)
                
                Case xcMarkerSymTriangle
                    uaPts() = DrawMarkerTriangle(x2, y2, lngColor)
    
                Case xcMarkerSymTrapezium
                    uaPts() = DrawMarkerTrapezium(x2, y2, lngColor)
                
                Case xcMarkerSymRhombus
                    uaPts() = DrawMarkerRhombus(x2, y2, lngColor)
    
            End Select
            'save coordinates of markers
            uColumnsHeight(X, 0) = uaPts(0).Y
            uColumnsHeight(X, 1) = uaPts(1).Y
            uColumnsBase(X, 0) = uaPts(0).X
            uColumnsBase(X, 1) = uaPts(1).X

            sngMarkYOff = uMarkerWidth * 1.5 * Screen.TwipsPerPixelY
            sngMarkXOff = uMarkerWidth * 1.5 * Screen.TwipsPerPixelX
            Select Case uMarkerLabelDirection
                Case MarkerDirectionConstants.xcMarkerDown
                    y2 = y2 + sngMarkYOff
                
                Case MarkerDirectionConstants.xcMarkerLeft
                    x2 = x2 - sngMarkXOff
                    y2 = y2 + sngMarkYOff
                
                Case MarkerDirectionConstants.xcMarkerRight
                    x2 = x2 + sngMarkXOff
                    y2 = y2 - sngMarkYOff
                
                Case MarkerDirectionConstants.xcMarkerUp
                    y2 = y2 - sngMarkYOff
            
            End Select
            .ForeColor = uMarkerLabelColor
            PrintRotText .hDC, oChartItem.Description, _
                            x2 / Screen.TwipsPerPixelX, _
                            y2 / Screen.TwipsPerPixelY, _
                            uMarkerLabelAngle
    
            If X = uSelectedColumn And uSelectable Then
                'display information
                Call DisplayInfo(CInt(X))
            End If
            
            'Add Legend item
            If Not bResize Then
                AddLegendItem oChartItem.LegendDescription, uMarkerColor, uLegendForeColor
            End If
            
        Next X
       
        'Print the x axis label
        If Len(uXAxisLabel) Then
            .FontSize = .FontSize - 1
            .CurrentY = .ScaleHeight - .TextHeight(uXAxisLabel) * 1.5
            .CurrentX = xMiddleArea - (.TextWidth(uXAxisLabel) / 2)
            .ForeColor = uXAxisLabelColor
            UserControl.Print uXAxisLabel
            .FontSize = .FontSize + 1
        End If
        
        'print the y axis label
        If Len(uYAxisLabel) > 0 Then
            .FontSize = .FontSize - 1
            .ForeColor = uYAxisLabelColor
            PrintRotText .hDC, uYAxisLabel, .TextHeight(uYAxisLabel) / Screen.TwipsPerPixelX, .ScaleHeight / 2 / Screen.TwipsPerPixelY, 90
            .FontSize = .FontSize + 1
        End If
    
        'in case the legend is displayed
        If bDisplayLegend = True Then
            picLegend.BackColor = uLegendBackColor
            picContainer.BackColor = uLegendBackColor
            If uSelectable And uSelectedColumn > -1 Then
                
                Dim perScreen As Integer
                Dim scrollValue As Integer
                            
                perScreen = Abs((picLegend.ScaleHeight / ((picBox(0).Height + (10 * Screen.TwipsPerPixelY)))) - 1)
                            
                If (uSelectedColumn + 1) > perScreen Then
                    scrollValue = ((uSelectedColumn + 1) * ((picBox(0).Height / Screen.TwipsPerPixelY) + 10)) - (picBox(perScreen).Top / Screen.TwipsPerPixelY)
                    If scrollValue > vsbContainer.Max Then scrollValue = vsbContainer.Max
                    vsbContainer.Value = scrollValue
                Else
                    vsbContainer.Value = 0
                End If
                picContainer.Line ((picBox(uSelectedColumn).Left - 3 * Screen.TwipsPerPixelX), (picBox(uSelectedColumn).Top - 3 * Screen.TwipsPerPixelY))-(picDescription(uSelectedColumn).Left + picDescription(uSelectedColumn).Width + 2 * Screen.TwipsPerPixelX, picBox(uSelectedColumn).Top + picBox(uSelectedColumn).Height + 2 * Screen.TwipsPerPixelY), uSelectedColor, B
            End If
        End If
    End With

End Sub

Public Property Let HotTracking(blnVal As Boolean)
    If blnVal <> uHotTracking Then
        uHotTracking = blnVal
        DrawChart
        PropertyChanged "HotTracking"
    End If
End Property

Public Property Get HotTracking() As Boolean
    HotTracking = uHotTracking
End Property


Public Function ShowLegend(Optional bHidden As Boolean = False)
    
    Dim stg As String

    picLegend.Line (0, 0)-(picLegend.ScaleWidth - Screen.TwipsPerPixelX, picLegend.ScaleHeight - Screen.TwipsPerPixelY), &HFFE0E0, B
    
    If bHidden Then bDisplayLegend = False Else bDisplayLegend = True
    
    If bDisplayLegend Then
        uRightMargin = uRightMargin + picLegend.ScaleWidth
        picLegend.Move UserControl.ScaleWidth - picLegend.Width + Screen.TwipsPerPixelX, 0, picLegend.Width, UserControl.ScaleHeight
        stg = Chr(187)
    Else
        uRightMargin = uRightMargin - picLegend.Width
        picLegend.Move UserControl.ScaleWidth
        stg = Chr(171)
    End If
    With picSplitter
        .Left = picLegend.Left - .Width
        .Height = picLegend.ScaleHeight
        .Cls
        picSplitter.Print stg
    End With

End Function

Private Sub AddLegendItem(sDescription As String, lngBackColor As OLE_COLOR, lngForeColor As OLE_COLOR)
    
    Dim X As Integer
    Dim sngX As Single
    Dim ShortDescript As String
    
    If bLegendAdded Then
        X = picBox.Count
        Load picBox(X)
        Load picDescription(X)
        
        picBox(X).BackColor = lngBackColor
        picBox(X).Top = picBox(X - 1).Top + picBox(X - 1).Height + 10 * Screen.TwipsPerPixelY
        picDescription(X).Top = picBox(X).Top
    Else
        X = 0
        picBox(X).BackColor = lngBackColor
        bLegendAdded = True
    End If
    
    ShortDescript = sDescription
    sngX = picDescription(X).Left
    While (Len(ShortDescript) > 0) And ((sngX + picContainer.TextWidth(ShortDescript)) > (picContainer.ScaleWidth - sngX - 5 * Screen.TwipsPerPixelX))
        ShortDescript = Left$(ShortDescript, Len(ShortDescript) - 1)
    Wend
    
    If Len(ShortDescript) < Len(sDescription) Then ShortDescript = ShortDescript & ".."
    With picDescription(X)
        .Width = picContainer.ScaleWidth - sngX - 5 * Screen.TwipsPerPixelX
        .BackColor = uLegendBackColor
        .ForeColor = lngForeColor
        'TAG is used to show tooltip
        If ShortDescript <> sDescription Then
            .Tag = sDescription & TooltipNeeded
        Else
            .Tag = sDescription
        End If
        .Cls
        picDescription(X).Print ShortDescript
        .Visible = True
    End With
            
    picBox(X).Visible = True
    picContainer.Height = ((picBox(0).Height + (10 * Screen.TwipsPerPixelY)) * picBox.Count - 1) + 10 * Screen.TwipsPerPixelY
    If picContainer.ScaleHeight > picLegend.ScaleHeight Then
        vsbContainer.Max = (picContainer.ScaleHeight / Screen.TwipsPerPixelY) - (picLegend.ScaleHeight / Screen.TwipsPerPixelY)
        If Not vsbContainer.Visible Then vsbContainer.Visible = True
    Else
        vsbContainer.Visible = False
    End If

End Sub

Private Function TooltipNeeded() As String

    TooltipNeeded = Chr$(0) & Chr$(255) & Chr$(9)

End Function


Private Sub ClearLegendItems()
    
    Dim X As Integer
    
    On Error Resume Next    'we are expecting an error for item 1
    
    If bLegendAdded Then
        bLegendAdded = False
        
        For X = 1 To picBox.Count
            Unload picBox(X)
            Unload picDescription(X)
            If Err.numer Then Err.Clear
            picBox(0).Visible = False
            picDescription(0).Visible = False
        Next X
        'vsbContainer.Value = 0
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    
    With PropBag
        uInfoItems = .ReadProperty("InfoItems")
        uSelectedColor = .ReadProperty("SelectedColor", vbCyan)
        uSelectable = .ReadProperty("uSelectable", False)
        uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
        uTopMargin = .ReadProperty("uTopMargin", 55)
        uBottomMargin = .ReadProperty("uBottomMargin", 55)
        uLeftMargin = .ReadProperty("uLeftMargin", 55)
        uRightMargin = .ReadProperty("uRightMargin", 55)
        uContentBorder = .ReadProperty("uContentBorder", True)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uAxisYOn = .ReadProperty("uAxisXOn", uAxisXOn)
        uAxisXOn = .ReadProperty("uAxisYOn", uAxisYOn)
        uIntersectMajorY = .ReadProperty("uIntersectMajorY", 10)
        uIntersectMinorY = .ReadProperty("uIntersectMinorY", 2)
        uIntersectMajorX = .ReadProperty("uIntersectMajorX", 10)
        uIntersectMinorX = .ReadProperty("uIntersectMinorX", 2)
        uMaxYValue = .ReadProperty("uMaxYValue", 100)
        uMaxXValue = .ReadProperty("uMaxXValue", 100)
        uQuadrantYValue = .ReadProperty("QuadrantY", 0)
        uQuadrantXValue = .ReadProperty("QuadrantX", 0)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        uXAxisLabel = .ReadProperty("uXAxisLabel", uXAxisLabel)
        uYAxisLabel = .ReadProperty("uYAxisLabel", uYAxisLabel)
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        UserControl.ForeColor = .ReadProperty("ForeColor", vbBlack)
        uMinYValue = .ReadProperty("MinY")
        uMinXValue = .ReadProperty("MinX")
        uMajorGridColor = .ReadProperty("MajorGridColor", vbBlack)
        uMinorGridColor = .ReadProperty("MinorGridColor", vbBlack)
        uLegendBackColor = .ReadProperty("LegendBackColor", vbWhite)
        uLegendForeColor = .ReadProperty("LegendForeColor", vbBlack)
        uInfoBackColor = .ReadProperty("InfoBackColor")
        uInfoForeColor = .ReadProperty("InfoForeColor")
        uInfoQuadrantBackColor = .ReadProperty("InfoQuadrantBackColor")
        uInfoQuadrantForeColor = .ReadProperty("InfoQuadrantForeColor")
        uXAxisLabelColor = .ReadProperty("XAxisLabelColor", vbBlack)
        uYAxisLabelColor = .ReadProperty("YAxisLabelColor", vbBlack)
        uXAxisItemsColor = .ReadProperty("XAxisItemsColor", vbBlack)
        uYAxisItemsColor = .ReadProperty("YAxisItemsColor", vbBlack)
        uChartTitleColor = .ReadProperty("ChartTitleColor", vbBlack)
        uChartSubTitleColor = .ReadProperty("ChartSubTitleColor", vbBlack)
        uMenuType = .ReadProperty("MenuType")
        uMenuItems = .ReadProperty("MenuItems")
        uCustomMenuItems = .ReadProperty("CustomMenuItems")
        uInfoItems = .ReadProperty("InfoItems")
        uSaveAsCaption = .ReadProperty("SaveAsCaption")
        uAutoRedraw = .ReadProperty("AutoRedraw", True)
        uMarkerSymbol = .ReadProperty("MarkerSymbol", xcMarkerSymBox)
        Set uPicture = .ReadProperty("Picture", Nothing)
        uPictureTile = .ReadProperty("PictureTile", False)
        uMinorGridOn = .ReadProperty("MinorGridOn", True)
        uMajorGridOn = .ReadProperty("MajorGridOn", True)
        uMarkerWidth = .ReadProperty("MarkerWidth", 1)
        uMarkerColor = .ReadProperty("MarkerColor", vbRed)
        uDataFormat = .ReadProperty("DataFormat")
        uPrinterFit = .ReadProperty("PrinterFit")
        uPrinterOrientation = .ReadProperty("PrinterOrientation")
        uLegendCaption = .ReadProperty("LegendCaption")
        uMarkerLabelAngle = .ReadProperty("MarkerLabelAngle")
        uMarkerLabelDirection = .ReadProperty("MarkerLabelDirection")
        uChartAsQuadrant = .ReadProperty("ChartAsQuadrant")
        uQuadrantDividerColor = .ReadProperty("QuadrantDividerColor")
        uQuadrantColorsOverridePicture = .ReadProperty("QuadrantColorsOverridePicture")
        uMarkerLabelColor = .ReadProperty("MarkerLabel")
        uQuadrantColors = .ReadProperty("QuadrantColors")
        uRightMarginOrg = uRightMargin
        uHotTracking = .ReadProperty("uHotTracking", False)
        uLegendPrintMode = .ReadProperty("LegendPrintMode", legPrintGraph)
        uInnerColor = .ReadProperty("InnerColor", vbWhite)
        uOldSelection = -1
    End With

End Sub

Private Sub UserControl_Resize()
    
    If bDisplayLegend Then
        picLegend.Left = UserControl.ScaleWidth - picLegend.Width
    Else
        picLegend.Left = UserControl.ScaleWidth
    End If
    picLegend.Height = UserControl.ScaleHeight
    vsbContainer.Height = picLegend.ScaleHeight
    With picSplitter
        .Left = picLegend.Left - picSplitter.Width
        .Height = picLegend.ScaleHeight
    End With
    FixLegendCaption
    picSplitter.Cls
    picSplitter.Print Chr(171)

    bResize = True
    DrawChart
    bResize = False

End Sub

Private Sub UserControl_Show()
    DrawChart
    FixMenu
    FixCustomMenu
End Sub

Private Sub UserControl_Terminate()
    Set cItems = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SelectedColor", uSelectedColor
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uSelectedColumn", uSelectedColumn
        .WriteProperty "uTopMargin", uTopMargin
        .WriteProperty "uBottomMargin", uBottomMargin
        .WriteProperty "uLeftMargin", uLeftMargin
        .WriteProperty "uRightMargin", uRightMargin
        .WriteProperty "uContentBorder", uContentBorder
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uAxisXOn", uAxisXOn
        .WriteProperty "uAxisYOn", uAxisYOn
        .WriteProperty "uIntersectMajorY", uIntersectMajorY
        .WriteProperty "uIntersectMinorY", uIntersectMinorY
        .WriteProperty "uIntersectMajorX", uIntersectMajorX
        .WriteProperty "uIntersectMinorX", uIntersectMinorX
        .WriteProperty "uMaxYValue", uMaxYValue
        .WriteProperty "uMaxXValue", uMaxXValue
        .WriteProperty "QuadrantY", uQuadrantYValue
        .WriteProperty "QuadrantX", uQuadrantXValue
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "uXAxisLabel", uXAxisLabel
        .WriteProperty "uYAxislabel", uYAxisLabel
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
        .WriteProperty "MinY", uMinYValue
        .WriteProperty "MinX", uMinXValue
        .WriteProperty "MajorGridColor", uMajorGridColor
        .WriteProperty "MinorGridColor", uMinorGridColor
        .WriteProperty "LegendBackColor", uLegendBackColor
        .WriteProperty "LegendForeColor", uLegendForeColor
        .WriteProperty "InfoBackColor", uInfoBackColor
        .WriteProperty "InfoForeColor", uInfoForeColor
        .WriteProperty "InfoQuadrantBackColor", uInfoQuadrantBackColor
        .WriteProperty "InfoQuadrantForeColor", uInfoQuadrantForeColor
        .WriteProperty "XAxisLabelColor", uXAxisLabelColor
        .WriteProperty "YAxisLabelColor", uYAxisLabelColor
        .WriteProperty "XAxisItemsColor", uXAxisItemsColor
        .WriteProperty "YAxisItemsColor", uYAxisItemsColor
        .WriteProperty "ChartTitleColor", uChartTitleColor
        .WriteProperty "ChartSubTitleColor", uChartSubTitleColor
        .WriteProperty "MenuType", uMenuType
        .WriteProperty "MenuItems", uMenuItems
        .WriteProperty "CustomMenuItems", uCustomMenuItems
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SaveAsCaption", uSaveAsCaption
        .WriteProperty "AutoRedraw", uAutoRedraw
        .WriteProperty "MarkerSymbol", uMarkerSymbol
        .WriteProperty "Picture", uPicture, Nothing
        .WriteProperty "PictureTile", uPictureTile
        .WriteProperty "MinorGridOn", uMinorGridOn
        .WriteProperty "MajorGridOn", uMajorGridOn
        .WriteProperty "MarkerWidth", uMarkerWidth
        .WriteProperty "MarkerColor", uMarkerColor
        .WriteProperty "DataFormat", uDataFormat
        .WriteProperty "PrinterFit", uPrinterFit
        .WriteProperty "PrinterOrientation", uPrinterOrientation
        .WriteProperty "LegendCaption", uLegendCaption
        .WriteProperty "MarkerLabelAngle", uMarkerLabelAngle
        .WriteProperty "MarkerLabelDirection", uMarkerLabelDirection
        .WriteProperty "ChartAsQuadrant", uChartAsQuadrant
        .WriteProperty "QuadrantDividerColor", uQuadrantDividerColor
        .WriteProperty "QuadrantColorsOverridePicture", uQuadrantColorsOverridePicture
        .WriteProperty "MarkerLabel", uMarkerLabelColor
        .WriteProperty "QuadrantColors", uQuadrantColors
        .WriteProperty "HotTracking", uHotTracking
        .WriteProperty "LegendPrintMode", uLegendPrintMode
        .WriteProperty "InnerColor", uInnerColor
    End With

End Sub

Private Sub vsbContainer_Change()
    
    With picContainer
        .Visible = False
        .Top = -vsbContainer.Value * Screen.TwipsPerPixelY
        .Visible = True
    End With

End Sub

Private Sub vsbContainer_Scroll()
    
    With picContainer
        .Visible = False
        .Top = -vsbContainer.Value * Screen.TwipsPerPixelY
        .Visible = True
    End With

End Sub

Public Property Get MarkerColor() As OLE_COLOR
Attribute MarkerColor.VB_Description = "Returns/sets the color used to display the marker."
    MarkerColor = uMarkerColor
End Property

Public Property Let MarkerColor(lngVal As OLE_COLOR)
    If uMarkerColor <> lngVal Then
        uMarkerColor = lngVal
        DrawChart
        PropertyChanged "MarkerColor"
    End If
End Property

