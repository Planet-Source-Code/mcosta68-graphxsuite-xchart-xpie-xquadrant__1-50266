VERSION 5.00
Begin VB.UserControl XChart 
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
   ToolboxBitmap   =   "XChart.ctx":0000
   Begin VB.PictureBox picToPrinterLegend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1650
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   14
      Top             =   2220
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
      Left            =   1620
      ScaleHeight     =   555
      ScaleWidth      =   1005
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1005
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
      TabIndex        =   10
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox picCommands 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   60
      ScaleHeight     =   330
      ScaleWidth      =   1605
      TabIndex        =   4
      Top             =   60
      Width           =   1605
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "XChart.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   3
         Left            =   975
         Picture         =   "XChart.ctx":089C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   1
         Left            =   330
         Picture         =   "XChart.ctx":0E26
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   4
         Left            =   1290
         Picture         =   "XChart.ctx":13B0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   2
         Left            =   660
         Picture         =   "XChart.ctx":193A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   315
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   4
         Left            =   1470
         Picture         =   "XChart.ctx":1CC4
         Top             =   585
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "XChart.ctx":1E0E
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   2
         Left            =   930
         Picture         =   "XChart.ctx":1F58
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   1
         Left            =   660
         Picture         =   "XChart.ctx":20A2
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "XChart.ctx":21EC
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
            Left            =   330
            ScaleHeight     =   195
            ScaleWidth      =   765
            TabIndex        =   13
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
            Left            =   90
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   12
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
Attribute VB_Name = "XChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Private uColumns()        As Double       'array of column height values
                                          'used to determine hittest feature.

'--------------------------------------------------------------------------------
'added by M. Costa on 21/06/2002
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
Private dblMeanValue      As Double       'the mean value
Private uMeanOn           As Boolean      'marker indicating if the mean value must be displayed
Private uMeanColor        As Long         'the mean line color
Private Const MEAN_CAPTION = "Mean"
Private uMeanCaption      As String       'the mean caption used in the legend
Private uPicture          As StdPicture   'the background picture
Private uPictureTile      As Boolean      'marker indicating if the background picture must be tiled
                                          '(TRUE) or stretched (FALSE)
Private uBarPicture       As StdPicture   'the background picture
Private uBarPictureTile   As Boolean      'marker indicating if the bar picture must be tiled
Private uBarShadow        As Boolean      'marker indicating if the bar must have the shadow
                                          '(shadow takes effect only if line width is 1!)
Private uBarShadowColor   As Long         'the bar shadow color
Private uAutoRedraw       As Boolean      'marker indicating if the chart is auto-redrawn
                                          'upon every property change
Private uRangeY           As Integer      'the absolute range between Y-axis min. ad max. values
Private uDataType         As Integer      'indicates the data distribution in the Y axis
Private Const DT_BOTH = 0                 ' 0 = range(-Y0, +Y1)
Private Const DT_NEG = 1                  ' 1 = range(-Y0, -Y1)
Private Const DT_POS = 2                  ' 2 = range(+Y0, +Y1)

Private uMinYValue        As Double       'minimum y value
Private uLineColor        As Long         'the color of the line
Private uLineStyle        As Integer      'the line style
Private uBarSymbolColor   As Long         'the color of the symbol
Private uBarColor         As Long         'the backcolor of the bars
Private uBarFillStyle     As Integer      'the bars fill style
Private uSelectedBarColor As Long         'the selected bar backcolor
Private uMinorGridColor   As Long         'the minor intersect grid color
Private uMajorGridColor   As Long         'the major intersect grid color
Private uMinorGridOn      As Boolean      'marker indicating display of minor grid
Private uMajorGridOn      As Boolean      'marker indicating display of major grid
Private uLegendBackColor  As Long         'the legend background color
Private uLegendForeColor  As Long         'the legend foreground color
Private uInfoBackColor    As Long         'the information box background color
Private uInfoForeColor    As Long         'the information box foreground color
Private uXAxisLabelColor  As Long         'the X axis label color
Private uYAxisLabelColor  As Long         'the Y axis label color
Private uXAxisItemsColor  As Long         'the X axis items color
Private uYAxisItemsColor  As Long         'the Y axis items color
Private uChartTitleColor  As Long         'the chart title color
Private uChartSubTitleColor As Long       'the chart subtitle color
Private uSaveAsCaption    As String       'the SaveAs dialog box caption
Private uInfoItems        As String       'the information items (to be displayed in the info box)
Private Const INFO_ITEMS = "Value|Description|Mean"

Public Enum ChartMenuConstants             'the enumerated for menu type
    xcPopUpMenu = 0
    xcButtonMenu
End Enum

Private uMenuType         As ChartMenuConstants 'the menu type.
Private uMenuItems        As String       'the menu's items.
Private Const MENU_ITEMS = "&Save as...|&Print|&Copy|Selection &information|&Legend|&Hide"

Private uCustomMenuItems  As String       'the custom menu's items.
Private Const CUSTOM_MENU_ITEMS = Empty

Private uLegendCaption    As String       'the legend's tooltip string
Private Const LEGEND_CAPTION = "Display legend"

Private Const XC_BAR = 1
Private Const XC_SYMBOL = 2
Private Const XC_LINE = 4
Private Const XC_OVAL = 8
Private Const XC_TRIANGLE = 16
Private Const XC_RHOMBUS = 32
Private Const XC_TRAPEZIUM = 64
Public Enum ChartTypeConstants            'the enumerated for chart type
    xcBar = XC_BAR
    xcSymbol = XC_SYMBOL
    xcLine = XC_LINE
    xcBarLine = XC_BAR + XC_LINE
    xcSymbolLine = XC_SYMBOL + XC_LINE
    xcOval = XC_OVAL
    xcOvalLine = XC_OVAL + XC_LINE
    xcTriangle = XC_TRIANGLE
    xcTriangleLine = XC_TRIANGLE + XC_LINE
    xcRhombus = XC_RHOMBUS
    xcRhombusLine = XC_RHOMBUS + XC_LINE
    xcTrapezium = XC_TRAPEZIUM
    xcTrapeziumLine = XC_TRAPEZIUM + XC_LINE
End Enum

Private uChartType        As ChartTypeConstants 'the chart type.
Private uBarSymbol        As String * 1   'the symbol to be displayed when uChartType=xcSymbol
Private uBarWidthPercentage As Integer    'the column width (in percentage) just for bar type
Private uLineWidth        As Integer      'the line width (used when uChartType=xcLine and for bar border in case of uChartType=xcBar)

Private Const IDX_SAVE = 0                'the command buttons' indexs
Private Const IDX_PRINT = 1
Private Const IDX_COPY = 2
Private Const IDX_INFO = 3
Private Const IDX_LEGEND = 4
'--------------------------------------------------------------------------------

Private uColWidth         As Single       'the calculated width of each column
Private uRowHeight        As Single       'the calculated height of each column
Private uTopMargin        As Single       '--------------------------------------
Private uBottomMargin     As Single       'margins used around the chart content
Private uLeftMargin       As Single       '
Private uRightMargin      As Single
Private uRightMarginOrg   As Single       '--------------------------------------
Private uContentBorder    As Boolean      'border around the chart content?
Private uSelectable       As Boolean      'marker indicating whether user can select a column
Private uHotTracking      As Boolean      'marker indicating use of hot tracking
Private uSelectedColumn   As Integer      'marker indicating the selected column
Private uOldSelection     As Long
Private uDisplayDescript  As Boolean      'display description when selectable
Private uChartTitle       As String       'chart title
Private uChartSubTitle    As String       'chart sub title
Private uAxisXOn          As Boolean      'marker indicating display of x axis
Private uAxisYOn          As Boolean      'marker indicating display of y axis
Private uColorBars        As Boolean      'marker indicating use of different coloured bars
Private uIntersectMajor   As Single       'major intersect value
Private uIntersectMinor   As Single       'minor intersect value
Private uMaxYValue        As Double       'maximum y value
Private uXAxisLabel       As String       'label to be displayed below the X-Axis
Private uYAxisLabel       As String       'label to be displayed left of the Y-Axis
Private cItems            As Collection   'collection of chart items

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
    XAxisDescription As String
    Value As Double
End Type

Public Event ItemClick(cItem As ChartItem)
Public Event MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'-----------------------------------------------
' for Ballon ToolTip
'-----------------------------------------------
Private ttpBalloon As New Tooltip

Public Property Let LegendPrintMode(val As LegendPrintConstants)
    uLegendPrintMode = val
    PropertyChanged "LegendPrintMode"
End Property
Public Property Get LegendPrintMode() As LegendPrintConstants
    LegendPrintMode = uLegendPrintMode
End Property

Public Function AddItem(cItem As ChartItem) As Boolean
    
    Dim oChartItem As ChartItem
    
    If uMeanOn = True Then
        If cItems.Count > 0 Then
            cItems.Remove (cItems.Count)
        End If
    End If

    cItems.Add cItem
    
    If uMeanOn = True Then
        CalcMean
        If uMeanCaption = Empty Then uMeanCaption = MEAN_CAPTION
        With oChartItem
            .Value = dblMeanValue
            .ItemID = uMeanCaption
            .XAxisDescription = uMeanCaption
            .SelectedDescription = uMeanCaption
            .LegendDescription = uMeanCaption
        End With
        cItems.Add oChartItem
    End If
    
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

Public Property Get BarShadow() As Boolean
Attribute BarShadow.VB_Description = "Determines if the bar must have the shadow."
    BarShadow = uBarShadow
End Property

Public Property Get BarShadowColor() As OLE_COLOR
Attribute BarShadowColor.VB_Description = "Returns/sets the color used to display the bar shadow."
    BarShadowColor = uBarShadowColor
End Property

Public Property Let BarShadow(blnVal As Boolean)
    If blnVal <> uBarShadow Then
        uBarShadow = blnVal
        DrawChart
        PropertyChanged "BarShadow"
    End If
End Property

Public Property Let BarShadowColor(lngVal As OLE_COLOR)
    If lngVal <> uBarShadowColor Then
        uBarShadowColor = lngVal
        DrawChart
        PropertyChanged "BarShadowColor"
    End If
End Property

Private Sub CalcMean()
    
    On Error Resume Next
    
    Dim intIdx As Integer
    
    dblMeanValue = 0
    For intIdx = 1 To cItems.Count
        dblMeanValue = dblMeanValue + cItems.Item(intIdx).Value
    Next
    dblMeanValue = dblMeanValue / cItems.Count
    
End Sub

Public Property Get DataFormat() As String
Attribute DataFormat.VB_Description = "Determines the format which the Y-values are displayed with."
    DataFormat = uDataFormat
End Property

Public Property Get PrinterOrientation() As PrinterObjectConstants
    PrinterOrientation = uPrinterOrientation
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
                On Error GoTo DisplayInfo_error
        
                If uInfoItems = Empty Then uInfoItems = INFO_ITEMS
                varItems = Split(uInfoItems, "|")
                sDescription = CStr(varItems(0)) & ": " & Format(.Value, uDataFormat)
                If Len(.SelectedDescription) > 0 Then
                    sDescription = CStr(varItems(1)) & ": " & .SelectedDescription & vbCrLf & sDescription
                End If
                If (uMeanOn = True) And (intIdx < cItems.Count - 1) Then
                    sDescription = sDescription & vbCrLf & CStr(varItems(2)) & ": " & Format(dblMeanValue, uDataFormat)
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

DisplayInfo_error:
    uInfoItems = INFO_ITEMS
    Resume Next

End Sub

Private Sub DrawOval(sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, sngBase As Single, sngHeight As Single, lngBorderColor As Long)
    
    On Error Resume Next
    
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    Dim sngH As Single
    Dim sngW As Single
    Dim lngFillColor As Long

    x1 = sngX1
    y1 = sngY1
    x2 = sngX2
    y2 = sngY2
    sngW = sngBase
    sngH = sngHeight
    x1 = x1 + (sngW / 2)
    y1 = y1 + (sngH / 2)
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillStyle = vbFSSolid
        UserControl.FillColor = uBarShadowColor
        UserControl.Circle (x1, y1), sngH / 2, uBarShadowColor, , , _
                            IIf((sngH > sngW), (sngH / sngW), (sngW / sngH))
        UserControl.FillColor = lngFillColor
        UserControl.FillStyle = uBarFillStyle
        x1 = x1 - 2 * Screen.TwipsPerPixelX
        sngW = sngW - 2 * Screen.TwipsPerPixelX
        sngH = sngH - 2 * Screen.TwipsPerPixelX
    End If
    'the aspect ratio depend on whether the base is greater than the height
    UserControl.Circle (x1, y1), sngH / 2, lngBorderColor, , , _
                        IIf((sngH > sngW), (sngH / sngW), (sngW / sngH))
    
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

Private Sub DrawRectangle(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, lngBorderColor As Long, blnOverridePicture As Boolean)
        
    On Error Resume Next
    
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    
    x1 = sngX1
    y1 = sngY1
    x2 = sngX2
    y2 = sngY2
    If uBarShadow = True Then
        x2 = x2 - 2 * Screen.TwipsPerPixelX
    End If
    If (blnOverridePicture = True) Or (uBarPicture Is Nothing) Then
        UserControl.Line (x1 + 1 * Screen.TwipsPerPixelX, y1)-(x2 - 1 * Screen.TwipsPerPixelX, y2), , BF
    Else
        Call DrawPicture(x1, x2, y1, y2, uBarPictureTile, uBarPicture)
        'if the fill  style is solid, the image is overriden when drawing the outer box
        If UserControl.FillStyle = vbFSSolid Then _
            UserControl.FillStyle = vbFSTransparent
    End If
    UserControl.Line (x1, y1)-(x2 - 1 * Screen.TwipsPerPixelX, y2), lngBorderColor, B
    If uBarShadow = True Then
        If dblData >= 0 Then
            y1 = y1 + 2 * Screen.TwipsPerPixelX
            UserControl.Line (x2, y1)-(x2 + 2 * Screen.TwipsPerPixelX, y2), uBarShadowColor, BF
        Else
            y2 = y2 - 2 * Screen.TwipsPerPixelX
            UserControl.Line (x2, y1)-(x2 + 2 * Screen.TwipsPerPixelX, y2), uBarShadowColor, BF
        End If
    End If
    UserControl.FillStyle = uBarFillStyle

End Sub

Private Sub DrawRhombus(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim sngXTemp As Single
    Dim sngYTemp As Single
    Dim uaPts(3) As PointAPI
    Dim lngFillColor As Long
    Dim intScaleMode As Integer
    
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
    sngXTemp = sngX1 + ((sngX2 - sngX1) / 2)
    sngYTemp = sngY1 + ((sngY2 - sngY1) / 2)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(0).Y = sngYTemp / Screen.TwipsPerPixelY
    uaPts(1).X = sngXTemp / Screen.TwipsPerPixelX
    uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
    uaPts(2).X = sngX2 / Screen.TwipsPerPixelX
    uaPts(2).Y = sngYTemp / Screen.TwipsPerPixelY
    uaPts(3).X = sngXTemp / Screen.TwipsPerPixelX
    uaPts(3).Y = sngY2 / Screen.TwipsPerPixelY
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 4)
        UserControl.FillColor = lngFillColor
        'resize the Rhombus
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 3
        uaPts(3).X = uaPts(3).X - 2
        If dblData > 0 Then
            uaPts(1).Y = uaPts(1).Y + 2
            uaPts(3).Y = uaPts(3).Y - 2
        Else
            uaPts(1).Y = uaPts(1).Y - 2
            uaPts(3).Y = uaPts(3).Y + 2
        End If
    End If
    
    'draw the filled Rhombus
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub

Private Sub DrawTrapezium(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim sngXTemp As Single
    Dim sngYTemp As Single
    Dim lngFillColor As Long
    Dim uaPts(3) As PointAPI
    Dim intScaleMode As Integer
    
    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 4 points of the trapezio
    sngXTemp = (sngX2 - sngX1) / 4      'consider the 25% as X-offset
    'set the points anti-clockwise
    '     (1)-----(2)
    '    /           \
    '   /             \
    ' (0)-------------(3)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(1).X = (sngX1 + sngXTemp) / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX2 - sngXTemp) / Screen.TwipsPerPixelX
    uaPts(3).X = sngX2 / Screen.TwipsPerPixelX
    If dblData > 0 Then
        uaPts(0).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(3).Y = sngY2 / Screen.TwipsPerPixelY
    Else
        uaPts(0).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(3).Y = sngY1 / Screen.TwipsPerPixelY
    End If
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 4)
        UserControl.FillColor = lngFillColor
        'resize the trapezio
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 2
        uaPts(3).X = uaPts(3).X - 2
        If dblData > 0 Then
            uaPts(1).Y = uaPts(1).Y + 2
            uaPts(2).Y = uaPts(2).Y + 2
        Else
            uaPts(1).Y = uaPts(1).Y - 2
            uaPts(2).Y = uaPts(2).Y - 2
        End If
    End If
    
    'draw the filled trapezio
    lRet = Polygon(UserControl.hDC, uaPts(0), 4)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub


Private Sub DrawTriangle(dblData As Double, sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single)

    On Error Resume Next
    
    Dim lRet As Long
    Dim uaPts(2) As PointAPI
    Dim lngFillColor As Long
    Dim intScaleMode As Integer

    'the polygon function works only with pixels!
    intScaleMode = UserControl.ScaleMode
    UserControl.ScaleMode = vbPixels
    
    'setup the 3 points of the triangle anti-clockwise
    '     (1)
    '    /   \
    '   /     \
    ' (0)-----(2)
    uaPts(0).X = sngX1 / Screen.TwipsPerPixelX
    uaPts(1).X = sngX2 / Screen.TwipsPerPixelX
    uaPts(2).X = (sngX1 + ((sngX2 - sngX1) / 2)) / Screen.TwipsPerPixelX
    If dblData > 0 Then
        uaPts(0).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY2 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY1 / Screen.TwipsPerPixelY
    Else
        uaPts(0).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(1).Y = sngY1 / Screen.TwipsPerPixelY
        uaPts(2).Y = sngY2 / Screen.TwipsPerPixelY
    End If
    
    If uBarShadow = True Then
        lngFillColor = UserControl.FillColor
        UserControl.FillColor = uBarShadowColor
        lRet = Polygon(UserControl.hDC, uaPts(0), 3)
        UserControl.FillColor = lngFillColor
        'resize the triangle
        uaPts(1).X = uaPts(1).X - 2
        uaPts(2).X = uaPts(2).X - 2
        If dblData > 0 Then
            uaPts(2).Y = uaPts(2).Y + 2
        Else
            uaPts(2).Y = uaPts(2).Y - 2
        End If
    End If
    
    'draw the filled triangle
    lRet = Polygon(UserControl.hDC, uaPts(0), 3)
    
    'reset the scalemode
    UserControl.ScaleMode = intScaleMode
    
    'free the memory
    Erase uaPts

End Sub

Private Sub FixLegendCaption()
    uLegendCaption = IIf(uLegendCaption = Empty, LEGEND_CAPTION, uLegendCaption)
End Sub

Public Property Let LegendCaption(stgVal As String)
    uLegendCaption = stgVal
    FixLegendCaption
End Property

Public Property Get LineStyle() As DrawStyleConstants
Attribute LineStyle.VB_Description = "Determines the style of the line displayed in the chart."
    LineStyle = uLineStyle
End Property

Public Property Let LineStyle(intVal As DrawStyleConstants)
    If uLineStyle <> intVal Then
        uLineStyle = intVal
        DrawChart
        PropertyChanged "LineStyle"
    End If
End Property

Public Property Let LineWidth(intVal As Integer)
    If intVal <> uLineWidth Then
        If intVal > 0 And intVal <= 9 Then
            uLineWidth = intVal
            DrawChart
            PropertyChanged "LineWidth"
        End If
    End If
End Property

Public Property Get LineWidth() As Integer
Attribute LineWidth.VB_Description = "Returns/sets the width of the line displayed in the chart."
    LineWidth = uLineWidth
End Property
Public Property Get MeanOn() As Boolean
Attribute MeanOn.VB_Description = "Determines if the mean value must be calculated and displayed."
    MeanOn = uMeanOn
End Property

Public Property Get MeanCaption() As String
Attribute MeanCaption.VB_Description = "Returns/sets the caption used in the legend for the mean value."
    MeanCaption = uMeanCaption
End Property


Public Property Get MeanColor() As OLE_COLOR
Attribute MeanColor.VB_Description = "Returns/sets the color used to display the mean value bar."
    MeanColor = uMeanColor
End Property


Public Property Let MeanOn(blnVal As Boolean)
    If blnVal <> uMeanOn Then
        uMeanOn = blnVal
        DrawChart
        PropertyChanged "MeanOn"
    End If
End Property

Public Property Let MeanCaption(stgVal As String)
    If stgVal <> uMeanCaption Then
        uMeanCaption = stgVal
        DrawChart
        PropertyChanged "MeanCaption"
    End If
End Property


Public Property Let MeanColor(lngVal As OLE_COLOR)
    If lngVal <> uMeanColor Then
        uMeanColor = lngVal
        DrawChart
        PropertyChanged "MeanColor"
    End If
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


Public Property Set BarPicture(ByVal picVal As StdPicture)
    Set uBarPicture = picVal
    DrawChart
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed as background of the chart."
    Set Picture = uPicture
End Property

Public Property Get BarPicture() As Picture
Attribute BarPicture.VB_Description = "Returns/sets a graphic to be displayed in the bars."
    Set BarPicture = uBarPicture
End Property

Public Property Get BarWidthPercentage() As Integer
Attribute BarWidthPercentage.VB_Description = "Returns/sets a value that determines how the bar must be wide."
    BarWidthPercentage = uBarWidthPercentage
End Property

Public Property Get BarSymbol() As String
Attribute BarSymbol.VB_Description = "Returns/sets the character to be displayed in place of the bar."
    BarSymbol = uBarSymbol
End Property

Public Property Let BarSymbol(stgVal As String)
    If stgVal <> uBarSymbol Then
        uBarSymbol = stgVal
        DrawChart
        PropertyChanged "BarSymbol"
    End If
End Property

Public Property Let ChartType(intVal As ChartTypeConstants)
    If intVal <> uChartType Then
        uChartType = intVal
        DrawChart
        PropertyChanged "ChartType"
    End If
End Property

Public Property Let BarWidthPercentage(intVal As Integer)
    If intVal > 0 And intVal <= 100 Then
        If intVal <> uBarWidthPercentage Then
            uBarWidthPercentage = intVal
            DrawChart
            PropertyChanged "BarWidthPercentage"
        End If
    End If
End Property

Public Property Get ChartType() As ChartTypeConstants
Attribute ChartType.VB_Description = "Determines the type of chart to be displayed."
    ChartType = uChartType
End Property

Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

Private Sub FixData()

    If uMinYValue < 0 And uMaxYValue < 0 Then
        uDataType = DT_NEG
        uRangeY = (Abs(uMinYValue) - Abs(uMaxYValue))
    ElseIf uMinYValue >= 0 And uMaxYValue >= 0 Then
        uDataType = DT_POS
        uRangeY = (Abs(uMaxYValue) - Abs(uMinYValue))
    Else
        uDataType = DT_BOTH
        uRangeY = (Abs(uMaxYValue) + Abs(uMinYValue))
    End If

    If uRangeY = 0 Then uRangeY = 1
    If uIntersectMajor = 0 Then uIntersectMajor = uRangeY / 10
    If uIntersectMinor = 0 Then uIntersectMinor = uIntersectMajor / 5
    
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
        mnuMainViewLegend.Caption = CStr(varItems(4))
    Else
        mnuMainViewLegend.Caption = "&Legend"
    End If
    
    If varItems(5) <> Empty Then
        mnuLegendHide.Caption = CStr(varItems(5))
    Else
        mnuLegendHide.Caption = "&Hide"
    End If

    If uMenuType = xcButtonMenu Then
        picCommands.Visible = True
        picCommands.BackColor = UserControl.BackColor
        picCommands.Move 60, 60
        lblInfo.Move picCommands.Left + picCommands.ScaleWidth + 60, 60
    Else
        picCommands.Visible = False
        lblInfo.Move 60, 60
    End If
    Exit Sub
    
FixMenu_error:
    uMenuItems = MENU_ITEMS
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


Private Function InColumn(X As Single, Y As Single) As Integer

    Dim sngY As Single
    Dim sngY1 As Single
    Dim sngY2 As Single
    Dim intCol As Integer
    Dim intSelectedCol As Integer

    intSelectedCol = -1
    If (uChartType And XC_BAR) = XC_BAR _
    Or (uChartType And XC_OVAL) = XC_OVAL _
    Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
    Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
    Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (Y >= uTopMargin) _
        And (uSelectable = True) Then
            intCol = (X - uLeftMargin) \ (uColWidth)
            sngY1 = uColumns(intCol, 0)
            sngY2 = uColumns(intCol, 1)
            If sngY1 > sngY2 Then
                sngY = sngY1
                sngY1 = sngY2
                sngY2 = sngY
            End If
            If (Y >= sngY1 And Y <= sngY2) Then
                intSelectedCol = intCol
            End If
        End If
    End If
    InColumn = intSelectedCol

End Function

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

Public Property Let BarPictureTile(blnVal As Boolean)
    If blnVal <> uBarPictureTile Then
        uBarPictureTile = blnVal
        DrawChart
        PropertyChanged "BarPictureTile"
    End If
End Property

Public Property Get PictureTile() As Boolean
Attribute PictureTile.VB_Description = "Determines if the picture used as the background of the chart must be tiled."
    PictureTile = uPictureTile
End Property

Public Property Get BarPictureTile() As Boolean
Attribute BarPictureTile.VB_Description = "Determines if the picture used to fill the bars of the chart must be tiled."
    BarPictureTile = uBarPictureTile
End Property

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
                stg = .LegendDescription & " (" & Format(.Value, uDataFormat) & ")"
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

Public Property Let Selectable(blnVal As Boolean)
    If blnVal <> uSelectable Then
        uSelectable = blnVal
        DrawChart
        PropertyChanged "Selectable"
    End If
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



Public Property Get Selectable() As Boolean
Attribute Selectable.VB_Description = "Returns/sets a value that determines if a bar can be selected with the mouse."
    Selectable = uSelectable
End Property

Public Property Let HotTracking(blnVal As Boolean)
    If blnVal <> uHotTracking Then
        uHotTracking = blnVal
        DrawChart
        PropertyChanged "HotTracking"
    End If
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value that determines if a bar becomes selected while moving over the mouse."
    HotTracking = uHotTracking
End Property

Public Property Get LegendCaption() As String
    LegendCaption = uLegendCaption
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
            If (uMeanOn = True) And (uSelectedColumn = cItems.Count - 1) Then
                'do nothing in case of mean bar selected
            Else
                oItem = cItems(lngColumn + 1)
                RaiseEvent ItemClick(oItem)
            End If
        End If
    End If

End Property

Public Property Get SelectedColumn() As Long
Attribute SelectedColumn.VB_Description = "Returns/sets the number of selected bar."
    SelectedColumn = uSelectedColumn
End Property

Public Property Let ChartTitle(stgVal As String)
    If stgVal <> uChartTitle Then
        uChartTitle = stgVal
        DrawChart
        PropertyChanged "ChartTitle"
    End If
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

Public Property Get InfoItems() As String
Attribute InfoItems.VB_Description = "Determines the string values displayed when selection information is enabled (separated by |)."
    InfoItems = uInfoItems
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

Public Property Let IntersectMajor(sngVal As Single)
    If sngVal <> uIntersectMajor Then
        uIntersectMajor = sngVal
        DrawChart
        PropertyChanged "IntersectMajor"
    End If
End Property

Public Property Get IntersectMajor() As Single
Attribute IntersectMajor.VB_Description = "Determines the value which the major intersection line is displayed for."
    IntersectMajor = uIntersectMajor
End Property

Public Property Let IntersectMinor(sngVal As Single)
    If sngVal <> uIntersectMinor Then
        uIntersectMinor = sngVal
        DrawChart
        PropertyChanged "IntersectMinor"
    End If
End Property

Public Property Get IntersectMinor() As Single
Attribute IntersectMinor.VB_Description = "Determines the value which the minor intersection line is displayed for."
    IntersectMinor = uIntersectMinor
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
Public Property Let MinY(dblMin As Double)
Attribute MinY.VB_Description = "Returns/sets the minimum Y value."
    If dblMin < uMaxYValue Then
        uMinYValue = dblMin
        DrawChart
        PropertyChanged "MinY"
    End If
End Property

Public Property Get MinY() As Double
    MinY = uMinYValue
End Property


Public Property Get MaxY() As Double
    MaxY = uMaxYValue
End Property

Public Property Let SelectionInformation(blnVal As Boolean)
Attribute SelectionInformation.VB_Description = "Determines if the information box about the selected bar must be visible or hidden."
    If blnVal <> uDisplayDescript Then
        uDisplayDescript = blnVal
        DrawChart
        PropertyChanged "SelectionInformation"
    End If
End Property
Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
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



Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Returns/sets the color used to display the bar."
    BarColor = uBarColor
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
Public Property Let InfoForeColor(lngVal As OLE_COLOR)
Attribute InfoForeColor.VB_Description = "Returns/sets the selection information foreground color."
    If lngVal <> uInfoForeColor Then
        uInfoForeColor = lngVal
        DrawChart
        PropertyChanged "InfoForeColor"
    End If
End Property

Public Property Get InfoBackColor() As OLE_COLOR
    InfoBackColor = uInfoBackColor
End Property

Public Property Get InfoForeColor() As OLE_COLOR
    InfoForeColor = uInfoForeColor
End Property

Public Property Let LegendBackColor(lngVal As OLE_COLOR)
    If lngVal <> uLegendBackColor Then
        uLegendBackColor = lngVal
        DrawChart
        PropertyChanged "LegendBackColor"
    End If
End Property

Public Property Get SelectedBarColor() As OLE_COLOR
Attribute SelectedBarColor.VB_Description = "Returns/sets the color used to display the selected bar."
    SelectedBarColor = uSelectedBarColor
End Property


Public Property Let SelectedBarColor(lngVal As OLE_COLOR)
    If lngVal <> uSelectedBarColor Then
        uSelectedBarColor = lngVal
        PropertyChanged "SelectedBarColor"
    End If
End Property

Public Property Let BarColor(lngVal As OLE_COLOR)
    If lngVal <> uBarColor Then
        uBarColor = lngVal
        DrawChart
        PropertyChanged "BarColor"
    End If
End Property


Public Property Let ColorBars(blnVal As Boolean)
Attribute ColorBars.VB_Description = "Returns/sets a value that determines if the bar color is randomly generated."
    If blnVal <> uColorBars Then
        uColorBars = blnVal
        DrawChart
        PropertyChanged "ColorBars"
    End If
End Property
Public Property Get ColorBars() As Boolean
    ColorBars = uColorBars
End Property

Private Sub Swap(ByRef var1 As Variant, ByRef var2 As Variant)
    
    Dim varDummy As Variant
    
    varDummy = var1
    var1 = var2
    var2 = varDummy

End Sub

Private Function TooltipNeeded() As String

    TooltipNeeded = Chr$(0) & Chr$(255) & Chr$(9)

End Function

Private Sub cmdCmd_Click(Index As Integer)

    Select Case Index
        Case IDX_SAVE:      mnuMainSaveAs_Click
        Case IDX_PRINT:     mnuMainPrint_Click
        Case IDX_COPY:      mnuMainCopy_Click
        Case IDX_INFO:      mnuMainSelectionInfo_Click
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
Public Property Get PrinterFit() As PrinterFitConstants
    PrinterFit = uPrinterFit
End Property

Public Property Let PrinterFit(intVal As PrinterFitConstants)
    uPrinterFit = intVal
    PropertyChanged "PrinterFit"
End Property

Private Sub mnuProperties_Click()
    'frmProperties.Show vbModal
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
        .sFile = "XChart.bmp" & Space$(1024) & vbNullChar & vbNullChar
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
    uContentBorder = True
    uSelectable = False
    uHotTracking = False
    uSelectedColumn = -1
    uOldSelection = -1
    uChartTitle = UserControl.Name
    uChartSubTitle = ""
    uAxisYOn = True
    uAxisXOn = True
    uColorBars = False
    uIntersectMajor = 10
    uIntersectMinor = 2
    uMaxYValue = 100
    UserControl.BackColor = vbWindowBackground
    UserControl.ForeColor = vbWindowText
    '----------------------------------------------------
    'added by M. Costa on 21/06/2002
    uMinYValue = 0
    uBarColor = vbGreen
    uSelectedBarColor = vbYellow
    uMajorGridColor = vbWhite
    uMinorGridColor = vbBlack
    uLegendBackColor = UserControl.BackColor
    uLegendForeColor = UserControl.ForeColor
    uInfoBackColor = vbInfoBackground
    uInfoForeColor = vbInfoText
    uXAxisLabelColor = UserControl.ForeColor
    uYAxisLabelColor = UserControl.ForeColor
    uXAxisItemsColor = UserControl.ForeColor
    uYAxisItemsColor = UserControl.ForeColor
    uChartTitleColor = UserControl.ForeColor
    uChartSubTitleColor = UserControl.ForeColor
    uBarSymbolColor = uBarColor
    uLineColor = uBarColor
    uMenuType = xcPopUpMenu
    uChartType = xcBar
    uBarSymbol = "*"
    uBarWidthPercentage = 100
    uMenuItems = Empty
    uCustomMenuItems = Empty
    uInfoItems = Empty
    uSaveAsCaption = Empty
    uAutoRedraw = True
    Set uBarPicture = Nothing
    uBarPictureTile = False
    Set uPicture = Nothing
    uPictureTile = False
    uMinorGridOn = True
    uMajorGridOn = True
    uLineWidth = 1
    uBarFillStyle = vbFSSolid
    uBarFillStyle = vbCross
    uLineStyle = vbSolid
    uBarShadow = True
    uBarShadowColor = vbBlack
    uMeanOn = False
    uMeanCaption = Empty
    uDataFormat = Empty
    uPrinterFit = prtFitCentered
    uPrinterOrientation = vbPRORLandscape
    uLegendCaption = LEGEND_CAPTION
    uLegendPrintMode = legPrintGraph
    '----------------------------------------------------

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
                    If (uMeanOn = True) And (uSelectedColumn = cItems.Count - 1) Then
                        'do nothing in case of mean bar selected
                    Else
                        oItem = cItems(uSelectedColumn + 1)
                        RaiseEvent ItemClick(oItem)
                    End If
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
    uSelectedColumn = -1
    DrawChart
End Sub

Public Sub DrawChart()
    
    On Error Resume Next

    Dim x1              As Single
    Dim x2              As Single
    Dim y1              As Single
    Dim y2              As Single
    Dim xTemp           As Single
    Dim yTemp           As Single
    Dim xPrev           As Single
    Dim yPrev           As Single
    Dim sngRowHeight    As Single
    Dim CurrentColor    As Integer
    Dim iCols           As Integer
    Dim X               As Integer
    Dim oChartItem      As ChartItem
    Dim sngColWidth     As Single
    Dim xMiddle         As Single
    Dim varLabel        As Variant
    Dim intIdx          As Integer
    
    'do not redraw the chart if not required
    If uAutoRedraw = False Then Exit Sub

    'calculate the data distribution in the y-axis
    FixData
    
    picDescription(0).ForeColor = uLegendForeColor
    
    iCols = cItems.Count
    
    With lblInfo
        .ForeColor = uInfoForeColor
        .BackColor = uInfoBackColor
        .Visible = IIf((uDisplayDescript And uSelectedColumn > -1), True, False)
    End With
    mnuMainSelectionInfo.Checked = uDisplayDescript
    
    If Not bResize Then ClearLegendItems

    With UserControl
        uRowHeight = ((.ScaleHeight - (uTopMargin + uBottomMargin)) / uRangeY)
        If iCols Then
            uColWidth = ((.ScaleWidth - (uLeftMargin + uRightMargin)) / iCols)
        End If
        
        .Cls
        If uPicture Is Nothing Then
        Else
            'paint the background image
            Call DrawPicture(uLeftMargin, .ScaleWidth - uRightMargin, _
                             uTopMargin, .ScaleHeight - uBottomMargin, _
                             uPictureTile, uPicture)
        End If
    
        If iCols Then ReDim uColumns(iCols - 1, 1)
    
        On Error Resume Next
        
        'dump chart title
        If bDisplayLegend Then
            xMiddle = (picSplitter.Left / 2)
        Else
            xMiddle = (.ScaleWidth / 2)
        End If
        .ForeColor = uChartTitleColor
        .CurrentX = xMiddle - (.TextWidth(uChartTitle) / 2)
        .CurrentY = 0
        .FontBold = True
        UserControl.Print uChartTitle
        .FontBold = False
        
        'dump chart subtitle
        .ForeColor = uChartSubTitleColor
        .FontSize = .FontSize - 2
        .CurrentX = xMiddle - (.TextWidth(uChartSubTitle) / 2)
        UserControl.Print uChartSubTitle
        .FontSize = .FontSize + 2
        
        If uAxisYOn Then
            'draw Y axis
            .ForeColor = uYAxisItemsColor
            For X = uMinYValue To uMaxYValue
                x1 = uLeftMargin + (2 * Screen.TwipsPerPixelX)
                x2 = .ScaleWidth - uRightMargin
                y1 = (.ScaleHeight - uBottomMargin)
                If uDataType = DT_NEG Then
                    y1 = y1 + ((Abs(X) - Abs(uMinYValue)) * uRowHeight)
                Else
                    y1 = y1 - ((X - uMinYValue) * uRowHeight)
                End If
                If (X = uMinYValue) Or (X = uMaxYValue) Or ((X Mod uIntersectMajor) = 0) Then
                    If uMajorGridOn Then
                        UserControl.Line (x1, y1)-(x2, y1), uMajorGridColor
                    End If
                    .FontSize = .FontSize - 2
                    .CurrentX = uLeftMargin - .TextWidth(X) - (5 * Screen.TwipsPerPixelX)
                    .CurrentY = y1 - (.TextHeight("0") / 2)
                    UserControl.Print X
                    .FontSize = .FontSize + 2
                ElseIf ((uMaxYValue - X) Mod uIntersectMinor = 0) Then
                    If uMinorGridOn Then
                        UserControl.Line (x1, y1)-(x2, y1), uMinorGridColor
                    End If
                End If
            Next X
        End If
    
        On Error GoTo 0
        If uContentBorder Then
            UserControl.Line (uLeftMargin, uTopMargin)-(.ScaleWidth - uRightMargin, .ScaleHeight - uBottomMargin), uMajorGridColor, B
        End If
        
        'draw bars, lines, symbols,...
        For X = 0 To cItems.Count - 1
            oChartItem = cItems(X + 1)
            x1 = (X * uColWidth) + uLeftMargin + (2 * Screen.TwipsPerPixelX)    'increment by 2 pixs.
            x2 = x1 + uColWidth - (2 * Screen.TwipsPerPixelX)                   'decrement by 2 pixs.
            If uDataType = DT_POS Then
                sngRowHeight = uRowHeight * (oChartItem.Value - uMinYValue)
                y2 = .ScaleHeight - uBottomMargin
                y1 = y2 - sngRowHeight
            ElseIf uDataType = DT_NEG Then
                sngRowHeight = uRowHeight * (Abs(CDbl(oChartItem.Value)) - Abs(uMaxYValue))
                y1 = uTopMargin
                y2 = y1 + sngRowHeight
            Else
                sngRowHeight = (-CDbl(oChartItem.Value) * uRowHeight)
                y1 = .ScaleHeight - uBottomMargin
                y1 = y1 - uRowHeight * Abs(uMinYValue)
                y2 = y1 + sngRowHeight
            End If
            sngRowHeight = Abs(sngRowHeight)
            'be sure the y1 coordinate is always less than y2
            If y2 < y1 Then Call Swap(y1, y2)
    
            'save coordinates of bar (only Y since X is calculated)
            uColumns(X, 0) = y1
            uColumns(X, 1) = y2
    
            If ((uChartType And XC_BAR) = XC_BAR) _
            Or (uChartType And XC_OVAL) = XC_OVAL _
            Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
            Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
            Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                'draw the bars in the right shape
    
                'adjust x-coordinates depending on bar width percentage
                sngColWidth = uColWidth * uBarWidthPercentage / 100
                xTemp = x1 + ((uColWidth - sngColWidth) / 2)
                x2 = x2 - ((uColWidth - sngColWidth) / 2)
                'Selected bar outline
                .DrawWidth = uLineWidth
                .FillStyle = uBarFillStyle
                If X = uSelectedColumn And uSelectable Then
                    .FillColor = uSelectedBarColor
                    If (uChartType And XC_OVAL) = XC_OVAL Then
                        Call DrawOval(xTemp, x2, y1, y2, sngColWidth, sngRowHeight, uBarColor)
                    ElseIf (uChartType And XC_BAR) = XC_BAR Then
                        Call DrawRectangle(oChartItem.Value, xTemp, x2, y1, y2, uBarColor, (uMeanOn = True) And (X = cItems.Count - 1))
                    ElseIf (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                        Call DrawTriangle(oChartItem.Value, xTemp, x2, y1, y2)
                    ElseIf (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM Then
                        Call DrawTrapezium(oChartItem.Value, xTemp, x2, y1, y2)
                    ElseIf (uChartType And XC_RHOMBUS) = XC_RHOMBUS Then
                        Call DrawRhombus(oChartItem.Value, xTemp, x2, y1, y2)
                    End If
                    .DrawWidth = 1
                    .FillStyle = vbFSTransparent
    
                    'display information
                    Call DisplayInfo(X)
                Else
                    If (uMeanOn = True) And (X = cItems.Count - 1) Then
                        .FillColor = uMeanColor
                    Else
                        .FillColor = IIf(uColorBars, QBColor(CurrentColor), uBarColor)
                    End If
                    .FillStyle = uBarFillStyle
                    .DrawWidth = uLineWidth
                    If (uChartType And XC_OVAL) = XC_OVAL Then
                        Call DrawOval(xTemp, x2, y1, y2, sngColWidth, sngRowHeight, uSelectedBarColor)
                    ElseIf (uChartType And XC_BAR) = XC_BAR Then
                        Call DrawRectangle(oChartItem.Value, xTemp, x2, y1, y2, uSelectedBarColor, (uMeanOn = True) And (X = cItems.Count - 1))
                    ElseIf (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                        Call DrawTriangle(oChartItem.Value, xTemp, x2, y1, y2)
                    ElseIf (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM Then
                        Call DrawTrapezium(oChartItem.Value, xTemp, x2, y1, y2)
                    ElseIf (uChartType And XC_RHOMBUS) = XC_RHOMBUS Then
                        Call DrawRhombus(oChartItem.Value, xTemp, x2, y1, y2)
                    End If
                    .DrawWidth = 1
                    .FillStyle = vbFSTransparent
                End If
            End If
            If (uChartType And XC_SYMBOL) = XC_SYMBOL Then
                'draw the symbol in the higher (absolute) point
                If uDataType = DT_NEG Then
                    yTemp = y2
                ElseIf uDataType = DT_POS Then
                    yTemp = y1
                Else
                    yTemp = IIf((oChartItem.Value > 0), y1, y2)
                End If
                xTemp = x1 + (uColWidth / 2) - (.TextWidth(uBarSymbol) / 2)
                yTemp = yTemp - (.TextHeight(uBarSymbol) / 2)
                If (uMeanOn = True) And (X = cItems.Count - 1) Then
                    .ForeColor = uMeanColor
                Else
                    .ForeColor = uBarSymbolColor
                End If
                .CurrentX = xTemp
                .CurrentY = yTemp
                .FontSize = .FontSize + 2
                UserControl.Print uBarSymbol
                .FontSize = .FontSize - 2
            End If
            If (uChartType And XC_LINE) = XC_LINE Then
                'draw the lines
                If uDataType = DT_NEG Then
                    yTemp = y2
                ElseIf uDataType = DT_POS Then
                    yTemp = y1
                Else
                    yTemp = IIf((oChartItem.Value > 0), y1, y2)
                End If
                xTemp = x1 + (uColWidth / 2)
                'check if it's the first data: if it is, do not draw the line
                If (X > 0) Then
                    If (uMeanOn = True And X = cItems.Count - 1) Then
                        'do nothing
                    Else
                        .DrawStyle = uLineStyle
                        .DrawWidth = uLineWidth
                        UserControl.Line (xPrev, yPrev)-(xTemp, yTemp), uLineColor
                        .DrawWidth = 1
                        .DrawStyle = vbSolid
                    End If
                End If
                xPrev = xTemp
                yPrev = yTemp
            End If
    
            'display X-axis labels and ticks
            If uAxisXOn Then
                .ForeColor = uXAxisItemsColor
                .FontSize = .FontSize - 1
                varLabel = Split(oChartItem.XAxisDescription, vbCrLf)
                'calculate X offset
                xTemp = (((x2 - x1) / 2) + x1) / Screen.TwipsPerPixelX
                'subtract displacement depending on how many lines are
                xTemp = xTemp - (((.TextHeight("A")) / Screen.TwipsPerPixelX) * (UBound(varLabel))) / 2
                For intIdx = UBound(varLabel) To 0 Step -1
                    'Y coordinate is the center of the rectangle used to draw text
                    yTemp = (.ScaleHeight - uBottomMargin + .TextWidth(varLabel(intIdx)) / 2) / Screen.TwipsPerPixelY + 5
                    PrintRotText .hDC, varLabel(intIdx), xTemp, yTemp, 270
                    'move X coordinate foreward
                    xTemp = xTemp + .TextHeight("A") / Screen.TwipsPerPixelX
                Next
                xTemp = (((x2 - x1) / 2) + x1) / Screen.TwipsPerPixelX
                yTemp = (.ScaleHeight - uBottomMargin) + Screen.TwipsPerPixelX
                UserControl.Line (xTemp * Screen.TwipsPerPixelX, yTemp)-(xTemp * Screen.TwipsPerPixelX, yTemp + 2 * Screen.TwipsPerPixelX), uMajorGridColor
                .FontSize = .FontSize + 1
            End If
            'Add Legend item
            If Not bResize Then
                If (uMeanOn = True) And (X = cItems.Count - 1) Then
                    .FillColor = uMeanColor
                ElseIf ((uChartType And XC_BAR) = XC_BAR) _
                Or (uChartType And XC_OVAL) = XC_OVAL _
                Or (uChartType And XC_RHOMBUS) = XC_RHOMBUS _
                Or (uChartType And XC_TRAPEZIUM) = XC_TRAPEZIUM _
                Or (uChartType And XC_TRIANGLE) = XC_TRIANGLE Then
                    'do nothing, since FillColor is already set
                ElseIf (uChartType And XC_LINE) = XC_LINE Then
                    .FillColor = uLineColor
                ElseIf (uChartType And XC_SYMBOL) = XC_SYMBOL Then
                    .FillColor = uBarSymbolColor
                End If
                AddLegendItem oChartItem.LegendDescription, .FillColor, uLegendForeColor
            End If
            
            If uColorBars = True Then
                CurrentColor = CurrentColor + 1
                If CurrentColor >= 15 Then CurrentColor = 0
            End If
        Next X
    
        'Print the x axis label
        If Len(uXAxisLabel) Then
            .FontSize = .FontSize - 1
            .CurrentY = .ScaleHeight - .TextHeight(uXAxisLabel) * 1.5
            .CurrentX = xMiddle - (.TextWidth(uXAxisLabel) / 2)
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
                picContainer.Line ((picBox(uSelectedColumn).Left - 3 * Screen.TwipsPerPixelX), (picBox(uSelectedColumn).Top - 3 * Screen.TwipsPerPixelY))-(picDescription(uSelectedColumn).Left + picDescription(uSelectedColumn).Width + 2 * Screen.TwipsPerPixelX, picBox(uSelectedColumn).Top + picBox(uSelectedColumn).Height + 2 * Screen.TwipsPerPixelY), uSelectedBarColor, B
            End If
        End If
    End With

End Sub

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

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    
    With PropBag
        uTopMargin = .ReadProperty("uTopMargin", 55)
        uBottomMargin = .ReadProperty("uBottomMargin", 55)
        uLeftMargin = .ReadProperty("uLeftMargin", 55)
        uRightMargin = .ReadProperty("uRightMargin", 55)
        uContentBorder = .ReadProperty("uContentBorder", True)
        uSelectable = .ReadProperty("uSelectable", False)
        uHotTracking = .ReadProperty("uHotTracking", False)
        uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uAxisYOn = .ReadProperty("uAxisXOn", uAxisXOn)
        uAxisXOn = .ReadProperty("uAxisYOn", uAxisYOn)
        uColorBars = .ReadProperty("uColorBars", False)
        uIntersectMajor = .ReadProperty("uIntersectMajor", 10)
        uIntersectMinor = .ReadProperty("uIntersectMinor", 2)
        uMaxYValue = .ReadProperty("uMaxYValue", 100)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        uXAxisLabel = .ReadProperty("uXAxisLabel", uXAxisLabel)
        uYAxisLabel = .ReadProperty("uYAxisLabel", uYAxisLabel)
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        UserControl.ForeColor = .ReadProperty("ForeColor", vbBlack)
        '----------------------------------------------------
        'added by M. Costa on 21/06/2002
        uMinYValue = .ReadProperty("MinY")
        uBarColor = .ReadProperty("BarColor", vbBlue)
        uSelectedBarColor = .ReadProperty("SelectedBarColor", vbCyan)
        uMajorGridColor = .ReadProperty("MajorGridColor", vbBlack)
        uMinorGridColor = .ReadProperty("MinorGridColor", vbBlack)
        uLegendBackColor = .ReadProperty("LegendBackColor", vbWhite)
        uLegendForeColor = .ReadProperty("LegendForeColor", vbBlack)
        uInfoBackColor = .ReadProperty("InfoBackColor")
        uInfoForeColor = .ReadProperty("InfoForeColor")
        uXAxisLabelColor = .ReadProperty("XAxisLabelColor", vbBlack)
        uYAxisLabelColor = .ReadProperty("YAxisLabelColor", vbBlack)
        uXAxisItemsColor = .ReadProperty("XAxisItemsColor", vbBlack)
        uYAxisItemsColor = .ReadProperty("YAxisItemsColor", vbBlack)
        uChartTitleColor = .ReadProperty("ChartTitleColor", vbBlack)
        uChartSubTitleColor = .ReadProperty("ChartSubTitleColor", vbBlack)
        uChartType = .ReadProperty("ChartType")
        uMenuType = .ReadProperty("MenuType")
        uMenuItems = .ReadProperty("MenuItems")
        uCustomMenuItems = .ReadProperty("CustomMenuItems")
        uInfoItems = .ReadProperty("InfoItems")
        uSaveAsCaption = .ReadProperty("SaveAsCaption")
        uAutoRedraw = .ReadProperty("AutoRedraw", True)
        uBarWidthPercentage = .ReadProperty("BarWidthPercentage", 100)
        uBarSymbol = .ReadProperty("BarSymbol", "*")
        Set uBarPicture = .ReadProperty("BarPicture", Nothing)
        uBarPictureTile = .ReadProperty("BarPictureTile", False)
        Set uPicture = .ReadProperty("Picture", Nothing)
        uPictureTile = .ReadProperty("PictureTile", False)
        uMinorGridOn = .ReadProperty("MinorGridOn", True)
        uMajorGridOn = .ReadProperty("MajorGridOn", True)
        uLineWidth = .ReadProperty("LineWidth", 1)
        uLineColor = .ReadProperty("LineColor", vbRed)
        uBarSymbolColor = .ReadProperty("BarSymbolColor", vbRed)
        uBarFillStyle = .ReadProperty("BarFillStyle", vbFSSolid)
        uLineStyle = .ReadProperty("LineStyle")
        uBarShadow = .ReadProperty("BarShadow", True)
        uBarShadowColor = .ReadProperty("BarShadowColor", vbBlack)
        uMeanOn = .ReadProperty("MeanOn", False)
        uMeanColor = .ReadProperty("MeanColor", vbCyan)
        uMeanCaption = .ReadProperty("MeanCaption")
        uDataFormat = .ReadProperty("DataFormat")
        uPrinterFit = .ReadProperty("PrinterFit")
        uPrinterOrientation = .ReadProperty("PrinterOrientation")
        uLegendCaption = .ReadProperty("LegendCaption")
        uLegendPrintMode = .ReadProperty("LegendPrintMode", legPrintGraph)
        '----------------------------------------------------
        uOldSelection = -1
        uRightMarginOrg = uRightMargin
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
        .WriteProperty "uTopMargin", uTopMargin
        .WriteProperty "uBottomMargin", uBottomMargin
        .WriteProperty "uLeftMargin", uLeftMargin
        .WriteProperty "uRightMargin", uRightMargin
        .WriteProperty "uContentBorder", uContentBorder
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uHotTracking", uHotTracking
        .WriteProperty "uSelectedColumn", uSelectedColumn
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uAxisXOn", uAxisXOn
        .WriteProperty "uAxisYOn", uAxisYOn
        .WriteProperty "uColorBars", uColorBars
        .WriteProperty "uIntersectMajor", uIntersectMajor
        .WriteProperty "uIntersectMinor", uIntersectMinor
        .WriteProperty "uMaxYValue", uMaxYValue
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "uXAxisLabel", uXAxisLabel
        .WriteProperty "uYAxislabel", uYAxisLabel
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
        '----------------------------------------------------
        'added by M. Costa on 21/06/2002
        .WriteProperty "MinY", uMinYValue
        .WriteProperty "BarColor", uBarColor
        .WriteProperty "SelectedBarColor", uSelectedBarColor
        .WriteProperty "MajorGridColor", uMajorGridColor
        .WriteProperty "MinorGridColor", uMinorGridColor
        .WriteProperty "LegendBackColor", uLegendBackColor
        .WriteProperty "LegendForeColor", uLegendForeColor
        .WriteProperty "InfoBackColor", uInfoBackColor
        .WriteProperty "InfoForeColor", uInfoForeColor
        .WriteProperty "XAxisLabelColor", uXAxisLabelColor
        .WriteProperty "YAxisLabelColor", uYAxisLabelColor
        .WriteProperty "XAxisItemsColor", uXAxisItemsColor
        .WriteProperty "YAxisItemsColor", uYAxisItemsColor
        .WriteProperty "ChartTitleColor", uChartTitleColor
        .WriteProperty "ChartSubTitleColor", uChartSubTitleColor
        .WriteProperty "ChartType", uChartType
        .WriteProperty "MenuType", uMenuType
        .WriteProperty "MenuItems", uMenuItems
        .WriteProperty "CustomMenuItems", uCustomMenuItems
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SaveAsCaption", uSaveAsCaption
        .WriteProperty "AutoRedraw", uAutoRedraw
        .WriteProperty "BarWidthPercentage", uBarWidthPercentage
        .WriteProperty "BarSymbol", uBarSymbol
        .WriteProperty "BarPicture", uBarPicture, Nothing
        .WriteProperty "BarPictureTile", uBarPictureTile
        .WriteProperty "Picture", uPicture, Nothing
        .WriteProperty "PictureTile", uPictureTile
        .WriteProperty "MinorGridOn", uMinorGridOn
        .WriteProperty "MajorGridOn", uMajorGridOn
        .WriteProperty "LineWidth", uLineWidth
        .WriteProperty "LineColor", uLineColor
        .WriteProperty "BarSymbolColor", uBarSymbolColor
        .WriteProperty "BarFillStyle", uBarFillStyle
        .WriteProperty "LineStyle", uLineStyle
        .WriteProperty "BarShadow", uBarShadow
        .WriteProperty "BarShadowColor", uBarShadowColor
        .WriteProperty "MeanOn", uMeanOn
        .WriteProperty "MeanColor", uMeanColor
        .WriteProperty "MeanCaption", uMeanCaption
        .WriteProperty "DataFormat", uDataFormat
        .WriteProperty "PrinterFit", uPrinterFit
        .WriteProperty "PrinterOrientation", uPrinterOrientation
        .WriteProperty "LegendCaption", uLegendCaption
        .WriteProperty "LegendPrintMode", uLegendPrintMode
        '----------------------------------------------------
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

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Returns/sets the line color."
    LineColor = uLineColor
End Property

Public Property Let LineColor(lngVal As OLE_COLOR)
    If lngVal <> uLineColor Then
        uLineColor = lngVal
        DrawChart
        PropertyChanged "LineColor"
    End If
End Property

Public Property Get BarSymbolColor() As OLE_COLOR
Attribute BarSymbolColor.VB_Description = "Returns/sets the color used to display the symbol."
    BarSymbolColor = uBarSymbolColor
End Property

Public Property Let BarSymbolColor(lngVal As OLE_COLOR)
    If uBarSymbolColor <> lngVal Then
        uBarSymbolColor = lngVal
        DrawChart
        PropertyChanged "BarSymbolColor"
    End If
End Property

Public Property Get BarFillStyle() As FillStyleConstants
Attribute BarFillStyle.VB_Description = "Returns/sets the fill style of the bar."
    BarFillStyle = uBarFillStyle
End Property

Public Property Let BarFillStyle(intVal As FillStyleConstants)
    If uBarFillStyle <> intVal Then
        uBarFillStyle = intVal
        DrawChart
        PropertyChanged "BarFillStyle"
    End If
End Property
