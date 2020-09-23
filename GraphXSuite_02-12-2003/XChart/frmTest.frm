VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "*\AXChart.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   Caption         =   "ActiveChart Test"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPrtMode 
      Height          =   315
      Left            =   4530
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   5250
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show mean value"
      Height          =   195
      Index           =   6
      Left            =   1950
      TabIndex        =   40
      Top             =   7980
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tile bar picture"
      Height          =   195
      Index           =   5
      Left            =   3900
      TabIndex        =   39
      Top             =   7500
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show bar picture"
      Height          =   195
      Index           =   4
      Left            =   3900
      TabIndex        =   38
      Top             =   7740
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bar shadow"
      Height          =   195
      Index           =   3
      Left            =   3900
      TabIndex        =   36
      Top             =   7980
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.TextBox txtSymbol 
      Height          =   285
      Left            =   5010
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   6390
      Width           =   705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show back picture"
      Height          =   195
      Index           =   2
      Left            =   3900
      TabIndex        =   31
      Top             =   7260
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.TextBox txtLineWidth 
      Height          =   285
      Left            =   5010
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   6060
      Width           =   705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tile back picture"
      Height          =   195
      Index           =   1
      Left            =   3900
      TabIndex        =   28
      Top             =   7020
      Width           =   1845
   End
   Begin VB.TextBox txtBarPerc 
      Height          =   285
      Left            =   5010
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   5730
      Width           =   705
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   3870
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   4920
      Width           =   1845
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1170
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   7950
      Width           =   705
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   1170
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   7620
      Width           =   705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hot tracking"
      Height          =   195
      Index           =   0
      Left            =   1950
      TabIndex        =   20
      Top             =   7740
      Width           =   1845
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Buttons"
      Height          =   225
      Index           =   1
      Left            =   3900
      TabIndex        =   17
      Top             =   6750
      Width           =   885
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Popup"
      Height          =   225
      Index           =   0
      Left            =   4860
      TabIndex        =   16
      Top             =   6750
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply settings"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   8250
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   7200
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose color"
   End
   Begin ActiveChart.XChart XChart1 
      Height          =   4755
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   8387
      uTopMargin      =   600
      uBottomMargin   =   750
      uLeftMargin     =   750
      uRightMargin    =   750
      uContentBorder  =   -1  'True
      uSelectable     =   -1  'True
      uHotTracking    =   -1  'True
      uSelectedColumn =   -1
      uChartTitle     =   "Gain in 2002"
      uChartSubTitle  =   "Italy"
      uAxisXOn        =   -1  'True
      uAxisYOn        =   -1  'True
      uColorBars      =   0   'False
      uIntersectMajor =   200
      uIntersectMinor =   50
      uMaxYValue      =   1000
      uDisplayDescript=   -1  'True
      uXAxisLabel     =   "Months of the year"
      uYAxislabel     =   "(in Euro)"
      BackColor       =   8421376
      ForeColor       =   16776960
      MinY            =   -1000
      BarColor        =   49152
      SelectedBarColor=   16777088
      MajorGridColor  =   16777215
      MinorGridColor  =   0
      LegendBackColor =   4210688
      LegendForeColor =   16777215
      InfoBackColor   =   12648447
      InfoForeColor   =   16711680
      XAxisLabelColor =   16776960
      YAxisLabelColor =   16776960
      XAxisItemsColor =   4210688
      YAxisItemsColor =   4210688
      ChartTitleColor =   65535
      ChartSubTitleColor=   8454143
      ChartType       =   1
      MenuType        =   0
      MenuItems       =   "&Save as...|&Print|&Copy|Selection &information|&Legend|&Hide"
      CustomMenuItems =   ""
      InfoItems       =   ""
      SaveAsCaption   =   "Salva grafico"
      AutoRedraw      =   -1  'True
      BarWidthPercentage=   100
      BarSymbol       =   "*"
      BarPicture      =   "frmTest.frx":0000
      BarPictureTile  =   -1  'True
      Picture         =   "frmTest.frx":02A3
      PictureTile     =   0   'False
      MinorGridOn     =   0   'False
      MajorGridOn     =   -1  'True
      LineWidth       =   1
      LineColor       =   255
      BarSymbolColor  =   255
      BarFillStyle    =   0
      LineStyle       =   0
      BarShadow       =   -1  'True
      BarShadowColor  =   0
      MeanOn          =   -1  'True
      MeanColor       =   65535
      MeanCaption     =   ""
      DataFormat      =   "##.00"
      PrinterFit      =   0
      PrinterOrientation=   0
      LegendCaption   =   "Display legend"
      LegendPrintMode =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3690
      Left            =   5880
      TabIndex        =   0
      Top             =   4890
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6509
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Printing mode"
      Height          =   390
      Left            =   3870
      TabIndex        =   42
      Top             =   5280
      Width           =   660
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bar shadow color"
      Height          =   255
      Index           =   17
      Left            =   1950
      TabIndex        =   37
      Top             =   7320
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Symbol"
      Height          =   195
      Left            =   3870
      TabIndex        =   35
      Top             =   6420
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Symbol color"
      Height          =   255
      Index           =   16
      Left            =   1950
      TabIndex        =   33
      Top             =   7020
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line color"
      Height          =   255
      Index           =   15
      Left            =   1950
      TabIndex        =   32
      Top             =   6720
      Width           =   1845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Line width"
      Height          =   195
      Left            =   3870
      TabIndex        =   30
      Top             =   6090
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bar width (%)"
      Height          =   195
      Left            =   3870
      TabIndex        =   26
      Top             =   5790
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max. Y value"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   24
      Top             =   8010
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Y value"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Major grid color"
      Height          =   255
      Index           =   14
      Left            =   1950
      TabIndex        =   19
      Top             =   6420
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minor grid color"
      Height          =   255
      Index           =   13
      Left            =   1950
      TabIndex        =   18
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info foreground color"
      Height          =   255
      Index           =   12
      Left            =   1950
      TabIndex        =   15
      Top             =   5520
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info background color"
      Height          =   255
      Index           =   11
      Left            =   1950
      TabIndex        =   14
      Top             =   5820
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Legend background color"
      Height          =   255
      Index           =   10
      Left            =   1950
      TabIndex        =   13
      Top             =   5220
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Legend foreground color"
      Height          =   255
      Index           =   9
      Left            =   1950
      TabIndex        =   12
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bar color"
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   11
      Top             =   7320
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected bar color"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   10
      Top             =   7020
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y axis items color"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   9
      Top             =   6720
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y axis label color"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   8
      Top             =   6420
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X axis items color"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X axis label color"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   5820
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtitle color"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   5520
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Title color"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   5220
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Background color"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   4920
      Width           =   1845
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PrepareData()
    
    Dim X As Integer
    Dim intSign As Integer
    Dim oChartItem As ChartItem
    Dim varMonths As Variant
    Dim varMonthsExt As Variant
    
    varMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    varMonthsExt = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

    Randomize
    grd.Rows = 1
    With XChart1
        .AutoRedraw = True
        .Clear
        .CustomMenuItems = "&Hello|&World|Print legend text"
        For X = 1 To 12
            If .MinY < 0 And .MaxY >= 0 Then
                intSign = CInt(Rnd * 1)
                If intSign = 0 Then
                    oChartItem.Value = CInt(Rnd * .MaxY)
                Else
                    oChartItem.Value = -CInt(Rnd * Abs(.MinY))
                End If
            ElseIf .MinY >= 0 And .MaxY >= 0 Then
                oChartItem.Value = .MinY + CInt(Rnd * (.MaxY - .MinY))
            ElseIf .MinY < 0 And .MaxY < 0 Then
                oChartItem.Value = .MaxY - CInt(Rnd * (Abs(.MinY) - Abs(.MaxY)))
            End If
            oChartItem.ItemID = X
            oChartItem.XAxisDescription = "Month" & vbCrLf & varMonths(X - 1)
            oChartItem.SelectedDescription = varMonthsExt(X - 1)
            oChartItem.LegendDescription = "Month " & varMonthsExt(X - 1)
            .AddItem oChartItem
    
            grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
        Next X
    End With
    
End Sub

Private Sub RefreshData()
    
    Dim intIdx As Integer
    
    With XChart1
        Label1(0).BackColor = .BackColor
        Label1(1).BackColor = .ChartTitleColor
        Label1(2).BackColor = .ChartSubTitleColor
        Label1(3).BackColor = .AxisLabelXColor
        Label1(4).BackColor = .AxisItemsXColor
        Label1(5).BackColor = .AxisLabelYColor
        Label1(6).BackColor = .AxisItemsYColor
        Label1(7).BackColor = .SelectedBarColor
        Label1(8).BackColor = .BarColor
        Label1(9).BackColor = .LegendForeColor
        Label1(10).BackColor = .LegendbackColor
        Label1(11).BackColor = .InfoForeColor
        Label1(12).BackColor = .InfoBackColor
        Label1(13).BackColor = .MinorGridColor
        Label1(14).BackColor = .MajorGridColor
        Label1(15).BackColor = .LineColor
        Label1(16).BackColor = .BarSymbolColor
        Label1(17).BackColor = .BarShadowColor
        Option1(.MenuType).Value = True
        Check1(0).Value = IIf((.HotTracking = True), vbChecked, vbUnchecked)
        Check1(1).Value = IIf((.PictureTile = True), vbChecked, vbUnchecked)
        Check1(3).Value = IIf((.BarShadow = True), vbChecked, vbUnchecked)
        Check1(5).Value = IIf((.BarPictureTile = True), vbChecked, vbUnchecked)
        Check1(6).Value = IIf((.MeanOn = True), vbChecked, vbUnchecked)
        txtMin.Text = .MinY
        txtMax.Text = .MaxY
        txtBarPerc.Text = CStr(.BarWidthPercentage)
        txtLineWidth.Text = CStr(.LineWidth)
        txtSymbol.Text = .BarSymbol
        For intIdx = 0 To cboType.ListCount - 1
            If cboType.ItemData(intIdx) = .ChartType Then
                cboType.ListIndex = intIdx
                Exit For
            End If
        Next
        For intIdx = 0 To cboPrtMode.ListCount - 1
            If cboPrtMode.ItemData(intIdx) = .printerfit Then
                cboPrtMode.ListIndex = intIdx
                Exit For
            End If
        Next
    End With

End Sub

Private Sub Command1_Click()

    With XChart1
        .AutoRedraw = False
        .BackColor = Label1(0).BackColor
        .ChartTitleColor = Label1(1).BackColor
        .ChartSubTitleColor = Label1(2).BackColor
        .AxisLabelXColor = Label1(3).BackColor
        .AxisItemsXColor = Label1(4).BackColor
        .AxisLabelYColor = Label1(5).BackColor
        .AxisItemsYColor = Label1(6).BackColor
        .SelectedBarColor = Label1(7).BackColor
        .BarColor = Label1(8).BackColor
        .LegendForeColor = Label1(9).BackColor
        .LegendbackColor = Label1(10).BackColor
        .InfoForeColor = Label1(11).BackColor
        .InfoBackColor = Label1(12).BackColor
        .MinorGridColor = Label1(13).BackColor
        .MajorGridColor = Label1(14).BackColor
        .LineColor = Label1(15).BackColor
        .BarSymbolColor = Label1(16).BackColor
        .BarShadowColor = Label1(17).BackColor
        If Option1(0).Value = True Then
            .MenuType = xcPopUpMenu
        Else
            .MenuType = xcButtonMenu
        End If
        .HotTracking = IIf((Check1(0).Value = vbChecked), True, False)
        .PictureTile = IIf((Check1(1).Value = vbChecked), True, False)
        .BarShadow = IIf((Check1(3).Value = vbChecked), True, False)
        .BarPictureTile = IIf((Check1(5).Value = vbChecked), True, False)
        .MeanOn = IIf((Check1(6).Value = vbChecked), True, False)
        
        .LineWidth = CInt(txtLineWidth.Text)
        .BarWidthPercentage = CInt(txtBarPerc.Text)
        .MinY = CDbl(txtMin.Text)
        .MaxY = CDbl(txtMax.Text)
        PrepareData
        .ChartType = cboType.ItemData(cboType.ListIndex)
        If Check1(2).Value = vbUnchecked Then
            Set .Picture = Nothing
        Else
            Set .Picture = LoadPicture(App.Path & "\stonehng.jpg")
        End If
        If Check1(4).Value = vbUnchecked Then
            Set .BarPicture = Nothing
        Else
            Set .BarPicture = LoadPicture(App.Path & "\tile1.jpg")
        End If
        .BarSymbol = Left$(txtSymbol.Text, 1)
        .printerfit = cboPrtMode.ItemData(cboPrtMode.ListIndex)
        .AutoRedraw = True
    End With
    RefreshData

End Sub

Private Sub Label1_Click(Index As Integer)
    
    dlgColor.Color = Label1(Index).BackColor
    dlgColor.ShowColor
    If dlgColor.Color <> Label1(Index).BackColor Then
        Label1(Index).BackColor = dlgColor.Color
    End If

End Sub

Private Sub xchart1_ItemClick(cItem As ActiveChart.ChartItem)
    grd.SelectionMode = flexSelectionByRow
    grd.Row = cItem.ItemID
    grd.ColSel = 2
End Sub

Private Sub Form_Load()
    
    With cboType
        .Clear
        .AddItem "Bar":             .ItemData(.NewIndex) = xcBar
        .AddItem "Symbol":          .ItemData(.NewIndex) = xcSymbol
        .AddItem "Line":            .ItemData(.NewIndex) = xcLine
        .AddItem "BarLine":         .ItemData(.NewIndex) = xcBarLine
        .AddItem "SymbolLine":      .ItemData(.NewIndex) = xcSymbolLine
        .AddItem "Oval":            .ItemData(.NewIndex) = xcOval
        .AddItem "OvalLine":        .ItemData(.NewIndex) = xcOvalLine
        .AddItem "Triangle":        .ItemData(.NewIndex) = xcTriangle
        .AddItem "TriangleLine":    .ItemData(.NewIndex) = xcTriangleLine
        .AddItem "Rhombus":         .ItemData(.NewIndex) = xcRhombus
        .AddItem "RhombusLine":     .ItemData(.NewIndex) = xcRhombusLine
        .AddItem "Trapezium":       .ItemData(.NewIndex) = xcTrapezium
        .AddItem "TrapeziumLine":   .ItemData(.NewIndex) = xcTrapeziumLine
    End With
    
    With cboPrtMode
        .Clear
        .AddItem "Stretched":      .ItemData(.NewIndex) = prtFitStretched
        .AddItem "Centered":       .ItemData(.NewIndex) = prtFitCentered
        .AddItem "TopLeft":        .ItemData(.NewIndex) = prtFitTopLeft
        .AddItem "TopRight":       .ItemData(.NewIndex) = prtFitTopRight
        .AddItem "BottomLeft":     .ItemData(.NewIndex) = prtFitBottomLeft
        .AddItem "BottomRight":    .ItemData(.NewIndex) = prtFitBottomRight
    End With
    
    PrepareData
        
    With grd
        .FixedRows = 1
        .TextMatrix(0, 0) = "Item"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Value"
        .ColWidth(0) = 800
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
    End With

    RefreshData
    
End Sub

Private Sub Form_Resize()
'    grd.Width = Me.ScaleWidth
'    XChart1.Width = Me.ScaleWidth

'    grd.ColWidth(0) = 960
'    grd.ColWidth(1) = Me.ScaleWidth - 960 - 2025
'    grd.ColWidth(2) = 2025
End Sub

Private Sub grd_Click()
    DoEvents
    XChart1.SelectedColumn = grd.Row - 1
End Sub



Private Sub XChart1_MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
    
    MsgBox "You clicked " & intMenuItemIndex & ":" & stgMenuItemCaption, _
            vbOKOnly, "CustomMenuItem"

    If intMenuItemIndex = 2 Then
        XChart1.PrintLegend
    End If
    
End Sub


