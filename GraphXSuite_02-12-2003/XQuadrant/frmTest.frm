VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\A..\XQUADR~1\XQuadrant.vbp"
Begin VB.Form frmTest 
   Caption         =   "ActiveQuadrant Test"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPrtMode 
      Height          =   315
      Left            =   6750
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   5640
      Width           =   915
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   6750
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5280
      Width           =   915
   End
   Begin VB.TextBox txtMinX 
      Height          =   285
      Left            =   3060
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   7110
      Width           =   705
   End
   Begin VB.TextBox txtMaxX 
      Height          =   285
      Left            =   3060
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   7440
      Width           =   705
   End
   Begin ActiveQuadrant.XQuadrant XQuadrant1 
      Height          =   4545
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   8017
      InfoItems       =   "XY values|Description"
      SelectedColor   =   255
      uSelectable     =   -1  'True
      uSelectedColumn =   -1
      uTopMargin      =   750
      uBottomMargin   =   825
      uLeftMargin     =   825
      uRightMargin    =   825
      uContentBorder  =   -1  'True
      uChartTitle     =   "Quadrants"
      uChartSubTitle  =   "(data processed on...)"
      uAxisXOn        =   -1  'True
      uAxisYOn        =   -1  'True
      uIntersectMajorY=   1
      uIntersectMinorY=   0,5
      uIntersectMajorX=   1
      uIntersectMinorX=   0,5
      uMaxYValue      =   5
      uMaxXValue      =   5
      QuadrantY       =   3
      QuadrantX       =   3
      uDisplayDescript=   0   'False
      uXAxisLabel     =   "Technical"
      uYAxislabel     =   "Business"
      BackColor       =   16777152
      ForeColor       =   4194368
      MinY            =   0
      MinX            =   0
      MajorGridColor  =   4194304
      MinorGridColor  =   16776960
      LegendBackColor =   16777088
      LegendForeColor =   16711680
      InfoBackColor   =   -2147483624
      InfoForeColor   =   -2147483625
      InfoQuadrantBackColor=   12648447
      InfoQuadrantForeColor=   4210688
      XAxisLabelColor =   4194368
      YAxisLabelColor =   4194368
      XAxisItemsColor =   8388608
      YAxisItemsColor =   8388608
      ChartTitleColor =   8421376
      ChartSubTitleColor=   4210688
      MenuType        =   0
      MenuItems       =   "&Save as...|&Print|&Copy|Selection &information|&Quadrant information|&Legend|&Hide"
      CustomMenuItems =   ""
      InfoItems       =   "XY values|Description"
      SaveAsCaption   =   ""
      AutoRedraw      =   -1  'True
      MarkerSymbol    =   0
      PictureTile     =   0   'False
      MinorGridOn     =   -1  'True
      MajorGridOn     =   -1  'True
      MarkerWidth     =   1
      MarkerColor     =   255
      DataFormat      =   ""
      PrinterFit      =   5
      PrinterOrientation=   2
      LegendCaption   =   "Display legend"
      MarkerLabelAngle=   45
      MarkerLabelDirection=   2
      ChartAsQuadrant =   -1  'True
      QuadrantDividerColor=   16711935
      QuadrantColorsOverridePicture=   -1  'True
      MarkerLabel     =   16576
      QuadrantColors  =   "13602382|13678156|12177484|10932300"
      HotTracking     =   0   'False
      LegendPrintMode =   1
      InnerColor      =   128
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show back picture"
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   29
      Top             =   7590
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtMarkerWidth 
      Height          =   285
      Left            =   6750
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4920
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tile back picture"
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   22
      Top             =   7290
      Width           =   1815
   End
   Begin VB.TextBox txtMaxY 
      Height          =   285
      Left            =   1170
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   7470
      Width           =   705
   End
   Begin VB.TextBox txtMinY 
      Height          =   285
      Left            =   1170
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   7140
      Width           =   705
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Buttons menu"
      Height          =   225
      Index           =   1
      Left            =   3840
      TabIndex        =   15
      Top             =   6990
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Popup menu"
      Height          =   225
      Index           =   0
      Left            =   3840
      TabIndex        =   14
      Top             =   6690
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply settings"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   7830
      Width           =   7515
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   7200
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose color"
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3270
      Left            =   7770
      TabIndex        =   1
      Top             =   4890
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   5768
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Printing mode"
      Height          =   390
      Left            =   5790
      TabIndex        =   27
      Top             =   5670
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info quadrant backcolor"
      Height          =   255
      Index           =   18
      Left            =   3870
      TabIndex        =   39
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info quadrant forecolor"
      Height          =   255
      Index           =   17
      Left            =   3870
      TabIndex        =   38
      Top             =   5820
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected color"
      Height          =   255
      Index           =   15
      Left            =   3870
      TabIndex        =   37
      Top             =   5520
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quad. div. color"
      Height          =   255
      Index           =   8
      Left            =   3870
      TabIndex        =   36
      Top             =   5220
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marker label color"
      Height          =   255
      Index           =   7
      Left            =   3870
      TabIndex        =   35
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. X value"
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   34
      Top             =   7170
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max. X value"
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   33
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Marker symbol"
      Height          =   390
      Left            =   5790
      TabIndex        =   25
      Top             =   5250
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marker color"
      Height          =   255
      Index           =   16
      Left            =   1950
      TabIndex        =   30
      Top             =   6720
      Width           =   1845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Marker width"
      Height          =   195
      Left            =   5790
      TabIndex        =   23
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max. Y value"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   21
      Top             =   7530
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Min. Y value"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   20
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Major grid color"
      Height          =   255
      Index           =   14
      Left            =   1950
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   4920
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
    Dim varYears As Variant
    Dim varYearsExt As Variant
    
    varYears = Array("2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015")
    varYearsExt = Array("2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015")

    Randomize
    grd.Rows = 1
    With XQuadrant1
        .Clear
        .CustomMenuItems = "&Hello|&World|Print legend text"
        For X = 1 To UBound(varYears) + 1
            Randomize
            If .MinY < 0 And .MaxY >= 0 Then
                intSign = CInt(Rnd * 1)
                If intSign = 0 Then
                    oChartItem.YValue = CDbl(Rnd * .MaxY)
                Else
                    oChartItem.YValue = -CDbl(Rnd * Abs(.MinY))
                End If
            ElseIf .MinY >= 0 And .MaxY >= 0 Then
                oChartItem.YValue = .MinY + CDbl(Rnd * (.MaxY - .MinY))
            ElseIf .MinY < 0 And .MaxY < 0 Then
                oChartItem.YValue = .MaxY - CDbl(Rnd * (Abs(.MinY) - Abs(.MaxY)))
            End If
            
            Randomize
            If .Minx < 0 And .Maxx >= 0 Then
                intSign = CInt(Rnd * 1)
                If intSign = 0 Then
                    oChartItem.XValue = CDbl(Rnd * .Maxx)
                Else
                    oChartItem.XValue = -CDbl(Rnd * Abs(.Minx))
                End If
            ElseIf .Minx >= 0 And .Maxx >= 0 Then
                oChartItem.XValue = .Minx + CDbl(Rnd * (.Maxx - .Minx))
            ElseIf .Minx < 0 And .Maxx < 0 Then
                oChartItem.XValue = .Maxx - CDbl(Rnd * (Abs(.Minx) - Abs(.Maxx)))
            End If
            
            With oChartItem
                .XValue = Round(oChartItem.XValue, 2)
                .YValue = Round(oChartItem.YValue, 2)
                .ItemID = X
                .Description = varYears(X - 1)
                .SelectedDescription = "Year" & vbCrLf & varYearsExt(X - 1)
                .LegendDescription = "Year " & varYearsExt(X - 1)
            End With
            .AddItem oChartItem
    
            grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.XValue & "/" & oChartItem.YValue
        Next X
    End With
    
End Sub

Private Sub RefreshData()
    
    Dim intIdx As Integer
    
    With XQuadrant1
        Label1(0).BackColor = .BackColor
        Label1(1).BackColor = .ChartTitleColor
        Label1(2).BackColor = .ChartSubTitleColor
        Label1(3).BackColor = .AxisLabelXColor
        Label1(4).BackColor = .AxisItemsXColor
        Label1(5).BackColor = .AxisLabelYColor
        Label1(6).BackColor = .AxisItemsYColor
        Label1(7).BackColor = .MarkerLabelColor
        Label1(8).BackColor = .QuadrantDividerColor
        Label1(9).BackColor = .LegendForeColor
        Label1(10).BackColor = .LegendbackColor
        Label1(11).BackColor = .InfoForeColor
        Label1(12).BackColor = .InfoBackColor
        Label1(13).BackColor = .MinorGridColor
        Label1(14).BackColor = .MajorGridColor
        Label1(15).BackColor = .SelectedColor
        Label1(16).BackColor = .MarkerColor
        Label1(17).BackColor = .InfoQuadrantforeColor
        Label1(18).BackColor = .InfoQuadrantBackColor
        Option1(.MenuType).Value = True
        Check1(1).Value = IIf((.PictureTile = True), vbChecked, vbUnchecked)
        txtMinY.Text = .MinY
        txtMaxY.Text = .MaxY
        txtMinX.Text = .Minx
        txtMaxX.Text = .Maxx
        txtMarkerWidth.Text = CStr(.markerWidth)
        For intIdx = 0 To cboType.ListCount - 1
            If cboType.ItemData(intIdx) = .MarkerSymbol Then
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

    With XQuadrant1
        .AutoRedraw = False
        .BackColor = Label1(0).BackColor
        .ChartTitleColor = Label1(1).BackColor
        .ChartSubTitleColor = Label1(2).BackColor
        .AxisLabelXColor = Label1(3).BackColor
        .AxisItemsXColor = Label1(4).BackColor
        .AxisLabelYColor = Label1(5).BackColor
        .AxisItemsYColor = Label1(6).BackColor
        .LegendForeColor = Label1(9).BackColor
        .LegendbackColor = Label1(10).BackColor
        .InfoForeColor = Label1(11).BackColor
        .InfoBackColor = Label1(12).BackColor
        .MinorGridColor = Label1(13).BackColor
        .MajorGridColor = Label1(14).BackColor
        .SelectedColor = Label1(15).BackColor
        .MarkerColor = Label1(16).BackColor
        .MarkerLabelColor = Label1(7).BackColor
        .QuadrantDividerColor = Label1(8).BackColor
        .InfoQuadrantforeColor = Label1(17).BackColor
        .InfoQuadrantBackColor = Label1(18).BackColor
        If Option1(0).Value = True Then
            .MenuType = xcPopUpMenu
        Else
            .MenuType = xcButtonMenu
        End If
        .PictureTile = IIf((Check1(1).Value = vbChecked), True, False)
        
        .markerWidth = CInt(txtMarkerWidth.Text)
        .MinY = CDbl(txtMinY.Text)
        .MaxY = CDbl(txtMaxY.Text)
        .Minx = CDbl(txtMinX.Text)
        .Maxx = CDbl(txtMaxX.Text)
        PrepareData
        If Check1(2).Value = vbUnchecked Then
            Set .Picture = Nothing
        Else
            Set .Picture = LoadPicture(App.Path & "\stonehng.jpg")
        End If
        .MarkerSymbol = cboType.ItemData(cboType.ListIndex)
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

Private Sub XQuadrant1_ItemClick(cItem As ActiveQuadrant.ChartItem)
    grd.SelectionMode = flexSelectionByRow
    grd.Row = cItem.ItemID
    grd.ColSel = 2
End Sub

Private Sub Form_Load()
    
    PrepareData
        
    With cboType
        .Clear
        .AddItem "Box":             .ItemData(.NewIndex) = xcMarkerSymBox
        .AddItem "Circle":          .ItemData(.NewIndex) = xcMarkerSymCircle
        .AddItem "Triangle":        .ItemData(.NewIndex) = xcMarkerSymTriangle
        .AddItem "Trapezium":       .ItemData(.NewIndex) = xcMarkerSymTrapezium
        .AddItem "Rhombus":         .ItemData(.NewIndex) = xcMarkerSymRhombus
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
'    XQuadrant1.Width = Me.ScaleWidth

'    grd.ColWidth(0) = 960
'    grd.ColWidth(1) = Me.ScaleWidth - 960 - 2025
'    grd.ColWidth(2) = 2025
End Sub

Private Sub grd_Click()
    DoEvents
End Sub



Private Sub XQuadrant1_MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
    
    MsgBox "You clicked " & intMenuItemIndex & ":" & stgMenuItemCaption, _
            vbOKOnly, "CustomMenuItem"

    If intMenuItemIndex = 2 Then
        XQuadrant1.PrintLegend
    End If

End Sub


