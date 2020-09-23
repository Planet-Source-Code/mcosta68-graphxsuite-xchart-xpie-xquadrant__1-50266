VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\AXpie.vbp"
Begin VB.Form frmTest 
   Caption         =   "ActivePie Test"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply (Static)"
      Height          =   345
      Left            =   2010
      TabIndex        =   23
      Top             =   7830
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hot tracking"
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   22
      Top             =   7560
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Group segments"
      Height          =   195
      Index           =   0
      Left            =   1980
      TabIndex        =   20
      Top             =   7290
      Width           =   1845
   End
   Begin ActivePie.XPie XPie1 
      Height          =   4665
      Left            =   90
      TabIndex        =   19
      Top             =   90
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   8229
      InfoItems       =   ""
      SelectedColor   =   32768
      uSelectable     =   -1  'True
      uSelected       =   -1
      MarginTop       =   750
      MarginBottom    =   300
      MarginLeft      =   300
      MarginRight     =   300
      uChartTitle     =   "XPie"
      uChartSubTitle  =   "Test of XPie OCX"
      uDisplayDescript=   -1  'True
      BackColor       =   16777152
      ForeColor       =   -2147483640
      LegendBackColor =   12632064
      LegendForeColor =   16777215
      InfoBackColor   =   8454016
      InfoForeColor   =   16384
      InfoPieBackColor=   0
      InfoPieForeColor=   0
      ChartTitleColor =   -2147483640
      ChartSubTitleColor=   -2147483640
      GroupExplodeBackColor=   8454143
      GroupExplodeForeColor=   8388608
      GroupExplodeTitleColor=   128
      GroupExplodeMenuItems=   "&OK|&Print|Others...|Item&1|Item&2"
      MenuType        =   0
      MenuItems       =   "&Save as...|&Print|&Copy|Selection &information|&Legend|&Hide"
      CustomMenuItems =   ""
      InfoItems       =   ""
      SaveAsCaption   =   ""
      AutoRedraw      =   -1  'True
      PictureTile     =   0   'False
      MarkerColor     =   16384
      PieBorderColor  =   4210688
      DataFormat      =   ""
      PrinterFit      =   1
      PrinterOrientation=   2
      LegendCaption   =   "Display legend"
      HotTracking     =   -1  'True
      LegendPrintMode =   1
      GroupSegment    =   0   'False
      GroupExplodeOnClick=   -1  'True
      GroupExplodeAllowCommands=   -1  'True
   End
   Begin VB.ComboBox cboPrtMode 
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   7350
      Width           =   1125
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show back picture"
      Height          =   195
      Index           =   2
      Left            =   1980
      TabIndex        =   14
      Top             =   7020
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tile back picture"
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   11
      Top             =   6750
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Buttons menu"
      Height          =   225
      Index           =   1
      Left            =   1980
      TabIndex        =   10
      Top             =   6450
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Popup menu"
      Height          =   225
      Index           =   0
      Left            =   1980
      TabIndex        =   9
      Top             =   6150
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply (Random)"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   7830
      Width           =   1815
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
      Left            =   3990
      TabIndex        =   0
      Top             =   4890
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   5768
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pie border color"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   21
      Top             =   5820
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Printing mode"
      Height          =   390
      Left            =   60
      TabIndex        =   12
      Top             =   7350
      Width           =   660
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info Pie backcolor"
      Height          =   255
      Index           =   18
      Left            =   60
      TabIndex        =   18
      Top             =   7020
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info Pie forecolor"
      Height          =   255
      Index           =   17
      Left            =   60
      TabIndex        =   17
      Top             =   6720
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected color"
      Height          =   255
      Index           =   15
      Left            =   60
      TabIndex        =   16
      Top             =   6420
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marker color"
      Height          =   255
      Index           =   16
      Left            =   60
      TabIndex        =   15
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info background color"
      Height          =   255
      Index           =   12
      Left            =   1950
      TabIndex        =   8
      Top             =   5520
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info foreground color"
      Height          =   255
      Index           =   11
      Left            =   1950
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtitle color"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   4920
      Width           =   1845
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PrepareData(intMode As Integer)
    
    Dim X As Integer
    Dim intSign As Integer
    Dim varMonths As Variant
    Dim varMonthsExt As Variant
    Dim oChartItem As ActivePie.PieSegment
    
    varMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    varMonthsExt = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    grd.Rows = 1
    With XPie1
        .Clear
        .CustomMenuItems = "&Hello|&World|Print legend text"
        If intMode = 0 Then
            '-----------------------------------------------------------------------------
            'these settings are useful to see how percentage is calculated and distributed
            'onto single segments and groups
            For X = 1 To 8
                With oChartItem
                    .Value = 1
                    .ItemID = X
                    .Description = X
                    .SelectedDescription = X
                    .LegendDescription = X
                    .Color = vbYellow
                    .Group = "1st"
                    .GroupColor = vbYellow
                End With
                .AddItem oChartItem
    
                grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
                grd.RowData(grd.Rows - 1) = oChartItem.GroupColor
            Next
            For X = 1 To 8
                With oChartItem
                    .Value = 1
                    .ItemID = X
                    .Description = X
                    .SelectedDescription = X
                    .LegendDescription = X
                    .Color = vbRed
                    .Group = "2nd"
                    .GroupColor = vbRed
                End With
                .AddItem oChartItem
    
                grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
                grd.RowData(grd.Rows - 1) = oChartItem.GroupColor
            Next
            For X = 1 To 3
                With oChartItem
                    .Value = 1
                    .ItemID = X
                    .Description = X
                    .SelectedDescription = X
                    .LegendDescription = X
                    .Color = vbBlue
                    .Group = "3rd"
                    .GroupColor = vbBlue
                End With
                .AddItem oChartItem
    
                grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
                grd.RowData(grd.Rows - 1) = oChartItem.GroupColor
            Next
            '-----------------------------------------------------------------------------
        Else
            Randomize
            For X = 1 To UBound(varMonths) + 1
                With oChartItem
                    .Value = CDbl(Rnd(UBound(varMonths) + 1) * 1000)
                    .Value = Round(.Value, 2)
                    .ItemID = X
                    .Description = varMonths(X - 1)
                    .SelectedDescription = "Month " & varMonthsExt(X - 1)
                    .LegendDescription = "Month " & varMonthsExt(X - 1)
                    If Check1(0).Value = vbUnchecked Then
                        .Group = ""
                    End If
                    If X >= 1 And X <= 3 Then
                        .Color = vbBlue
                        .Group = "1st quarter"
                        .GroupColor = vbBlue
                    ElseIf X >= 4 And X <= 6 Then
                        .Color = vbGreen
                        .Group = "2nd quarter"
                        .GroupColor = vbGreen
                    ElseIf X >= 7 And X <= 9 Then
                        .Color = vbRed
                        .Group = "3rd quarter"
                        .GroupColor = vbRed
                    Else
                        .Color = vbYellow
                        .Group = "4th quarter"
                        .GroupColor = vbYellow
                    End If
                End With
                .AddItem oChartItem
        
                grd.AddItem X & vbTab & oChartItem.SelectedDescription & vbTab & oChartItem.Value
                grd.RowData(grd.Rows - 1) = oChartItem.GroupColor
            Next X
        End If
        .DrawChart
    End With
    grd.RowSel = 0

End Sub

Private Sub RefreshData()
    
    Dim intIdx As Integer
    
    With XPie1
        Label1(0).BackColor = .BackColor
        Label1(1).BackColor = .ChartTitleColor
        Label1(2).BackColor = .ChartSubTitleColor
        Label1(3).BackColor = .pieborderColor
        Label1(9).BackColor = .LegendForeColor
        Label1(10).BackColor = .LegendbackColor
        Label1(11).BackColor = .InfoForeColor
        Label1(12).BackColor = .InfoBackColor
        Label1(15).BackColor = .SelectedColor
        Label1(16).BackColor = .MarkerColor
        Label1(17).BackColor = .InfoPieforeColor
        Label1(18).BackColor = .InfoPieBackColor
        Option1(.MenuType).Value = True
        Check1(1).Value = IIf((.PictureTile = True), vbChecked, vbUnchecked)
        Check1(0).Value = IIf((.GroupSegment = True), vbChecked, vbUnchecked)
        Check1(3).Value = IIf((.HotTRacking = True), vbChecked, vbUnchecked)
        For intIdx = 0 To cboPrtMode.ListCount - 1
            If cboPrtMode.ItemData(intIdx) = .PrinterFit Then
                cboPrtMode.ListIndex = intIdx
                Exit For
            End If
        Next
    End With
    
End Sub

Private Sub Check1_Click(Index As Integer)

'    Dim lngRow As Long
'    Dim lngCol As Long
'
'    If Index = 0 Then
'        With XPie1
'            .GroupSegment = IIf(Check1(0).Value = vbChecked, True, False)
'            .SelectedColumn = -1
'        End With
'        For lngRow = grd.FixedRows To grd.Rows - 1
'            grd.Row = lngRow
'            For lngCol = grd.FixedCols To grd.Cols - 1
'                grd.Col = lngCol
'                grd.CellBackColor = vbWhite
'            Next
'        Next
'    End If

End Sub

Private Sub Command1_Click()

    With XPie1
        .AutoRedraw = False
        .BackColor = Label1(0).BackColor
        .ChartTitleColor = Label1(1).BackColor
        .ChartSubTitleColor = Label1(2).BackColor
        .pieborderColor = Label1(3).BackColor
        .LegendForeColor = Label1(9).BackColor
        .LegendbackColor = Label1(10).BackColor
        .InfoForeColor = Label1(11).BackColor
        .InfoBackColor = Label1(12).BackColor
        .SelectedColor = Label1(15).BackColor
        .MarkerColor = Label1(16).BackColor
        .InfoPieforeColor = Label1(17).BackColor
        .InfoPieBackColor = Label1(18).BackColor
        If Option1(0).Value = True Then
            .MenuType = xcPopUpMenu
        Else
            .MenuType = xcButtonMenu
        End If
        .PictureTile = IIf((Check1(1).Value = vbChecked), True, False)
        
        PrepareData (1)
        If Check1(2).Value = vbUnchecked Then
            Set .Picture = Nothing
        Else
            Set .Picture = LoadPicture(App.Path & "\STONEHNG.JPG")
        End If
        .GroupSegment = IIf(Check1(0).Value = vbUnchecked, False, True)
        .HotTRacking = IIf(Check1(3).Value = vbUnchecked, False, True)
        .PrinterFit = cboPrtMode.ItemData(cboPrtMode.ListIndex)
        .SelectedColumn = -1
        .AutoRedraw = True
    End With
    RefreshData

End Sub

Private Sub Command2_Click()
    
    With XPie1
        .AutoRedraw = False
        .BackColor = Label1(0).BackColor
        .ChartTitleColor = Label1(1).BackColor
        .ChartSubTitleColor = Label1(2).BackColor
        .pieborderColor = Label1(3).BackColor
        .LegendForeColor = Label1(9).BackColor
        .LegendbackColor = Label1(10).BackColor
        .InfoForeColor = Label1(11).BackColor
        .InfoBackColor = Label1(12).BackColor
        .SelectedColor = Label1(15).BackColor
        .MarkerColor = Label1(16).BackColor
        .InfoPieforeColor = Label1(17).BackColor
        .InfoPieBackColor = Label1(18).BackColor
        If Option1(0).Value = True Then
            .MenuType = xcPopUpMenu
        Else
            .MenuType = xcButtonMenu
        End If
        .PictureTile = IIf((Check1(1).Value = vbChecked), True, False)
        
        PrepareData (0)
        If Check1(2).Value = vbUnchecked Then
            Set .Picture = Nothing
        Else
            Set .Picture = LoadPicture(App.Path & "\STONEHNG.JPG")
        End If
        .GroupSegment = IIf(Check1(0).Value = vbUnchecked, False, True)
        .HotTRacking = IIf(Check1(3).Value = vbUnchecked, False, True)
        .PrinterFit = cboPrtMode.ItemData(cboPrtMode.ListIndex)
        .SelectedColumn = -1
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

Private Sub XPie1_GroupMenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)

    MsgBox "You clicked in the group explosion form the item #" & intMenuItemIndex & ":" & stgMenuItemCaption, _
            vbOKOnly, "GroupMenuItem"

End Sub

Private Sub XPie1_ItemClick(cItem As ActivePie.PieSegment)
        
    Dim lngRow As Long
    Dim lngCol As Long
    
    For lngRow = grd.FixedRows To grd.Rows - 1
        grd.Row = lngRow
        For lngCol = grd.FixedCols To grd.Cols - 1
            grd.Col = lngCol
            If lngRow = cItem.ItemID Then
                grd.CellBackColor = &H8000000D
                grd.CellForeColor = vbWhite
            Else
                grd.CellBackColor = vbWhite
                grd.CellForeColor = vbBlack
            End If
        Next
    Next
    
End Sub

Private Sub Form_Load()
    
    PrepareData (0)
        
    With cboPrtMode
        .Clear
        .AddItem "Stretched":      .ItemData(.NewIndex) = PrinterFitConstants.prtFitStretched
        .AddItem "Centered":       .ItemData(.NewIndex) = PrinterFitConstants.prtFitCentered
        .AddItem "TopLeft":        .ItemData(.NewIndex) = PrinterFitConstants.prtFitTopLeft
        .AddItem "TopRight":       .ItemData(.NewIndex) = PrinterFitConstants.prtFitTopRight
        .AddItem "BottomLeft":     .ItemData(.NewIndex) = PrinterFitConstants.prtFitBottomLeft
        .AddItem "BottomRight":    .ItemData(.NewIndex) = PrinterFitConstants.prtFitBottomRight
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
'    XPie1.Width = Me.ScaleWidth

'    grd.ColWidth(0) = 960
'    grd.ColWidth(1) = Me.ScaleWidth - 960 - 2025
'    grd.ColWidth(2) = 2025
End Sub

Private Sub grd_Click()
    DoEvents
End Sub



Private Sub XPie1_ItemGroupClick(cItem As ActivePie.PieGroup)
    
    Dim lngRow As Long
    Dim lngCol As Long
    
    For lngRow = grd.FixedRows To grd.Rows - 1
        grd.Row = lngRow
        For lngCol = grd.FixedCols To grd.Cols - 1
            grd.Col = lngCol
            If grd.RowData(lngRow) = cItem.Color Then
                grd.CellBackColor = &H8000000D
                grd.CellForeColor = vbWhite
            Else
                grd.CellBackColor = vbWhite
                grd.CellForeColor = vbBlack
            End If
        Next
    Next

End Sub

Private Sub XPie1_MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
    
    MsgBox "You clicked the custom menu item #" & intMenuItemIndex & ":" & stgMenuItemCaption, _
            vbOKOnly, "CustomMenuItem"

    If intMenuItemIndex = 2 Then
        XPie1.PrintLegend
    End If

End Sub


