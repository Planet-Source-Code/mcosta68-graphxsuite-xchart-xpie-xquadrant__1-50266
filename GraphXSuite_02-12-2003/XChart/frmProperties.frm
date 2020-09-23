VERSION 5.00
Begin VB.Form frmProperties 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chart properties"
   ClientHeight    =   3180
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCmd 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Index           =   1
      Left            =   6960
      TabIndex        =   32
      Top             =   2700
      Width           =   885
   End
   Begin VB.Frame fraColors 
      Caption         =   " Colors "
      Height          =   2985
      Left            =   2280
      TabIndex        =   17
      Top             =   90
      Width           =   4485
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Background color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   30
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Title color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   29
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Subtitle color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   28
         Top             =   900
         Width           =   915
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "X axis label color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   27
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "X axis items color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   26
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Y axis label color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   25
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Y axis items color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   24
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Legend foreground color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   2070
         TabIndex        =   23
         Top             =   300
         Width           =   1740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Legend background color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   2070
         TabIndex        =   22
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Info background color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   2070
         TabIndex        =   21
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Info foreground color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   2070
         TabIndex        =   20
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Minor grid color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   2070
         TabIndex        =   19
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Major grid color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   2070
         TabIndex        =   18
         Top             =   1800
         Width           =   1080
      End
   End
   Begin VB.Frame fraPicture 
      Caption         =   " Picture "
      Height          =   1155
      Left            =   6840
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
      Begin VB.CheckBox chkData 
         Caption         =   "Tile back picture"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   300
         Width           =   1815
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Show picture"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   570
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame fraType 
      Caption         =   " Chart "
      Height          =   2985
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   2115
      Begin VB.CheckBox chkData 
         Caption         =   "Hot tracking"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   31
         Top             =   2670
         Width           =   1815
      End
      Begin VB.TextBox txtBarPerc 
         Height          =   285
         Left            =   1290
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2000
         Width           =   705
      End
      Begin VB.TextBox txtLineWidth 
         Height          =   285
         Left            =   1290
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2300
         Width           =   705
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bar width (%)"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   2000
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Line width"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   2300
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Selected bar color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   9
         Top             =   1100
         Width           =   1290
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Bar color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   8
         Top             =   800
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Line color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   7
         Top             =   1400
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Symbol color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   6
         Top             =   1700
         Width           =   900
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   " Menu "
      Height          =   1155
      Left            =   6840
      TabIndex        =   1
      Top             =   90
      Width           =   2055
      Begin VB.OptionButton optData 
         Caption         =   "Popup menu"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Buttons menu"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   570
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   7950
      TabIndex        =   0
      Top             =   2700
      Width           =   885
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub RefreshData()
    
    Dim intIdx As Integer
    
    Label1(0).BackColor = uBackColor
    Label1(1).BackColor = uChartTitleColor
    Label1(2).BackColor = uChartSubTitleColor
    Label1(3).BackColor = uAxisLabelXColor
    Label1(4).BackColor = uAxisItemsXColor
    Label1(5).BackColor = uAxisLabelYColor
    Label1(6).BackColor = uAxisItemsYColor
    Label1(7).BackColor = uSelectedBarColor
    Label1(8).BackColor = uBarColor
    Label1(9).BackColor = uLegendForeColor
    Label1(10).BackColor = uLegendBackColor
    Label1(11).BackColor = uInfoForeColor
    Label1(12).BackColor = uInfoBackColor
    Label1(13).BackColor = uMinorGridColor
    Label1(14).BackColor = uMajorGridColor
    Label1(15).BackColor = uLineColor
    Label1(16).BackColor = uBarSymbolColor
    optData(uMenuType).Value = True
    chkData(0).Value = IIf((uHotTracking = True), vbChecked, vbUnchecked)
    chkData(1).Value = IIf((uPictureTile = True), vbChecked, vbUnchecked)
    txtBarPerc.Text = CStr(uBarWidthPercentage)
    txtLineWidth.Text = CStr(uLineWidth)
'    txtSymbol.Text = uBarSymbol
    For intIdx = 0 To cboType.ListCount - 1
        If cboType.ItemData(intIdx) = uChartType Then
            cboType.ListIndex = intIdx
            Exit For
        End If
    Next

End Sub



Private Sub cmdCmd_Click(Index As Integer)

    Unload Me

End Sub

Private Sub Form_Load()

    cboType.Clear
    cboType.AddItem "Bar"
    cboType.ItemData(cboType.NewIndex) = xcBar
    cboType.AddItem "Symbol"
    cboType.ItemData(cboType.NewIndex) = xcSymbol
    cboType.AddItem "Line"
    cboType.ItemData(cboType.NewIndex) = xcLine
    cboType.AddItem "BarLine"
    cboType.ItemData(cboType.NewIndex) = xcBarLine
    cboType.AddItem "SymbolLine"
    cboType.ItemData(cboType.NewIndex) = xcSymbolLine
    cboType.AddItem "Oval"
    cboType.ItemData(cboType.NewIndex) = xcOval
    cboType.AddItem "OvalLine"
    cboType.ItemData(cboType.NewIndex) = xcOvalLine
    RefreshData
    
End Sub

Private Sub Label1_Click(Index As Integer)
    
'    dlgColor.Color = Label1(Index).BackColor
'    dlgColor.ShowColor
'    If dlgColor.Color <> Label1(Index).BackColor Then
'        Label1(Index).BackColor = dlgColor.Color
'    End If

End Sub


