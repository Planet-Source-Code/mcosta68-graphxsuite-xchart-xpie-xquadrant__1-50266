VERSION 5.00
Begin VB.UserControl XPie 
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
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   ToolboxBitmap   =   "XPie.ctx":0000
   Begin VB.PictureBox picCommands 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   1605
      TabIndex        =   10
      Top             =   90
      Width           =   1605
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   2
         Left            =   660
         Picture         =   "XPie.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   4
         Left            =   1290
         Picture         =   "XPie.ctx":069C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   1
         Left            =   330
         Picture         =   "XPie.ctx":0C26
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   3
         Left            =   975
         Picture         =   "XPie.ctx":11B0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCmd 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "XPie.ctx":173A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   315
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "XPie.ctx":1CC4
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   1
         Left            =   660
         Picture         =   "XPie.ctx":1E0E
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   2
         Left            =   930
         Picture         =   "XPie.ctx":1F58
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "XPie.ctx":20A2
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgCmd 
         Height          =   240
         Index           =   4
         Left            =   1470
         Picture         =   "XPie.ctx":21EC
         Top             =   585
         Width           =   240
      End
   End
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
      TabIndex        =   9
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
      TabIndex        =   6
      Top             =   2460
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picInfoPie 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   900
      ScaleHeight     =   405
      ScaleWidth      =   375
      TabIndex        =   5
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
      ScaleHeight     =   5415
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   0
      Width           =   75
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
            TabIndex        =   8
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
            TabIndex        =   7
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
Attribute VB_Name = "XPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Const PI As Double = 3.14159265358979

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

Private uPieBorderColor   As Long         'the color of the pie border
Private uMarkerColor      As Long         'the color of the symbol
Private uLegendBackColor  As Long         'the legend background color
Private uLegendForeColor  As Long         'the legend foreground color
Private uInfoBackColor    As Long         'the information picBox background color
Private uInfoForeColor    As Long         'the information picBox foreground color
Private uChartTitleColor  As Long         'the chart title color
Private uChartSubTitleColor As Long       'the chart subtitle color
Private uSelected         As Integer      'marker indicating the selected column
Private uSelectable       As Boolean      'marker indicating whether user can select a column
Private uSelectedColor    As Long         'the selected marker forecolor
Private uSaveAsCaption    As String       'the SaveAs dialog picBox caption
Private uOldSelection     As Long
Private uGroupSegment     As Boolean      'indicates if the pie's segments must be grouped
                                          '(Group property of PieSegment type) or not
Private uGroupExplode           As Boolean 'indicates if the group must be exploded (meaningful only if uGroupSegment=True)
Private uGroupExplodeBackColor  As Long   'the group explosion form's backcolor
Private uGroupExplodeForeColor  As Long   'the group explosion form's forecolor
Private uGroupExplodeTitleColor As Long   'the group explosion form's title color
Private uGroupExplodeMenuItems  As String 'the menu's items in the group explosion form
Private uGroupExplodeAllowCommands As Boolean 'indicates if the command buttons in the group form are allowed (meaningful only if uGroupExplode=True)
Private Const GROUP_EXPLODE_MENU_ITEMS = "&OK|&Print"

Private uInfoItems        As String       'the information items (to be displayed in the info picBox)
Private Const INFO_ITEMS = "Value|Description"

Private uInfoPieBackColor       As Long    'the group information picBox background color
Private uInfoPieForeColor       As Long    'the group information picBox foreground color

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

Private Const IDX_SAVE = 0                'the command buttons' indexs
Private Const IDX_PRINT = 1
Private Const IDX_COPY = 2
Private Const IDX_INFO = 3
Private Const IDX_LEGEND = 4

Private uTopMargin        As Single       '--------------------------------------
Private uBottomMargin     As Single       'margins used around the pie content
Private uLeftMargin       As Single       '
Private uRightMargin      As Single
Private uRightMarginOrg   As Single       '--------------------------------------
Private uDisplayDescript  As Boolean      'display description when selectable
Private uChartTitle       As String       'chart title
Private uChartSubTitle    As String       'chart sub title
Private uHotTracking      As Boolean      'marker indicating use of hot tracking

Private offsetX           As Long
Private offsetY           As Long
Private sngRadius         As Single
Private sngPieXCenter     As Single
Private sngPieYCenter     As Single
Private dblPieTotal       As Double

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean
Private bResizeLegend     As Boolean

Private bProcessingOver   As Boolean      'marker to speed up mouse over effects

Public Type CoordConv
    X As Single
    Y As Single
    Angle As Double
    Radius As Double
End Type

Public Type PieSegment
    ItemID As String
    Group As String
    GroupColor As Long
    SelectedDescription As String
    LegendDescription As String
    Description As String
    Value As Double
    Color As Long
End Type
Private cItems As Collection   'collection of chart items

Public Type PieSegmentAttributes
    AngleFrom As Single
    AngleTo As Single
    Percentage As Double
End Type
Private cItemsAttributes As Collection

Public Type PieGroup
    Value As Double
    Name As String
    Color As Long
End Type
Private cGroups As Collection   'collection of pie groups

Public Type PieGroupAttributes
    AngleFrom As Single
    AngleTo As Single
    Percentage As Double
End Type
Private cGroupsAttributes As Collection

Public Event ItemClick(cItem As PieSegment)
Public Event ItemGroupClick(cItem As PieGroup)
Public Event MenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
Public Event GroupMenuItemClick(intMenuItemIndex As Integer, stgMenuItemCaption As String)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'-----------------------------------------------
' for Ballon ToolTip
'-----------------------------------------------
Private ttpBalloon As New Tooltip

Private Sub HandleSegments()
    
    Dim sngAngle As Single
    Dim intCount As Integer
    Dim dblPercentage As Double
    Dim dblSegmentTotal As Double
    Dim dblPercentageTotal As Double
    Dim dblPercentageRemainder As Double
    Dim oSAttr As PieSegmentAttributes
    
    On Error GoTo errHandle
    
    'define segments' attributes
    dblSegmentTotal = 100
    dblPercentageTotal = 0
    dblPercentageRemainder = 0
    Set cItemsAttributes = Nothing
    Set cItemsAttributes = New Collection
    With cItems
        For intCount = 1 To .Count
            'total segment percentage is decreased only if we are after the first segment
            If (intCount > 1) Then
                oSAttr.AngleFrom = cItemsAttributes(intCount - 1).AngleTo
                dblSegmentTotal = dblSegmentTotal - cItemsAttributes(intCount - 1).Percentage
            Else
                oSAttr.AngleFrom = 0
            End If
            'calculate percentage/angle and add to the proper collection
            If intCount = .Count Then
                'last item's percentage corresponds to the remainder (100-assigned)
                dblPercentage = 100 - dblPercentageTotal
            Else
                'item by item, add the previous remainder value
                '(difference between double value and rounded value)
                dblPercentage = ((.Item(intCount).Value / dblPieTotal) * 100) + dblPercentageRemainder
                If (dblPercentage - Round(dblPercentage, 2)) > 0 Then
                    'the remainder is saved only in case the value was not rounded by excess
                    dblPercentageRemainder = dblPercentage - Round(dblPercentage, 2)
                Else
                    dblPercentageRemainder = 0
                End If
                'increment total percentage
                dblPercentageTotal = dblPercentageTotal + dblPercentage
            End If
            sngAngle = 360 * dblPercentage / 100
            oSAttr.AngleTo = oSAttr.AngleFrom + sngAngle
            oSAttr.Percentage = Round(dblPercentage, 2)
            cItemsAttributes.Add oSAttr
        Next intCount
    End With
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Function Radians(ByVal Degrees As Single) As Double
    Radians = CDbl(Degrees) * PI / 180
End Function

Private Function ConvertCoordinates(X As Single, Y As Single) As CoordConv
    
    Dim sngY1 As Single
    Dim sngX1 As Single
    Dim sngAngle As Single
    Dim oConv As CoordConv
    
    sngX1 = X / Screen.TwipsPerPixelX
    sngY1 = Y / Screen.TwipsPerPixelY
    sngX1 = (sngPieXCenter - sngX1)     'calculate in respect to center
    sngY1 = (sngPieYCenter - sngY1)     'calculate in respect to center
    sngAngle = CSng(Atn(sngY1 / IIf(sngX1 = 0, 0.00000001, sngX1))) * 180 / PI
    'these calculations are valid in case the graduation is as follows:
    '
    '              0
    '              |
    '              |
    ' 90-----------+----------270
    '              |
    '              |
    '             180
    '
    If (sngX1 < 0) And (sngY1 >= 0) Then
        sngAngle = 270 + Abs(sngAngle)
    ElseIf (sngX1 >= 0) And (sngY1 >= 0) Then
        sngAngle = 90 - sngAngle
    ElseIf (sngX1 >= 0) And (sngY1 < 0) Then
        sngAngle = 90 + Abs(sngAngle)
    ElseIf (sngX1 < 0) And (sngY1 < 0) Then
        sngAngle = 270 - Abs(sngAngle)
    End If
        
'    'these calculations are valid in case the graduation is as follows:
'    '
'    '              90
'    '              |
'    '              |
'    ' 180----------+----------0
'    '              |
'    '              |
'    '             270
'    '
'    If (sngX1 < 0) And (sngY1 >= 0) Then
'        sngAngle = Abs(sngAngle)
'    ElseIf (sngX1 >= 0) And (sngY1 >= 0) Then
'        sngAngle = 180 - sngAngle
'    ElseIf (sngX1 >= 0) And (sngY1 < 0) Then
'        sngAngle = 180 - sngAngle
'    ElseIf (sngX1 < 0) And (sngY1 < 0) Then
'        sngAngle = 360 - sngAngle
'    End If

    With oConv
        .X = sngX1
        .Y = sngY1
        .Angle = sngAngle
        .Radius = Sqr(sngX1 * sngX1 + sngY1 * sngY1)
    End With
    
    ConvertCoordinates = oConv

End Function

Private Sub DrawSegmentText(sngAngleFrom As Single, sngAngleTo As Single, intItemIdx As Integer)
    
    On Error GoTo errHandle

    Dim lngW As Long
    Dim lngH As Long
    Dim sngX1 As Single
    Dim sngY1 As Single
    Dim lngColor As Long
    Dim stgDesc As String
    Dim sngXoff As Single
    Dim sngYoff As Single
    Dim sngAngle As Single
    Dim sngAngleRad As Single

    sngAngle = sngAngleFrom + (sngAngleTo - sngAngleFrom) / 2
    sngAngleRad = Radians(sngAngle)
    sngXoff = Sin(sngAngleRad) * sngRadius
    sngYoff = Cos(sngAngleRad) * sngRadius

    With UserControl
        lngColor = .ForeColor
        If uGroupSegment = False Then
            stgDesc = cItems(intItemIdx).Description
            .ForeColor = IIf(((intItemIdx - 1) = uSelected And uSelectable), uSelectedColor, uMarkerColor)
        Else
            stgDesc = cGroups(intItemIdx).Name
            .ForeColor = IIf(((intItemIdx - 1) = uSelected And uSelectable), uSelectedColor, uMarkerColor)
        End If
    
        sngX1 = sngPieXCenter - sngXoff
        sngY1 = sngPieYCenter - sngYoff
        '
        '              0
        '              |
        '              |
        ' 90-----------+----------270
        '              |
        '              |
        '             180
        '
        lngW = .TextWidth(stgDesc) / Screen.TwipsPerPixelX / 2
        lngH = .TextHeight(stgDesc) / Screen.TwipsPerPixelY / 2
        If sngAngle >= 0 And sngAngle < 90 Then
            sngX1 = sngX1 - lngW
            sngY1 = sngY1 - lngH
        ElseIf sngAngle >= 90 And sngAngle < 180 Then
            sngX1 = sngX1 - lngW
            sngY1 = sngY1 + lngH
        ElseIf sngAngle >= 180 And sngAngle < 270 Then
            sngX1 = sngX1 + lngW
            sngY1 = sngY1 + lngH
        Else
            sngX1 = sngX1 + lngW
            sngY1 = sngY1 - lngH
        End If
        PrintRotText .hDC, stgDesc, sngX1, sngY1, 0
        .ForeColor = lngColor
    End With
    Exit Sub
    
errHandle:
    Exit Sub
    

End Sub

Private Sub GroupSegments()
    
    'important: before grouping segments, it has to be defined
    'segments' attributes!!!

    Dim intIdx As Integer
    Dim intCount As Integer
    Dim oGroup As PieGroup
    Dim stgGroup As String
    Dim stgGroups As String
    Dim stgColor As String
    Dim stgGroupsColor As String
    Dim dblSegmentTotal As Double
    Dim oGAttr As PieGroupAttributes
    Dim dblPercentage As Double
    Dim sngAngle As Single
    Dim dblPercentageTotal As Double
    Dim dblPercentageRemainder As Double
    
    'verify how may groups there are in segments
    dblSegmentTotal = 0
    stgGroups = Empty
    stgGroupsColor = Empty
    For intCount = 1 To cItems.Count
        stgGroup = cItems(intCount).Group
        If InStr(stgGroups, stgGroup) = 0 Then
            'add group name to the list
            stgGroups = stgGroups & stgGroup & Chr$(0)
            'and group color
            stgGroupsColor = stgGroupsColor & CStr(cItems(intCount).GroupColor) & Chr$(0)
        End If
    Next
    
    'set new groups' collection
    Set cGroups = Nothing
    Set cGroups = New Collection
    intIdx = 1
    While intIdx > 0
        stgGroup = TokenByPos(stgGroups, intIdx, Chr$(0))
        If stgGroup = "" Then
            'no more items in the groups' list
            intIdx = -1
        Else
            dblSegmentTotal = 0
            stgColor = TokenByPos(stgGroupsColor, intIdx, Chr$(0))
            For intCount = 1 To cItems.Count
                If cItems(intCount).Group = stgGroup Then
                    dblSegmentTotal = dblSegmentTotal + cItems(intCount).Value
                End If
            Next
            With oGroup
                .Value = dblSegmentTotal
                .Name = stgGroup
                .Color = CLng(stgColor)
            End With
            cGroups.Add oGroup
            intIdx = intIdx + 1
        End If
    Wend
    
    'groups' attributes
    dblSegmentTotal = 100
    Set cGroupsAttributes = Nothing
    Set cGroupsAttributes = New Collection
    dblPercentageTotal = 0
    dblPercentageRemainder = 0
    With cGroups
        For intCount = 1 To .Count
            'total segment percentage is decreased only if we are after the first segment
            If intCount > 1 Then
                oGAttr.AngleFrom = cGroupsAttributes(intCount - 1).AngleTo
                dblSegmentTotal = dblSegmentTotal - cGroupsAttributes(intCount - 1).Percentage
            Else
                oGAttr.AngleFrom = 0
            End If
            'calculate percentage, basing on percentage of single segments, belonging to the group
            dblPercentage = 0
            For intIdx = 1 To cItems.Count
                If cItems(intIdx).Group = .Item(intCount).Name Then
                    dblPercentage = dblPercentage + cItemsAttributes(intIdx).Percentage
                End If
            Next
            sngAngle = 360 * dblPercentage / 100
            oGAttr.AngleTo = oGAttr.AngleFrom + sngAngle
            oGAttr.Percentage = dblPercentage
            cGroupsAttributes.Add oGAttr
        Next intCount
    End With

End Sub

Public Property Let LegendPrintMode(val As LegendPrintConstants)
    uLegendPrintMode = val
    PropertyChanged "LegendPrintMode"
End Property

Public Property Get LegendPrintMode() As LegendPrintConstants
    LegendPrintMode = uLegendPrintMode
End Property

Private Function Tracking(X As Single, Y As Single) As Integer
    
    Dim gitem As PieGroup
    Dim oItem As PieSegment
    Dim intSelectedCol As Integer
    
    If Not bProcessingOver Then
        bProcessingOver = True
        intSelectedCol = InSegment(X, Y)
        If (intSelectedCol >= 0) And (intSelectedCol <> uOldSelection) Then
            uSelected = intSelectedCol
            DrawChart
            uOldSelection = uSelected
            If uGroupSegment = False Then
                oItem = cItems(uSelected + 1)
                RaiseEvent ItemClick(oItem)
            Else
                gitem = cGroups(uSelected + 1)
                RaiseEvent ItemGroupClick(gitem)
            End If
        End If
        bProcessingOver = False
    End If
    Tracking = intSelectedCol
    Exit Function
    
Tracking_error:
    Tracking = -1
    Exit Function
    
End Function

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


Public Function AddItem(cItem As PieSegment) As Boolean
    
    cItems.Add cItem
    
    dblPieTotal = dblPieTotal + cItem.Value
    Call HandleSegments
    Call GroupSegments
    
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

    Dim varItems As Variant
    Dim sDescription As String
    Dim dblPercentage As Double
    
    'it's important to let the info label invisible at beginning to avoid flickering effect
    lblInfo.Visible = False
    If uDisplayDescript Then
        If intIdx > -1 Then
            If uInfoItems = Empty Then uInfoItems = INFO_ITEMS
            varItems = Split(uInfoItems, "|")
            'this kind of error trapping is useful in case the user
            'did not define any item in the menu items string, so the default is used
            On Error GoTo DrawChart_error
            If uGroupSegment = False Then
                With cItems.Item(intIdx + 1)
                    sDescription = CStr(varItems(0)) & ": " & Format(.Value, uDataFormat)
                    If Len(.SelectedDescription) > 0 Then
                        sDescription = CStr(varItems(1)) & ": " & .SelectedDescription & vbCrLf & sDescription
                        dblPercentage = cItemsAttributes.Item(intIdx + 1).Percentage
                    End If
                End With
            Else
                With cGroups.Item(intIdx + 1)
                    sDescription = CStr(varItems(0)) & ": " & Format(.Value, uDataFormat)
                    If Len(.Name) > 0 Then
                        sDescription = CStr(varItems(1)) & ": " & .Name & vbCrLf & sDescription
                        dblPercentage = cGroupsAttributes.Item(intIdx + 1).Percentage
                    End If
                End With
            End If
        End If
        If sDescription <> Empty Then
            With lblInfo
                sDescription = sDescription & " (" & Format$(dblPercentage, "#0.00\%") & ")"
                .Caption = sDescription
                .Width = UserControl.TextWidth(sDescription) + 5 * Screen.TwipsPerPixelX
                .Height = UserControl.TextHeight(sDescription) * 1.2
                .Visible = True
            End With
        End If
    End If
    Exit Sub

DrawChart_error:
    uInfoItems = INFO_ITEMS
    Resume Next

End Sub

Private Sub DrawPicture(sngX1 As Single, sngX2 As Single, sngY1 As Single, sngY2 As Single, blnTile As Boolean, pic As StdPicture)

    On Error Resume Next
    
    Dim x1 As Single
    Dim X2 As Single
    Dim Y1 As Single
    Dim Y2 As Single
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
        Y1 = sngY1
        Y2 = sngY2
        X2 = sngX2
        Do While Y1 < Y2
            x1 = sngX1
            Do While x1 < X2
                If (x1 + sngW) > X2 Then
                    xTemp = (X2 - x1)
                Else
                    xTemp = sngW
                End If
                xTemp = IIf(xTemp < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, xTemp)
                If (Y1 + sngH) > Y2 Then
                    yTemp = (Y2 - Y1)
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
                            x1, Y1, _
                            xTemp, _
                            yTemp, _
                            0, 0, xTemp, yTemp
                x1 = (x1 + sngW)
            Loop
            Y1 = (Y1 + sngH)
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

Public Property Set Picture(ByVal picVal As StdPicture)
    Set uPicture = picVal
    DrawChart
End Property


Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed as background of the chart."
    Set Picture = uPicture
End Property

Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

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
        If lblInfo.Visible = False Then
            lblInfo.Move picCommands.Left + picCommands.ScaleWidth + 60, 60
        End If
    Else
        picCommands.Visible = False
        If lblInfo.Visible = False Then
            lblInfo.Move 60, 60
        End If
    End If
    If picInfoPie.Visible = False Then
        picInfoPie.Move lblInfo.Left, lblInfo.Top + lblInfo.Height + 60
    End If
    Exit Sub
    
FixMenu_error:
    uMenuItems = MENU_ITEMS
    Resume Next

End Sub
Private Sub FixGroupExplodeMenu()
    
    'this kind of error trapping is useful in case the user
    'did not define any item in the menu items string, so the default is used
    On Error GoTo FixMenu_error
    
    Dim stgData As String
    Dim intIdx As Integer
    Dim varItems As Variant
    
    If uGroupExplodeMenuItems = Empty Then
        uGroupExplodeMenuItems = GROUP_EXPLODE_MENU_ITEMS
    Else
        stgData = Empty
        varItems = Split(uGroupExplodeMenuItems, "|")
        
        If varItems(0) <> Empty Then
            stgData = CStr(varItems(0))
        Else
            stgData = "&OK"
        End If
        
        If varItems(1) <> Empty Then
            stgData = stgData & "|" & CStr(varItems(1))
        Else
            stgData = stgData & "|&Print"
        End If
        
        'new errore handling
        On Error Resume Next
        For intIdx = 2 To UBound(varItems)
            stgData = stgData & "|" & CStr(varItems(intIdx))
        Next
        uGroupExplodeMenuItems = stgData
    End If
    Exit Sub
    
FixMenu_error:
    uGroupExplodeMenuItems = GROUP_EXPLODE_MENU_ITEMS
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

Public Sub PrintPie()
    
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
    Dim oPieSegment As ActivePie.PieSegment
    
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
        If uGroupSegment = False Then
            For intIdx = 1 To cItems.Count
                With cItems(intIdx)
                    stg = .LegendDescription & " (" & Format(.Value, uDataFormat) & " - " & Format(cItemsAttributes(intIdx).Value, "#0.00\%") & ")"
                End With
                Printer.Print stg
            Next
        Else
            For intIdx = 1 To cGroups.Count
                With cGroups(intIdx)
                    stg = .Group & " (" & Format(.Value, uDataFormat) & " - " & Format(cGroupsAttributes(intIdx).Value, "#0.00\%") & ")"
                End With
                Printer.Print stg
            Next
        End If
        
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

Public Property Get ChartTitle() As String
Attribute ChartTitle.VB_Description = "Determines the title of the chart."
    ChartTitle = uChartTitle
End Property

Public Property Let MenuItems(stgVal As String)
    uMenuItems = stgVal
    FixMenu
    PropertyChanged "MenuItems"
End Property

Public Property Let GroupExplodeMenuItems(stgVal As String)
    
    'menu items to be accessed when exploding group
    'form is: item1|item2|[item3]|[item4]|....[itemN]
    ' mandatory:    item1 is for OK command button
    '               item2 is for PRINT command button
    ' optional:     item3 is for OTHERS command button
    '               item4...itemN are the others commands available (10 commands must be set)
    
    uGroupExplodeMenuItems = stgVal
    FixGroupExplodeMenu
    PropertyChanged "GroupExplodeMenuItems"

End Property


Public Property Let CustomMenuItems(stgVal As String)
    uCustomMenuItems = stgVal
    FixCustomMenu
    PropertyChanged "CustomMenuItems"
End Property


Public Property Let InfoItems(stgVal As String)
Attribute InfoItems.VB_Description = "Determines the string values displayed when selection information is enabled (separated by |)."
    uInfoItems = stgVal
    PropertyChanged "InfoItems"
End Property
Public Property Get InfoItems() As String
    InfoItems = uInfoItems
End Property
Public Property Get MenuItems() As String
Attribute MenuItems.VB_Description = "Determines the string values displayed when popup menu is enabled (separated by |)."
    MenuItems = uMenuItems
End Property
Public Property Get GroupExplodeMenuItems() As String
    GroupExplodeMenuItems = uGroupExplodeMenuItems
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

Public Property Let SelectionInformation(blnVal As Boolean)
Attribute SelectionInformation.VB_Description = "Determines if the information box about the selected bar must be visible or hidden."
    If blnVal <> uDisplayDescript Then
        uDisplayDescript = blnVal
        DrawChart
        PropertyChanged "SelectionInformation"
    End If
End Property
Public Property Let GroupSegment(blnVal As Boolean)
    uGroupSegment = blnVal
    DrawChart
    PropertyChanged "GroupSegment"
End Property

Public Property Let GroupExplodeOnClick(blnVal As Boolean)
    If blnVal <> uGroupExplode Then
        uGroupExplode = blnVal
        PropertyChanged "GroupExplode"
    End If
End Property
Public Property Let GroupExplodeAllowCommands(blnVal As Boolean)
    If blnVal <> uGroupExplodeAllowCommands Then
        uGroupExplodeAllowCommands = blnVal
        PropertyChanged "GroupExplodeAllowCommands"
    End If
End Property



Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
End Property

Public Property Let BackColor(lngVal As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns/sets the color of the chart background."
    If lngVal <> UserControl.BackColor Then
        UserControl.BackColor = lngVal
        DrawChart
        PropertyChanged "BackColor"
    End If
End Property
Public Property Let GroupExplodeBackColor(lngVal As OLE_COLOR)
    If lngVal <> uGroupExplodeBackColor Then
        uGroupExplodeBackColor = lngVal
        PropertyChanged "GroupExplodeBackColor"
    End If
End Property
Public Property Let GroupExplodeForeColor(lngVal As OLE_COLOR)
    If lngVal <> uGroupExplodeForeColor Then
        uGroupExplodeForeColor = lngVal
        PropertyChanged "GroupExplodeForeColor"
    End If
End Property
Public Property Let GroupExplodeTitleColor(lngVal As OLE_COLOR)
    If lngVal <> uGroupExplodeTitleColor Then
        uGroupExplodeTitleColor = lngVal
        PropertyChanged "GroupExplodeTitleColor"
    End If
End Property

Public Property Let PieBorderColor(lngVal As OLE_COLOR)
    If lngVal <> uPieBorderColor Then
        uPieBorderColor = lngVal
        DrawChart
        PropertyChanged "PieBorderColor"
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Get PieBorderColor() As OLE_COLOR
    PieBorderColor = uPieBorderColor
End Property


Public Property Get ChartTitleColor() As OLE_COLOR
Attribute ChartTitleColor.VB_Description = "Returns/sets the color used to display the chart title."
    ChartTitleColor = uChartTitleColor
End Property
Public Property Get GroupExplodeBackColor() As OLE_COLOR
    GroupExplodeBackColor = uGroupExplodeBackColor
End Property
Public Property Get GroupExplodeTitleColor() As OLE_COLOR
    GroupExplodeTitleColor = uGroupExplodeTitleColor
End Property

Public Property Get GroupExplodeForeColor() As OLE_COLOR
    GroupExplodeForeColor = uGroupExplodeForeColor
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
Public Property Let InfoPieBackColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoPieBackColor Then
        uInfoPieBackColor = lngVal
        DrawChart
        PropertyChanged "InfoPieBackColor"
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
Public Property Let InfoPieForeColor(lngVal As OLE_COLOR)
    If lngVal <> uInfoPieForeColor Then
        uInfoPieForeColor = lngVal
        DrawChart
        PropertyChanged "InfoPieForeColor"
    End If
End Property


Public Property Get InfoBackColor() As OLE_COLOR
    InfoBackColor = uInfoBackColor
End Property

Public Property Get InfoPieBackColor() As OLE_COLOR
    InfoPieBackColor = uInfoPieBackColor
End Property


Public Property Get InfoForeColor() As OLE_COLOR
    InfoForeColor = uInfoForeColor
End Property

Public Property Get InfoPieForeColor() As OLE_COLOR
    InfoPieForeColor = uInfoPieForeColor
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
            uSelected = Index
            uOldSelection = uSelected
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
    SelectedColumn = uSelected
End Property
Public Property Let SelectedColumn(lngColumn As Long)
    
    On Error Resume Next
    
    Dim oItem As PieSegment
    Dim gitem As PieGroup
    
    If lngColumn <> uSelected Then
        uSelected = lngColumn
        DrawChart
        PropertyChanged "SelectedColumn"
        
        If Err.Number Then
            uSelected = -1
        Else
            If uGroupSegment = False Then
                If uSelected >= 0 And uSelected <= cItems.Count - 1 Then
                    oItem = cItems(lngColumn + 1)
                    RaiseEvent ItemClick(oItem)
                End If
            Else
                If uSelected >= 0 And uSelected <= cGroups.Count - 1 Then
                    gitem = cGroups(lngColumn + 1)
                    RaiseEvent ItemGroupClick(gitem)
                End If
            End If
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
    
    Call PrintPie

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
        .sFile = "XPie.bmp" & Space$(1024) & vbNullChar & vbNullChar
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
    Call DisplayInfo(uSelected)
    
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
   
    If ttpBalloon.ObjName <> ("picDescription" & Index) Then
        If Right$(picDescription(Index).Tag, Len(TooltipNeeded)) = TooltipNeeded Then
            ttpBalloon.Title = ""
            ttpBalloon.TipText = picDescription(Index).Tag
            ttpBalloon.Icon = TTIconInfo
            ttpBalloon.Style = TTBalloon
            ttpBalloon.Centered = False
            Set ttpBalloon.ParentControl = picDescription(Index)
            ttpBalloon.ObjName = "picDescription" & Index
            ttpBalloon.Create
        End If
    End If

End Sub

Private Sub picInfoPie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        picInfoPie.Drag
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
    Set cGroups = New Collection
End Sub

Private Sub UserControl_InitProperties()
    
    Dim X As Integer
    
    uTopMargin = 50 * Screen.TwipsPerPixelY
    uBottomMargin = 55 * Screen.TwipsPerPixelY
    uLeftMargin = 55 * Screen.TwipsPerPixelX
    uRightMargin = 55 * Screen.TwipsPerPixelX
    uRightMarginOrg = uRightMargin
    uSelected = -1
    uChartTitle = UserControl.Name
    uChartSubTitle = ""
    UserControl.BackColor = vbWindowBackground
    UserControl.ForeColor = vbWindowText
    uLegendBackColor = UserControl.BackColor
    uLegendForeColor = UserControl.ForeColor
    uInfoBackColor = vbInfoBackground
    uInfoForeColor = vbInfoText
    uInfoPieBackColor = vbInfoBackground
    uInfoPieForeColor = vbInfoText
    uChartTitleColor = UserControl.ForeColor
    uChartSubTitleColor = UserControl.ForeColor
    uGroupExplodeBackColor = vbInfoBackground
    uGroupExplodeForeColor = vbInfoText
    uGroupExplodeTitleColor = vbInfoText
    uMenuType = xcPopUpMenu
    uGroupExplodeMenuItems = Empty
    uMenuItems = Empty
    uCustomMenuItems = Empty
    uInfoItems = Empty
    uSaveAsCaption = Empty
    uAutoRedraw = True
    Set uPicture = Nothing
    uPictureTile = False
    uDataFormat = Empty
    uPrinterFit = prtFitCentered
    uPrinterOrientation = vbPRORLandscape
    uLegendCaption = LEGEND_CAPTION
    uPieBorderColor = UserControl.ForeColor
    uSelectedColor = vbCyan
    uHotTracking = False
    uOldSelection = -1
    uLegendPrintMode = legPrintGraph
    uGroupSegment = False
    uGroupExplode = False
    uGroupExplodeAllowCommands = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo TrackExit
    
    Dim stgDesc As String
    Dim stgGroup As String
    Dim intCount As Integer
    Dim intSelected As Integer
    Dim dblPercentage As Double

    If Button = vbLeftButton Then
        intSelected = Tracking(X, Y)
        If intSelected >= 0 Then
            If (uGroupSegment = True) And (uGroupExplode) = True Then
                Call FixGroupExplodeMenu
                stgGroup = Trim(cGroups(intSelected + 1).Name)
                With frmGroupExplode
                    stgDesc = Empty
                    For intCount = 1 To cItems.Count
                        If Trim(cItems(intCount).Group) = stgGroup Then
                            dblPercentage = cItemsAttributes(intCount).Percentage
                            stgDesc = stgDesc & Trim(cItems(intCount).SelectedDescription) & " (" & Format$(dblPercentage, "#0.00\%") & ")" & Chr$(0)
                        End If
                    Next
                    .GroupList = stgDesc
                    .Title = Trim(cGroups.Item(intSelected + 1).Name)
                    .InfoBackColor = uGroupExplodeBackColor
                    .InfoForeColor = uGroupExplodeForeColor
                    .TitleColor = uGroupExplodeTitleColor
                    .CommandItems = uGroupExplodeMenuItems
                    .AllowCommands = uGroupExplodeAllowCommands
                    .Show vbModal
                    If .ItemClicked >= 0 Then
                        RaiseEvent GroupMenuItemClick(.ItemClicked, .ItemCaption)
                    End If
                End With
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

Private Function InSegment(X As Single, Y As Single) As Integer

    Dim sngY1 As Single
    Dim sngY2 As Single
    Dim sngX1 As Single
    Dim sngX2 As Single
    Dim sngAngle As Single
    Dim sngTemp As Single
    Dim oConv As CoordConv
    Dim intSegment As Integer
    Dim intSelected As Integer

    intSelected = -1
    sngY1 = Y / Screen.TwipsPerPixelY
    sngX1 = X / Screen.TwipsPerPixelX
    If (sngX1 >= sngPieXCenter - sngRadius) And (sngX1 <= sngPieXCenter + sngRadius) _
    And (sngY1 >= sngPieYCenter - sngRadius) And (sngY1 <= sngPieYCenter + sngRadius) _
    And (uSelectable = True) Then
        oConv = ConvertCoordinates(X, Y)
        If oConv.Radius <= sngRadius Then
            If uGroupSegment = False Then
                For intSegment = 1 To cItemsAttributes.Count
                    With cItemsAttributes(intSegment)
                        If oConv.Angle >= .AngleFrom And oConv.Angle <= .AngleTo Then
                            intSelected = intSegment - 1
                            Call DrawSegmentText(.AngleFrom, .AngleTo, intSegment)
                            Exit For
                        End If
                    End With
                Next
            Else
                For intSegment = 1 To cGroupsAttributes.Count
                    With cGroupsAttributes(intSegment)
                        If oConv.Angle >= .AngleFrom And oConv.Angle <= .AngleTo Then
                            intSelected = intSegment - 1
                            Call DrawSegmentText(.AngleFrom, .AngleTo, intSegment)
                            Exit For
                        End If
                    End With
                Next
            End If
        End If
    End If
    InSegment = intSelected

End Function
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (uHotTracking = True) Then
        Call Tracking(X, Y)
    ElseIf (Button = vbLeftButton) Then
        Call UserControl_MouseDown(vbLeftButton, Shift, X, Y)
    End If
    
End Sub

Public Sub Refresh()
    DrawChart
End Sub

Public Sub Clear()
    
    Set cItems = Nothing
    Set cItems = New Collection
    Set cGroups = Nothing
    Set cGroups = New Collection
    
    dblPieTotal = 0
    ClearLegendItems
    'the following forces the drawing chart routine to not enhance the description
    'in the legend (if it is visible); the legend items were already deleted!
    DrawChart

End Sub

Public Sub DrawChart()
    
    On Error Resume Next

    Dim lngW            As Long
    Dim lngH            As Long
    Dim intCount        As Integer
    Dim xMiddleArea     As Single
    Dim lngColor        As Long
    
    'do not redraw the chart if not required
    If uAutoRedraw = False Then Exit Sub

    With lblInfo
        .ForeColor = uInfoForeColor
        .BackColor = uInfoBackColor
        .Visible = IIf((uDisplayDescript And uSelected > -1), True, False)
    End With
    mnuMainSelectionInfo.Checked = uDisplayDescript
    picDescription(0).ForeColor = uLegendForeColor
    
    If Not bResize Then ClearLegendItems

    With UserControl
        .Cls
        If uPicture Is Nothing Then
        Else
            'paint the background image
            Call DrawPicture(uLeftMargin, .ScaleWidth - uRightMargin, _
                             uTopMargin, .ScaleHeight - uBottomMargin, _
                             uPictureTile, uPicture)
        End If
    
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
    
        'draw the pie
        'calculate the radius, depending on the drawing area
        If bDisplayLegend Then
            lngW = picSplitter.Left
        Else
            lngW = .ScaleWidth - uRightMargin
        End If
        lngW = lngW - uLeftMargin
        lngH = .ScaleHeight - uBottomMargin - uTopMargin
        If (lngW) > (lngH) Then
            sngRadius = lngH / 2 / Screen.TwipsPerPixelY
        Else
            sngRadius = lngW / 2 / Screen.TwipsPerPixelX
        End If
        'the center
        sngPieXCenter = xMiddleArea / Screen.TwipsPerPixelX
        sngPieYCenter = (uTopMargin + lngH / 2) / Screen.TwipsPerPixelY
        Call DrawPie(sngPieXCenter, sngPieYCenter, sngRadius)

        'add legend, depending on wheter data are displayed using groups or not,
        'and display information
        If uGroupSegment = False Then
            With cItems
                For intCount = 0 To .Count - 1
                    If intCount = uSelected And uSelectable Then Call DisplayInfo(intCount)
                    'Add Legend item
                    If Not bResize Then AddLegendItem .Item(intCount + 1).LegendDescription, .Item(intCount + 1).Color, uLegendForeColor
                Next intCount
            End With
        Else
            With cGroups
                For intCount = 0 To .Count - 1
                    If intCount = uSelected And uSelectable Then Call DisplayInfo(intCount)
                    'Add Legend item
                    If Not bResize Then AddLegendItem .Item(intCount + 1).Name, .Item(intCount + 1).Color, uLegendForeColor
                Next intCount
            End With
        End If

        'in case the legend is displayed
        If bDisplayLegend = True Then
            picLegend.BackColor = uLegendBackColor
            picContainer.BackColor = uLegendBackColor
            If uSelectable And uSelected > -1 Then
                
                Dim perScreen As Integer
                Dim scrollValue As Integer
                            
                perScreen = Abs((picLegend.ScaleHeight / ((picBox(0).Height + (10 * Screen.TwipsPerPixelY)))) - 1)
                            
                If (uSelected + 1) > perScreen Then
                    scrollValue = ((uSelected + 1) * ((picBox(0).Height / Screen.TwipsPerPixelY) + 10)) - (picBox(perScreen).Top / Screen.TwipsPerPixelY)
                    If scrollValue > vsbContainer.Max Then scrollValue = vsbContainer.Max
                    vsbContainer.Value = scrollValue
                Else
                    vsbContainer.Value = 0
                End If
                picContainer.Line ((picBox(uSelected).Left - 3 * Screen.TwipsPerPixelX), (picBox(uSelected).Top - 3 * Screen.TwipsPerPixelY))-(picDescription(uSelected).Left + picDescription(uSelected).Width + 2 * Screen.TwipsPerPixelX, picBox(uSelected).Top + picBox(uSelected).Height + 2 * Screen.TwipsPerPixelY), uSelectedColor, B
            End If
        End If
    End With

End Sub

Private Sub DrawPie(sngXcenter As Single, _
                    sngYcenter As Single, _
                    sngRadius As Single)
        
    Dim obj As Object
    Dim objAtt As Object
    Dim lngColor As Long
    Dim intCount As Integer
    Dim dblSegmentTotal As Double

    On Error GoTo errHandle

    'save current forecolor and set new forecolor
    lngColor = UserControl.ForeColor
    UserControl.ForeColor = uPieBorderColor

    dblSegmentTotal = 100
    If uGroupSegment = False Then
        Set obj = cItems
        Set objAtt = cItemsAttributes
    Else
        Set obj = cGroups
        Set objAtt = cGroupsAttributes
    End If
    With obj
        For intCount = 1 To .Count
            If intCount > 1 Then
                dblSegmentTotal = dblSegmentTotal - objAtt(intCount - 1).Percentage
            End If
            'Create and draw the segment
            If (intCount - 1) = uSelected And uSelectable Then
                DrawSegment dblSegmentTotal, uSelectedColor, sngRadius, sngXcenter, sngYcenter
            Else
                DrawSegment dblSegmentTotal, .Item(intCount).Color, sngRadius, sngXcenter, sngYcenter
            End If
            'print the text
            Call DrawSegmentText(objAtt(intCount).AngleFrom, objAtt(intCount).AngleTo, intCount)
        Next intCount
    End With
    UserControl.ForeColor = lngColor
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub DrawSegment(ByVal intPerc As Integer, _
                        ByVal lngColour As Long, _
                        ByVal sngRadius As Single, _
                        ByVal sngXcenter As Single, _
                        ByVal sngYcenter As Single, _
                        Optional blnIsShadow As Boolean = False)
    
    Dim x1, Y1, X2, Y2, X3, Y3, x4, y4, theta, beta As Double
    Dim lngPie As Long
    Dim lngBrush As Long

    On Error GoTo errHandle
    
    If blnIsShadow = False Then
        x1 = sngXcenter: Y1 = sngYcenter         'Initial Circle points
    Else
        x1 = sngXcenter + 5: Y1 = sngYcenter + 5 'Initial Circle points
    End If
    
    x1 = x1 - sngRadius
    Y1 = Y1 - sngRadius
    X2 = x1 + (sngRadius * 2): Y2 = Y1 + (sngRadius * 2)
    X3 = x1 + (X2 - x1) / 2: Y3 = Y1 'Initial the first point being at North point on the circle
    If intPerc = 100 Then
        'it's the first segment, corresponding to 100%
        x4 = X3
        y4 = Y3
    Else
        'it's another segment, corresponding to <100%
        theta = (intPerc / 100) * 360  'Get theta from the percentage of the pie segment passed as a parameter
        beta = 180 - theta - 90 'This gets the missing angle from the RHT assuming the segment is <90 degrees
        x4 = x1 + sngRadius + ((sngRadius * (Sin(theta * (PI / 180)))) * 180 / PI)
        y4 = Y1 + sngRadius - ((sngRadius * (Sin(beta * (PI / 180)))) * 180 / PI) 'Converts from radians and gets y4 point
    End If
    
    'Automatically fill the segment
    lngBrush = CreateSolidBrush(lngColour)
    SelectObject UserControl.hDC, lngBrush
    
    'Draw the segment
    lngPie = Pie(UserControl.hDC, CLng(x1), CLng(Y1), CLng(X2), CLng(Y2), CLng(x4), CLng(y4), CLng(X3), CLng(Y3)) 'Draw Pie Swapped x3y3 for x4y4 because i want the smaller segment
    Exit Sub

errHandle:
    Exit Sub

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
Public Property Get GroupSegment() As Boolean
    GroupSegment = uGroupSegment
End Property

Public Property Get GroupExplodeOnClick() As Boolean
    GroupExplodeOnClick = uGroupExplode
End Property
Public Property Get GroupExplodeAllowCommands() As Boolean
    GroupExplodeAllowCommands = uGroupExplodeAllowCommands
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

Private Sub AddLegendItem(stgDescription As String, lngBackColor As OLE_COLOR, lngForeColor As OLE_COLOR)
    
    Dim intX As Integer
    Dim sngX As Single
    Dim stgShortDesc As String
    
    If bLegendAdded Then
        intX = picBox.Count
        Load picBox(intX)
        Load picDescription(intX)
        
        picBox(intX).BackColor = lngBackColor
        picBox(intX).Top = picBox(intX - 1).Top + picBox(intX - 1).Height + 10 * Screen.TwipsPerPixelY
        picDescription(intX).Top = picBox(intX).Top
    Else
        intX = 0
        picBox(intX).BackColor = lngBackColor
        bLegendAdded = True
    End If
    
    stgShortDesc = stgDescription
    sngX = picDescription(intX).Left
    While (Len(stgShortDesc) > 0) And ((sngX + picContainer.TextWidth(stgShortDesc)) > (picContainer.ScaleWidth - sngX - 5 * Screen.TwipsPerPixelX))
        stgShortDesc = Left$(stgShortDesc, Len(stgShortDesc) - 1)
    Wend
    
    If Len(stgShortDesc) < Len(stgDescription) Then stgShortDesc = stgShortDesc & ".."
    With picDescription(intX)
        .Width = picContainer.ScaleWidth - sngX - 5 * Screen.TwipsPerPixelX
        .BackColor = uLegendBackColor
        .ForeColor = lngForeColor
        'TAG is used to show tooltip
        If stgShortDesc <> stgDescription Then
            .Tag = stgDescription & TooltipNeeded
        Else
            .Tag = stgDescription
        End If
        .Cls
        picDescription(intX).Print stgShortDesc
        .Visible = True
    End With
            
    picBox(intX).Visible = True
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
        uSelected = .ReadProperty("uSelected", -1)
        uTopMargin = .ReadProperty("MarginTop", 55)
        uBottomMargin = .ReadProperty("MarginBottom", 55)
        uLeftMargin = .ReadProperty("MarginLeft", 55)
        uRightMargin = .ReadProperty("MarginRight", 55)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        UserControl.ForeColor = .ReadProperty("ForeColor", vbBlack)
        uLegendBackColor = .ReadProperty("LegendBackColor", vbWhite)
        uLegendForeColor = .ReadProperty("LegendForeColor", vbBlack)
        uInfoBackColor = .ReadProperty("InfoBackColor")
        uInfoForeColor = .ReadProperty("InfoForeColor")
        uInfoPieBackColor = .ReadProperty("InfoPieBackColor")
        uInfoPieForeColor = .ReadProperty("InfoPieForeColor")
        uChartTitleColor = .ReadProperty("ChartTitleColor", vbBlack)
        uChartSubTitleColor = .ReadProperty("ChartSubTitleColor", vbBlack)
        uGroupExplodeBackColor = .ReadProperty("GroupExplodeBackColor")
        uGroupExplodeForeColor = .ReadProperty("GroupExplodeForeColor")
        uGroupExplodeTitleColor = .ReadProperty("GroupExplodeTitleColor")
        uGroupExplodeMenuItems = .ReadProperty("GroupExplodeMenuItems")
        uMenuType = .ReadProperty("MenuType")
        uMenuItems = .ReadProperty("MenuItems")
        uCustomMenuItems = .ReadProperty("CustomMenuItems")
        uInfoItems = .ReadProperty("InfoItems")
        uSaveAsCaption = .ReadProperty("SaveAsCaption")
        uAutoRedraw = .ReadProperty("AutoRedraw", True)
        Set uPicture = .ReadProperty("Picture", Nothing)
        uPictureTile = .ReadProperty("PictureTile", False)
        uMarkerColor = .ReadProperty("MarkerColor", vbRed)
        uPieBorderColor = .ReadProperty("PieBorderColor", vbRed)
        uDataFormat = .ReadProperty("DataFormat")
        uPrinterFit = .ReadProperty("PrinterFit")
        uPrinterOrientation = .ReadProperty("PrinterOrientation")
        uLegendCaption = .ReadProperty("LegendCaption")
        uRightMarginOrg = uRightMargin
        uHotTracking = .ReadProperty("HotTracking", False)
        uLegendPrintMode = .ReadProperty("LegendPrintMode", legPrintGraph)
        uGroupSegment = .ReadProperty("GroupSegment", False)
        uGroupExplode = .ReadProperty("GroupExplodeOnClick", False)
        uGroupExplodeAllowCommands = .ReadProperty("GroupExplodeAllowCommands", False)
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
    Set cGroups = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SelectedColor", uSelectedColor
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uSelected", uSelected
        .WriteProperty "MarginTop", uTopMargin
        .WriteProperty "MarginBottom", uBottomMargin
        .WriteProperty "MarginLeft", uLeftMargin
        .WriteProperty "MarginRight", uRightMargin
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
        .WriteProperty "LegendBackColor", uLegendBackColor
        .WriteProperty "LegendForeColor", uLegendForeColor
        .WriteProperty "InfoBackColor", uInfoBackColor
        .WriteProperty "InfoForeColor", uInfoForeColor
        .WriteProperty "InfoPieBackColor", uInfoPieBackColor
        .WriteProperty "InfoPieForeColor", uInfoPieForeColor
        .WriteProperty "ChartTitleColor", uChartTitleColor
        .WriteProperty "ChartSubTitleColor", uChartSubTitleColor
        .WriteProperty "GroupExplodeBackColor", uGroupExplodeBackColor
        .WriteProperty "GroupExplodeForeColor", uGroupExplodeForeColor
        .WriteProperty "GroupExplodeTitleColor", uGroupExplodeTitleColor
        .WriteProperty "GroupExplodeMenuItems", uGroupExplodeMenuItems
        .WriteProperty "MenuType", uMenuType
        .WriteProperty "MenuItems", uMenuItems
        .WriteProperty "CustomMenuItems", uCustomMenuItems
        .WriteProperty "InfoItems", uInfoItems
        .WriteProperty "SaveAsCaption", uSaveAsCaption
        .WriteProperty "AutoRedraw", uAutoRedraw
        .WriteProperty "Picture", uPicture, Nothing
        .WriteProperty "PictureTile", uPictureTile
        .WriteProperty "MarkerColor", uMarkerColor
        .WriteProperty "PieBorderColor", uPieBorderColor
        .WriteProperty "DataFormat", uDataFormat
        .WriteProperty "PrinterFit", uPrinterFit
        .WriteProperty "PrinterOrientation", uPrinterOrientation
        .WriteProperty "LegendCaption", uLegendCaption
        .WriteProperty "HotTracking", uHotTracking
        .WriteProperty "LegendPrintMode", uLegendPrintMode
        .WriteProperty "GroupSegment", uGroupSegment
        .WriteProperty "GroupExplodeOnClick", uGroupExplode
        .WriteProperty "GroupExplodeAllowCommands", uGroupExplodeAllowCommands
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

