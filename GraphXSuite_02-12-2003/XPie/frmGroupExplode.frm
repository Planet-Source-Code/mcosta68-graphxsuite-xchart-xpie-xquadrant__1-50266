VERSION 5.00
Begin VB.Form frmGroupExplode 
   BorderStyle     =   0  'None
   ClientHeight    =   1155
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGroupExplode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCmd 
      Caption         =   "Others..."
      Height          =   315
      Index           =   2
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCmd 
      Caption         =   "Print"
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdCmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Line linSep 
      X1              =   14
      X2              =   58
      Y1              =   38
      Y2              =   38
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   585
   End
End
Attribute VB_Name = "frmGroupExplode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'Sets window region; used after setting the form's shape
Private Declare Function CreateRoundRectRgn Lib "gdi32" _
                                (ByVal x1 As Long, ByVal Y1 As Long, _
                                 ByVal X2 As Long, ByVal Y2 As Long, _
                                 ByVal X3 As Long, ByVal Y3 As Long) As Long
                                 
'SetWindowRgn is used when setting the form's shape (rounded corners) so
'Windows knows what the window's region is. That's the area in the window
'where Windows permits drawing, and it won't show any part of the window
'that is outside the window region. hWnd is the handle of the window we're
'working with, hRgn is the region's handle, and bRedraw is the redraw flag.
Private Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hRgn As Long, ByVal bRedraw As Long) As Long

'Used to shape form (round corners)
Private Declare Function RoundRect Lib "gdi32" _
    (ByVal hDC As Long, _
     ByVal x1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long, _
     ByVal X3 As Long, ByVal Y3 As Long) As Long

Private stgTitle As String
Private stgGroupList As String
Private lngInfoBackColor As OLE_COLOR
Private lngInfoForeColor As OLE_COLOR
Private lngTitleColor As OLE_COLOR
Private varCmdItems As Variant
Private stgCmdItems As String
Private stgCmdOthers As String
Private blnAllowCommands As Boolean
Private intItemClicked As Integer
Private stgItemCaption As String

Private Const PRINT_X_OFF = 7       'in pixels
Public Property Get ItemClicked() As Integer
    ItemClicked = intItemClicked
End Property

Public Property Get ItemCaption() As String
    ItemCaption = stgItemCaption
End Property

Public Property Let InfoBackColor(lngVal As OLE_COLOR)

    lngInfoBackColor = lngVal

End Property

Public Property Let InfoForeColor(lngVal As OLE_COLOR)

    lngInfoForeColor = lngVal

End Property


Public Property Let Title(stgVal As String)

    stgTitle = stgVal

End Property
Public Property Let CommandItems(stgVal As String)

    Dim intIdx As Integer
    
    stgCmdItems = stgVal
    varCmdItems = Split(stgCmdItems, "|")
    
    stgCmdOthers = Empty
    For intIdx = 3 To UBound(varCmdItems)
        stgCmdOthers = stgCmdOthers & CStr(varCmdItems(intIdx)) & "|"
    Next

End Property
Public Property Let TitleColor(lngVal As OLE_COLOR)

    lngTitleColor = lngVal

End Property
Public Property Let AllowCommands(blnVal As Boolean)

    blnAllowCommands = blnVal

End Property



Public Property Let GroupList(stgVal As String)

    stgGroupList = stgVal

End Property

Private Sub cmdCmd_Click(Index As Integer)
    
    Dim sngX As Single
    Dim sngY As Single
    Dim sngW As Single
    Dim intPos As Integer
    Dim intIdx As Integer
    Dim stgDesc As String
    Dim stgPerc As String
    Dim intScale As Integer
    Dim dblPercentage As Double
    
    Select Case Index
        Case 0
            'OK button
            Unload Me
        
        Case 1
            'print the group's list
            With Printer
                intScale = .ScaleMode
                'use vbTwips scale since it can rely on it (vbPixels isn't)
                .ScaleMode = vbTwips
                .FontBold = True
                Printer.Print lblTitle.Caption & vbCrLf
                .FontBold = False
                dblPercentage = 0
                intIdx = 1
                While intIdx > 0
                    stgDesc = TokenByPos(stgGroupList, intIdx, vbNullChar)
                    If stgDesc <> Empty Then
                        intPos = InStrRev(stgDesc, "(")
                        stgPerc = Right$(stgDesc, Len(stgDesc) - intPos + 1)
                        stgDesc = Left$(stgDesc, intPos - 1)
                        stgPerc = Replace(Replace(stgPerc, "(", ""), ")", "")
                        dblPercentage = dblPercentage + CDbl(Replace(Replace(Replace(stgPerc, "(", ""), ")", ""), "%", ""))
                        .CurrentX = PRINT_X_OFF * .TwipsPerPixelX
                        Printer.Print stgDesc;
                        .CurrentX = .ScaleX(Me.Width, Me.ScaleMode, .ScaleMode) - .TextWidth(stgPerc) - PRINT_X_OFF * .TwipsPerPixelX
                        Printer.Print stgPerc
                        intIdx = intIdx + 1
                    Else
                        intIdx = -1
                    End If
                Wend
                'print the separator line
                Printer.Line (PRINT_X_OFF * .TwipsPerPixelX, .CurrentY + 3 * .TwipsPerPixelY)-(.ScaleX(Me.Width, Me.ScaleMode, .ScaleMode) - PRINT_X_OFF * .TwipsPerPixelX, .CurrentY + 3 * .TwipsPerPixelY)
                'print the total
                stgPerc = Format$(dblPercentage, "#0.00\%")
                .CurrentY = .CurrentY + 6 * .TwipsPerPixelY
                .CurrentX = .ScaleX(Me.Width, Me.ScaleMode, .ScaleMode) - .TextWidth(stgPerc) - PRINT_X_OFF * .TwipsPerPixelX
                Printer.Print stgPerc
                .ScaleMode = intScale
                .EndDoc
            End With
        
        Case 2
            If stgCmdOthers <> Empty Then
                With frmGroupExplodeMenu
                    .CommandItems = stgCmdOthers
                    .CurrX = Me.Left + cmdCmd(Index).Left * Screen.TwipsPerPixelX
                    .CurrY = Me.Top + (cmdCmd(Index).Top + cmdCmd(Index).Height) * Screen.TwipsPerPixelY
                    .Show vbModal
                    If .ItemClicked >= 0 Then
                        intItemClicked = .ItemClicked
                        stgItemCaption = .ItemCaption
                        Unload Me
                    End If
                End With
            End If
            
    End Select

End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()

    Dim lngH As Long
    Dim lngW As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim ctl As Control
    Dim lngWCmd As Long
    Dim C As PointAPI
    Dim stgDesc As String
    Dim intIdx As Integer
    
    intItemClicked = -1
    stgItemCaption = Empty
    lblTitle.Caption = IIf((stgTitle = Empty), "Group data", stgTitle)
    
    'verify if the form must be resized
    intIdx = 1
    lngW = 0
    lngH = lblTitle.Top + lblTitle.Height + 5
    While intIdx > 0
        stgDesc = TokenByPos(stgGroupList, intIdx, vbNullChar)
        If stgDesc <> Empty Then
            lngH = lngH + TextHeight(stgDesc)
            If TextWidth(stgDesc) > lngW Then lngW = TextWidth(stgDesc)
            intIdx = intIdx + 1
        Else
            intIdx = -1
        End If
    Wend
    'add displacement for line separator and command buttons (if allowed)
    lngH = lngH + 8 + TextHeight("A") + 8 + IIf(blnAllowCommands = True, cmdCmd(0).Height + 4, 0)
    lngW = lngW + 55
    
    If blnAllowCommands = True Then
        'calculate X-space occupied by command buttons
        lngWCmd = 0
        For Each ctl In cmdCmd
            With ctl
                If .Index <= UBound(varCmdItems) Then
                    lngWCmd = lngWCmd + TextWidth(varCmdItems(.Index)) + 12
                End If
            End With
        Next
        If lngW < lngWCmd Then lngW = lngWCmd + 24
    End If
    If lngW < 160 Then lngW = 160
    
    'calculate X,Y coordinates where showing the form
    GetCursorPos C
    lngX = C.X * Screen.TwipsPerPixelX
    lngY = C.Y * Screen.TwipsPerPixelY
    'check if form is inside screen
    While (lngY + lngH > Screen.Height) And (lngY > 0)
        lngY = lngY - Screen.TwipsPerPixelY
    Wend
    While (lngX + lngW > Screen.Width) And (lngX > 0)
        lngX = lngX - Screen.TwipsPerPixelX
    Wend
    'move form
    Move lngX, lngY, lngW * Screen.TwipsPerPixelX, lngH * Screen.TwipsPerPixelY
    BackColor = lngInfoBackColor
    ForeColor = lngInfoForeColor
    
    'move the title to the center
    With lblTitle
        .ForeColor = lngTitleColor
        .Left = (ScaleWidth - lblTitle.Width) / 2
    End With
    
End Sub

Private Sub Form_Paint()
    
    On Error Resume Next

    Dim ctl As Control
    Dim lngX As Long
    Dim lngY As Long
    Dim lngW As Long
    Dim intPos As Integer
    Dim intIdx As Integer
    Dim stgDesc As String
    Dim stgPerc As String
    Dim dblPercentage As Double

    'display the group's list
    Cls
    CurrentY = lblTitle.Top + lblTitle.Height + 5
    dblPercentage = 0
    intIdx = 1
    While intIdx > 0
        stgDesc = TokenByPos(stgGroupList, intIdx, vbNullChar)
        If stgDesc <> Empty Then
            intPos = InStrRev(stgDesc, "(")
            stgPerc = Right$(stgDesc, Len(stgDesc) - intPos + 1)
            stgDesc = Left$(stgDesc, intPos - 1)
            stgPerc = Replace(Replace(stgPerc, "(", ""), ")", "")
            dblPercentage = dblPercentage + CDbl(Replace(Replace(Replace(stgPerc, "(", ""), ")", ""), "%", ""))
            CurrentX = PRINT_X_OFF
            Print stgDesc;
            CurrentX = ScaleWidth - TextWidth(stgPerc) - PRINT_X_OFF
            Print stgPerc
            intIdx = intIdx + 1
        Else
            intIdx = -1
        End If
    Wend
    'place the separator
    With linSep
        .Y1 = CurrentY + 3
        .Y2 = .Y1
        .x1 = PRINT_X_OFF
        .X2 = ScaleWidth - PRINT_X_OFF
    End With
    'print the total
    CurrentY = CurrentY + 6
    stgPerc = Format$(dblPercentage, "#0.00\%")
    CurrentX = ScaleWidth - TextWidth(stgPerc) - PRINT_X_OFF
    Print stgPerc
    
    If blnAllowCommands = True Then
        'display command buttons
        lngY = CurrentY + 6
        lngX = ScaleWidth - 4
        For Each ctl In cmdCmd
            With ctl
                If .Index <= UBound(varCmdItems) Then
                    stgDesc = varCmdItems(.Index)
                    lngW = TextWidth(stgDesc) + 12
                    lngX = lngX - lngW - 3
                    .Move lngX, lngY, lngW
                    .Caption = stgDesc
                    .Visible = True
                Else
                    .Visible = False
                End If
            End With
        Next
    Else
        For Each ctl In cmdCmd
            ctl.Visible = False
        Next
    End If
    
    RoundRect hDC, 0, 0, ScaleWidth, ScaleHeight, 26, 26
    
End Sub

Private Sub Form_Resize()

    Dim hRgn As Long
    Dim lRes As Long
   
    'Round the corners of this form to make it look "tool-tippy"
    hRgn = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, 28, 28)
    lRes = SetWindowRgn(hWnd, hRgn, True)

End Sub

Private Sub lblTitle_Click()
    Unload Me
End Sub


