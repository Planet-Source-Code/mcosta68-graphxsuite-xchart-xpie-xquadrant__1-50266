VERSION 5.00
Begin VB.Form frmGroupExplodeMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3045
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCmd 
      Caption         =   "Command1"
      Height          =   345
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "frmGroupExplodeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngCurrX As Long
Private lngCurrY As Long
Private stgCmdItems As String
Private intItemClicked As Integer
Private stgItemCaption As String
Public Property Let CommandItems(stgVal As String)
    stgCmdItems = stgVal
End Property
Public Property Get ItemClicked() As Integer
    ItemClicked = intItemClicked
End Property

Public Property Get ItemCaption() As String
    ItemCaption = stgItemCaption
End Property

Public Property Let CurrX(lngX As Long)
    lngCurrX = lngX
End Property
Public Property Let CurrY(lngY As Long)
    lngCurrY = lngY
End Property

Private Sub cmdCmd_Click(Index As Integer)
   
    intItemClicked = Index - 1
    stgItemCaption = cmdCmd(Index).Caption
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        intItemClicked = -1
        stgItemCaption = Empty
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Dim lngH As Long
    Dim lngW As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim stgCmd As String
    Dim intIdx As Integer
    Dim varCmdItems As Variant
    
    intItemClicked = -1
    stgItemCaption = Empty
    varCmdItems = Split(stgCmdItems, "|")
    
    'find the maximum width of the command buttons and the height of the form
    lngW = 0
    lngH = 0
    For intIdx = 0 To UBound(varCmdItems)
        stgCmd = CStr(varCmdItems(intIdx))
        If stgCmd <> Empty Then
            lngH = lngH + TextHeight(stgCmd) + 10
            If TextWidth(stgCmd) > lngW Then lngW = TextWidth(stgCmd) + 12
        End If
    Next
    If lngW < 60 Then lngW = 60
    lngH = lngH + 4
    
    'display command buttons
    lngX = 0
    lngY = 0
    For intIdx = 0 To UBound(varCmdItems)
        stgCmd = CStr(varCmdItems(intIdx))
        If stgCmd <> Empty Then
            Load cmdCmd(intIdx + 1)
            With cmdCmd(intIdx + 1)
                .Move lngX, lngY, lngW - 6, TextHeight(stgCmd) + 8
                .Caption = stgCmd
                .Visible = True
                lngY = lngY + .Height + 2
            End With
        End If
    Next
    
    'calculate X,Y coordinates where showing the form
    lngX = lngCurrX
    lngY = lngCurrY
    'check if form is inside screen
    While (lngY + lngH > Screen.Height) And (lngY > 0)
        lngY = lngY - Screen.TwipsPerPixelY
    Wend
    While (lngX + lngW > Screen.Width) And (lngX > 0)
        lngX = lngX - Screen.TwipsPerPixelX
    Wend
    'move form
    Move lngX, lngY, lngW * Screen.TwipsPerPixelX, lngH * Screen.TwipsPerPixelY

End Sub


