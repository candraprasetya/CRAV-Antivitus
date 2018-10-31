VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":8D25A
   ScaleHeight     =   3810
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrSplash 
      Interval        =   5000
      Left            =   6000
      Top             =   2880
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Antivirus"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2640
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nid As NOTIFYICONDATA

Private Sub Form_Load()
SysTrayCRAV
End Sub

Private Sub tmrSplash_Timer()
If tmrSplash.Interval = 1000 Then
lblStatus.Caption = "Starting Antivirus"
ElseIf tmrSplash.Interval = 3000 Then
lblStatus.Caption = "Load Database Antivirus"
ElseIf tmrSplash.Interval = 5000 Then
Me.Hide
tmrSplash.Enabled = False
End If
End Sub


Public Function SysTrayCRAV()
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon 'the icon will be your Form1 project icon
nid.szTip = "CRAV Antivirus 2014" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long
Dim sFilter As String
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
PopupMenu frmMain.mnTray
Case WM_RBUTTONUP
'Me.Show
Case WM_RBUTTONDBLCLK
End Select
End Sub

