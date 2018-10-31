VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRTP 
   Caption         =   "Real Time Protector"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRTP.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrControls 
      Interval        =   100
      Left            =   8400
      Top             =   720
   End
   Begin VB.CommandButton cmdQuarantine 
      Caption         =   "Quarantine Selected"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Delete Selected"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin ComctlLib.ListView lvMalware 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Virus Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size(s)"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Label cmdAbaikan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore All Viruses"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   7560
      MouseIcon       =   "frmRTP.frx":7A48A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3480
      Width           =   1350
   End
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbaikan_Click()
Me.Hide
frmMain.tmrRTP.Enabled = True
tmrControls.Enabled = True
End Sub

Private Sub Form_Load()
ListviewFlat lvMalware
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Height = 4605
Me.Width = 9375
End Sub

Private Sub tmrControls_Timer()
If lvMalware.ListItems.Count <> 0 Then
Me.Show
frmMain.tmrRTP.Enabled = False
tmrControls.Enabled = False
End If
End Sub
