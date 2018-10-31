VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CRAV Antivirus 2014"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
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
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7245
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Index           =   0
      Left            =   2280
      ScaleHeight     =   5175
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   1800
      Width           =   8055
      Begin VB.Timer tmrRTP 
         Interval        =   100
         Left            =   6600
         Top             =   3360
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Disabled"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   56
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": 01 June 2015"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   55
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Active"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   54
         Top             =   1080
         Width           =   660
      End
      Begin VB.Image imgcontoh 
         Height          =   240
         Index           =   2
         Left            =   2280
         Picture         =   "frmMain.frx":F6C76
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USB Auto Scan"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   53
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Image imgcontoh 
         Height          =   240
         Index           =   1
         Left            =   2280
         Picture         =   "frmMain.frx":F7200
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Update"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   52
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Real Time Protection"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   51
         Top             =   1080
         Width           =   1725
      End
      Begin VB.Image imgcontoh 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmMain.frx":F778A
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   720
         Picture         =   "frmMain.frx":F7D14
         Top             =   480
         Width           =   1500
      End
   End
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Index           =   1
      Left            =   2280
      ScaleHeight     =   5175
      ScaleWidth      =   8055
      TabIndex        =   13
      Top             =   1800
      Width           =   8055
      Begin VB.PictureBox PicScan 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   0
         Left            =   240
         ScaleHeight     =   4455
         ScaleWidth      =   7575
         TabIndex        =   31
         Top             =   480
         Width           =   7575
         Begin VB.CommandButton cmdRefDir 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   5640
            TabIndex        =   35
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdScan 
            Caption         =   "Scan Now"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5640
            TabIndex        =   34
            Top             =   120
            Width           =   1815
         End
         Begin Project1.DirTree DirScan 
            Height          =   4335
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   7646
         End
         Begin VB.TextBox txtPath 
            Height          =   375
            Left            =   600
            TabIndex        =   33
            Top             =   3480
            Visible         =   0   'False
            Width           =   2655
         End
      End
      Begin VB.PictureBox PicScan 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   1
         Left            =   240
         ScaleHeight     =   4455
         ScaleWidth      =   7575
         TabIndex        =   36
         Top             =   480
         Width           =   7575
         Begin VB.Timer tmrProgress 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2760
            Top             =   3000
         End
         Begin ComctlLib.ProgressBar Prog 
            Height          =   255
            Left            =   2040
            TabIndex        =   58
            Top             =   2520
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Progressbar:"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   57
            Top             =   2520
            Width           =   960
         End
         Begin VB.Label lblfolder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2040
            TabIndex        =   46
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Directory Count:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label lblscan 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   44
            Top             =   840
            Width           =   5940
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scan Processess:"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   8.25
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lblvirus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2040
            TabIndex        =   42
            Top             =   2160
            Width           =   90
         End
         Begin VB.Label lblFile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   2040
            TabIndex        =   41
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File(s) Count:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Malware Detected:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   39
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CRAV Scan Information"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   12
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   2355
         End
         Begin VB.Label lblAbout 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information Scanning"
            BeginProperty Font 
               Name            =   "Segoe UI Light"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   240
            TabIndex        =   37
            Top             =   600
            Width           =   1575
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FCF4EC&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00E0E0E0&
            Height          =   1095
            Left            =   0
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.PictureBox PicScan 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   2
         Left            =   240
         ScaleHeight     =   4455
         ScaleWidth      =   7575
         TabIndex        =   48
         Top             =   480
         Width           =   7575
         Begin VB.CommandButton cmdHapusVir 
            Caption         =   "Hapus Selected"
            Height          =   375
            Left            =   6000
            TabIndex        =   50
            Top             =   4080
            Width           =   1575
         End
         Begin ComctlLib.ListView lvMalware 
            Height          =   3855
            Left            =   0
            TabIndex        =   49
            Top             =   120
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   6800
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
      End
      Begin ComctlLib.TabStrip TabScan 
         Height          =   4935
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Choose a Location"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Scanner Status"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Virus Detected"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Index           =   2
      Left            =   2280
      ScaleHeight     =   5175
      ScaleWidth      =   8055
      TabIndex        =   14
      Top             =   1800
      Width           =   8055
      Begin VB.PictureBox PicTools 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   0
         Left            =   240
         ScaleHeight     =   4455
         ScaleWidth      =   7575
         TabIndex        =   23
         Top             =   480
         Width           =   7575
         Begin ComctlLib.ListView lvProcess 
            Height          =   3855
            Left            =   0
            TabIndex        =   24
            Top             =   120
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   6800
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Process Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "PID"
               Object.Width           =   1129
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Size(b)"
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Location"
               Object.Width           =   7832
            EndProperty
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   4560
            TabIndex        =   26
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdTerminate 
            Caption         =   "Terminate"
            Height          =   375
            Left            =   6120
            TabIndex        =   25
            Top             =   4080
            Width           =   1455
         End
      End
      Begin VB.PictureBox PicTools 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4455
         Index           =   1
         Left            =   240
         ScaleHeight     =   4455
         ScaleWidth      =   7575
         TabIndex        =   27
         Top             =   480
         Width           =   7575
         Begin ComctlLib.ListView lvStartup 
            Height          =   3855
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   6800
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
               Text            =   "Object Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Location"
               Object.Width           =   7832
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Registry"
               Object.Width           =   1658
            EndProperty
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Remove Selected"
            Height          =   375
            Left            =   5880
            TabIndex        =   29
            Top             =   4080
            Width           =   1695
         End
      End
      Begin ComctlLib.TabStrip TabTools 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Process Manager"
               Key             =   ""
               Object.Tag             =   "process"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Startup Manager"
               Key             =   ""
               Object.Tag             =   "startup"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox PicMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Index           =   3
      Left            =   2280
      ScaleHeight     =   5175
      ScaleWidth      =   8055
      TabIndex        =   15
      Top             =   1800
      Width           =   8055
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designer UI:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Muh.Isfahani Ghiyath.YM"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   20
         Top             =   1320
         Width           =   2010
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Candra Ramadhan"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   19
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEO && Programmer:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Antivirus Information"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About CRAV"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1275
      End
   End
   Begin Project1.rtp_mode rtp_mode1 
      Index           =   0
      Left            =   11040
      Top             =   6600
      _ExtentX        =   1296
      _ExtentY        =   450
   End
   Begin VB.Label lblMenu 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   120
      MouseIcon       =   "frmMain.frx":100096
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   120
      MouseIcon       =   "frmMain.frx":1001E8
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmMain.frx":10033A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmMain.frx":10048C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label LblDeskripsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   8
      Top             =   5520
      Width           =   915
   End
   Begin VB.Label lblItemMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   5160
      Width           =   630
   End
   Begin VB.Label LblDeskripsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More Application"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblItemMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label LblDeskripsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Malware"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblItemMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scanner"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label LblDeskripsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overview Crav"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Image ImgMenu 
      Height          =   870
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":1005DE
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image ImgMenu 
      Height          =   870
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":1073C8
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image ImgMenu 
      Height          =   870
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":10E1B2
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblItemMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image ImgMenu 
      Height          =   870
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":114F9C
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Menu mnTray 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnOpen 
         Caption         =   "Open CRAV"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit CRAV"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents SCANPROC As cScanProcesses
Attribute SCANPROC.VB_VarHelpID = -1
Dim m_Time As Single
Dim WithEvents ShellIE As SHDocVw.ShellWindows
Attribute ShellIE.VB_VarHelpID = -1
Private Declare Function InitCommonControls Lib "Comctl32" () As Long
Public statusRTP As Boolean
Dim isCompatch As Boolean
Dim eWindow As InternetExplorer
Private cWindow As New ShellWindows

Private Sub cmdHapusVir_Click()
Dim i As Long
For i = 1 To lvMalware.ListItems.Count
If lvMalware.ListItems(i).Selected = True Then
Kill lvMalware.ListItems(i).SubItems(2)
lvMalware.ListItems.Remove lvMalware.ListItems(i).Selected = True
End If
Next
End Sub

Private Sub cmdRefDir_Click()
DirScan.LoadTreeDir False
End Sub

Private Sub cmdRefresh_Click()
lvProcess.ListItems.Clear
SCANPROC.SystemProcesses = True
m_Time = Timer
SCANPROC.BeginScanning
End Sub

Private Sub GetAllRun()
    On Error Resume Next
    Dim X As ListItem, hkey As Long, lCount As Long, i As Long
    lvStartup.ListItems.Clear
    hkey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hkey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartup.ListItems.Add(, , EnumValue(hkey, i))
        X.SubItems(1) = GetKeyValue(hkey, EnumValue(hkey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        Set X = Nothing
    Next
    hkey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hkey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartup.ListItems.Add(, , EnumValue(hkey, i))
        X.SubItems(1) = GetKeyValue(hkey, EnumValue(hkey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE (Service)"
        Set X = Nothing
    Next
    hkey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hkey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartup.ListItems.Add(, , EnumValue(hkey, i))
        X.SubItems(1) = GetKeyValue(hkey, EnumValue(hkey, i))
        X.SubItems(2) = "HKEY_CURRENT_USER"
        Set X = Nothing
    Next
    Dim fso As New FileSystemObject
    Dim sFolder As Folder
    Dim sFiles As Files
    Dim sFile As File
    Set sFolder = fso.GetFolder("C:\Windows\Tasks")
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            Set X = Me.lvStartup.ListItems.Add(, , sFile.Name)
            X.SubItems(1) = sFile.path
            X.SubItems(2) = "Tasks"
            Set X = Nothing
        Next
    End If
    Set sFolder = fso.GetFolder("C:\Documents and Settings\All Users\Start Menu\Programs\Startup")
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            Set X = lvStartup.ListItems.Add(, , sFile.Name)
            X.SubItems(1) = sFile.path
            X.SubItems(2) = "All User Startup"
            Set X = Nothing
        Next
    End If
End Sub

Private Sub cmdScan_Click()
Call bersihkanlog
StopScan = False
BufferStop = False
If txtPath.Text = "" Then
MsgBox "Directory to Scan not found!", vbCritical
Else
TabScan.Tabs(2).Selected = True
tmrProgress.Enabled = True
BufferAntivirus txtPath.Text, True
cmdStop.Enabled = True
EngineAntivirus txtPath.Text, True
StopScan = True
BufferStop = True
tmrProgress.Enabled = False
End If
End Sub

Sub bersihkanlog()
lblscan.Caption = "-": lblvirus.Caption = "0": lblfolder.Caption = "0": lblFile.Caption = "0"
End Sub

Private Sub cmdStop_Click()
MsgBox "Scan Finished!", vbInformation
cmdStop.Enabled = False
BufferStop = True
StopScan = True
End Sub

Private Sub cmdTerminate_Click()
On Error Resume Next
Dim i As Long
For i = 1 To lvProcess.ListItems.Count
If lvProcess.SelectedItem.Selected = False Then MsgBox "Not Found!", vbCritical: Exit Sub
If lvProcess.ListItems(i).Selected = True Then
SCANPROC.TerminateProcess lvProcess.SelectedItem.SubItems(1)
lvProcess.ListItems.Remove lvProcess.SelectedItem.Index
End If
Next
End Sub

Private Sub Command1_Click()
 With lvStartup
    Dim i As Long
    Dim TMP As Long
    Dim fso As New FileSystemObject
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Selected = True Then
            If .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE (Service)" Then 'run service startup
                DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices", .ListItems.Item(i).Text
            ElseIf .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE" Or .ListItems.Item(i).SubItems(2) = "HKEY_CURRENT_USER" Then 'normal startup
                DeleteStartup GetClassKey(.ListItems.Item(i).SubItems(2)), "Software\Microsoft\Windows\CurrentVersion\Run", .ListItems.Item(i).Text
            Else
                fso.DeleteFile .ListItems.Item(i).SubItems(1), True
            End If
        End If
    Next
    Set fso = Nothing
    End With
    Call GetAllRun
End Sub

Private Sub Form_Load()
Set SCANPROC = New cScanProcesses
Call CompactObject
frmRTP.tmrControls.Enabled = True
ListviewFlat lvProcess
ListviewFlat lvStartup
ListviewFlat lvMalware
lvProcess.ListItems.Clear
SCANPROC.SystemProcesses = False
m_Time = Timer
SCANPROC.BeginScanning
Call GetAllRun
Call lblMenu_Click(0)
DirScan.LoadTreeDir False

frmMain.Height = 7755
frmMain.Width = 10620
End Sub

Private Sub Form_Resize()
On Error Resume Next
frmMain.Height = 7755
frmMain.Width = 10620
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub lblMenu_Click(Index As Integer)
    On Error Resume Next
    Dim i As Byte
    Const t1 = 1800, l1 = 2280, h1 = 5175, w1 = 8055
    For i = lblMenu.LBound To lblMenu.UBound
        If i <> Index Then
            PicMenu(i).Visible = False
            ImgMenu(i).Visible = False
        End If
    Next
    With PicMenu(Index)
        .Top = t1: .Left = l1: .Height = h1: .Width = w1
    End With
    ImgMenu(Index).Visible = True
    PicMenu(Index).Visible = True
End Sub

Private Sub mnExit_Click()
Unload Me
Unload frmRTP
Unload frmSplash
End Sub

Private Sub mnOpen_Click()
Me.Show
End Sub

Private Sub SCANPROC_CurrentProcess(File As String, path As String, ID As Long, Terminate As Boolean)
On Error Resume Next
    Set lsv = lvProcess.ListItems.Add(, , File)
    lsv.SubItems(1) = ID
    lsv.SubItems(2) = FileLen(ProcessPathByPID(ID))
    lsv.SubItems(3) = ProcessPathByPID(ID) 'Path & File
End Sub

Private Sub SCANPROC_DoneScanning(TotalProcesses As Integer)
    Dim p_Elapsed As Single
    p_Elapsed = Timer - m_Time
    Debug.Print "Total Number of Process Detected: " & TotalProcesses & vbNewLine & "Total Scan Time: " & p_Elapsed & vbNewLine
    NUMPROC = TotalProcesses
End Sub

Private Sub TabScan_Click()
If TabScan.SelectedItem.Index = 1 Then
PicScan(0).Visible = True: PicScan(1).Visible = False: PicScan(2).Visible = False
ElseIf TabScan.SelectedItem.Index = 2 Then
PicScan(1).Visible = True: PicScan(0).Visible = False: PicScan(2).Visible = False
ElseIf TabScan.SelectedItem.Index = 3 Then
PicScan(2).Visible = True: PicScan(1).Visible = False: PicScan(0).Visible = False
End If
End Sub

Private Sub TabTools_Click()
If TabTools.SelectedItem.Index = 1 Then
PicTools(0).Visible = True: PicTools(1).Visible = False
ElseIf TabTools.SelectedItem.Index = 2 Then
PicTools(1).Visible = True: PicTools(0).Visible = False
End If
End Sub

Private Sub tmrProgress_Timer()
On Error Resume Next
 If JumlahBuffer <> 0 Then
    Dim persen As Single: persen = Round(CSng((jumlahFile / JumlahBuffer) * 100), 1)
    If Prog.value <> persen Then: Prog.value = persen
 End If
End Sub

Private Sub rtp_mode1_PathChange(Index As Integer, strPath As String)
tmrRTP.Enabled = True
End Sub

Private Sub ShellIE_WindowRegistered(ByVal lCookie As Long) ' user membuka explorer baru
   Call MulaiRTP ' jika TRUE ajh
End Sub
Private Sub MulaiRTP()
On Error Resume Next
If isCompatch = False Then
   Dim i As Integer, CNT As Integer
   CNT = ShellIE.Count - 1
   For i = 0 To CNT
       If (rtp_mode1.Count - 1) < CNT Then
          AddIEObj i
       End If
          If FindID(ShellIE(i).hwnd) = False Then
             rtp_mode1(i).EnabledMonitoring True
             rtp_mode1(i).AddSubClass ShellIE(i)
          End If
   Next i
End If
End Sub

Sub AddIEObj(Index As Integer)
On Error GoTo salah
    Load rtp_mode1(Index)
salah:
End Sub

Function FindID(ID As Long) As Boolean
On Error GoTo salah
    Dim i As Integer
    For i = 0 To rtp_mode1.Count - 1
        If rtp_mode1(i).IEKey = ID Then
           FindID = True
        End If
    Next i
salah:
End Function

Private Sub CompactObject() ' untuk aktifkan rtp
On Error Resume Next
isCompatch = True
   Dim i As Integer, CNT As Integer
   For i = 0 To rtp_mode1.Count - 1
       rtp_mode1(i).SetIENothing
   Next i
       
   Set ShellIE = Nothing
   For i = 1 To rtp_mode1.Count - 1
        Unload rtp_mode1(i)
   Next i
   
   Set ShellIE = New SHDocVw.ShellWindows
   CNT = ShellIE.Count - 1
   For i = 0 To CNT
       If i > 0 Then
          AddIEObj i
       End If
          rtp_mode1(i).AddSubClass ShellIE(i)
   Next i
isCompatch = False
End Sub

Private Sub tmrRTP_Timer()
Dim buffer As String
Dim cLocData As String
Dim Files As Collection
Dim clocation As String

    tmrRTP.Enabled = False
    
    For Each eWindow In cWindow
    
        DoEvents
        If eWindow.Busy Then
            GoTo winBusy
        End If
        
        clocation = eWindow.LocationURL
        cLocData = InStr(1, buffer, eWindow.LocationName & "|" & eWindow.LocationURL & "|")
        
        If cLocData = 0 Then
            If Mid$(clocation, 1, 7) = "file://" Then
                 clocation = Replace(clocation, "file:///", "") 'jadikan url /// kosonk
                 clocation = Replace(clocation, "%20", " ") 'jadikan jadi space
                 clocation = Replace(clocation, "/", "\") 'ubah left to right
                 EngineAntivirusForRTP clocation, False
                 Debug.Print clocation
            End If
        End If
        
winBusy:
        
    Next
    tmrRTP.Enabled = True
    On Error GoTo 0
    'MsgBox "error"
    tmrRTP.Enabled = False 'matikan biar gak ne fload , bisa juga aktifin trus tapi kinerja selalu on menimbulkan bugging
End Sub
