Attribute VB_Name = "ModListview"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const OnTopFlags = &H2 Or &H1
Public Const HWND_TOPMOST = -1

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = &H1000 + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = &H1000 + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1 ' untuk membuat gridlines
Public Const LVS_EX_CHECKBOXES As Long = &H4 ' untuk penambahan checkbox
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
Public Const LVS_EX_SUBITEMIMAGES = &H2

Public Const LVIF_STATE = &H8
 
Public Const LVM_SETITEMSTATE = (&H1000 + 43)
Public Const LVM_GETITEMSTATE As Long = (&H1000 + 44)
Public Const LVM_GETITEMTEXT As Long = (&H1000 + 45)
Private Const GWL_STYLE        As Long = (-16)
Private Const LVM_GETHEADER    As Long = (&H1000 + 31)
Private Const LVM_ARRANGE      As Long = (&H1000 + 22)
Private Const HDS_BUTTONS      As Long = 2

Public Const LVIS_STATEIMAGEMASK As Long = &HF000

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Const LVM_GETCOLUMN = (&H1000 + 25)
Public Const LVM_GETCOLUMNORDERARRAY = (&H1000 + 59)
Public Const LVCF_TEXT = &H4

Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

Public Const LVM_FIRST As Long = &H1000

Public Type LV_ITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type
Function ListviewCheck(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES Xor LVS_EX_CHECKBOXES
    r = SendMessageLong(lvStyle.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function
Function ListviewFlat(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES 'Xor LVS_EX_CHECKBOXES
    r = SendMessageLong(lvStyle.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function
Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)
 
 Dim c As ColumnHeader
 If Column Is Nothing Then
  For Each c In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, c.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub



