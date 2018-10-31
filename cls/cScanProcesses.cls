VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Author         : Noel A. Dacara (noeldacara@yahoo.com)
'Filename       : Scan Processes.cls (cScanProcesses Class Module)
'Description    : Scan through all opened processes on your system
'Date           : Monday, May 30, 2005, 09:15 AM
'Last Update    : Thursday, November 24, 2005, 12:43 PM

'You can freely use and distribute this class or upload these codes on any site
'provided that the original credits are kept unmodified.

'Keep note that :
'Class was tested only on the following operating systems: Win98SE, WinNT, and WinXP.

'I consider processes as system processes only if their path is cannot be determined by
'using the function I have here...because sometimes, viruses, worms, etc.. makes use of
'system processes' filenames as theirs so that in some way it will be difficult for a user
'to terminate the process using the built-in Windows® Task Manager. Try it and you'll see.

'Credits goes to:
'Luprix (luprixnet@hotmail.com) for his "List Processes and Ports with Native API" project
'The Terminator at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=62342&lngWId=1

'Scanning Tips  :
'To experience the maximum scanning speed of your machine with this module,
'it is advisable that you close other unnecessary programs running in the system.

'API Declarations
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32.dll" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hsnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hsnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long

'Modified API Declaration
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long
Private Declare Function TerminateProcess32 Lib "kernel32.dll" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'API Constants
Private Const ANYSIZE_ARRAY             As Long = 1
Private Const MAX_PATH                  As Long = 260
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_VM_READ           As Long = (&H10)
Private Const SE_DEBUG_NAME             As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED      As Long = &H2
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const TH32CS_SNAPPROCESS        As Long = &H2
Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const TOKEN_QUERY               As Long = &H8
Private Const PROCESS_ALL_ACCESS        As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

'API Types
Private Type LARGE_INTEGER
    lowpart     As Long
    highpart    As Long
End Type

Private Type LUID
    lowpart     As Long
    highpart    As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid       As LUID
    Attributes  As Long
End Type

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * MAX_PATH
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(ANYSIZE_ARRAY)   As LUID_AND_ATTRIBUTES
End Type

'Variable Declarations
Private m_Cancel        As Boolean
Private m_FullPathName  As String
Private m_NumProcesses  As Integer
Private m_ProcessID     As Long
Private m_Scanning      As Boolean
Private m_SystemProcess As Boolean

Event CurrentProcess(File As String, Path As String, ID As Long, Terminate As Boolean)
'Occurs everytime a process is scanned.
Event DoneScanning(TotalProcesses As Integer)
'Occures after scanning all processes.

#If False Then
    'Trick to preserve casing of these variables when used in VB IDE
    Private File, Path, ID, Terminate, TotalProcesses
#End If

Property Get SystemProcesses() As Boolean
'Returns/sets whether to include scanning for system processes.
    SystemProcesses = m_SystemProcess
End Property

Property Let SystemProcesses(Value As Boolean)
    m_SystemProcess = Value
End Property

Property Get CurrentProcessID() As Long
'Returns the Process ID of the currently scanned file.
    CurrentProcessID = m_ProcessID
End Property

Property Get CurrentProcessPath() As String
'Returns the path of the currently scanned process.
    CurrentProcessPath = m_FullPathName
End Property

Property Get Scanning() As Boolean
'Returns True if class is currently on scanning state.
    Scanning = m_Scanning
End Property

Property Get TotalProcesses() As Integer
'Returns the current or the total number of processes scanned.
    TotalProcesses = m_NumProcesses
End Property

'Procedures
Sub BeginScanning()
'Start scanning for running processes in the system.
    m_Cancel = False
    m_FullPathName = ""
    m_NumProcesses = 0
    
    m_Scanning = True
    Call ScanForProcesses
    m_Scanning = False
    
    m_Cancel = False
    m_FullPathName = ""
    m_ProcessID = -1
    RaiseEvent DoneScanning((m_NumProcesses))
End Sub

Sub CancelScanning()
'Abort scanning of running processes.
    m_Cancel = True
End Sub

Function TerminateProcess(Optional ProcessID As Long = -1) As Boolean
'Terminate a running process using the specified process ID.
'If ProcessID is not given, the function will attempt to terminate the currently scanned process.
    If ProcessID = -1 Then
        ProcessID = Me.CurrentProcessID 'Get ID of currently scanned process
        
        If ProcessID = -1 Then
            Exit Function
        End If
    End If
    
    Dim p_Privileges    As TOKEN_PRIVILEGES
    Dim p_Token         As Long
    
    'Windows NT/2000 requires special treatment to ensure that the
    'calling process has enough privileges to shutdown the system.
    
    If GetVersion >= 0 Then
        'Open token of the defined process
        If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_Token) = 0 Then
            GoTo End_Function
        End If
        
        'Get LUID used to locally represent the specified privilege name
        If LookupPrivilegeValue("", SE_DEBUG_NAME, p_Privileges.Privileges(ANYSIZE_ARRAY).pLuid) = 0 Then
            GoTo End_Function
        End If
        
        p_Privileges.PrivilegeCount = 1
        p_Privileges.Privileges(ANYSIZE_ARRAY).Attributes = SE_PRIVILEGE_ENABLED
        
        'Attempt to acquire debug privilege for the process
        If AdjustTokenPrivileges(p_Token, 0&, p_Privileges, 0&, 0&, 0&) = 0 Then
            GoTo End_Function
        End If
    End If
    
    'Finally, open the defined process
    Dim p_Process   As Long
    Dim p_Result    As Long
    p_Process = OpenProcess(PROCESS_ALL_ACCESS, 0&, ProcessID)
    
    If Not p_Process = 0 Then
        'Attempt to terminate the process
        p_Result = TerminateProcess32(p_Process, 0&)
        DoEvents
        TerminateProcess = (Not p_Result = 0)
        CloseHandle p_Process
    End If
    
    If GetVersion >= 0 Then
        'Restore original privilege
        p_Privileges.Privileges(ANYSIZE_ARRAY).Attributes = 0
        AdjustTokenPrivileges p_Token, 0&, p_Privileges, 0&, 0&, 0&
        
End_Function:
        If Not p_Token = 0 Then
            Call CloseHandle(p_Token)
        End If
    End If
End Function

'Private Procedures
Private Sub ScanForProcesses()
    DoEvents 'Process other system events before scanning for processes
    
    Dim p_Snapshot As Long
    p_Snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    
    If Not p_Snapshot = 0 Then
        Dim p_NullPos       As Integer
        Dim p_Process       As Long
        Dim p_ProcessEntry  As PROCESSENTRY32
        Dim p_ProcessID     As Long
        Dim p_ProcessName   As String
        Dim p_ProcessPath   As String
        Dim p_Terminate     As Boolean
        
        p_ProcessEntry.dwSize = Len(p_ProcessEntry)
        p_Process = Process32First(p_Snapshot, p_ProcessEntry)
        
        Do While p_Process 'Iterate through all processes
            If m_Cancel Then
                Exit Do 'Abort the whole scanning process
            End If
            
            DoEvents 'Ensure that the system processes other events
            p_ProcessName = p_ProcessEntry.szExeFile
            
            p_NullPos = InStr(1, p_ProcessName, vbNullChar)
            If Not p_NullPos = 0 Then
                p_ProcessName = Left(p_ProcessName, p_NullPos - 1)
            End If
            
            p_ProcessID = p_ProcessEntry.th32ProcessID
            p_ProcessPath = ProcessPath(p_ProcessID)
            
            If Not m_SystemProcess Then
                If p_ProcessPath = "SYSTEM" Then
                    GoTo Next_Process 'System processes not included. Get next process instead.
                End If
            End If
            
            m_ProcessID = p_ProcessID
            m_NumProcesses = m_NumProcesses + 1
            
            p_Terminate = False
            RaiseEvent CurrentProcess((p_ProcessName), (p_ProcessPath), (p_ProcessID), p_Terminate)
            If p_Terminate Then
                Call TerminateProcess(p_ProcessID)
            End If
Next_Process:
            p_ProcessEntry.szExeFile = String(MAX_PATH, 0)
            p_Process = Process32Next(p_Snapshot, p_ProcessEntry)
        Loop
    End If
End Sub

'Private Functions
Private Function ProcessPath(ByVal ID As Long) As String
    If m_Cancel Then
        Exit Function 'Abort the whole scanning process
    End If
    
    Dim p_Process As Long
    p_Process = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, ID)
    
    If Not p_Process = 0 Then
        Dim p_Modules(1 To 200) As Long
        Dim p_Needed            As Long
        Dim p_Ret               As Long
        
        p_Ret = EnumProcessModules(p_Process, p_Modules(1), 200, p_Needed)
        
        If Not p_Ret = 0 Then
            Dim p_ModuleName    As String * MAX_PATH
            Dim p_Size          As Long
            
            p_Size = 500
            p_Ret = GetModuleFileNameEx(p_Process, p_Modules(1), p_ModuleName, p_Size)
            
            ProcessPath = Left$(p_ModuleName, p_Ret)
            
            ProcessPath = Replace$(ProcessPath, "\??\", Empty)
            ProcessPath = Replace$(ProcessPath, "\SystemRoot\", WindowsDirectory & "\")
            
            m_FullPathName = ProcessPath 'Save full path with filename
            
            ProcessPath = Left$(ProcessPath, InStrRev(ProcessPath, "\")) 'Return without the filename
        End If
        
        p_Ret = CloseHandle(p_Process)
    End If
    
    If ProcessPath = Empty Then
        ProcessPath = "SYSTEM"
    End If
End Function

'Private Properties
Private Property Get WindowsDirectory()
    Dim p_Buffer    As String * MAX_PATH
    Dim p_NullPos   As Integer
    
    If Not GetWindowsDirectory(p_Buffer, MAX_PATH) = 0 Then
        p_NullPos = InStr(1, p_Buffer, vbNullChar)
        
        WindowsDirectory = Left(p_Buffer, p_NullPos - 1)
    End If
End Property

'Class Procedures
Private Sub Class_Initialize()
    m_FullPathName = Empty
    m_ProcessID = -1
    m_SystemProcess = False
End Sub

'Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines


