Attribute VB_Name = "ModScan"

Public StopScan, StopRTP, BufferStop As Boolean
Public jumlahFile, JumlahBuffer As Long
Public jumlahfolder As Long

Public Sub BufferAntivirus(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, DirCount As Long
    Dim RealPath As String, GetViri As String
    
    GetViri = ""
    DirCount = -1
    
    ScanInfo = "Scan File"
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If BufferStop = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    jumlahfolder = jumlahfolder + 1
                    If SubDirs = True Then
                        DirCount = DirCount + 1
                        ReDim Preserve tempDir(DirCount) As String
                        tempDir(DirCount) = RealPath & CurItem
                    End If
                Else
                    JumlahBuffer = JumlahBuffer + 1
                    frmMain.lblFile.Caption = JumlahBuffer
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
        
        If SubDirs = True Then
            If DirCount <> -1 Then
                For i = 0 To DirCount
                    BufferAntivirus tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub
Public Sub EngineAntivirus(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, DirCount As Long
    Dim RealPath As String, GetViri As String
    
    GetViri = ""
    DirCount = -1
    
    ScanInfo = "Scan File"
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If StopScan = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    jumlahDir = jumlahDir + 1
                    frmMain.lblfolder.Caption = jumlahDir
                    If SubDirs = True Then
                        DirCount = DirCount + 1
                        ReDim Preserve tempDir(DirCount) As String
                        tempDir(DirCount) = RealPath & CurItem
                    End If
                Else
                    jumlahFile = jumlahFile + 1
                    'frmMain.lblFile.Caption = jumlahFile
                    If Len(RealPath) > 50 Then 'jika panjang nama file > 50
                       If Len(CurItem) < 15 Then
                          frmMain.lblscan.Caption = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & CurItem
                       Else
                          frmMain.lblscan.Caption = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & "..." & Right(CurItem, 15)
                       End If
                    End If
                       If isProperFile(CStr(RealPath & CurItem), 3, "COM") = True Then
                          If InStr(1, OpenTeks(RealPath & CurItem), "EICAR") > 0 Then
                          Dim LV As ListItem
                          Set LV = frmMain.lvMalware.ListItems.Add(, , "EICAR-VIRUS-TEST!!")
                          LV.SubItems(1) = FileLen(RealPath & CurItem) & " b"
                          LV.SubItems(2) = RealPath & CurItem
                          End If
                       End If
                   If frmMain.lvMalware.ListItems.Count <> 0 Then
                   frmMain.lblvirus.Caption = frmMain.lvMalware.ListItems.Count
                   End If
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
        
        If SubDirs = True Then
            If DirCount <> -1 Then
                For i = 0 To DirCount
                    EngineAntivirus tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub
Public Sub EngineAntivirusForRTP(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, DirCount As Long
    Dim RealPath As String, GetViri As String
    
    GetViri = ""
    DirCount = -1
    
    ScanInfo = "Scan File"
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If StopRTP = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    'jumlahDir = jumlahDir + 1
                    If SubDirs = True Then
                        DirCount = DirCount + 1
                        ReDim Preserve tempDir(DirCount) As String
                        tempDir(DirCount) = RealPath & CurItem
                    End If
                Else
                    'jumlahFile = jumlahFile + 1
                       If isProperFile(CStr(RealPath & CurItem), 3, "COM") = True Then
                          If InStr(1, OpenTeks(RealPath & CurItem), "EICAR") > 0 Then
                          Dim LV As ListItem
                          Set LV = frmRTP.lvMalware.ListItems.Add(, , "EICAR-VIRUS-TEST!!")
                          LV.SubItems(1) = FileLen(RealPath & CurItem) & " b"
                          LV.SubItems(2) = RealPath & CurItem
                          End If
                       End If
                   If frmRTP.lvMalware.ListItems.Count <> 0 Then
                   frmRTP.Caption = "Real Time Protector >> " & frmRTP.lvMalware.ListItems.Count & " Virus"
                   End If
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
        
        If SubDirs = True Then
            If DirCount <> -1 Then
                For i = 0 To DirCount
                    EngineAntivirusForRTP tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub

