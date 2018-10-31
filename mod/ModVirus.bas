Attribute VB_Name = "ModVirus"
Public Function isProperFile(sPath As String, limitSizeMB As Integer, sExt As String) As Boolean
On Error Resume Next
If (limitSizeMB * 1024 * 1024) > FileLen(sPath) Then
    If InStr(1, UCase(sExt), UCase(Right(sPath, 3))) > 0 Then
        isProperFile = True
    Else
        isProperFile = False
    End If
Else
    isProperFile = False
End If
End Function
Public Function OpenTeks(sFile As String) As String
Dim sTemp As String
Open sFile For Binary As #1
    sTemp = Space(LOF(1))
    Get #1, , sTemp
Close #1
OpenTeks = sTemp
End Function
