Attribute VB_Name = "mIOIPAOTreeView"
Option Explicit

Public Type IPAOHookStructTreeView
    lpVTable    As Long
    IPAOReal    As IOleInPlaceActiveObject
    Ctl         As ucTreeView
    ThisPointer As Long
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As GUID, iid2 As GUID) As Long

Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Private Const S_FALSE               As Long = 1
Private Const S_OK                  As Long = 0

Private IID_IOleInPlaceActiveObject As GUID
Private m_IPAOVTable(9)             As Long

Public Sub InitIPAO(IPAOHookStruct As IPAOHookStructTreeView, Ctl As ucTreeView)
    
  Dim IPAO As IOleInPlaceActiveObject
    
    With IPAOHookStruct
        Set IPAO = Ctl
        Call CopyMemory(.IPAOReal, IPAO, 4)
        Call CopyMemory(.Ctl, Ctl, 4)
        .lpVTable = GetVTable
        .ThisPointer = VarPtr(IPAOHookStruct)
    End With
End Sub

Public Sub TerminateIPAO(IPAOHookStruct As IPAOHookStructTreeView)
    With IPAOHookStruct
        Call CopyMemory(.IPAOReal, 0&, 4)
        Call CopyMemory(.Ctl, 0&, 4)
    End With
End Sub

Private Function GetVTable() As Long

    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = AddressOfFunction(AddressOf QueryInterface)
        m_IPAOVTable(1) = AddressOfFunction(AddressOf AddRef)
        m_IPAOVTable(2) = AddressOfFunction(AddressOf Release)
        m_IPAOVTable(3) = AddressOfFunction(AddressOf GetWindow)
        m_IPAOVTable(4) = AddressOfFunction(AddressOf ContextSensitiveHelp)
        m_IPAOVTable(5) = AddressOfFunction(AddressOf TranslateAccelerator)
        m_IPAOVTable(6) = AddressOfFunction(AddressOf OnFrameWindowActivate)
        m_IPAOVTable(7) = AddressOfFunction(AddressOf OnDocWindowActivate)
        m_IPAOVTable(8) = AddressOfFunction(AddressOf ResizeBorder)
        m_IPAOVTable(9) = AddressOfFunction(AddressOf EnableModeless)
        '--- init guid
        With IID_IOleInPlaceActiveObject
            .Data1 = &H117&
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function

Private Function AddressOfFunction(lpfn As Long) As Long
    AddressOfFunction = lpfn
End Function

Private Function AddRef(This As IPAOHookStructTreeView) As Long
    AddRef = This.IPAOReal.AddRef
End Function

Private Function Release(This As IPAOHookStructTreeView) As Long
    Release = This.IPAOReal.Release
End Function

Private Function QueryInterface(This As IPAOHookStructTreeView, riid As GUID, pvObj As Long) As Long
    If (IsEqualGUID(riid, IID_IOleInPlaceActiveObject)) Then
        pvObj = This.ThisPointer
        AddRef This
        QueryInterface = 0
      Else
        QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
    End If
End Function

Private Function GetWindow(This As IPAOHookStructTreeView, phwnd As Long) As Long
End Function

Private Function ContextSensitiveHelp(This As IPAOHookStructTreeView, ByVal fEnterMode As Long) As Long
End Function

Private Function TranslateAccelerator(This As IPAOHookStructTreeView, lpMsg As Msg) As Long
    If (This.Ctl.frTranslateAccel(lpMsg)) Then
        TranslateAccelerator = S_OK
      Else
        TranslateAccelerator = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
    End If
End Function

Private Function OnFrameWindowActivate(This As IPAOHookStructTreeView, ByVal fActivate As Long) As Long
    OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function OnDocWindowActivate(This As IPAOHookStructTreeView, ByVal fActivate As Long) As Long
    OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function ResizeBorder(This As IPAOHookStructTreeView, prcBorder As RECT, ByVal puiWindow As IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
    ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
End Function

Private Function EnableModeless(This As IPAOHookStructTreeView, ByVal fEnable As Long) As Long
    EnableModeless = This.IPAOReal.EnableModeless(fEnable)
End Function


