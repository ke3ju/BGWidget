Attribute VB_Name = "modWidgetPlacement"
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type MONITORINFOEX
    cbSize    As Long
    rcMonitor As RECT
    rcWork    As RECT
    dwFlags   As Long
    szDevice  As String * 32
End Type

Private Type MONITORDATA
    intDisplayNumber As Integer
    boolPrimary      As Boolean
    lngLeft          As Long
    lngTop           As Long
    lngRight         As Long
    lngBottom        As Long
End Type

Private Declare Function EnumDisplayMonitors Lib "user32" ( _
    ByVal hdc As Long, _
    ByVal lprcClip As Long, _
    ByVal lpfnEnum As Long, _
    ByVal dwData As Long) As Long

Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
    ByVal hMonitor As Long, _
    lpmi As MONITORINFOEX) As Long

Private colMirrorWidgets As Collection

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = &HFFFFFFEC
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&

Private Const MONITORINFOF_PRIMARY = &H1
Private Const MAX_MONITORS = 16

Private udtMonitors(1 To MAX_MONITORS) As MONITORDATA
Private intMonitorCount               As Integer

Private strPlacementMonitor           As String
Private strPlacementPosition          As String
Private lngPlacementMarginX           As Long
Private lngPlacementMarginY           As Long


Public Function ReadPlacementINI(ByVal strINIFile As String) As Boolean
    Dim objPlacementINI As INIFile

    On Error GoTo ErrorHandler

    Set objPlacementINI = New INIFile

    objPlacementINI.Load strINIFile, "Placement"

    strPlacementMonitor = Trim$(CStr(objPlacementINI.Value("Monitor")))
    strPlacementPosition = Trim$(CStr(objPlacementINI.Value("Position")))
    lngPlacementMarginX = CLng(objPlacementINI.Value("MarginX"))
    lngPlacementMarginY = CLng(objPlacementINI.Value("MarginY"))

    objPlacementINI.UnLoad

    Set objPlacementINI = Nothing

    ReadPlacementINI = True
    Exit Function

ErrorHandler:
    Set objPlacementINI = Nothing
    ReadPlacementINI = False
End Function


Public Function ApplyWidgetPlacement(ByVal frmTarget As Form) As Boolean
    Dim intDisplayNumber As Integer

    On Error GoTo ErrorHandler

    If LCase$(Trim$(strPlacementMonitor)) = "all" Then
        intDisplayNumber = 1
    Else
        intDisplayNumber = CInt(strPlacementMonitor)
    End If

    ApplyWidgetPlacement = ApplyWidgetPlacementToMonitor(frmTarget, intDisplayNumber)

    Exit Function

ErrorHandler:
    MsgBox "ApplyWidgetPlacement Error " & Err.Number & vbCrLf & Err.Description, vbExclamation, "BGWidget"
End Function

Public Function ApplyWidgetPlacementToMonitor(ByVal frmTarget As Form, ByVal intDisplayNumber As Integer) As Boolean
    Dim intMonitorIndex As Integer
    Dim lngFormWidth    As Long
    Dim lngFormHeight   As Long
    Dim lngLeft         As Long
    Dim lngTop          As Long

    On Error GoTo ErrorHandler

    Call EnumerateMonitors

    intMonitorIndex = FindMonitorIndex(intDisplayNumber)

    If intMonitorIndex = 0 Then Exit Function

    lngFormWidth = frmTarget.Width \ Screen.TwipsPerPixelX
    lngFormHeight = frmTarget.Height \ Screen.TwipsPerPixelY

    Select Case LCase$(Trim$(strPlacementPosition))
        Case "topleft"
            lngLeft = udtMonitors(intMonitorIndex).lngLeft + lngPlacementMarginX
            lngTop = udtMonitors(intMonitorIndex).lngTop + lngPlacementMarginY

        Case "topright"
            lngLeft = udtMonitors(intMonitorIndex).lngRight - lngFormWidth - lngPlacementMarginX
            lngTop = udtMonitors(intMonitorIndex).lngTop + lngPlacementMarginY

        Case "bottomleft"
            lngLeft = udtMonitors(intMonitorIndex).lngLeft + lngPlacementMarginX
            lngTop = udtMonitors(intMonitorIndex).lngBottom - lngFormHeight - lngPlacementMarginY

        Case "bottomright"
            lngLeft = udtMonitors(intMonitorIndex).lngRight - lngFormWidth - lngPlacementMarginX
            lngTop = udtMonitors(intMonitorIndex).lngBottom - lngFormHeight - lngPlacementMarginY

        Case "center"
            lngLeft = udtMonitors(intMonitorIndex).lngLeft + _
                      ((udtMonitors(intMonitorIndex).lngRight - udtMonitors(intMonitorIndex).lngLeft - lngFormWidth) \ 2)

            lngTop = udtMonitors(intMonitorIndex).lngTop + _
                     ((udtMonitors(intMonitorIndex).lngBottom - udtMonitors(intMonitorIndex).lngTop - lngFormHeight) \ 2)

            lngLeft = lngLeft + lngPlacementMarginX
            lngTop = lngTop + lngPlacementMarginY

        Case Else
            Exit Function
    End Select

    frmTarget.Left = CLng(lngLeft) * CLng(Screen.TwipsPerPixelX)
    frmTarget.Top = CLng(lngTop) * CLng(Screen.TwipsPerPixelY)

    ApplyWidgetPlacementToMonitor = True
    Exit Function

ErrorHandler:
    ApplyWidgetPlacementToMonitor = False
End Function

Public Function CreateMirrorWidgets(ByVal frmMaster As frmSplash) As Boolean
    Dim intIndex  As Integer
    Dim objMirror As frmSplash

    On Error GoTo ErrorHandler

    If LCase$(Trim$(strPlacementMonitor)) <> "all" Then
        CreateMirrorWidgets = True
        Exit Function
    End If

    Call EnumerateMonitors

    Set colMirrorWidgets = New Collection

    For intIndex = 1 To intMonitorCount
        If udtMonitors(intIndex).intDisplayNumber <> 1 Then

            Set objMirror = New frmSplash

            objMirror.boolMirrorMode = True
            objMirror.intMirrorMonitor = udtMonitors(intIndex).intDisplayNumber

            colMirrorWidgets.Add objMirror

            objMirror.Show
        End If
    Next

    CreateMirrorWidgets = True
    Exit Function

ErrorHandler:
    CreateMirrorWidgets = False
End Function

Public Function SyncMirrorWidget(ByVal frmMaster As frmSplash, ByVal frmMirror As frmSplash) As Boolean

    frmMirror.lblCurrentGlucose.Caption = frmMaster.lblCurrentGlucose.Caption
    frmMirror.lblCurrentGlucose.ForeColor = frmMaster.lblCurrentGlucose.ForeColor
    frmMirror.lblCurrentGlucose.Visible = frmMaster.lblCurrentGlucose.Visible

    frmMirror.lblA1cMGDL.Caption = frmMaster.lblA1cMGDL.Caption
    frmMirror.lblA1cMGDL.ForeColor = frmMaster.lblA1cMGDL.ForeColor
    frmMirror.lblA1cMGDL.Visible = frmMaster.lblA1cMGDL.Visible

    frmMirror.lblA1cMMOLL.Caption = frmMaster.lblA1cMMOLL.Caption
    frmMirror.lblA1cMMOLL.ForeColor = frmMaster.lblA1cMMOLL.ForeColor
    frmMirror.lblA1cMMOLL.Visible = frmMaster.lblA1cMMOLL.Visible

    frmMirror.lblRecordNumInDays.Caption = frmMaster.lblRecordNumInDays.Caption
    frmMirror.lblRecordNumInDays.ForeColor = frmMaster.lblRecordNumInDays.ForeColor
    frmMirror.lblRecordNumInDays.Visible = frmMaster.lblRecordNumInDays.Visible

    frmMirror.lblLastUpdate.Caption = frmMaster.lblLastUpdate.Caption
    frmMirror.lblLastUpdate.ForeColor = frmMaster.lblLastUpdate.ForeColor
    frmMirror.lblLastUpdate.Visible = frmMaster.lblLastUpdate.Visible

    SyncMirrorWidget = True
End Function

Public Function SyncMirrorWidgets(ByVal frmMaster As frmSplash) As Boolean
    Dim objMirror As frmSplash

    If colMirrorWidgets Is Nothing Then
        SyncMirrorWidgets = True
        Exit Function
    End If

    For Each objMirror In colMirrorWidgets
        Call SyncMirrorWidget(frmMaster, objMirror)
    Next

    SyncMirrorWidgets = True
End Function

Public Function MakeMirrorTransparent(ByVal frmTarget As Form) As Boolean

    frmTarget.BackColor = vbBlack

    SetWindowLong frmTarget.hwnd, _
                  GWL_EXSTYLE, _
                  GetWindowLong(frmTarget.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

    SetLayeredWindowAttributes frmTarget.hwnd, vbBlack, 0&, LWA_COLORKEY

    MakeMirrorTransparent = True
End Function

Private Function EnumerateMonitors() As Integer
    intMonitorCount = 0

    Call EnumDisplayMonitors(0, 0, AddressOf MonitorEnumProc, 0)

    EnumerateMonitors = intMonitorCount
End Function


Public Function MonitorEnumProc( _
    ByVal hMonitor As Long, _
    ByVal hdcMonitor As Long, _
    ByVal lprcMonitor As Long, _
    ByVal dwData As Long) As Long

    Dim udtInfo As MONITORINFOEX

    udtInfo.cbSize = Len(udtInfo)

    If GetMonitorInfo(hMonitor, udtInfo) <> 0 Then

        If intMonitorCount < MAX_MONITORS Then
            intMonitorCount = intMonitorCount + 1

            udtMonitors(intMonitorCount).intDisplayNumber = GetDisplayNumber(udtInfo.szDevice)
            udtMonitors(intMonitorCount).boolPrimary = ((udtInfo.dwFlags And MONITORINFOF_PRIMARY) <> 0)

            udtMonitors(intMonitorCount).lngLeft = udtInfo.rcWork.Left
            udtMonitors(intMonitorCount).lngTop = udtInfo.rcWork.Top
            udtMonitors(intMonitorCount).lngRight = udtInfo.rcWork.Right
            udtMonitors(intMonitorCount).lngBottom = udtInfo.rcWork.Bottom
        End If

    End If

    MonitorEnumProc = 1
End Function


Private Function FindMonitorIndex(ByVal intDisplayNumber As Integer) As Integer
    Dim intIndex As Integer

    For intIndex = 1 To intMonitorCount
        If udtMonitors(intIndex).intDisplayNumber = intDisplayNumber Then
            FindMonitorIndex = intIndex
            Exit Function
        End If
    Next

    For intIndex = 1 To intMonitorCount
        If udtMonitors(intIndex).boolPrimary Then
            FindMonitorIndex = intIndex
            Exit Function
        End If
    Next
End Function


Private Function GetDisplayNumber(ByVal strDevice As String) As Integer
    Dim intNull    As Integer
    Dim intDisplay As Integer

    intNull = InStr(1, strDevice, vbNullChar)

    If intNull > 0 Then
        strDevice = Left$(strDevice, intNull - 1)
    End If

    intDisplay = InStr(1, UCase$(strDevice), "DISPLAY")

    If intDisplay > 0 Then
        GetDisplayNumber = CInt(Val(Mid$(strDevice, intDisplay + 7)))
    End If
End Function
