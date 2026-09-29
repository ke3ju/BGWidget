Attribute VB_Name = "modHistoryGraph"
Option Explicit

Private Type HISTORYPOINT
    lngUnixTime As Long
    intGlucose  As Integer
End Type

Private udtHistory()       As HISTORYPOINT
Private lngHistoryCount    As Long
Private boolHistoryLoaded  As Boolean

Private boolShowHistory    As Boolean
Private intHistoryHours    As Integer
Private intMinGlucose      As Integer
Private intMaxGlucose      As Integer
Private strLineColor       As String
Private strGridColor       As String
Private intLineWidth       As Integer


Public Function ReadHistoryINI(ByVal strINIFile As String) As Boolean
    Dim objHistoryINI As INIFile

    On Error GoTo ErrorHandler

    Set objHistoryINI = New INIFile

    objHistoryINI.Load strINIFile, "HistoryGraph"

    boolShowHistory = CBool(objHistoryINI.Value("Show"))
    intHistoryHours = CInt(objHistoryINI.Value("Hours"))
    intMinGlucose = CInt(objHistoryINI.Value("MinGlucose"))
    intMaxGlucose = CInt(objHistoryINI.Value("MaxGlucose"))
    strLineColor = CStr(objHistoryINI.Value("LineColor"))
    strGridColor = CStr(objHistoryINI.Value("GridColor"))
    intLineWidth = CInt(objHistoryINI.Value("LineWidth"))

    objHistoryINI.UnLoad

    Set objHistoryINI = Nothing

    ReadHistoryINI = True
    Exit Function

ErrorHandler:
    Set objHistoryINI = Nothing
    ReadHistoryINI = False
End Function


Public Function UpdateHistoryCache(ByVal strResponse As String) As Boolean

    If Not boolShowHistory Then Exit Function

    If Not boolHistoryLoaded Then
        boolHistoryLoaded = LoadHistoryCache(strResponse)
        UpdateHistoryCache = boolHistoryLoaded
    Else
        UpdateHistoryCache = MergeHistoryCache(strResponse)
    End If

End Function


Private Function LoadHistoryCache(ByVal strResponse As String) As Boolean
    Dim arrLines As Variant
    Dim arrRecord As Variant
    Dim strRecord As String
    Dim lngIndex As Long
    Dim lngLatest As Long
    Dim lngCutoff As Long
    Dim lngCount As Long
    Dim lngUnixTime As Long
    Dim intGlucose As Integer

    arrLines = Split(strResponse, vbLf)

    For lngIndex = UBound(arrLines) To 0 Step -1
        strRecord = Replace$(CStr(arrLines(lngIndex)), vbTab, "|")
        arrRecord = Split(strRecord, "|")
        If UBound(arrRecord) >= 6 Then
            If IsNumeric(arrRecord(2)) And IsNumeric(arrRecord(6)) Then
                lngLatest = CLng(arrRecord(2))
                Exit For
            End If
        End If
    Next

    If lngLatest = 0 Then Exit Function

    lngCutoff = lngLatest - (CLng(intHistoryHours) * 3600&)

    For lngIndex = 0 To UBound(arrLines)
        strRecord = Replace$(CStr(arrLines(lngIndex)), vbTab, "|")
        arrRecord = Split(strRecord, "|")
        If UBound(arrRecord) >= 6 Then
            If IsNumeric(arrRecord(2)) And IsNumeric(arrRecord(6)) Then
                lngUnixTime = CLng(arrRecord(2))
                intGlucose = CInt(arrRecord(6))
                If lngUnixTime >= lngCutoff And intGlucose > 0 Then lngCount = lngCount + 1
            End If
        End If
    Next

    If lngCount = 0 Then Exit Function

    ReDim udtHistory(1 To lngCount)
    lngCount = 0

    For lngIndex = 0 To UBound(arrLines)
        strRecord = Replace$(CStr(arrLines(lngIndex)), vbTab, "|")
        arrRecord = Split(strRecord, "|")
        If UBound(arrRecord) >= 6 Then
            If IsNumeric(arrRecord(2)) And IsNumeric(arrRecord(6)) Then
                lngUnixTime = CLng(arrRecord(2))
                intGlucose = CInt(arrRecord(6))
                If lngUnixTime >= lngCutoff And intGlucose > 0 Then
                    lngCount = lngCount + 1
                    udtHistory(lngCount).lngUnixTime = lngUnixTime
                    udtHistory(lngCount).intGlucose = intGlucose
                End If
            End If
        End If
    Next

    lngHistoryCount = lngCount
    LoadHistoryCache = True
End Function

Private Function MergeHistoryCache(ByVal strResponse As String) As Boolean
    Dim arrLines As Variant
    Dim arrRecord As Variant
    Dim strRecord As String
    Dim lngIndex As Long
    Dim lngUnixTime As Long
    Dim intGlucose As Integer
    Dim boolAdded As Boolean

    arrLines = Split(strResponse, vbLf)

    For lngIndex = 0 To UBound(arrLines)
        strRecord = Replace$(CStr(arrLines(lngIndex)), vbTab, "|")
        arrRecord = Split(strRecord, "|")

        If UBound(arrRecord) >= 6 Then
            If IsNumeric(arrRecord(2)) And IsNumeric(arrRecord(6)) Then
                lngUnixTime = CLng(arrRecord(2))
                intGlucose = CInt(arrRecord(6))

                If intGlucose > 0 Then
                    If AddHistoryPoint(lngUnixTime, intGlucose) Then boolAdded = True
                End If
            End If
        End If
    Next

    If boolAdded Then Call PurgeHistoryCache

    MergeHistoryCache = boolAdded
End Function


Private Function AddHistoryPoint(ByVal lngUnixTime As Long, ByVal intGlucose As Integer) As Boolean

    Dim lngInsertPos As Long
    Dim lngIndex As Long

    On Error GoTo ErrorHandler

    AddHistoryPoint = False

    ' First point in the history.
    If lngHistoryCount = 0 Then
        lngHistoryCount = 1

        ReDim udtHistory(1 To lngHistoryCount)

        udtHistory(1).lngUnixTime = lngUnixTime
        udtHistory(1).intGlucose = intGlucose

        AddHistoryPoint = True
        Exit Function
    End If

    '
    ' Find where this reading belongs.
    '
    ' History MUST remain in ascending timestamp order because
    ' DrawHistoryGraph connects each array entry to the next one.
    '
    lngInsertPos = 0

    For lngIndex = lngHistoryCount To 1 Step -1

        ' Same timestamp already exists.
        ' Replace the glucose value instead of adding another point.
        If udtHistory(lngIndex).lngUnixTime = lngUnixTime Then

            udtHistory(lngIndex).intGlucose = intGlucose

            AddHistoryPoint = True
            Exit Function

        End If

        ' We found the reading immediately before the new one.
        If udtHistory(lngIndex).lngUnixTime < lngUnixTime Then

            lngInsertPos = lngIndex + 1
            Exit For

        End If

    Next lngIndex

    '
    ' If we never found an earlier timestamp, this new reading
    ' belongs at the very beginning of the history.
    '
    If lngInsertPos = 0 Then
        lngInsertPos = 1
    End If

    '
    ' Make room for the new point.
    '
    lngHistoryCount = lngHistoryCount + 1

    ReDim Preserve udtHistory(1 To lngHistoryCount)

    '
    ' Shift everything after the insertion position one slot
    ' toward the end of the array.
    '
    For lngIndex = lngHistoryCount To lngInsertPos + 1 Step -1

        udtHistory(lngIndex).lngUnixTime = _
            udtHistory(lngIndex - 1).lngUnixTime

        udtHistory(lngIndex).intGlucose = _
            udtHistory(lngIndex - 1).intGlucose

    Next lngIndex

    '
    ' Insert the new reading in its correct chronological position.
    '
    udtHistory(lngInsertPos).lngUnixTime = lngUnixTime
    udtHistory(lngInsertPos).intGlucose = intGlucose

    AddHistoryPoint = True
    Exit Function


ErrorHandler:

    AddHistoryPoint = False

End Function


Private Function PurgeHistoryCache() As Boolean
    Dim lngLatest       As Long
    Dim lngCutoff       As Long
    Dim lngFirstRecord  As Long
    Dim lngIndex        As Long
    Dim lngNewIndex     As Long
    Dim udtNewHistory() As HISTORYPOINT

    If lngHistoryCount = 0 Then Exit Function

    lngLatest = udtHistory(lngHistoryCount).lngUnixTime
    lngCutoff = lngLatest - (CLng(intHistoryHours) * 3600&)

    For lngIndex = 1 To lngHistoryCount
        If udtHistory(lngIndex).lngUnixTime >= lngCutoff Then
            lngFirstRecord = lngIndex
            Exit For
        End If
    Next

    If lngFirstRecord <= 1 Then
        PurgeHistoryCache = True
        Exit Function
    End If

    ReDim udtNewHistory(1 To lngHistoryCount - lngFirstRecord + 1)

    For lngIndex = lngFirstRecord To lngHistoryCount
        lngNewIndex = lngNewIndex + 1

        udtNewHistory(lngNewIndex) = udtHistory(lngIndex)
    Next

    ReDim udtHistory(1 To lngNewIndex)

    For lngIndex = 1 To lngNewIndex
        udtHistory(lngIndex) = udtNewHistory(lngIndex)
    Next

    lngHistoryCount = lngNewIndex

    PurgeHistoryCache = True
End Function


Public Function HistoryCount() As Long
    HistoryCount = lngHistoryCount
End Function


Public Function HistoryUnixTime(ByVal lngIndex As Long) As Long

    If lngIndex < 1 Or lngIndex > lngHistoryCount Then Exit Function

    HistoryUnixTime = udtHistory(lngIndex).lngUnixTime

End Function


Public Function HistoryGlucose(ByVal lngIndex As Long) As Integer

    If lngIndex < 1 Or lngIndex > lngHistoryCount Then Exit Function

    HistoryGlucose = udtHistory(lngIndex).intGlucose

End Function


Public Function ShowHistoryGraph() As Boolean
    ShowHistoryGraph = boolShowHistory
End Function


Public Function HistoryHours() As Integer
    HistoryHours = intHistoryHours
End Function


Public Function HistoryMinGlucose() As Integer
    HistoryMinGlucose = intMinGlucose
End Function


Public Function HistoryMaxGlucose() As Integer
    HistoryMaxGlucose = intMaxGlucose
End Function


Public Function HistoryLineColor() As Long
    HistoryLineColor = HexRGBToVBColor(strLineColor)
End Function


Public Function HistoryGridColor() As Long
    HistoryGridColor = HexRGBToVBColor(strGridColor)
End Function


Public Function HistoryLineWidth() As Integer
    HistoryLineWidth = intLineWidth
End Function


Private Function HexRGBToVBColor(ByVal strColor As String) As Long
    Dim lngColor As Long

    On Error GoTo ErrorHandler

    strColor = Replace$(Trim$(strColor), "#", "")

    If Left$(UCase$(strColor), 2) = "&H" Then
        strColor = Mid$(strColor, 3)
    End If

    If Len(strColor) <> 6 Then GoTo ErrorHandler

    lngColor = CLng("&H" & strColor)

    HexRGBToVBColor = RGB( _
        (lngColor \ &H10000) And &HFF&, _
        (lngColor \ &H100&) And &HFF&, _
        lngColor And &HFF&)

    Exit Function

ErrorHandler:
    HexRGBToVBColor = vbWhite
End Function

Public Function DrawHistoryGraphs() As Boolean
    Dim frmItem As Form

    On Error GoTo ErrorHandler

    If Not boolShowHistory Then Exit Function
    If lngHistoryCount < 2 Then Exit Function

    For Each frmItem In Forms
        If TypeName(frmItem) = "frmSplash" Then
            Call DrawHistoryGraph(frmItem)
        End If
    Next

    DrawHistoryGraphs = True
    Exit Function

ErrorHandler:
    DrawHistoryGraphs = False
End Function


Public Function DrawHistoryGraph(ByVal frmTarget As Form) As Boolean
    Dim lngIndex       As Long

    Dim dblGraphLeft   As Double
    Dim dblGraphTop    As Double
    Dim dblGraphRight  As Double
    Dim dblGraphBottom As Double

    Dim dblX1          As Double
    Dim dblY1          As Double
    Dim dblX2          As Double
    Dim dblY2          As Double

    Dim lngLatestTime  As Long
    Dim lngStartTime   As Long

    Dim intGlucose     As Integer
    Dim intGridValue   As Integer

    On Error GoTo ErrorHandler

    If Not boolShowHistory Then Exit Function
    If lngHistoryCount < 2 Then Exit Function

    frmTarget.AutoRedraw = True
    frmTarget.Cls

    dblGraphLeft = frmTarget.ScaleWidth * 0.08
    dblGraphRight = frmTarget.ScaleWidth * 0.92

    dblGraphTop = frmTarget.ScaleHeight * 0.15
    dblGraphBottom = frmTarget.ScaleHeight * 0.85

    lngLatestTime = udtHistory(lngHistoryCount).lngUnixTime
    lngStartTime = lngLatestTime - (CLng(intHistoryHours) * 3600&)

    frmTarget.DrawWidth = 1
    frmTarget.DrawStyle = vbSolid

    For intGridValue = intMinGlucose To intMaxGlucose Step 50

        dblY1 = GraphY( _
            intGridValue, _
            dblGraphTop, _
            dblGraphBottom)

        frmTarget.Line _
            (dblGraphLeft, dblY1)- _
            (dblGraphRight, dblY1), _
            HistoryGridColor

    Next

    frmTarget.DrawWidth = intLineWidth

    For lngIndex = 2 To lngHistoryCount

        If udtHistory(lngIndex - 1).lngUnixTime >= lngStartTime Then

            dblX1 = GraphX( _
                udtHistory(lngIndex - 1).lngUnixTime, _
                lngStartTime, _
                lngLatestTime, _
                dblGraphLeft, _
                dblGraphRight)

            dblX2 = GraphX( _
                udtHistory(lngIndex).lngUnixTime, _
                lngStartTime, _
                lngLatestTime, _
                dblGraphLeft, _
                dblGraphRight)

            intGlucose = udtHistory(lngIndex - 1).intGlucose

            dblY1 = GraphY( _
                intGlucose, _
                dblGraphTop, _
                dblGraphBottom)

            intGlucose = udtHistory(lngIndex).intGlucose

            dblY2 = GraphY( _
                intGlucose, _
                dblGraphTop, _
                dblGraphBottom)

            frmTarget.Line _
                (dblX1, dblY1)- _
                (dblX2, dblY2), _
                HistoryLineColor

        End If

    Next

    frmTarget.DrawWidth = 1
    frmTarget.DrawStyle = vbSolid

    frmTarget.Refresh

    DrawHistoryGraph = True
    Exit Function

ErrorHandler:
    DrawHistoryGraph = False
End Function


Private Function GraphX( _
    ByVal lngUnixTime As Long, _
    ByVal lngStartTime As Long, _
    ByVal lngEndTime As Long, _
    ByVal dblLeft As Double, _
    ByVal dblRight As Double) As Double

    Dim dblPercent As Double

    If lngEndTime <= lngStartTime Then
        GraphX = dblLeft
        Exit Function
    End If

    dblPercent = _
        CDbl(lngUnixTime - lngStartTime) / _
        CDbl(lngEndTime - lngStartTime)

    If dblPercent < 0 Then dblPercent = 0
    If dblPercent > 1 Then dblPercent = 1

    GraphX = dblLeft + _
             ((dblRight - dblLeft) * dblPercent)

End Function


Private Function GraphY( _
    ByVal intGlucose As Integer, _
    ByVal dblTop As Double, _
    ByVal dblBottom As Double) As Double

    Dim dblPercent As Double

    If intGlucose < intMinGlucose Then intGlucose = intMinGlucose
    If intGlucose > intMaxGlucose Then intGlucose = intMaxGlucose

    dblPercent = _
        CDbl(intGlucose - intMinGlucose) / _
        CDbl(intMaxGlucose - intMinGlucose)

    GraphY = dblBottom - _
             ((dblBottom - dblTop) * dblPercent)

End Function
