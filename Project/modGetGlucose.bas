Attribute VB_Name = "modGetGlucose"
'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME:
'
' AUTHOR: Edward A. Ludwig , EDSIT Enterprises
' DATE  : 3/5/2025
'
' COMMENT:
'
'==========================================================================
Option Explicit

Dim objURL                As Object
Dim strURL                As String
Dim strResponse           As String
Dim strRecord             As String
Dim dblCombinedA1c        As Double
Dim dblNumberOfRecords    As Double
Dim intCurrentGlucose     As Integer

Public strINIFile         As String
Public boolPostLaunch     As Boolean
Public objReadWrite       As Object
Public objINI             As Object
Public intA1cInterval     As Integer
Public boolDebug          As Boolean

Dim dblSeconds            As Double
Dim dblMinutes            As Double
Dim dblHours              As Double
Dim dtLastReading         As Date
Dim intDays               As Integer
Dim intDaysOfData         As Integer
Dim intLastBG             As Integer
Dim intDelta              As Integer
Dim intGetGlucoseInterval As Integer
Dim strDelta              As String
Dim varRateOfChange       As Variant
Dim varLastA1cTime        As Variant

Dim strCurrentGlucose     As String
Dim strhbA1c              As String
Dim strRecNum             As String
Dim strLastUpdate         As String
Dim strDataPath           As String

Dim boolShowGlucose       As Boolean
Dim boolShowAverageGlucose As Boolean
Dim boolShowA1c           As Boolean
Dim boolShowDaysOfRecords As Boolean
Dim boolShowLastUpdate    As Boolean
Dim boolUseRelativeTime   As Boolean

Dim strGlucoseColor       As String
Dim strA1cColor           As String
Dim strDaysOfRecordsColor As String
Dim strLastUpdateColor    As String

Const BOTH = -2

Sub Main()
    MsgBox GetGlucose
End Sub

Function ReadINI(strINIFile As String, strSection As String) As String
    objINI.Load strINIFile, strSection

    strDataPath = objINI.Value("DataPath")

    intGetGlucoseInterval = CInt(objINI.Value("GetGlucoseInterval"))
    intA1cInterval = CInt(objINI.Value("IntervalToCalculateA1cInMinutes"))

    boolShowGlucose = CBool(objINI.Value("ShowGlucose"))
    boolShowAverageGlucose = CBool(objINI.Value("ShowAverageGlucose"))
    boolShowA1c = CBool(objINI.Value("ShowA1c"))
    boolShowDaysOfRecords = CBool(objINI.Value("ShowDaysOfRecords"))
    boolShowLastUpdate = CBool(objINI.Value("ShowLastUpdate"))
    boolUseRelativeTime = CBool(objINI.Value("UseRelativeTime"))

    objINI.UnLoad

    objINI.Load strINIFile, "GlucoseColor"
    strGlucoseColor = CStr(objINI.Value("FontColor"))
    objINI.UnLoad

    objINI.Load strINIFile, "hbA1cColor"
    strA1cColor = CStr(objINI.Value("FontColor"))
    objINI.UnLoad

    objINI.Load strINIFile, "DaysOfRecordsColor"
    strDaysOfRecordsColor = CStr(objINI.Value("FontColor"))
    objINI.UnLoad

    objINI.Load strINIFile, "LastUpdateColor"
    strLastUpdateColor = CStr(objINI.Value("FontColor"))
    objINI.UnLoad
End Function

Private Function GetTrendArrow() As String
    Dim dblRate As Double

    If Not IsNumeric(varRateOfChange) Then
        GetTrendArrow = ""
        Exit Function
    End If

    dblRate = CDbl(varRateOfChange)

    If dblRate >= 2 Then
        GetTrendArrow = ChrW$(&H2191)
    ElseIf dblRate >= 1 Then
        GetTrendArrow = ChrW$(&H2197)
    ElseIf dblRate > -1 Then
        GetTrendArrow = ChrW$(&H2192)
    ElseIf dblRate > -2 Then
        GetTrendArrow = ChrW$(&H2198)
    Else
        GetTrendArrow = ChrW$(&H2193)
    End If
End Function

Function ApplySettings()
    frmSplash.tmrGetCurrentGlucose.Interval = CLng(intGetGlucoseInterval) * 1000&
    frmSplash.tmrA1cTimer.Interval = 60000

    frmSplash.lblCurrentGlucose.Visible = boolShowGlucose
    frmSplash.lblA1cMGDL.Visible = boolShowAverageGlucose
    frmSplash.lblA1cMMOLL.Visible = boolShowA1c
    frmSplash.lblRecordNumInDays.Visible = boolShowDaysOfRecords
    frmSplash.lblLastUpdate.Visible = boolShowLastUpdate
End Function

Function ApplyColors()
    frmSplash.lblCurrentGlucose.ForeColor = HexRGBToVBColor(strGlucoseColor)

    frmSplash.lblA1cMGDL.ForeColor = HexRGBToVBColor(strA1cColor)
    frmSplash.lblA1cMMOLL.ForeColor = HexRGBToVBColor(strA1cColor)

    frmSplash.lblRecordNumInDays.ForeColor = HexRGBToVBColor(strDaysOfRecordsColor)
    frmSplash.lblLastUpdate.ForeColor = HexRGBToVBColor(strLastUpdateColor)
End Function

Private Function HexRGBToVBColor(ByVal strColor As String) As Long
    Dim lngColor As Long

    On Error GoTo ErrorHandler

    strColor = Replace$(Trim$(strColor), "#", "")
    If Left$(UCase$(strColor), 2) = "&H" Then strColor = Mid$(strColor, 3)

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

Function GetGlucose() As String
    Dim boolNewReading As Boolean

    boolDebug = False
    Set objURL = CreateObject("Msxml2.ServerXMLHTTP.6.0")

    intDaysOfData = 90
    intDays = intDaysOfData
    dblHours = intDays * 24
    dblMinutes = dblHours * 60
    dblSeconds = dblMinutes * 60

    If Not boolPostLaunch Then
        If boolDebug Then MsgBox "Not boolPostLaunch"

        dblCombinedA1c = 0
        dblNumberOfRecords = 0

        strURL = strDataPath & "/x/stream?header&duration=" & dblSeconds
        boolNewReading = GetA1c(strURL)

        If intCurrentGlucose > 0 Then strCurrentGlucose = CStr(intCurrentGlucose)

        strhbA1c = CStr(CalculateA1c)
        strRecNum = CStr(intDaysOfData)

        If boolNewReading Then strLastUpdate = Time

        varLastA1cTime = Time

        If Len(strCurrentGlucose) > 0 Then
            If IsNumeric(strCurrentGlucose) Then
                If CInt(strCurrentGlucose) > 0 Then
                    intLastBG = CInt(strCurrentGlucose)
                    intDelta = 0
                    strDelta = "0"
                End If
            End If
        End If
    Else
        If boolDebug Then MsgBox "Else boolPostLaunch"

        strURL = strDataPath & "/x/stream?header&duration=60"

        If boolDebug Then MsgBox "Call GetA1c()"

        boolNewReading = GetA1c(strURL)

        If boolNewReading Then
            strCurrentGlucose = CStr(intCurrentGlucose)

            If intLastBG = 0 Then intLastBG = CInt(strCurrentGlucose)

            intDelta = CInt(strCurrentGlucose) - intLastBG

            If intDelta > 0 Then
                strDelta = "+" & CStr(intDelta)
            Else
                strDelta = CStr(intDelta)
            End If

            intLastBG = CInt(strCurrentGlucose)
            strLastUpdate = Time
        End If
    End If

    boolPostLaunch = True

    frmSplash.lblCurrentGlucose = strCurrentGlucose & " mg/dL " & strDelta
    frmSplash.lblA1cMGDL = "Avg BG " & Split(strhbA1c, "|")(0) & " mg/dL"
    frmSplash.lblA1cMMOLL = "hbA1c " & Split(strhbA1c, "|")(2) & "%"
    frmSplash.lblRecordNumInDays = "Total Days: " & strRecNum
    frmSplash.lblLastUpdate = FormatLastReadingTime()

    Call modWidgetPlacement.SyncMirrorWidgets(frmSplash)
End Function

Function CalculateA1c()
    Dim dblAverageGlucose         As Double
    Dim dblGMI                    As Double
    Dim dblA1c                   As Double

    Dim strLogPath               As String
    Dim intFile                  As Integer

    Dim strLine                  As Variant
    Dim arrRecord                As Variant
    Dim strCleanLine             As String
    Dim strDateKey               As String
    Dim strYYYYMMDD              As String
    Dim strCutoffKey             As String

    Dim lngTotalLines            As Long
    Dim lngDataLines             As Long
    Dim lngAccepted              As Long
    Dim lngTooOld                As Long
    Dim lngMalformed             As Long

    Dim dblResponseSum           As Double
    Dim dblResponseAverage       As Double
    Dim dblResponseA1c           As Double

    Dim dtRecord                 As Date
    Dim dtOldest                 As Date
    Dim dtNewest                 As Date
    Dim boolHaveDate             As Boolean
    Dim dblActualDays            As Double

    Dim intMGDL                  As Integer

    Dim objDays                  As Object
    Dim arrKeys                  As Variant
    Dim lngI                    As Long

    '------------------------------------------------------------
    ' Keep the existing widget calculation exactly as it was.
    '------------------------------------------------------------
    On Error Resume Next

    dblAverageGlucose = (dblCombinedA1c / dblNumberOfRecords)
    dblGMI = 3.31 + 0.02392 * dblAverageGlucose
    dblA1c = (dblAverageGlucose + 46.7) / 28.7

    On Error GoTo 0

    '------------------------------------------------------------
    ' Diagnostic only.
    '
    ' This does NOT alter dblCombinedA1c, dblNumberOfRecords,
    ' the glucose value, the graph, or anything displayed.
    '------------------------------------------------------------
    On Error Resume Next

    Set objDays = CreateObject("Scripting.Dictionary")

    strCutoffKey = Replace$(GetDate90DaysAgo(), "-", "")

    For Each strLine In Split(strResponse, vbLf)

        lngTotalLines = lngTotalLines + 1

        strCleanLine = Trim$(Replace$(CStr(strLine), vbTab, "|"))

        If Len(strCleanLine) > 0 Then

            arrRecord = Split(strCleanLine, "|")

            If UBound(arrRecord) >= 6 Then

                If Len(CStr(arrRecord(3))) >= 10 Then

                    strDateKey = Left$(CStr(arrRecord(3)), 10)
                    strYYYYMMDD = Replace$(strDateKey, "-", "")

                    If IsNumeric(strYYYYMMDD) And _
                       IsNumeric(arrRecord(6)) Then

                        lngDataLines = lngDataLines + 1

                        If CLng(strYYYYMMDD) > CLng(strCutoffKey) Then

                            intMGDL = CInt(arrRecord(6))

                            lngAccepted = lngAccepted + 1
                            dblResponseSum = dblResponseSum + CDbl(intMGDL)

                            '------------------------------------
                            ' Count readings for each calendar day.
                            '------------------------------------
                            If objDays.Exists(strDateKey) Then
                                objDays(strDateKey) = _
                                    CLng(objDays(strDateKey)) + 1
                            Else
                                objDays.Add strDateKey, 1
                            End If

                            '------------------------------------
                            ' Convert Juggluco timestamp:
                            '
                            ' yyyy-mm-dd-HH:MM:SS
                            '------------------------------------
                            Err.Clear

                            dtRecord = _
                                DateSerial( _
                                    CInt(Mid$(arrRecord(3), 1, 4)), _
                                    CInt(Mid$(arrRecord(3), 6, 2)), _
                                    CInt(Mid$(arrRecord(3), 9, 2))) + _
                                TimeSerial( _
                                    CInt(Mid$(arrRecord(3), 12, 2)), _
                                    CInt(Mid$(arrRecord(3), 15, 2)), _
                                    CInt(Mid$(arrRecord(3), 18, 2)))

                            If Err.Number = 0 Then

                                If Not boolHaveDate Then
                                    dtOldest = dtRecord
                                    dtNewest = dtRecord
                                    boolHaveDate = True
                                Else
                                    If dtRecord < dtOldest Then
                                        dtOldest = dtRecord
                                    End If

                                    If dtRecord > dtNewest Then
                                        dtNewest = dtRecord
                                    End If
                                End If

                            End If

                            Err.Clear

                        Else

                            lngTooOld = lngTooOld + 1

                        End If

                    Else

                        lngMalformed = lngMalformed + 1

                    End If

                Else

                    lngMalformed = lngMalformed + 1

                End If

            End If

        End If

    Next strLine

    '------------------------------------------------------------
    ' Independently calculate what THIS response says the
    ' average/A1c should be.
    '------------------------------------------------------------
    If lngAccepted > 0 Then
        dblResponseAverage = dblResponseSum / CDbl(lngAccepted)
        dblResponseA1c = (dblResponseAverage + 46.7) / 28.7
    End If

    If boolHaveDate Then
        dblActualDays = CDbl(dtNewest - dtOldest)
    End If

    '------------------------------------------------------------
    ' Write a fresh diagnostic log.
    ' It is overwritten each time A1c is recalculated.
    '------------------------------------------------------------
    If Right$(App.Path, 1) = "\" Then
        strLogPath = App.Path & "BGWidget_A1c.log"
    Else
        strLogPath = App.Path & "\BGWidget_A1c.log"
    End If

    intFile = FreeFile

    Open strLogPath For Output As #intFile

    Print #intFile, "BGWidget A1C DIAGNOSTIC"
    Print #intFile, "========================"
    Print #intFile, ""
    Print #intFile, "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #intFile, ""
    Print #intFile, "REQUEST"
    Print #intFile, "-------"
    Print #intFile, "Requested Days: " & CStr(intDaysOfData)
    Print #intFile, "Cutoff Date: " & GetDate90DaysAgo()
    Print #intFile, "URL: " & strURL
    Print #intFile, ""
    Print #intFile, "RESPONSE"
    Print #intFile, "--------"
    Print #intFile, "Total Lines: " & CStr(lngTotalLines)
    Print #intFile, "Valid Data Lines: " & CStr(lngDataLines)
    Print #intFile, "Accepted Records: " & CStr(lngAccepted)
    Print #intFile, "Rejected As Too Old: " & CStr(lngTooOld)
    Print #intFile, "Malformed/Invalid: " & CStr(lngMalformed)
    Print #intFile, ""

    If boolHaveDate Then
        Print #intFile, "Oldest Accepted: " & _
                        Format$(dtOldest, "yyyy-mm-dd hh:nn:ss")

        Print #intFile, "Newest Accepted: " & _
                        Format$(dtNewest, "yyyy-mm-dd hh:nn:ss")

        Print #intFile, "Actual Span In Days: " & _
                        Format$(dblActualDays, "0.000")
    Else
        Print #intFile, "Oldest Accepted: NONE"
        Print #intFile, "Newest Accepted: NONE"
        Print #intFile, "Actual Span In Days: 0"
    End If

    Print #intFile, ""
    Print #intFile, "INDEPENDENT RESPONSE CALCULATION"
    Print #intFile, "--------------------------------"
    Print #intFile, "Response Glucose Sum: " & _
                    Format$(dblResponseSum, "0")

    Print #intFile, "Response Record Count: " & _
                    CStr(lngAccepted)

    Print #intFile, "Response Average Glucose: " & _
                    Format$(dblResponseAverage, "0.00")

    Print #intFile, "Response Calculated A1c: " & _
                    Format$(dblResponseA1c, "0.00") & "%"

    Print #intFile, ""
    Print #intFile, "ACTUAL WIDGET ACCUMULATORS"
    Print #intFile, "--------------------------"
    Print #intFile, "dblCombinedA1c: " & _
                    Format$(dblCombinedA1c, "0")

    Print #intFile, "dblNumberOfRecords: " & _
                    Format$(dblNumberOfRecords, "0")

    Print #intFile, "Widget Average Glucose: " & _
                    Format$(dblAverageGlucose, "0.00")

    Print #intFile, "Widget Calculated A1c: " & _
                    Format$(dblA1c, "0.00") & "%"

    Print #intFile, ""
    Print #intFile, "RECORDS BY DAY"
    Print #intFile, "--------------"

    If Not objDays Is Nothing Then

        If objDays.Count > 0 Then

            arrKeys = objDays.Keys

            For lngI = 0 To objDays.Count - 1
                Print #intFile, _
                    CStr(arrKeys(lngI)) & " = " & _
                    CStr(objDays(arrKeys(lngI)))
            Next lngI

        End If

    End If

    Close #intFile

    Set objDays = Nothing

    On Error GoTo 0

    '------------------------------------------------------------
    ' Original return value.
    '------------------------------------------------------------
    On Error Resume Next

    CalculateA1c = _
        FormatNumber(CStr(dblAverageGlucose), 0) & "|" & _
        FormatNumber(CStr(dblGMI), 1) & "|" & _
        FormatNumber(CStr(dblA1c), 1)

    On Error GoTo 0
End Function

Function GetA1c(strURL) As Boolean
    Dim intMGDL As Integer
    Dim intLastValidMGDL As Integer
    Dim arrRecord As Variant
    Dim strRecord As Variant
    Dim strDateTime As String
    Dim varRate As Variant
    Dim varLastValidRate As Variant
    Dim varYYYYMMDD As Variant
    Dim var90DaysAgo As Variant
    Dim dtLastValidReading As Date
    Dim boolFoundValid As Boolean
    Dim boolLatestSampleValid As Boolean

    GetA1c = False
    On Error GoTo GetA1cError

    var90DaysAgo = Replace$(GetDate90DaysAgo(), "-", "")

    objURL.Open "GET", strURL, False
    objURL.Send
    strResponse = objURL.responseText

    If modHistoryGraph.UpdateHistoryCache(strResponse) Then Call modHistoryGraph.DrawHistoryGraphs

    For Each strRecord In Split(strResponse, vbLf)
        strRecord = Replace$(CStr(strRecord), vbCr, "")
        strRecord = Replace$(CStr(strRecord), vbTab, "|")

        If Len(Trim$(CStr(strRecord))) > 0 Then
            arrRecord = Split(CStr(strRecord), "|")

            If UBound(arrRecord) >= 7 Then
                If Len(Trim$(CStr(arrRecord(3)))) >= 19 Then
                    varYYYYMMDD = Replace$(Left$(Trim$(CStr(arrRecord(3))), 10), "-", "")

                    If IsNumeric(varYYYYMMDD) Then
                        If CDbl(varYYYYMMDD) > CDbl(var90DaysAgo) Then
                            If IsNumeric(Trim$(CStr(arrRecord(6)))) Then
                                intMGDL = CInt(Trim$(CStr(arrRecord(6))))

                                If intMGDL > 0 Then
                                    varRate = Trim$(CStr(arrRecord(7)))
                                    strDateTime = Trim$(CStr(arrRecord(3)))

                                    dblCombinedA1c = dblCombinedA1c + CDbl(intMGDL)
                                    dblNumberOfRecords = dblNumberOfRecords + 1

                                    intLastValidMGDL = intMGDL
                                    varLastValidRate = varRate
                                    dtLastValidReading = DateSerial(CInt(Left$(strDateTime, 4)), CInt(Mid$(strDateTime, 6, 2)), CInt(Mid$(strDateTime, 9, 2)))
                                    dtLastValidReading = dtLastValidReading + TimeSerial(CInt(Mid$(strDateTime, 12, 2)), CInt(Mid$(strDateTime, 15, 2)), CInt(Mid$(strDateTime, 18, 2)))

                                    boolFoundValid = True
                                    boolLatestSampleValid = True
                                Else
                                    boolLatestSampleValid = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next strRecord

    If boolFoundValid Then
        intCurrentGlucose = intLastValidMGDL
        varRateOfChange = varLastValidRate
        dtLastReading = dtLastValidReading
    End If

    GetA1c = (boolFoundValid And boolLatestSampleValid)
    Exit Function

GetA1cError:
    strLastUpdate = "Error " & CStr(Err.Number) & ": " & Err.Description & " at " & Time
    Debug.Print "GetA1c Error " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    GetA1c = False
End Function

Private Function FormatLastReadingTime() As String
    Dim lngSeconds As Long
    Dim lngMinutes As Long

    If dtLastReading = 0 Then
        FormatLastReadingTime = "--"
        Exit Function
    End If

    If Not boolUseRelativeTime Then
        FormatLastReadingTime = Format$(dtLastReading, "h:mm:ss AM/PM")
        Exit Function
    End If

    lngSeconds = DateDiff("s", dtLastReading, Now)
    If lngSeconds < 0 Then lngSeconds = 0

    If lngSeconds < 60 Then
        FormatLastReadingTime = "Just Now"
        Exit Function
    End If

    lngMinutes = lngSeconds \ 60

    If lngMinutes = 1 Then
        FormatLastReadingTime = "1 Min Ago"
    Else
        FormatLastReadingTime = CStr(lngMinutes) & " Mins Ago"
    End If
End Function

Function GetNewReading(strURL) As String
    Dim intMGDL      As Integer
    Dim arrRecord    As Variant
    Dim strRecord    As Variant
    Dim varRate      As Variant

    'Sensorid    nr      UnixTime    YYYY-mm-dd-HH:MM:SS    TZ  Min     mg/dL  Rate    ChangeLabel
    'XX0HMCDRNCC 8199    1741292217  2025-03-06-15:16:57    -5  8198    106    +0.1    STABLE

    On Error Resume Next
        Do
            objURL.Open "GET", strURL, False
            objURL.Send

            strResponse = objURL.responseText
            If Err.Number <> 0 Then strLastUpdate = "Error: " & Err.Number & " at " & Time
        Loop Until Err.Number = 0
    On Error GoTo 0

    For Each strRecord In Split(strResponse, vbLf)
        strRecord = Replace(strRecord, vbTab, "|")
        If InStr(1, strRecord, "|", vbTextCompare) > 6 Then
            arrRecord = Split(strRecord, "|", -1, vbTextCompare)
            If IsNumeric(arrRecord(6)) Then
                intMGDL = CInt(arrRecord(6))
            End If
        End If
        If IsNumeric(intMGDL) Then
            dblCombinedA1c = dblCombinedA1c + CInt(intMGDL)
            dblNumberOfRecords = dblNumberOfRecords + 1
        End If
    Next

    intCurrentGlucose = intMGDL
End Function

Function GetDate90DaysAgo() As String
  ' Get the current date.
  Dim currentDate As Date
  currentDate = Date

  ' Subtract 90 days.
  Dim ninetyDaysAgo As Date
  ninetyDaysAgo = DateAdd("d", -90, currentDate) '

  ' Format the date as "yyyy-mm-dd".
  GetDate90DaysAgo = Format(ninetyDaysAgo, "yyyy-mm-dd") '

End Function

Function FindInJSON(strPattern, intLevel, strJSON) As String
    Dim arrTabbedJSON     As Variant
    Dim strTabbedJSONFile As Variant

    arrTabbedJSON = Split(strJSON, vbCrLf)

    For Each strTabbedJSONFile In arrTabbedJSON
        If Len(strTabbedJSONFile) - Len(Replace(strTabbedJSONFile, vbTab, "")) = intLevel Then
            If Len(strTabbedJSONFile) - Len(Replace(strTabbedJSONFile, vbTab, "")) < intLevel + 1 Then
                If InStr(1, strTabbedJSONFile, strPattern, vbTextCompare) Then
                    FindInJSON = FindInJSON & Replace(strTabbedJSONFile, vbTab, "") & vbCrLf
                End If
            End If
        End If
    Next
End Function

Function ParseJSONFile(strRawJSON) As String
    Dim boolRetractFirstTab As Boolean
    Dim strGetJSON          As String
    Dim arrGetJSON          As Variant
    Dim strGetJSONLine      As Variant
    Dim intTabs             As Integer
    Dim strNewGetIssues     As Integer

    strGetJSON = Replace(strRawJSON, ",", ",^")    ' Insert Carraige Returns
    strGetJSON = Replace(strGetJSON, "{", "{^")
    strGetJSON = Replace(strGetJSON, "}", "^}")
    strGetJSON = Replace(strGetJSON, "[", "[^")
    strGetJSON = Replace(strGetJSON, "]", "]^")
    strGetJSON = Replace(strGetJSON, "[^]^", "[]")
    arrGetJSON = Split(strGetJSON, "^")

    For Each strGetJSONLine In arrGetJSON
        If Not boolRetractFirstTab Then intTabs = -1
        ParseJSONFile = ParseJSONFile & strNewGetIssues & Tabs(intTabs) & strGetJSONLine & vbCrLf

        If InStr(1, strGetJSONLine, "[", vbTextCompare) <> 0 Then intTabs = intTabs + 1
        If InStr(1, strGetJSONLine, "]", vbTextCompare) <> 0 Then intTabs = intTabs - 1
        If InStr(1, strGetJSONLine, "{", vbTextCompare) <> 0 Then intTabs = intTabs + 1
        If InStr(1, strGetJSONLine, "}", vbTextCompare) <> 0 Then intTabs = intTabs - 1

        boolRetractFirstTab = True
    Next
End Function

Function Tabs(intTabs) As String
    Dim intTab As Integer

    For intTab = 0 To intTabs
        Tabs = Tabs & vbTab
    Next
End Function

