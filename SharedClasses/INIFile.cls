VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'v13.02.25
Private strINIData       As String
Private strFinalData     As String
Private boolFoundSection As Boolean
Private boolPathSet      As Boolean

Public ReadRaw           As Boolean
Public ListFilter        As String

Public Function Value(strValueName)
    Dim i          As Integer
    Dim arrINIData As Variant

    arrINIData = Split(strINIData, vbCrLf)
    For i = 0 To UBound(arrINIData)
        If InStr(arrINIData(i), "=") > 0 Then
            If Trim(LCase(Split(arrINIData(i), "=", 2)(0))) = LCase(strValueName) Then
                Value = CStrBool(Trim(Split(arrINIData(i), "=", 2)(1)))
                If IsNumeric(Value) = True And Len(Value) < 5 Then
                    Value = CInt(Value)
                End If
            End If
        End If
    Next
End Function
'v13.02.25
Public Function Load(ByVal strINIFile As String, ByVal strSection As String) As String '   Reads and parses the specific in the INI file
    Dim arrINIFileAll         As Variant
    Dim strData               As String

    Dim index                 As Integer
    Dim index1                As Integer
    Dim index2                As Integer

    Dim objShell              As Object
    Dim objFile               As Object

    Dim boolIncludedFilesMode As Boolean

    Set objShell = CreateObject("WScript.Shell")
    Set objFile = New ReadWrite

    boolFoundSection = False
    arrINIFileAll = Split(objFile.ReadFile(strINIFile), vbCrLf)

    For index1 = 0 To UBound(arrINIFileAll) '   Loop until specified section if found
        index2 = index1 + 1
        If InStr(arrINIFileAll(index1), "[") > 0 And strSection = "$ListSections$" And InStr(arrINIFileAll(index1), ";") = 0 Then
            If ListFilter <> "" Then
                If InStr(arrINIFileAll(index1), ListFilter) > 0 Then
                    strINIData = strINIData & arrINIFileAll(index1) & vbCrLf
                End If
            Else
                strINIData = strINIData & arrINIFileAll(index1) & vbCrLf
            End If
        Else
            If InStr(arrINIFileAll(index1), "[" & strSection & "]") > 0 Or strSection = "$All$" Then '   Look for specified section
                If strSection <> "$All$" Then boolFoundSection = True
                If strSection = "$All$" Then strINIData = strINIData & arrINIFileAll(index1) & vbCrLf
                Do Until index2 = UBound(arrINIFileAll) + 1 '   Loop until next section or end of file is found
                    If index2 <= UBound(arrINIFileAll) Then
                        If ReadRaw = True Then
                            If InStr(Left(arrINIFileAll(index2), 1), "[") > 0 Then
                                Exit Do
                            End If
                            strData = strData & Trim(objShell.ExpandEnvironmentStrings(arrINIFileAll(index2))) & vbCrLf
                        Else
                            If InStr(Left(arrINIFileAll(index2), 1), "[") > 0 And strSection <> "$All$" Then
                                Exit Do
                            End If
                            If InStr(arrINIFileAll(index2), "=") > 0 Then
                                If Left(arrINIFileAll(index2), 1) <> ";" And InStr(arrINIFileAll(index2), ";") = 0 Then
                                    strData = strData & Trim(objShell.ExpandEnvironmentStrings(arrINIFileAll(index2))) & vbCrLf
                                ElseIf Left(arrINIFileAll(index2), 1) <> ";" And InStr(arrINIFileAll(index2), ";") > 0 Then
                                    If Split(arrINIFileAll(index2), ";")(0) <> "" Then
                                        strData = strData & Trim(objShell.ExpandEnvironmentStrings(Split(arrINIFileAll(index2), ";")(0))) & vbCrLf
                                    End If
                                End If
                            End If
                        End If
                    End If
                    index2 = index2 + 1
                Loop
                Exit For
            End If
        End If
    Next

    If boolFoundSection = False Then
        If boolIncludedFilesMode <> True Then
            boolIncludedFilesMode = True
            strINIFile = CStr(strINIFile)
            strSection = CStr(strSection)
            Call CheckIncludedFiles(strINIFile, strSection)
        End If
    End If

    If strData <> "" Then
        strFinalData = strData
    End If

    Load = strFinalData
    strINIData = strINIData & strData
    If boolFoundSection = True Then ReadRaw = False

    If strSection = "$All$" Or strSection = "$ListSections$" Then Load = strINIData
End Function

Public Function UnLoad()
    strINIData = ""
    strFinalData = ""
End Function

Public Function WriteValue(ByVal strINIFile As String, ByVal strSection As String, ByVal strValueName As String, ByVal strValue As String) As String
    Dim objFile        As ReadWrite
    Dim objFSO         As Object
    Dim arrINIFile     As Variant
    Dim strLine        As String
    Dim strData        As String
    Dim strCurrentKey  As String
    Dim strResult      As String
    Dim intIndex       As Integer
    Dim boolInSection  As Boolean
    Dim boolFoundSection As Boolean
    Dim boolFoundValue As Boolean

    Set objFile = New ReadWrite
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists(strINIFile) Then
        arrINIFile = Split(objFile.ReadFile(strINIFile), vbCrLf)

        For intIndex = 0 To UBound(arrINIFile)
            strLine = arrINIFile(intIndex)

            If Left(Trim(strLine), 1) = "[" And Right(Trim(strLine), 1) = "]" Then
                If boolInSection = True And boolFoundValue = False Then
                    strData = strData & strValueName & "=" & strValue & vbCrLf
                    boolFoundValue = True
                End If

                If LCase(Trim(strLine)) = LCase("[" & strSection & "]") Then
                    boolInSection = True
                    boolFoundSection = True
                Else
                    boolInSection = False
                End If
            End If

            If boolInSection = True And boolFoundValue = False Then
                If Left(Trim(strLine), 1) <> ";" And InStr(strLine, "=") > 0 Then
                    strCurrentKey = Trim(Split(strLine, "=", 2)(0))

                    If LCase(strCurrentKey) = LCase(strValueName) Then
                        strLine = Left(strLine, InStr(strLine, "=")) & strValue
                        boolFoundValue = True
                    End If
                End If
            End If

            strData = strData & strLine

            If intIndex < UBound(arrINIFile) Then
                strData = strData & vbCrLf
            End If
        Next

        If boolFoundSection = True And boolFoundValue = False Then
            If Right(strData, 2) <> vbCrLf Then strData = strData & vbCrLf
            strData = strData & strValueName & "=" & strValue
        ElseIf boolFoundSection = False Then
            If Right(strData, 2) <> vbCrLf Then strData = strData & vbCrLf
            strData = strData & vbCrLf & "[" & strSection & "]" & vbCrLf
            strData = strData & strValueName & "=" & strValue
        End If
    Else
        strData = "[" & strSection & "]" & vbCrLf
        strData = strData & strValueName & "=" & strValue
    End If

    strResult = objFile.WriteFile(strINIFile, strData, True, True)

    If strResult = "0" Then
        WriteValue = "0"
    Else
        WriteValue = strResult
    End If

    Set objFSO = Nothing
    Set objFile = Nothing
End Function

Private Function CheckIncludedFiles(strINIFile As String, strSection As String) '   Version 1.0
    Dim arrIncludedFiles As Variant
    Dim strIncludedFile  As Variant
    Dim objFSO           As Object

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    arrIncludedFiles = IncludedFiles(strINIFile)

    If IsArray(arrIncludedFiles) Then
        For Each strIncludedFile In arrIncludedFiles
            strIncludedFile = CStr(strIncludedFile)
            If objFSO.FileExists(strIncludedFile) Then
                strIncludedFile = CStr(strIncludedFile)
                Call Load(strIncludedFile, strSection)
            End If
        Next
    Else
        Exit Function
    End If

    Set objFSO = Nothing
End Function

Private Function IncludedFiles(strINIFile)  '   Version 1.4
    Dim objFSO   As Object
    Dim objShell As Object

    Dim strPath  As String
    Dim itemData As Variant
    Dim strData  As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If strINIFile = "" Then Exit Function

    strPath = Replace(strINIFile, Split(strINIFile, "\")(UBound(Split(strINIFile, "\"))), "", 1, -1, vbTextCompare)

    Set objFile = New ReadWrite

    For Each itemData In Split(objFile.ReadFile(strINIFile), vbCrLf)
        If Left(LCase(itemData), Len("$Include")) = LCase("$Include") Then
            If objFSO.FileExists(objShell.ExpandEnvironmentStrings(Trim(Split(itemData, "$Include", 2, vbTextCompare)(1)))) Then
                strData = strData & objShell.ExpandEnvironmentStrings(Trim(Split(itemData, "$Include", 2, vbTextCompare)(1))) & vbCrLf
            Else
                strData = strData & strPath & objShell.ExpandEnvironmentStrings(Trim(Split(itemData, "$Include", 2, vbTextCompare)(1))) & vbCrLf
            End If
        End If
    Next

    IncludedFiles = Split(strData, vbCrLf)

    Set objShell = Nothing
    Set objFSO = Nothing
End Function

Private Function CStrBool(strBool)  '   Version 2.3
    If Trim(LCase(strBool)) = LCase("True") Or Trim(LCase(strBool)) = LCase("Yes") Then
        CStrBool = True
        strBool = True
    ElseIf Trim(LCase(strBool)) = LCase("False") Or Trim(LCase(strBool)) = LCase("No") Then
        CStrBool = False
        strBool = False
    Else
        CStrBool = strBool
    End If
End Function
