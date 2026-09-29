VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'V26.09.20
Private Const GENERIC_READ As Long = &H80000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const MAX_PATH As Long = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public CheckSemaphore As Boolean

Public Function ReadFileEnc(strFileName As String) As String   ' v26.09.20

    ReadFileEnc = ReadFile(strFileName)

End Function

Public Function ReadFile(strFileName As String) As String   ' v26.09.20

    Dim objStream As Object

    Dim vntBytes As Variant

    Dim strCharset As String
    Dim strResult As String
    Dim strErrorDescription As String

    Dim lngFileSize As Long
    Dim lngPosition As Long
    Dim lngRequiredBytes As Long
    Dim lngCheck As Long
    Dim lngByte As Long
    Dim lngErrorNumber As Long

    Dim boolValidUTF8 As Boolean

    ReadFile = ""

    If FileExists(strFileName) = False Then

        ReadFile = _
            "[File Not Found] - " & _
            strFileName

        Exit Function

    End If

    If CheckSemaphore Then

        Do Until Not IsFileInUse(strFileName)
            DoEvents
        Loop

    End If

    On Error GoTo ReadError

    '
    '==========================================================
    ' READ RAW BYTES
    '==========================================================
    '

    Set objStream = CreateObject("ADODB.Stream")

    objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile strFileName

    lngFileSize = objStream.Size

    If lngFileSize <= 0 Then

        objStream.Close
        Set objStream = Nothing

        ReadFile = ""
        Exit Function

    End If

    vntBytes = objStream.Read

    objStream.Close
    Set objStream = Nothing

    '
    '==========================================================
    ' DETECT ENCODING
    '==========================================================
    '

    strCharset = "windows-1252"

    '
    ' UTF-8 BOM
    '
    If lngFileSize >= 3 Then

        If vntBytes(0) = &HEF And _
           vntBytes(1) = &HBB And _
           vntBytes(2) = &HBF Then

            strCharset = "utf-8"

        End If

    End If

    '
    ' UTF-16 LE BOM
    '
    If strCharset = "windows-1252" And _
       lngFileSize >= 2 Then

        If vntBytes(0) = &HFF And _
           vntBytes(1) = &HFE Then

            strCharset = "unicode"

        End If

    End If

    '
    ' UTF-16 BE BOM
    '
    If strCharset = "windows-1252" And _
       lngFileSize >= 2 Then

        If vntBytes(0) = &HFE And _
           vntBytes(1) = &HFF Then

            strCharset = "unicodeFFFE"

        End If

    End If

    '
    '==========================================================
    ' NO BOM:
    '
    ' Validate the byte stream as UTF-8.
    '
    ' ASCII is also valid UTF-8, so ordinary ASCII text will
    ' naturally come through this path without changing.
    '==========================================================
    '

    If strCharset = "windows-1252" Then

        boolValidUTF8 = True
        lngPosition = 0

        Do While lngPosition < lngFileSize

            lngByte = CLng(vntBytes(lngPosition))
            lngRequiredBytes = 0

            If lngByte <= &H7F Then

                lngPosition = lngPosition + 1

            ElseIf lngByte >= &HC2 And _
                   lngByte <= &HDF Then

                lngRequiredBytes = 1

            ElseIf lngByte >= &HE0 And _
                   lngByte <= &HEF Then

                lngRequiredBytes = 2

            ElseIf lngByte >= &HF0 And _
                   lngByte <= &HF4 Then

                lngRequiredBytes = 3

            Else

                boolValidUTF8 = False
                Exit Do

            End If

            If lngRequiredBytes > 0 Then

                If lngPosition + lngRequiredBytes >= lngFileSize Then

                    boolValidUTF8 = False
                    Exit Do

                End If

                For lngCheck = 1 To lngRequiredBytes

                    lngByte = _
                        CLng( _
                            vntBytes( _
                                lngPosition + lngCheck))

                    If lngByte < &H80 Or _
                       lngByte > &HBF Then

                        boolValidUTF8 = False
                        Exit For

                    End If

                Next lngCheck

                If boolValidUTF8 = False Then Exit Do

                '
                ' Reject overlong UTF-8, surrogate code points,
                ' and values above Unicode U+10FFFF.
                '
                lngByte = CLng(vntBytes(lngPosition))

                If lngRequiredBytes = 2 Then

                    If lngByte = &HE0 Then

                        If CLng(vntBytes(lngPosition + 1)) < &HA0 Then
                            boolValidUTF8 = False
                            Exit Do
                        End If

                    ElseIf lngByte = &HED Then

                        If CLng(vntBytes(lngPosition + 1)) > &H9F Then
                            boolValidUTF8 = False
                            Exit Do
                        End If

                    End If

                ElseIf lngRequiredBytes = 3 Then

                    If lngByte = &HF0 Then

                        If CLng(vntBytes(lngPosition + 1)) < &H90 Then
                            boolValidUTF8 = False
                            Exit Do
                        End If

                    ElseIf lngByte = &HF4 Then

                        If CLng(vntBytes(lngPosition + 1)) > &H8F Then
                            boolValidUTF8 = False
                            Exit Do
                        End If

                    End If

                End If

                lngPosition = _
                    lngPosition + _
                    lngRequiredBytes + 1

            End If

        Loop

        If boolValidUTF8 Then
            strCharset = "utf-8"
        End If

    End If

    '
    '==========================================================
    ' DECODE RAW FILE USING DETECTED CHARACTER SET
    '==========================================================
    '

    Set objStream = CreateObject("ADODB.Stream")

    objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile strFileName

    objStream.Position = 0
    objStream.Type = 2
    objStream.Charset = strCharset
    objStream.Position = 0

    strResult = objStream.ReadText(-1)

    objStream.Close
    Set objStream = Nothing

    '
    ' ADODB.Stream normally consumes the BOM, but remove a
    ' remaining Unicode BOM character defensively.
    '
    If Len(strResult) > 0 Then

        If AscW(Left$(strResult, 1)) = &HFEFF Then
            strResult = Mid$(strResult, 2)
        End If

    End If

    ReadFile = strResult

    Exit Function


ReadError:

    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description

    On Error Resume Next

    If Not objStream Is Nothing Then
        objStream.Close
    End If

    Set objStream = Nothing

    On Error GoTo 0

    ReadFile = _
        "[Read Error " & _
        CStr(lngErrorNumber) & _
        "] - " & _
        strErrorDescription & _
        " - " & _
        strFileName

End Function

Public Function WriteFile(strFileName As String, ByVal strData As String, boolOverWrite As Boolean, boolCreatePath As Boolean) As String   ' v26.09.20

    Dim objTextStream As Object
    Dim objBinaryStream As Object

    Dim strOutput As String
    Dim strExisting As String
    Dim strErrorDescription As String

    Dim lngErrorNumber As Long

    WriteFile = ""

    If boolCreatePath Then
        Call CreateFolder(strFileName)
    End If

    If FileExists(strFileName) Then

        If CheckSemaphore Then

            Do Until Not IsFileInUse(strFileName)
                DoEvents
            Loop

        End If

    End If

    '
    '==========================================================
    ' PRESERVE EXISTING APPEND BEHAVIOR
    '
    ' Instead of mixing ANSI and UTF-8 bytes in the same file,
    ' read the existing text correctly, append the new text in
    ' Unicode, then write the complete result as clean UTF-8.
    '==========================================================
    '

    If boolOverWrite = False And _
       FileExists(strFileName) Then

        strExisting = ReadFile(strFileName)

        If Left$(strExisting, 12) = "[Read Error " Then

            WriteFile = _
                "Failed to read existing file before append: " & _
                strExisting

            Exit Function

        End If

        strOutput = _
            strExisting & _
            strData

    Else

        strOutput = strData

    End If

    On Error GoTo WriteError

    '
    '==========================================================
    ' CREATE UTF-8 TEXT STREAM
    '==========================================================
    '

    Set objTextStream = CreateObject("ADODB.Stream")

    objTextStream.Type = 2
    objTextStream.Charset = "utf-8"
    objTextStream.Open

    objTextStream.WriteText strOutput

    '
    ' ADODB.Stream inserts the UTF-8 BOM:
    '
    '     EF BB BF
    '
    ' WFF/XML does not need it, so strip those first three
    ' bytes before writing the final file.
    '==========================================================
    '

    objTextStream.Position = 0
    objTextStream.Type = 1

    If objTextStream.Size >= 3 Then
        objTextStream.Position = 3
    Else
        objTextStream.Position = objTextStream.Size
    End If

    Set objBinaryStream = CreateObject("ADODB.Stream")

    objBinaryStream.Type = 1
    objBinaryStream.Open

    objTextStream.CopyTo objBinaryStream

    '
    ' 2 = adSaveCreateOverWrite
    '
    objBinaryStream.SaveToFile strFileName, 2

    objBinaryStream.Close
    objTextStream.Close

    Set objBinaryStream = Nothing
    Set objTextStream = Nothing

    WriteFile = "0"

    Exit Function


WriteError:

    lngErrorNumber = Err.Number
    strErrorDescription = Err.Description

    On Error Resume Next

    If Not objBinaryStream Is Nothing Then
        objBinaryStream.Close
    End If

    If Not objTextStream Is Nothing Then
        objTextStream.Close
    End If

    Set objBinaryStream = Nothing
    Set objTextStream = Nothing

    On Error GoTo 0

    WriteFile = _
        "Failed to write to " & _
        strFileName & _
        "...with Error [" & _
        CStr(lngErrorNumber) & _
        "]: " & _
        strErrorDescription & _
        "..."

End Function

Private Function CreateFolder(strFolderPath As String)  '   v15.07.08
    Dim objFSO        As Object
    Dim splPath       As Variant
    Dim itemFolder    As Variant
    Dim colFolders    As String
    Dim strFolderOnly As String
    If strFolderPath = "" Then Exit Function
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    splPath = Split(strFolderPath, "\")
    strFolderOnly = Replace(strFolderPath, "\" & splPath(UBound(splPath)), "")
    If objFSO.FolderExists(strFolderOnly) Then
        Set objFSO = Nothing
        Exit Function
    End If
    For Each itemFolder In splPath
        colFolders = colFolders & itemFolder & "\"
        If objFSO.FolderExists(Left(colFolders, Len(colFolders) - 1)) = False And itemFolder <> splPath(UBound(splPath)) And itemFolder <> "" And InStr(colFolders, "\\\") = 0 Then
            On Error Resume Next
            objFSO.CreateFolder Left(colFolders, Len(colFolders) - 1)
            On Error GoTo 0
        End If
    Next
    Set objFSO = Nothing
End Function

Private Function FileExists(strSource As String) As Boolean  '   v15.07.08
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long

    hFile = FindFirstFile(strSource, WFD)
    FileExists = hFile <> INVALID_HANDLE_VALUE

    Call FindClose(hFile)
End Function

Private Function IsFileInUse(strFile As String) As Boolean  '   v15.07.08
    Dim hFile As Long
   
    hFile = CreateFile(strFile, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    IsFileInUse = hFile = INVALID_HANDLE_VALUE
   
    CloseHandle hFile
End Function
