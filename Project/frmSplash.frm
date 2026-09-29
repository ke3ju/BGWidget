VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4365
   ClientLeft      =   1350
   ClientTop       =   0
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Leelawadee UI"
      Size            =   56.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   291
   ScaleMode       =   0  'User
   ScaleWidth      =   330
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrA1cTimer 
      Interval        =   60000
      Left            =   5640
      Top             =   2160
   End
   Begin VB.Timer tmrGetCurrentGlucose 
      Interval        =   60000
      Left            =   5640
      Top             =   1320
   End
   Begin VB.Label lblA1cMMOLL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6.3 mmol/l"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   0
      TabIndex        =   4
      Top             =   1950
      Width           =   6594
   End
   Begin VB.Label lblLastUpdate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Last Update: 12:49p PM"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   6603
   End
   Begin VB.Label lblRecordNumInDays 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "90 Day A1c/GMI History"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      TabIndex        =   2
      Top             =   2925
      Width           =   6603
   End
   Begin VB.Label lblA1cMGDL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "134 mg/dL"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "hbA1c"
      Top             =   1275
      Width           =   6603
   End
   Begin VB.Label lblCurrentGlucose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "105 mg/dL"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Current Glucose Reading"
      Top             =   0
      Width           =   6603
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_MONITORPOWER As Long = &HF170
Private Const MONITOR_OFF As Long = 2


Dim intMinutesCount As Integer

Public boolMirrorMode   As Boolean
Public intMirrorMonitor As Integer

Private Sub Form_Load()

    If boolMirrorMode Then
        tmrGetCurrentGlucose.Enabled = False
        tmrA1cTimer.Enabled = False

        Call modWidgetPlacement.ApplyWidgetPlacementToMonitor(Me, intMirrorMonitor)
        Call modWidgetPlacement.MakeMirrorTransparent(Me)
        Call modWidgetPlacement.SyncMirrorWidget(frmSplash, Me)

        Exit Sub
    End If

    Set objReadWrite = New ReadWrite
    Set objINI = New INIFile

    frmSplash.BackColor = vbWhite
    MakeFormTransparent

    Call modGetGlucose.ReadINI(App.Path & "\" & App.EXEName & ".ini", "Main")
    Call modGetGlucose.ApplySettings
    Call modGetGlucose.ApplyColors

    Call modHistoryGraph.ReadHistoryINI(App.Path & "\" & App.EXEName & ".ini")

    Call modWidgetPlacement.ReadPlacementINI(App.Path & "\" & App.EXEName & ".ini")
    Call modWidgetPlacement.ApplyWidgetPlacement(frmSplash)

    tmrGetCurrentGlucose.Enabled = True

    modGetGlucose.GetGlucose

    Call modWidgetPlacement.CreateMirrorWidgets(frmSplash)

End Sub

Public Sub MakeFormTransparent()
    Const GWL_STYLE         As Long = &HFFFFFFF0
    Const GWL_EXSTYLE       As Long = &HFFFFFFEC
    Const WS_EX_LAYERED     As Long = &H80000
    Const LWA_COLORKEY      As Long = &H1&
    Const LWA_ALPHA         As Long = &H2&
    '
    frmSplash.BackColor = vbBlack
    SetWindowLong frmSplash.hwnd, GWL_EXSTYLE, GetWindowLong(frmSplash.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes frmSplash.hwnd, vbBlack, 0&, LWA_COLORKEY
End Sub

Private Sub lblA1cMGDL_Click()
    If MsgBox("Exit BG Widget?", vbYesNo + vbQuestion, "Exit BG Widget?") = vbYes Then ExitProcess 0
End Sub

Private Sub lblA1cMMOLL_Click()
    TurnOffMonitor
End Sub

Sub TurnOffMonitor()
    Dim hwnd As Long
    hwnd = GetDesktopWindow()
    SendMessage hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF
End Sub

Private Sub tmrA1cTimer_Timer()
    intMinutesCount = intMinutesCount + 1

    If intMinutesCount > (modGetGlucose.intA1cInterval - 1) Then
        modGetGlucose.boolPostLaunch = False
        intMinutesCount = 0
    End If
End Sub

Private Sub tmrGetCurrentGlucose_Timer()
    modGetGlucose.GetGlucose
End Sub
