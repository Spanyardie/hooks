VERSION 5.00
Begin VB.Form FTestHook 
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Journalling"
      Height          =   1215
      Left            =   6120
      TabIndex        =   25
      Top             =   7920
      Width           =   3375
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdPlayback 
         Caption         =   "Playback"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "Record"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scope"
      Height          =   975
      Left            =   2880
      TabIndex        =   22
      Top             =   840
      Width           =   1455
      Begin VB.OptionButton optApp 
         Caption         =   "Application"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optGlobal 
         Caption         =   "Global"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraWnd 
      Caption         =   "Window creation"
      Height          =   1215
      Left            =   360
      TabIndex        =   19
      Top             =   7920
      Width           =   5535
      Begin VB.CheckBox chkDisableWndCreate 
         Alignment       =   1  'Right Justify
         Caption         =   "Use hook to disable window activation"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdCreateWnd 
         Caption         =   "&Create Window"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Application status"
      Height          =   855
      Left            =   6120
      TabIndex        =   17
      Top             =   840
      Width           =   3375
      Begin VB.Label lblAppStatus 
         Caption         =   "Window activation is enabled!"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3135
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Hook Status"
      Height          =   5895
      Left            =   6120
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
      Begin VB.CheckBox chkWH_JOURNALRECORD 
         Caption         =   "          WH_JOURNALRECORD"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   5280
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_JOURNALPLAYBACK 
         Caption         =   "          WH_JOURNALPLAYBACK"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4800
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_SHELL 
         Caption         =   "          WH_SHELL"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   4320
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_MOUSE 
         Caption         =   "          WH_MOUSE"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_MSGFILTER 
         Caption         =   "          WH_MESSAGE"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_KEYBOARD 
         Caption         =   "          WH_KEYBOARD"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_GETMESSAGE 
         Caption         =   "          WH_GETMESSAGE"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_FOREGROUNDIDLE 
         Caption         =   "          WH_FOREGROUNDIDLE"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_DEBUG 
         Caption         =   "          WH_DEBUG"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_CBT 
         Caption         =   "          WH_CBT"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chkWH_CALLWNDPROC 
         Caption         =   "          WH_CALLWNDPROC"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.TextBox txtMsg 
      Height          =   5775
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   5535
   End
   Begin VB.CommandButton cmdReleaseHook 
      Caption         =   "&Release Hook"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetHook 
      Caption         =   "&Set Hook"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboHookType 
      Height          =   315
      ItemData        =   "FTestHook.frx":0000
      Left            =   240
      List            =   "FTestHook.frx":0024
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Application scope Hook test application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Message notifications:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select required hook:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "FTestHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents gHookDLL As CHooks
Attribute gHookDLL.VB_VarHelpID = -1


Private Sub chkDisableWndCreate_Click()

If chkDisableWndCreate.Value = 0 Then
    lblAppStatus = "Window activation is enabled!"
Else
    lblAppStatus = "Window activation is DISABLED!"
End If

End Sub

Private Sub cmdCreateWnd_Click()

If cmdCreateWnd.Caption = "&Create Window" Then
    FTest.Show
    cmdCreateWnd.Caption = "&Kill Window"
Else
    Unload FTest
    cmdCreateWnd.Caption = "&Create Window"
End If

End Sub

Private Sub cmdPlayback_Click()

gHookDLL.SetUpHook WH_JOURNALPLAYBACK, HI_GLOBAL
cmdRecord.Enabled = False

End Sub

Private Sub cmdRecord_Click()

'set up the journalling hook
gHookDLL.SetUpHook WH_JOURNALRECORD, HI_GLOBAL

cmdPlayback.Enabled = False

End Sub

Private Sub cmdReleaseHook_Click()

Dim heErrRet As HOOKERRORS

On Error GoTo ReleaseHook_Error

heErrRet = gHookDLL.ReleaseHook(cboHookType.ItemData(cboHookType.ListIndex))
If heErrRet <> ERROR_SUCCESS Then
    MsgBox "The following error has occured: " & heErrRet & "  therefore could not release hook!"
End If

'set the status checkbox for the hook set
Dim oCtl As Control

For Each oCtl In Me.Controls
    If InStr(UCase(oCtl.Name), Trim(UCase(cboHookType.Text))) Then
        oCtl.Value = 0
        Exit For
    End If
Next oCtl

Exit Sub

ReleaseHook_Error:
    MsgBox "The following error has occured: " & vbCr & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description
    
End Sub

Private Sub cmdSetHook_Click()

Dim heErrRet As HOOKERRORS
Dim sMsg As String
Dim lHookType As HookTypes

On Error GoTo SetHook_Error

If Trim(cboHookType.Text) = "" Then Exit Sub

lHookType = cboHookType.ItemData(cboHookType.ListIndex)

If lHookType = WH_JOURNALRECORD Then
    Kill App.Path & "\Journal.jrn"
End If

If gbDLL_LOADED Then

    heErrRet = gHookDLL.SetUpHook(lHookType, IIf(optGlobal.Value, HI_GLOBAL, HI_APPLICATION))

    If heErrRet <> ERROR_SUCCESS Then
        gHookDLL.ResolveError heErrRet, sMsg
        MsgBox "Error '" & sMsg & "' occured!"
        Exit Sub
    End If
    
    'set the status checkbox for the hook set
    Dim oCtl As Control
    
    For Each oCtl In Me.Controls
        If InStr(UCase(oCtl.Name), Trim(UCase(cboHookType.Text))) Then
            oCtl.Value = 1
            Exit For
        End If
    Next oCtl
Else
    MsgBox "Unable to set global hook because DLL module failed to load!"
End If

Exit Sub

SetHook_Error:
    MsgBox "The following error has occured: " & vbCr & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description

End Sub

Private Sub Form_Load()

Dim lRet As Long

Set gHookDLL = New CHooks

On Error Resume Next
gHookDLL.Init
If Err.Number <> 0 Then
    MsgBox "Error occured initialising global hook DLL!" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description & vbCr & vbCr & _
    "This application will now close down...", vbCritical + vbOKOnly, "Fatal application error"
    gHookDLL.UnHookAll
    Set gHookDLL = Nothing
    End
Else
    gbDLL_LOADED = True
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)


'ensure I have removed all hooks
With gHookDLL
    .RestoreWndProc
    .UnHookAll
End With


Set gHookDLL = Nothing


End Sub

Private Function ResolveHookType(ByVal HookType As HookTypes) As String

Select Case HookType

Case WH_CALLWNDPROC
    ResolveHookType = "WH_CALLWNDPROC"
    
Case WH_CBT
    ResolveHookType = "WH_CBT"
    
Case WH_DEBUG
    ResolveHookType = "WH_DEBUG"
    
Case WH_KEYBOARD
    ResolveHookType = "WH_KEYBOARD"
    
Case WH_GETMESSAGE
    ResolveHookType = "WH_GETMESSAGE"
    
Case WH_MSGFILTER
    ResolveHookType = "WH_MSGFILTER"
    
Case WH_MOUSE
    ResolveHookType = "WH_MOUSE"
    
Case WH_SHELL
    ResolveHookType = "WH_SHELL"
    
End Select


End Function

Private Sub gHookDLL_CallWndProcEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'lparam contains a pointer to a CWPSTRUCT
Dim tWndStruct As CWPSTRUCT
Dim sMsg As String

On Error GoTo CallWndProcEvent_Error

'flag return value
ReturnVal = 0

'CopyMemory tWndStruct, ByVal lParam, Len(tWndStruct)
With gHookDLL.CWPData
    'wParam is non-zero if sent by this thread(it should be for application scope hooks)
    If .wParam Then
        txtMsg = txtMsg & "Message received on this thread..." & vbCrLf
    Else
        txtMsg = txtMsg & "Message received from another thread..." & vbCrLf
    End If
    
    If .wParam = 0 Then
        sMsg = ResolveWindowsMsg(.message)
        txtMsg = txtMsg & "** RECEIVED CALLWNDPROC MSG **" & vbCrLf & _
        "Handle: " & Hex(.hWnd) & "  Msg: " & IIf(sMsg = "", "Message not currently handled!", sMsg) & "  wParam: " & Hex(.wParam) & _
        "  lParam: " & Hex(.lParam) & vbCrLf & vbCrLf
    End If
End With

Exit Sub

CallWndProcEvent_Error:
    txtMsg = txtMsg & "**Error occured in CallWndProcEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description
    
End Sub

Private Sub gHookDLL_CBTEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

Dim eHookCode As CBTHookCodes
Dim tActivateWnd As CBTACTIVATESTRUCT
Dim tMouseInfo As MOUSEHOOKSTRUCT
Dim tWndRect As RECT
Dim tKeyInfo As KeyStrokeInfo
Dim tCreateWnd As CBT_CREATEWND
On Error GoTo CBTEvent_Error

'flag return value
ReturnVal = 0

eHookCode = nCode

Select Case eHookCode

Case HCBT_ACTIVATE
    'in this case wParam is a handle to the window about to be activated
    'lParam is a pointer to a CBTACTIVATESTRUCT
    'CopyMemory tActivateWnd, ByVal lParam, Len(tActivateWnd)
    With gHookDLL.CBTActivateData
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_ACTIVATE **" & vbCrLf & _
        "hWnd about to be activated: " & Hex(wParam) & vbCrLf & _
        IIf(.fMouse, "Window activated by mouse click", "Window activated by application") & _
        vbCrLf & "Handle to active window: " & IIf(.hWndActive, Hex(.hWndActive), "Active window is in another process!") & vbCrLf & vbCrLf
    End With
    
    If chkDisableWndCreate.Value = 1 Then
        ReturnVal = 1
    End If
    
Case HCBT_CLICKSKIPPED
    'wParam indicates the mouse msg retrieved from the queue
    'lParam contains pointer to MOUSEHOOKSTRUCT
    'CopyMemory tMouseInfo, ByVal lParam, Len(tMouseInfo)
    With gHookDLL.CBTMouseData
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_CLICKSKIPPED **" & vbCrLf & _
        "Extra info: " & Hex(.dwExtraInfo) & "  hWnd receiving msg: " & Hex(.hWnd) & vbCrLf & _
        "Hit test code: " & Hex(.wHitTestCode) & "  PT X: " & .pt.x & "  PT Y: " & .pt.y & vbCrLf & vbCrLf
    End With
    
Case HCBT_CREATEWND
    'wParam specifies handle to new window
    'lParam pointer to CBT_CREATEWND
    txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_CREATEWND **" & vbCrLf & _
    "hWnd of window about to be created: " & Hex(wParam) & vbCrLf & vbCrLf
    
Case HCBT_DESTROYWND
    'wParam specifies window handle about to be destroyed
    txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_DESTROYWND **" & vbCrLf & _
    "hWnd of window about to be destroyed: " & Hex(wParam) & vbCrLf & vbCrLf
    
Case HCBT_KEYSKIPPED
    '** THIS WILL ONLY WORK IF WH_KEYBOARD IS HOOKED ALSO **
    'wParam virtual keycode
    'lParam key information
    
    With gHookDLL.CBTKeySkipped
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_KEYSKIPPED **" & vbCrLf & _
        "Virtual keycode: " & Hex(wParam) & "  ASCII: " & Chr(wParam) & vbCrLf & _
        "Context code: " & .ContCode & "  Ext key: " & .ExtKey & "  PrevState: " & .PreviousKeyState & vbCrLf & _
        "Repeatcount: " & Hex(.RepeatCount) & "  Scancode: " & Hex(.ScanCode) & "  Transition: " & .TransitionState & vbCrLf & vbCrLf
    End With
    
Case HCBT_MINMAX
    'wParam specifies handle to window
    'lParam contains SW_ value
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_MINMAX **" & vbCrLf & _
        "hWnd: " & Hex(wParam) & "  Show status: " & ResolveSW(gHookDLL.CBTMinMaxData) & vbCrLf & vbCrLf
    
    
Case HCBT_MOVESIZE
    'wParam handle of window moved/sized
    'lParam pointer to RECT
    'CopyMemory tWndRect, ByVal lParam, Len(tWndRect)
    With gHookDLL.CBTMoveSizeData
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_MOVESIZE **" & vbCrLf & _
        "hWnd: " & Hex(wParam) & vbCrLf & _
        "Bottom: " & .Bottom & "  Left: " & .Left & "  Top: " & .Top & "  Right: " & .Right & vbCrLf & vbCrLf
    End With

    'here I am attempting to stop the movement of what ever window was moved
    'setting returnval to 1 prevents the message being passed on
    ReturnVal = 1
    Exit Sub
    
Case HCBT_SETFOCUS
    'wParam handle to window gaining keyboard focus
    'lParam handle to window losing keyboard focus
    txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_SETFOCUS **" & vbCrLf & _
    "hWnd gaining focus: " & Hex(wParam) & "  hWnd losing focus: " & IIf(lParam = 0, "Window in another process", Hex(lParam)) & vbCrLf & vbCrLf
    
Case HCBT_SYSCOMMAND
    With gHookDLL.CBTSysCommandData
        txtMsg = txtMsg & "** RECEIVED CBTPROC MSG - HCBT_SYSCOMMAND **" & vbCrLf & _
        "System command: " & .scSysCommand & "  X: " & .x & "  Y: " & .y & vbCrLf & vbCrLf
    End With
    ReturnVal = 1
    
End Select

Exit Sub

CBTEvent_Error:
    txtMsg = txtMsg & "**Error occured in CBTEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf
    
End Sub

Private Sub gHookDLL_DebugEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

On Error GoTo DebugEvent_Error

ReturnVal = 0

'wParam contains the Hook type
'lParam pointer to DEBUGHOOKINFO structure

With gHookDLL.DebugData
    txtMsg = txtMsg & "** RECEIVED DEBUGPROC MSG **" & vbCrLf
    txtMsg = txtMsg & "HookType: " & ResolveHookType(wParam) & vbCrLf
    txtMsg = txtMsg & "ThreadID: " & .lThreadID & "  ThreadInstallerID: " & Hex(.lThreadInstallerID) & _
    "  wParam: " & Hex(.wParam) & "  lParam: " & Hex(.lParam) & "  nCode: " & .nCode & vbCrLf & vbCrLf
End With

Exit Sub

DebugEvent_Error:
    txtMsg = txtMsg & "**Error occured in DebugEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf

End Sub

Private Sub gHookDLL_ForeGroundIdleEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'not going to implement this hook at this time

End Sub

Private Sub gHookDLL_GetMessageEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'nCode is Hook code
'wParam is Removal option
'lParam pointer to MSG structure
Dim eMsgDets As Msg
Dim sMsg As String

On Error GoTo GetMessageEvent_Error

ReturnVal = 0

txtMsg = txtMsg & "** RECEIVED GETMESSAGEPROC MSG ***" & vbCrLf
txtMsg = txtMsg & "Removal option: " & ResolveRemovalCode(wParam) & vbCrLf

CopyMemory eMsgDets, ByVal lParam, Len(eMsgDets)

With eMsgDets
    sMsg = ResolveWindowsMsg(.message)
    txtMsg = txtMsg & "hWnd: " & .hWnd & "  wMsg: " & IIf(sMsg = "", "Message not currently handled!", sMsg) & "  wParam: " & .wParam & _
    "  lParam: " & .lParam & vbCrLf & vbCrLf
End With

Exit Sub

GetMessageEvent_Error:
    txtMsg = txtMsg & "**Error occured in GetMessageEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf & vbCrLf

End Sub

Private Sub gHookDLL_JournalPlayBackEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'nCode is Hook code
'wParam is not used
'lParam  contains a ptr to the playback data (EVENTMSG)
Static fFile As Long
Static bOpenFile As Boolean

Dim lMsg As Long
Dim lparamL As Long
Dim lparamH As Long
Dim lTime As Long
Dim lhWnd As Long

If nCode = HC_GETNEXT Then

If Not bOpenFile Then
    bOpenFile = True
    fFile = FreeFile
    Open App.Path & "Journal.jrn" For Input As #fFile
    Input #fFile, lMsg, lparamL, lparamH, lTime, lhWnd
    With gHookDLL.JournalEventData
        .hWnd = lhWnd
        .message = lMsg
        .paramH = lparamH
        .paramL = lparamL
        .Time = lTime
    End With
Else
    If Not EOF(fFile) Then
        Input #fFile, lMsg, lparamL, lparamH, lTime, lhWnd
        With gHookDLL.JournalEventData
            .hWnd = lhWnd
            .message = lMsg
            .paramH = lparamH
            .paramL = lparamL
            .Time = lTime
        End With
    Else
        Close #fFile
        ReturnVal = 1
    End If
End If
Else
End If


End Sub

Private Sub gHookDLL_JournalRecordEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'nCode is Hook code
'wParam is not used
'lParam ptr to EVENTMSG structure (not these appear as a property of the BattNotes DLL

Dim fFile As Long

On Error GoTo JournalRecordEvent_Error

txtMsg = txtMsg & "** RECEIVED JOURNALRECORD MSG ***" & vbCrLf

With gHookDLL.JournalEventData
    txtMsg = txtMsg & "hWnd: " & .hWnd & " wMsg: " & ResolveWindowsMsg(.message) & " paramH: " & .paramH & " paramL: " & .paramL & _
    " time: " & .Time
End With

'here I am going to attempt to save the data
fFile = FreeFile

Open App.Path & "\Journal.jrn" For Append As #fFile

With gHookDLL.JournalEventData
    Print #fFile, Trim(.hWnd), Trim(.message), Trim(.paramH), Trim(.paramL), Trim(.Time)
End With

Close #fFile

Exit Sub

JournalRecordEvent_Error:
    txtMsg = txtMsg & "**Error occured in JournalRecordEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf & vbCrLf

End Sub

Private Sub gHookDLL_KeyBoardEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)
'wParam specifies virtual keycode
'lParam specifies keystroke data
Dim eKeyInfo As KeyStrokeInfo

On Error GoTo KeyBoardEvent_Error

ReturnVal = 0

eKeyInfo = DecodeKeyInfo(lParam)

With eKeyInfo
    txtMsg = txtMsg & "** RECEIVED KEYBOARDPROC MSG **" & vbCrLf & _
    "Virtual keycode: " & Hex(wParam) & "  ASCII: " & Chr(wParam) & vbCrLf & _
    "Context code: " & .ContCode & "  Ext key: " & .ExtKey & "  PrevState: " & .PreviousKeyState & vbCrLf & _
    "Repeatcount: " & Hex(.RepeatCount) & "  Scancode: " & Hex(.ScanCode) & "  Transition: " & .TransitionState & vbCrLf & vbCrLf
End With



Exit Sub

KeyBoardEvent_Error:
    txtMsg = txtMsg & "**Error occured in KeyBoardEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf & vbCrLf

End Sub

Private Sub gHookDLL_MouseEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'wParam contains mouse message identifier
'lParam pointer to MOUSEHOOKSTRUCT

Dim eMouseInfo As MOUSEHOOKSTRUCT
Dim sMsg As String

On Error GoTo MouseEvent_Error

ReturnVal = 0

txtMsg = txtMsg & "** RECEIVED MOUSEPROC MSG **" & vbCrLf
sMsg = ResolveWindowsMsg(wParam)
txtMsg = txtMsg & "Mouse message: " & IIf(sMsg = "", "Message not currently handled!", sMsg) & ":" & Hex(wParam) & vbCrLf

CopyMemory eMouseInfo, ByVal lParam, Len(eMouseInfo)

With eMouseInfo
    txtMsg = txtMsg & "Extra: " & Hex(.dwExtraInfo) & "  hWnd: " & Hex(.hWnd) & vbCrLf & _
    "HitTest: " & ResolveHitTest(.wHitTestCode) & "  X Co-ord: " & .pt.x & "  Y Co-ord: " & .pt.y & vbCrLf & vbCrLf
End With


Exit Sub

MouseEvent_Error:
    txtMsg = txtMsg & "**Error occured in MouseEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf & vbCrLf

End Sub

Private Function ResolveHitTest(ByVal HitTestCode As HitTestCodes) As String

Select Case HitTestCode

Case HTBORDER
    ResolveHitTest = "HTBORDER"
    
Case HTBOTTOM
    ResolveHitTest = "HTBOTTOM"
    
Case HTBOTTOMLEFT
    ResolveHitTest = "HTBOTTOMLEFT"
    
Case HTBOTTOMRIGHT
    ResolveHitTest = "HTBOTTOMRIGHT"
    
Case HTCAPTION
    ResolveHitTest = "HTCAPTION"
    
Case HTCLIENT
    ResolveHitTest = "HTCLIENT"
    
Case HTERROR
    ResolveHitTest = "HTERROR"
    
Case HTGROWBOX
    ResolveHitTest = "HTGROWBOX"
    
Case HTHSCROLL
    ResolveHitTest = "HTHSCROLL"
    
Case HTLEFT
    ResolveHitTest = "HTLEFT"
    
Case HTMAXBUTTON
    ResolveHitTest = "HTMAXBUTTON"
    
Case HTMENU
    ResolveHitTest = "HTMENU"
    
Case HTMINBUTTON
    ResolveHitTest = "HTMINBUTTON"
    
Case HTNOWHERE
    ResolveHitTest = "HTNOWHERE"
    
Case HTREDUCE
    ResolveHitTest = "HTREDUCE"
    
Case HTRIGHT
    ResolveHitTest = "HTRIGHT"
    
Case HTSIZE
    ResolveHitTest = "HTSIZE"
    
Case HTSIZEFIRST
    ResolveHitTest = "HTSIZEFIRST"
    
Case HTSIZELAST
    ResolveHitTest = "HTSIZELAST"
    
Case HTSYSMENU
    ResolveHitTest = "HTSYSMENU"
    
Case HTTOP
    ResolveHitTest = "HTTOP"
    
Case HTTOPLEFT
    ResolveHitTest = "HTTOPLEFT"
    
Case HTTOPRIGHT
    ResolveHitTest = "HTTOPRIGHT"
    
Case HTTRANSPARENT
    ResolveHitTest = "HTTRANSPARENT"
    
Case HTVSCROLL
    ResolveHitTest = "HTVSCROLL"
    
Case HTZOOM
    ResolveHitTest = "HTZOOM"
    
End Select


End Function

Private Function ResolveMouseMsg(ByVal MouseMsg As WindowsMessages) As String

Select Case MouseMsg

Case WM_LBUTTONDBLCLK
    ResolveMouseMsg = "WM_LBUTTONDBLCLK"
    
Case WM_LBUTTONDOWN
    ResolveMouseMsg = "WM_LBUTTONDOWN"
    
Case WM_LBUTTONUP
    ResolveMouseMsg = "WM_LBUTTONUP"
    
Case WM_MBUTTONDBLCLK
    ResolveMouseMsg = "WM_MBUTTONDBLCLK"
    
Case WM_MBUTTONDOWN
    ResolveMouseMsg = "WM_MBUTTONDOWN"
    
Case WM_MBUTTONUP
    ResolveMouseMsg = "WM_MBUTTONUP"
    
Case WM_MOUSEMOVE, WM_MOUSEFIRST
    ResolveMouseMsg = "WM_MOUSEMOVE/MOUSEFIRST"
    
Case WM_RBUTTONDBLCLK
    ResolveMouseMsg = "WM_RBUTTONDBLCLK"
    
Case WM_RBUTTONDOWN
    ResolveMouseMsg = "WM_RBUTTONDOWN"
    
Case WM_RBUTTONUP
    ResolveMouseMsg = "WM_RBUTTONUP"
    
Case WM_NCMOUSEMOVE
    ResolveMouseMsg = "WM_NCMOUSEMOVE"
    
Case WM_NCLBUTTONDBLCLK
    ResolveMouseMsg = "WM_NCLBUTTONDBLCLK"
    
Case WM_NCLBUTTONDOWN
    ResolveMouseMsg = "WM_NCLBUTTONDOWN"
    
Case WM_NCLBUTTONUP
    ResolveMouseMsg = "WM_NCLBUTTONUP"
    
Case WM_NCMBUTTONDBLCLK
    ResolveMouseMsg = "WM_NCMBUTTONDBLCLK"
    
Case WM_NCMBUTTONDOWN
    ResolveMouseMsg = "WM_NCMBUTTONDOWN"
    
Case WM_NCMBUTTONUP
    ResolveMouseMsg = "WM_NCMBUTTONUP"
    
Case WM_NCHITTEST
    ResolveMouseMsg = "WM_NCHITTEST"
    
End Select

End Function

Private Sub gHookDLL_MsgFilterEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)

'nCode is MsgFilterCode
'wParam is NULL
'lParam pointer to MSG
Dim eMsg As Msg
Dim sMsg As String

ReturnVal = 0

On Error GoTo MsgFilterEvent_Error

txtMsg = txtMsg & "** RECEIVED MSGFILTERPROC MSG **" & vbCrLf
txtMsg = txtMsg & "FilterCode: " & ResolveFilterCode(nCode) & vbCrLf

CopyMemory eMsg, ByVal lParam, Len(eMsg)


With gHookDLL.MessageProcData
    sMsg = ResolveWindowsMsg(.message)
    txtMsg = txtMsg & "hWnd: " & .hWnd & "  wMsg: " & IIf(sMsg = "", "Message not currently handled!", sMsg) & "  wParam: " & .wParam & _
    "  lParam: " & .lParam & vbCrLf & "  ptX: " & .ptX & "  ptY: " & .ptY & "  Time: " & .Time & "  MsgFilter: " & .MsgFilter & vbCrLf & vbCrLf
End With

DoEvents

Exit Sub

MsgFilterEvent_Error:
    txtMsg = txtMsg & "**Error occured in MsgFilterEvent:" & vbCrLf & _
    "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description & vbCrLf & vbCrLf
    
End Sub

Private Function ResolveRemovalCode(ByVal RemovalOpt As RemovalOptions) As String

Select Case RemovalOpt

Case PM_NOREMOVE
    ResolveRemovalCode = "PM_NOREMOVE"
    
Case PM_REMOVE
    ResolveRemovalCode = "PM_REMOVE"

End Select


End Function

Private Sub gHookDLL_ShellEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ReturnVal As Long)
'nCode specifies ShellHookCode
'wParam handle depending on nCode
'lParam ignored for Shell code we are processing
Dim eShell As ShellHookCodes
Dim tRect As RECT

eShell = nCode

ReturnVal = 0

Select Case nCode

Case HSHELL_ACTIVATESHELLWINDOW, HSHELL_WINDOWCREATED, HSHELL_WINDOWDESTROYED
    txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & "  hWnd: " & Hex(wParam) & vbCrLf

Case HSHELL_WINDOWACTIVATED
    txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & "  hWnd: " & Hex(wParam) & "  Maximized: " & CBool(lParam) & vbCrLf
    
Case HSHELL_GETMINRECT
    'CopyMemory tRect, ByVal lParam, Len(tRect)
    With gHookDLL.ShellGetMinRectData
        txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & "  hWnd: " & Hex(wParam) & _
        " Top: " & .Top & " Left: " & .Left & " Bottom: " & .Bottom & " Right: " & .Right & vbCrLf
    End With

Case HSHELL_LANGUAGE
    txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & "  hWnd: " & Hex(wParam) & " hKeybd: " & Hex(lParam) & vbCrLf

Case HSHELL_REDRAW
    txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & "  hWnd: " & Hex(wParam) & " Flashing: " & CBool(lParam) & vbCrLf
    
Case HSHELL_TASKMAN
    txtMsg = txtMsg & "ShellCode: " & ResolveShellCode(nCode) & vbCrLf
    
End Select

End Sub

Private Function ResolveShellCode(ByVal ShellCode As ShellHookCodes) As String

Select Case ShellCode

Case HSHELL_WINDOWCREATED
    ResolveShellCode = "HSHELL_WINDOWCREATED"
    
Case HSHELL_WINDOWDESTROYED
    ResolveShellCode = "HSHELL_WINDOWDESTROYED"

Case HSHELL_ACTIVATESHELLWINDOW
    ResolveShellCode = "HSHELL_ACTIVATESHELLWINDOW"
    
Case HSHELL_GETMINRECT
    ResolveShellCode = "HSHELL_GETMINRECT"
    
Case HSHELL_LANGUAGE
    ResolveShellCode = "HSHELL_LANGUAGE"
    
Case HSHELL_REDRAW
    ResolveShellCode = "HSHELL_REDRAW"
    
Case HSHELL_TASKMAN
    ResolveShellCode = "HSHELL_TASKMAN"
    
Case HSHELL_WINDOWACTIVATED
    ResolveShellCode = "HSHELL_WINDOWACTIVATED"
    
End Select


End Function


Private Function ResolveWindowsMsg(ByVal WndMsg As WindowsMessages) As String

ResolveWindowsMsg = ""

Select Case WndMsg

Case WM_ACTIVATE
    ResolveWindowsMsg = "WM_ACTIVATE"
    
Case WM_ACTIVATEAPP
    ResolveWindowsMsg = "WM_ACTIVATEAPP"
    
Case WM_ASKCBFORMATNAME
    ResolveWindowsMsg = "WM_ASKCBFORMATNAME"
    
Case WM_CANCELJOURNAL
    ResolveWindowsMsg = "WM_CANCELJOURNAL"
    
Case WM_CANCELMODE
    ResolveWindowsMsg = "WM_CANCELMODE"
    
Case WM_CHANGECBCHAIN
    ResolveWindowsMsg = "WM_CHANGECBCHAIN"
    
Case WM_CHAR
    ResolveWindowsMsg = "WM_CHAR"
    
Case WM_CHARTOITEM
    ResolveWindowsMsg = "WM_CHARTOITEM"
    
Case WM_CHILDACTIVATE
    ResolveWindowsMsg = "WM_CHILDACTIVATE"
    
Case WM_USER
    ResolveWindowsMsg = "WM_USER"
    
Case WM_CHOOSEFONT_GETLOGFONT
    ResolveWindowsMsg = "WM_CHOOSEFONT_GETLOGFONT"
    
Case WM_CHOOSEFONT_SETFLAGS
    ResolveWindowsMsg = "WM_CHOOSEFONT_SETFLAGS"
    
Case WM_CHOOSEFONT_SETLOGFONT
    ResolveWindowsMsg = "WM_CHOOSEFONT_SETLOGFONT"
    
Case WM_CLEAR
    ResolveWindowsMsg = "WM_CLEAR"
    
Case WM_CLOSE
    ResolveWindowsMsg = "WM_CLOSE"
    
Case WM_COMMAND
    ResolveWindowsMsg = "WM_COMMAND"
    
Case WM_COMMNOTIFY
    ResolveWindowsMsg = "WM_COMMNOTIFY"
    
Case WM_COMPACTING
    ResolveWindowsMsg = "WM_COMPACTING"
    
Case WM_COMPAREITEM
    ResolveWindowsMsg = "WM_COMPAREITEM"
    
Case WM_CONVERTREQUESTEX
    ResolveWindowsMsg = "WM_CONVERTREQUESTEX"
    
Case WM_COPY
    ResolveWindowsMsg = "WM_COPY"
    
Case WM_COPYDATA
    ResolveWindowsMsg = "WM_COPYDATA"
    
Case WM_CREATE
    ResolveWindowsMsg = "WM_CREATE"
    
Case WM_CTLCOLORBTN
    ResolveWindowsMsg = "WM_CTLCOLORBTN"
    
Case WM_CTLCOLORDLG
    ResolveWindowsMsg = "WM_CTLCOLORDLG"
    
Case WM_CTLCOLOREDIT
    ResolveWindowsMsg = "WM_CTLCOLOREDIT"
    
Case WM_CTLCOLORLISTBOX
    ResolveWindowsMsg = "WM_CTLCOLORLISTBOX"
    
Case WM_CTLCOLORMSGBOX
    ResolveWindowsMsg = "WM_CTLCOLORMSGBOX"
    
Case WM_CTLCOLORSCROLLBAR
    ResolveWindowsMsg = "WM_CTLCOLORSCROLLBAR"
    
Case WM_CTLCOLORSTATIC
    ResolveWindowsMsg = "WM_CTLCOLORSTATIC"
    
Case WM_CUT
    ResolveWindowsMsg = "WM_CUT"
    
Case WM_DDE_FIRST
    ResolveWindowsMsg = "WM_DDE_FIRST"
    
Case WM_DDE_ACK
    ResolveWindowsMsg = "WM_DDE_ACK"
    
Case WM_DDE_ADVISE
    ResolveWindowsMsg = "WM_DDE_ADVISE"
    
Case WM_DDE_DATA
    ResolveWindowsMsg = "WM_DDE_DATA"
    
Case WM_DDE_EXECUTE
    ResolveWindowsMsg = "WM_DDE_EXECUTE"
    
Case WM_DDE_INITIATE
    ResolveWindowsMsg = "WM_DDE_INITIATE"
    
Case WM_DDE_LAST
    ResolveWindowsMsg = "WM_DDE_LAST"
    
Case WM_DDE_POKE
    ResolveWindowsMsg = "WM_DDE_POKE"
    
Case WM_DDE_REQUEST
    ResolveWindowsMsg = "WM_DDE_REQUEST"
    
Case WM_DDE_TERMINATE
    ResolveWindowsMsg = "WM_DDE_TERMINATE"
    
Case WM_DDE_UNADVISE
    ResolveWindowsMsg = "WM_DDE_UNADVISE"
    
Case WM_DEADCHAR
    ResolveWindowsMsg = "WM_DEADCHAR"
    
Case WM_DELETEITEM
    ResolveWindowsMsg = "WM_DELETEITEM"
    
Case WM_DESTROY
    ResolveWindowsMsg = "WM_DESTROY"
    
Case WM_DESTROYCLIPBOARD
    ResolveWindowsMsg = "WM_DESTROYCLIPBOARD"
    
Case WM_DEVMODECHANGE
    ResolveWindowsMsg = "WM_DEVMODECHANGE"
    
Case WM_DRAWCLIPBOARD
    ResolveWindowsMsg = "WM_DRAWCLIPBOARD"
    
Case WM_DRAWITEM
    ResolveWindowsMsg = "WM_DRAWITEM"
    
Case WM_DROPFILES
    ResolveWindowsMsg = "WM_DROPFILES"
    
Case WM_ENABLE
    ResolveWindowsMsg = "WM_ENABLE"
    
Case WM_ENDSESSION
    ResolveWindowsMsg = "WM_ENDSESSION"
    
Case WM_ENTERIDLE
    ResolveWindowsMsg = "WM_ENTERIDLE"
    
Case WM_ENTERMENULOOP
    ResolveWindowsMsg = "WM_ENTERMENULOOP"
    
Case WM_ERASEBKGND
    ResolveWindowsMsg = "WM_ERASEBKGND"
    
Case WM_EXITMENULOOP
    ResolveWindowsMsg = "WM_EXITMENULOOP"
    
Case WM_FONTCHANGE
    ResolveWindowsMsg = "WM_FONTCHANGE"
    
Case WM_GETDLGCODE
    ResolveWindowsMsg = "WM_GETDLGCODE"
    
Case WM_GETFONT
    ResolveWindowsMsg = "WM_GETFONT"
    
Case WM_GETHOTKEY
    ResolveWindowsMsg = "WM_GETHOTKEY"
    
Case WM_GETMINMAXINFO
    ResolveWindowsMsg = "WM_GETMINMAXINFO"
    
Case WM_GETTEXT
    ResolveWindowsMsg = "WM_GETTEXT"
    
Case WM_GETTEXTLENGTH
    ResolveWindowsMsg = "WM_GETTEXTLENGTH"
    
Case WM_HOTKEY
    ResolveWindowsMsg = "WM_HOTKEY"
    
Case WM_HSCROLL
    ResolveWindowsMsg = "WM_HSCROLL"
    
Case WM_HSCROLLCLIPBOARD
    ResolveWindowsMsg = "WM_HSCROLLCLIPBOARD"
    
Case WM_ICONERASEBKGND
    ResolveWindowsMsg = "WM_ICONERASEBKGND"
    
Case WM_IME_CHAR
    ResolveWindowsMsg = "WM_IME_CHAR"
    
Case WM_IME_COMPOSITION
    ResolveWindowsMsg = "WM_IME_COMPOSITION"
    
Case WM_IME_COMPOSITIONFULL
    ResolveWindowsMsg = "WM_IME_COMPOSITIONFULL"
    
Case WM_IME_CONTROL
    ResolveWindowsMsg = "WM_IME_CONTROL"
    
Case WM_IME_ENDCOMPOSITION
    ResolveWindowsMsg = "WM_IME_ENDCOMPOSITION"
    
Case WM_IME_KEYDOWN
    ResolveWindowsMsg = "WM_IME_KEYDOWN"
    
Case WM_IME_KEYLAST
    ResolveWindowsMsg = "WM_IME_KEYLAST"
    
Case WM_IME_KEYUP
    ResolveWindowsMsg = "WM_IME_KEYUP"
    
Case WM_IME_NOTIFY
    ResolveWindowsMsg = "WM_IME_NOTIFY"
    
Case WM_IME_SELECT
    ResolveWindowsMsg = "WM_IME_SELECT"
    
Case WM_IME_SETCONTEXT
    ResolveWindowsMsg = "WM_IME_SETCONTEXT"
    
Case WM_IME_STARTCOMPOSITION
    ResolveWindowsMsg = "WM_IME_STARTCOMPOSITION"
    
Case WM_INITDIALOG
    ResolveWindowsMsg = "WM_INITDIALOG"
    
Case WM_INITMENU
    ResolveWindowsMsg = "WM_INITMENU"
    
Case WM_INITMENUPOPUP
    ResolveWindowsMsg = "WM_INITMENUPOPUP"
    
Case WM_KEYDOWN
    ResolveWindowsMsg = "WM_KEYDOWN"
    
Case WM_KEYFIRST
    ResolveWindowsMsg = "WM_KEYFIRST"
    
Case WM_KEYLAST
    ResolveWindowsMsg = "WM_KEYLAST"
    
Case WM_KEYUP
    ResolveWindowsMsg = "WM_KEYUP"
    
Case WM_KILLFOCUS
    ResolveWindowsMsg = "WM_KILLFOCUS"
    
Case WM_LBUTTONDBLCLK
    ResolveWindowsMsg = "WM_LBUTTONDBLCLK"
    
Case WM_LBUTTONDOWN
    ResolveWindowsMsg = "WM_LBUTTONDOWN"
    
Case WM_LBUTTONUP
    ResolveWindowsMsg = "WM_LBUTTONUP"
    
Case WM_MBUTTONDBLCLK
    ResolveWindowsMsg = "WM_MBUTTONDBLCLK"
    
Case WM_MBUTTONDOWN
    ResolveWindowsMsg = "WM_MBUTTONDOWN"
    
Case WM_MBUTTONUP
    ResolveWindowsMsg = "WM_MBUTTONUP"
    
Case WM_MDIACTIVATE
    ResolveWindowsMsg = "WM_MDIACTIVATE"
    
Case WM_MDICASCADE
    ResolveWindowsMsg = "WM_MDICASCADE"
    
Case WM_MDICREATE
    ResolveWindowsMsg = "WM_MDICREATE"
    
Case WM_MDIDESTROY
    ResolveWindowsMsg = "WM_MDIDESTROY"
    
Case WM_MDIGETACTIVE
    ResolveWindowsMsg = "WM_MDIGETACTIVE"
    
Case WM_MDIICONARRANGE
    ResolveWindowsMsg = "WM_MDIICONARRANGE"
    
Case WM_MDIMAXIMIZE
    ResolveWindowsMsg = "WM_MDIMAXIMIZE"
    
Case WM_MDINEXT
    ResolveWindowsMsg = "WM_MDINEXT"
    
Case WM_MDIREFRESHMENU
    ResolveWindowsMsg = "WM_MDIREFRESHMENU"
    
Case WM_MDIRESTORE
    ResolveWindowsMsg = "WM_MDIRESTORE"
    
Case WM_MDISETMENU
    ResolveWindowsMsg = "WM_MDISETMENU"
    
Case WM_MDITILE
    ResolveWindowsMsg = "WM_MDITILE"
    
Case WM_MEASUREITEM
    ResolveWindowsMsg = "WM_MEASUREITEM"
    
Case WM_MENUCHAR
    ResolveWindowsMsg = "WM_MENUCHAR"
    
Case WM_MENUSELECT
    ResolveWindowsMsg = "WM_MENUSELECT"
    
Case WM_MOUSEACTIVATE
    ResolveWindowsMsg = "WM_MOUSEACTIVATE"
    
Case WM_MOUSEFIRST
    ResolveWindowsMsg = "WM_MOUSEFIRST"
    
Case WM_MOUSELAST
    ResolveWindowsMsg = "WM_MOUSELAST"
    
Case WM_MOUSEMOVE
    ResolveWindowsMsg = "WM_MOUSEMOVE"
    
Case WM_MOVE
    ResolveWindowsMsg = "WM_MOVE"
    
Case WM_NCACTIVATE
    ResolveWindowsMsg = "WM_NCACTIVATE"
    
Case WM_NCCALCSIZE
    ResolveWindowsMsg = "WM_NCCALCSIZE"
    
Case WM_NCCREATE
    ResolveWindowsMsg = "WM_NCCREATE"
    
Case WM_NCDESTROY
    ResolveWindowsMsg = "WM_NCDESTROY"
    
Case WM_NCHITTEST
    ResolveWindowsMsg = "WM_NCHITTEST"
    
Case WM_NCLBUTTONDBLCLK
    ResolveWindowsMsg = "WM_NCLBUTTONDBLCLK"
    
Case WM_NCLBUTTONDOWN
    ResolveWindowsMsg = "WM_NCLBUTTONDOWN"
    
Case WM_NCLBUTTONUP
    ResolveWindowsMsg = "WM_NCLBUTTONUP"
    
Case WM_NCMBUTTONDBLCLK
    ResolveWindowsMsg = "WM_NCMBUTTONDBLCLK"
    
Case WM_NCMBUTTONDOWN
    ResolveWindowsMsg = "WM_NCMBUTTONDOWN"
    
Case WM_NCMBUTTONUP
    ResolveWindowsMsg = "WM_NCMBUTTONUP"
    
Case WM_NCMOUSEMOVE
    ResolveWindowsMsg = "WM_NCMOUSEMOVE"
    
Case WM_NCPAINT
    ResolveWindowsMsg = "WM_NCPAINT"
    
Case WM_NCRBUTTONDBLCLK
    ResolveWindowsMsg = "WM_NCRBUTTONDBLCLK"
    
Case WM_NCRBUTTONDOWN
    ResolveWindowsMsg = "WM_NCRBUTTONDOWN"
    
Case WM_NCRBUTTONUP
    ResolveWindowsMsg = "WM_NCRBUTTONUP"
    
Case WM_NEXTDLGCTL
    ResolveWindowsMsg = "WM_NEXTDLGCTL"
    
Case WM_NULL
    ResolveWindowsMsg = "WM_NULL"
    
Case WM_OTHERWINDOWCREATED
    ResolveWindowsMsg = "WM_OTHERWINDOWCREATED"
    
Case WM_OTHERWINDOWDESTROYED
    ResolveWindowsMsg = "WM_OTHERWINDOWDESTROYED"
    
Case WM_PAINT
    ResolveWindowsMsg = "WM_PAINT"
    
Case WM_PAINTCLIPBOARD
    ResolveWindowsMsg = "WM_PAINTCLIPBOARD"
    
Case WM_PAINTICON
    ResolveWindowsMsg = "WM_PAINTICON"
    
Case WM_PALETTECHANGED
    ResolveWindowsMsg = "WM_PALETTECHANGED"
    
Case WM_PALETTEISCHANGING
    ResolveWindowsMsg = "WM_PALETTEISCHANGING"
    
Case WM_PARENTNOTIFY
    ResolveWindowsMsg = "WM_PARENTNOTIFY"
    
Case WM_PASTE
    ResolveWindowsMsg = "WM_PASTE"
    
Case WM_PENWINFIRST
    ResolveWindowsMsg = "WM_PENWINFIRST"
    
Case WM_PENWINLAST
    ResolveWindowsMsg = "WM_PENWINLAST"
    
Case WM_POWER
    ResolveWindowsMsg = "WM_POWER"
    
Case WM_PSD_ENVSTAMPRECT
    ResolveWindowsMsg = "WM_PSD_ENVSTAMPRECT"
    
Case WM_PSD_FULLPAGERECT
    ResolveWindowsMsg = "WM_PSD_FULLPAGERECT"
    
Case WM_PSD_GREEKTEXTRECT
    ResolveWindowsMsg = "WM_PSD_GREEKTEXTRECT"
    
Case WM_PSD_MARGINRECT
    ResolveWindowsMsg = "WM_PSD_MARGINRECT"
    
Case WM_PSD_MINMARGINRECT
    ResolveWindowsMsg = "WM_PSD_MINMARGINRECT"
    
Case WM_PSD_PAGESETUPDLG
    ResolveWindowsMsg = "WM_PSD_PAGESETUPDLG"
    
Case WM_PSD_YAFULLPAGERECT
    ResolveWindowsMsg = "WM_PSD_YAFULLPAGERECT"
    
Case WM_QUERYDRAGICON
    ResolveWindowsMsg = "WM_QUERYDRAGICON"
    
Case WM_QUERYENDSESSION
    ResolveWindowsMsg = "WM_QUERYENDSESSION"
    
Case WM_QUERYNEWPALETTE
    ResolveWindowsMsg = "WM_QUERYNEWPALETTE"
    
Case WM_QUERYOPEN
    ResolveWindowsMsg = "WM_QUERYOPEN"
    
Case WM_QUEUESYNC
    ResolveWindowsMsg = "WM_QUEUESYNC"
    
Case WM_QUIT
    ResolveWindowsMsg = "WM_QUIT"
    
Case WM_RBUTTONDBLCLK
    ResolveWindowsMsg = "WM_RBUTTONDBLCLK"
    
Case WM_RBUTTONDOWN
    ResolveWindowsMsg = "WM_RBUTTONDOWN"
    
Case WM_RBUTTONUP
    ResolveWindowsMsg = "WM_RBUTTONUP"
    
Case WM_RENDERALLFORMATS
    ResolveWindowsMsg = "WM_RENDERALLFORMATS"
    
Case WM_RENDERFORMAT
    ResolveWindowsMsg = "WM_RENDERFORMAT"
    
Case WM_SETCURSOR
    ResolveWindowsMsg = "WM_SETCURSOR"
    
Case WM_SETFOCUS
    ResolveWindowsMsg = "WM_SETFOCUS"
    
Case WM_SETFONT
    ResolveWindowsMsg = "WM_SETFONT"
    
Case WM_SETHOTKEY
    ResolveWindowsMsg = "WM_SETHOTKEY"
    
Case WM_SETREDRAW
    ResolveWindowsMsg = "WM_SETREDRAW"
    
Case WM_SETTEXT
    ResolveWindowsMsg = "WM_SETTEXT"
    
Case WM_SHOWWINDOW
    ResolveWindowsMsg = "WM_SHOWWINDOW"
    
Case WM_SIZE
    ResolveWindowsMsg = "WM_SIZE"
    
Case WM_SIZECLIPBOARD
    ResolveWindowsMsg = "WM_SIZECLIPBOARD"
    
Case WM_SPOOLERSTATUS
    ResolveWindowsMsg = "WM_SPOOLERSTATUS"
    
Case WM_SYSCHAR
    ResolveWindowsMsg = "WM_SYSCHAR"
    
Case WM_SYSCOLORCHANGE
    ResolveWindowsMsg = "WM_SYSCOLORCHANGE"
    
Case WM_SYSCOMMAND
    ResolveWindowsMsg = "WM_SYSCOMMAND"
    
Case WM_SYSDEADCHAR
    ResolveWindowsMsg = "WM_SYSDEADCHAR"
    
Case WM_SYSKEYDOWN
    ResolveWindowsMsg = "WM_SYSKEYDOWN"
    
Case WM_SYSKEYUP
    ResolveWindowsMsg = "WM_SYSKEYUP"
    
Case WM_TIMECHANGE
    ResolveWindowsMsg = "WM_TIMECHANGE"
    
Case WM_TIMER
    ResolveWindowsMsg = "WM_TIMER"
    
Case WM_UNDO
    ResolveWindowsMsg = "WM_UNDO"
    
Case WM_VKEYTOITEM
    ResolveWindowsMsg = "WM_VKEYTOITEM"
    
Case WM_VSCROLL
    ResolveWindowsMsg = "WM_VSCROLL"
    
Case WM_VSCROLLCLIPBOARD
    ResolveWindowsMsg = "WM_VSCROLLCLIPBOARD"
    
Case WM_WINDOWPOSCHANGED
    ResolveWindowsMsg = "WM_WINDOWPOSCHANGED"
    
Case WM_WINDOWPOSCHANGING
    ResolveWindowsMsg = "WM_WINDOWPOSCHANGING"
    
Case WM_WININICHANGE
    ResolveWindowsMsg = "WM_WININICHANGE"
    
End Select

End Function

