Attribute VB_Name = "MHookMod"
Option Explicit



'Hook declares for in-process hooks
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

'Hook declares for global hooks (reside in C++ DLL)
Public Declare Function SetUpHookGlobal Lib "WinHook.dll" (ByVal HookType As Integer) As Long
Public Declare Function UnHookGlobal Lib "WinHook.dll" (ByVal HookType As Integer) As Long
Public Declare Function RegisterNewMessageGlobal Lib "WinHook.dll" () As Integer
Public Declare Function InitDLL Lib "WinHook.dll" (ByVal hWnd As Long) As Long
Public Declare Function CloseDownDLL Lib "WinHook.dll" () As Long

'declares for loading library
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'declares for sub-classing
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'support declares
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'the following array holds info relating to the hooks
'whether they are application hooks or global, and whether they are active

Public gStatusOfHooks(11) As HookStatus

'****************************************************
'global vars
'the DLL module instance handle
Public ghInstDLL As Long

'the sub-class old window proc address
Public gOldWndProc As Long

'the new registered message, for the first instance
Public giNewMessage As Long

'subclass flag
Public gbSUBCLASSED As Boolean

'global reference to the CHooks class
Public goHook As CHooks

'file handle for logging
Public gfFile As Long

'***************************************
'proc addresses for global hooks
Public glpCallWndProc As Long
Public glpCBTProc As Long
Public glpDebugProc As Long
Public glpForeGroundIdleProc As Long
Public glpGetMsgProc As Long
Public glpJournalPlaybackProc As Long
Public glpJournalRecordProc As Long
Public glpKeyBoardProc As Long
Public glpMessageProc As Long
Public glpMouseProc As Long
Public glpShellProc As Long
Public glpSysMsgProc As Long

'constants

'sub-classing
Public Const GWL_WNDPROC = (-4)

'the DLL file and path
Public Const gsDLLFile As String = "WinHook.dll"

'DLL procedure names, these are the export names from the DLL
Public Const sCallWndProc As String = "CallWndProc"
Public Const sCBTProc As String = "CBTProc"
Public Const sDebugProc As String = "DebugProc"
Public Const sForeGroundIdleProc As String = "ForeGroundIdleProc"
Public Const sGetMsgProc As String = "GetMsgProc"
Public Const sJournalPlaybackProc As String = "JournalPlaybackProc"
Public Const sJournalRecordProc As String = "JournalRecordProc"
Public Const sKeyBoardProc As String = "KeyboardProc"
Public Const sMessageProc As String = "MessageProc"
Public Const sMouseProc As String = "MouseProc"
Public Const sShellProc As String = "ShellProc"
Public Const sSysMsgProc As String = "SysMsgProc"

'other constants
Public Const VK_CANCEL = &H3


Public Function CBTProc(ByVal nCode As CBTHookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo CBTProc_Error
    
    'lParam points to various structures depending on the message
    'wParam specifies various handles or codes depending upon the message
    
    'set up local required structs
    
    'For HCBT_ACTIVATE, wParam - Handle of window to be activated, lParam - ptr to CBTACTIVATESTRUCT
    
    'For HCBT_CLICKSKIPPED, wParam - Mouse message removed from queue, lParam - ptr to MOUSEHOOKSTRUCT
    
    'For HCBT_CREATEWND, wParam - Handle to new window, lParam - ptr to CBT_CREATEWND struct
    
    'For HCBT_DESTROYWND, wParam - Handle of window being destroyed, lParam is zero
    
    'For HCBT_KEYSKIPPED, wParam - Virtual key code, lParam - Specifies the repeat count, scan code,
    'key-transition code, previous key state, and context code
    
    'For HCBT_MINMAX, wParam - Handle of window being min/max, lParam - LOWORD showwindow value
    '**REMEMBER to use LOWORD function when assigning value to swSizeWindow
    
    'For HCBT_MOVESIZE, wParam - Handle of window being moved or sized, lParam - ptr to RECT
    
    'For HCBT_QS, both wParam and lParam are undefined and will be zero
    
    'For HCBT_SETFOCUS, wParam - Handle of wnd gaining keybd focus, lParam - Handle of wnd losing focus
    
    'For HCBT_SYSCOMMAND, wParam - System command value, lParam - Not used unless mouse used to select
    'system menu item, in which case loword contains x screen co-ordinate, hiword y screen co-ordinate
    '**REMEMBER to convert lParam into LOWORD and HIWORD values for screen co-ordinates
    
    Dim lReturnVal As Long
    
    'if no action is required then pass on to next filter in chain
    If nCode < HCBT_ACTIONREQUIRED Then
        CBTProc = CallNextHookEx(gStatusOfHooks(1).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    'process and extract required property values
    Select Case nCode
        Case HCBT_ACTIVATE
            Dim tActivateData As CBTACTIVATESTRUCT
            CopyMemory tActivateData, ByVal lParam, Len(tActivateData)
            goHook.CBTActivateData = tActivateData
            
            
        Case HCBT_CLICKSKIPPED
            Dim tMouseData As MOUSEHOOKSTRUCT
            CopyMemory tMouseData, ByVal lParam, Len(tMouseData)
            goHook.CBTMouseData = tMouseData

        Case HCBT_CREATEWND
            '**************************************************************************************************
            '   Copying memory whilst receiving the HCBT_CREATEWND code just seems to cause a GPF
            '   To retrieve window class information, the client will have to use the hWnd passed in wParam
            '**************************************************************************************************
'
'        Case HCBT_DESTROYWND
'            'nothing needs doing here since wParam holds the handle to the destroyed window, lParam is not used
'
        Case HCBT_KEYSKIPPED
            'remember that the WH_KEYBOARD hook must be in place for this code to be raised
            goHook.CBTKeySkipped = DecodeKeyInfo(lParam)
'
        Case HCBT_MINMAX
            goHook.CBTMinMaxData = LOWORD(lParam)

        Case HCBT_MOVESIZE
            Dim tMoveSize As RECT
            CopyMemory tMoveSize, ByVal lParam, Len(tMoveSize)
            goHook.CBTMoveSizeData = tMoveSize

'        Case HCBT_SETFOCUS
'            'nothing here since wParam is hWnd to window gaining kbd focus, lParam hWnd of window losing focus
'
        Case HCBT_SYSCOMMAND
            With goHook.CBTSysCommandData
                .scSysCommand = wParam
                .x = LOWORD(lParam) 'x and y are valid if a system menu command is selected with the mouse
                .y = HIWORD(lParam)
            End With
            
    End Select
    
        
    '....... Perform processing here
    LogEvent fcCBTProc, "Signalling CBTProc event via COM..."
    
    goHook.SignalEvent WH_CBT, nCode, wParam, lParam, lReturnVal
    
    LogEvent fcCBTProc, "...returned from signalling event"
    If nCode = HCBT_CREATEWND Then
        'apparently you have to exit the function here or the call to CallNextHookEx will cause a GPF
        Exit Function
    End If
    
    'callee returns non-zero if message is to be discarded
    If lReturnVal <> 0 Then
        CBTProc = 1
        Exit Function
    End If
    
    
    'remember to pass on to the next filter in the chain
    CBTProc = CallNextHookEx(gStatusOfHooks(1).HS_HOOKID, nCode, wParam, lParam)

    Exit Function
    
CBTProc_Error:
    'ensure an error doesn't stop the call to the next filter proc in the chain
    CBTProc = CallNextHookEx(gStatusOfHooks(1).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Function CallWndProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo CallWndProc_Error
    
    'wParam will be non-zero when message sent by current thread
    'lParam is a ptr to CWPSTRUCT
    Dim lReturnVal As Long
    Dim tCWPData As CWPSTRUCT
    
    'if no action is required then pass on to next filter in chain
    If nCode < HC_ACTIONREQUIRED Then
        CallWndProc = CallNextHookEx(gStatusOfHooks(0).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    'since we are only interested in the current thread, check that wParam is greater than zero
    If wParam > 0 Then
        'copy the data pointed to by lParam into my local
        CopyMemory tCWPData, ByVal lParam, Len(tCWPData)
        goHook.CWPData = tCWPData
        
        '...... Perform required processing here
        goHook.SignalEvent WH_CALLWNDPROC, nCode, wParam, lParam, lReturnVal
    
        'does returnval indicate not to propagate the message
        If lReturnVal = 0 Then
            CallWndProc = 0
            Exit Function
        End If
    End If
    
    'remember to pass on to the next filter in the chain
    CallWndProc = CallNextHookEx(gStatusOfHooks(0).HS_HOOKID, nCode, wParam, lParam)

    Exit Function
    
CallWndProc_Error:
    'ensure an error doesn't stop the call to the next filter proc in the chain
    CallWndProc = CallNextHookEx(gStatusOfHooks(0).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function DebugProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo DebugProc_Error
    
    'wParam is the hooktype showing hook proc about to be called
    'lParam - ptr to debughookinfo struct
    Dim dhDebugInfo As DEBUGHOOKINFO
    Dim lReturnVal As Long
    
    If nCode < HC_ACTIONREQUIRED Then
        DebugProc = CallNextHookEx(gStatusOfHooks(2).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    '....... Perform processing here
    'copy the data pointed to by lParam into my private local
    CopyMemory dhDebugInfo, ByVal lParam, Len(dhDebugInfo)
    
    With goHook
        .DebugData = dhDebugInfo
        .SignalEvent WH_DEBUG, nCode, wParam, lParam, lReturnVal
    End With
    'if return is not zero then prevent the hook procedure from being called by the system
    If lReturnVal <> 0 Then
        DebugProc = 1
        Exit Function
    End If
    
    '*****Return non-zero to prevent the hook procedure from being called, then exit
    
    'remember to pass on to the next filter in the chain
    DebugProc = CallNextHookEx(gStatusOfHooks(2).HS_HOOKID, nCode, wParam, lParam)

    Exit Function
    
DebugProc_Error:
    'ensure an error doesn't stop the call to the next filter proc in the chain
    DebugProc = CallNextHookEx(gStatusOfHooks(2).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function GetMsgProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo GetMsgProc_Error
    
    'wParam - removal code
    'lParam - ptr to a MSG structure
    Dim msMSG As Msg
    Dim lReturnVal As Long
    
    If nCode < HC_ACTIONREQUIRED Then
        GetMsgProc = CallNextHookEx(gStatusOfHooks(4).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    '........ Perform required processing here
    CopyMemory msMSG, ByVal lParam, Len(msMSG)
    
    With goHook.Get_MsgData
        .hWnd = msMSG.hWnd
        .lParam = msMSG.lParam
        .message = msMSG.message
        .ptX = msMSG.pt.x
        .ptY = msMSG.pt.y
        .RemovalOption = wParam
        .Time = msMSG.Time
        .wParam = msMSG.wParam
    End With
        
    goHook.SignalEvent WH_GETMESSAGE, nCode, wParam, lParam, lReturnVal
        
    'remember to pass on to the next filter in the chain
    GetMsgProc = CallNextHookEx(gStatusOfHooks(4).HS_HOOKID, nCode, wParam, lParam)

    Exit Function
    
GetMsgProc_Error:
    'ensure an error doesn't stop the call to the next filter proc in the chain
    GetMsgProc = CallNextHookEx(gStatusOfHooks(4).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Function Shellproc(ByVal nCode As ShellHookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo ShellProc_Error
    
    Dim tRect As RECT
    
    'THIS IS A NOTIFICATION HOOK ONLY, YOU CANNOT CHANGE ANY PARAMETERS HERE
    
    'For HSHELL_ACTIVATESHELLWINDOW, not used at this time
    
    'For HSHELL_WINDOWCREATED, wParam - Handle to created window, lParam not used
    
    'For HSHELL_WINDOWDESTROYED, wParam - Handle to window to be destroyed, lParam not used
    
    Dim lReturnVal As Long
    
    If nCode < HSHELL_ACTIONREQUIRED Then
        Shellproc = CallNextHookEx(gStatusOfHooks(10).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    LogEvent fcShellProc, "** RECEIVED WH_SHELL NOTIFICATION **"
    LogEvent fcShellProc, "nCode: " & nCode & " wParam: " & wParam & " lPAram: " & lParam
    
    If nCode = HSHELL_GETMINRECT Then
        CopyMemory tRect, lParam, Len(tRect)
        'unlike almost every other CopyMemory we've done up to now, we do not use ByVal with lParam
        'if we do then no data is retrieved - WHAT IS THE EXPLANATION FOR THIS!!??
        goHook.ShellGetMinRectData = tRect
        With tRect
            LogEvent fcShellProc, "Top: " & .Top & " Left: " & .Left & " Bottom: " & .Bottom & " Right: " & .Right
        End With
    End If
    
    '........ Perform required processing here
    goHook.SignalEvent WH_SHELL, nCode, wParam, lParam, lReturnVal
    
    'remember to pass on to the next filter in the chain
    Shellproc = CallNextHookEx(gStatusOfHooks(10).HS_HOOKID, nCode, wParam, lParam)

    Exit Function
    
ShellProc_Error:
    'do not fail without passing onto next hook filter
    Shellproc = CallNextHookEx(gStatusOfHooks(10).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function KeyboardProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    'For HC_ACTION, The wParam and lParam parameters contain information about a keystroke message
    
    'For HC_NOREMOVE, The wParam and lParam parameters contain information about a keystroke message,
    'and the keystroke message has not been removed from the message queue. (An application called the
    'PeekMessage function, specifying the PM_NOREMOVE flag)
    
    'wParam has virtual key code
    'lParam contains key information
    On Error GoTo KeyboardProc_Error
    
    Dim ksKeyInfo As KeyStrokeInfo
    Dim lReturnVal As Long
    
    If nCode < HC_ACTIONREQUIRED Then
        KeyboardProc = CallNextHookEx(gStatusOfHooks(7).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    '......... Perform required processing here
    With goHook
        .KeyboardData = DecodeKeyInfo(lParam)
    
        .SignalEvent WH_KEYBOARD, nCode, wParam, lParam, lReturnVal
    End With
    
    'callee returns positive when message is to be discarded
    If lReturnVal <> 0 Then
        KeyboardProc = 1
        Exit Function
    End If
    
    'remember to pass on to the next filter in the chain
    KeyboardProc = CallNextHookEx(gStatusOfHooks(7).HS_HOOKID, nCode, wParam, lParam)
    
    Exit Function
    
KeyboardProc_Error:
    'do not fail without passing onto next hook filter
    KeyboardProc = CallNextHookEx(gStatusOfHooks(7).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Function MouseProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    '**NOTE - DO NOT ATTEMPT TO INSTALL A JOURNALPLAYBACK HOOK FROM HERE
    
    On Error GoTo MouseProc_Error
    
    'if mouse message is WM_NCHITTEST then we need a local for hit test codes
    Dim htHitTest As HitTestCodes
    
    Dim lReturnVal As Long
    
    'lparam  - ptr to MOUSEHOOKSTRUCT
    Dim mhMouse As MOUSEHOOKSTRUCT
    
    If nCode < HC_ACTIONREQUIRED Then
        MouseProc = CallNextHookEx(gStatusOfHooks(8).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    '......... Perform required processing here
    CopyMemory mhMouse, ByVal lParam, Len(mhMouse)
    
    With goHook.MouseData
        .dwExtraInfo = mhMouse.dwExtraInfo
        .hWnd = mhMouse.hWnd
        .mIdentifier = wParam
        .ptX = mhMouse.pt.x
        .ptY = mhMouse.pt.y
        .wHitTestCode = mhMouse.wHitTestCode
    End With
        
    goHook.SignalEvent WH_MOUSE, nCode, wParam, lParam, lReturnVal
    
    If lReturnVal <> 0 Then
        MouseProc = 1
        Exit Function
    End If
    
    
    'remember to pass on to the next filter in the chain
    MouseProc = CallNextHookEx(gStatusOfHooks(8).HS_HOOKID, nCode, wParam, lParam)
    
    Exit Function
    
MouseProc_Error:
    'do not fail without passing onto next hook filter
    MouseProc = CallNextHookEx(gStatusOfHooks(8).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function MessageProc(ByVal nCode As MsgFilterCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    'wParam - not used
    'lParam - ptr to MSG struct
    Dim msMSG As Msg
    Dim lReturnVal As Long
    Dim tMsg As MessageData
    
    On Error GoTo MessageProc_Error
    
    If nCode < MSGF_ACTIONREQUIRED Then
        MessageProc = CallNextHookEx(gStatusOfHooks(9).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    '......... Perform required processing here
    CopyMemory msMSG, ByVal lParam, Len(msMSG)
    
    With tMsg
        .hWnd = msMSG.hWnd
        .lParam = msMSG.lParam
        .message = msMSG.message
        .MsgFilter = wParam
        .ptX = msMSG.pt.x
        .ptY = msMSG.pt.y
        .Time = msMSG.Time
        .wParam = msMSG.wParam
    End With
    
    goHook.MessageProcData = tMsg
    
    LogEvent fcMessageProc, "** RECEIVED WH_MSGFILTER MESSAGE **"
    With tMsg
        LogEvent fcMessageProc, "hWnd: " & .hWnd & "  lParam: " & .lParam & "  Msg: " & .message
        LogEvent fcMessageProc, "ptX: " & .ptX & "  ptY: " & .ptY & "  Time: " & .Time & "  wParam: " & .wParam
    End With
    
    goHook.SignalEvent WH_MSGFILTER, nCode, wParam, lParam, lReturnVal
    
    'callee returns non-zero if message is to be discarded
    If lReturnVal <> 0 Then
        MessageProc = 1
        Exit Function
    End If
    
    
    'remember to pass on to the next filter in the chain
    MessageProc = CallNextHookEx(gStatusOfHooks(9).HS_HOOKID, nCode, wParam, lParam)
    
    Exit Function
    
MessageProc_Error:
    'do not fail without passing onto next hook filter
    MessageProc = CallNextHookEx(gStatusOfHooks(9).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function LOWORD(ByVal lValue As Long) As Long

'retrieves the low-order word of a value
LOWORD = lValue And &HFFFF

End Function

Public Function HIWORD(ByVal lValue As Long) As Long

'retrieves the high-order word of a value
HIWORD = lValue And &HFFFF0000

End Function


Public Function DecodeKeyInfo(ByVal ksKeyStroke As Long) As KeyStrokeInfo

'decodes the keystroke information in a long
Dim ksKeyInfo As KeyStrokeInfo

With ksKeyInfo
    If ksKeyStroke < 0 Then ksKeyStroke = ksKeyStroke + 2147483647
    .RepeatCount = ksKeyStroke And &HFFFF                      'bits 0-15
    .ScanCode = ksKeyStroke And &HFF0000                       'bits 16-23
    .ExtKey = IIf(ksKeyStroke And &H1000000, 1, 0)             'bit 24
    .Reserved = ksKeyStroke And &H1E000000                     'bits 25-28
    .ContCode = IIf(ksKeyStroke And &H20000000, 1, 0)          'bit 29
    .PreviousKeyState = IIf(ksKeyStroke And &H40000000, 1, 0)  'bit 30
    .TransitionState = IIf(ksKeyStroke And &H80000000, 1, 0)   'bit 31
End With

'return struct with info
DecodeKeyInfo = ksKeyInfo

End Function

Public Function ForeGroundIdleProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    'neither wParam nor lParam are used, this is an idle notification only
    
    On Error GoTo ForeGroundIdleProc_Error
    
    'for reqd parameter to signalevent, it is not used
    Dim ReturnVal As Long
    
    If nCode < HC_ACTIONREQUIRED Then
        ForeGroundIdleProc = CallNextHookEx(gStatusOfHooks(3).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If
    
    
    goHook.SignalEvent WH_FOREGROUNDIDLE, nCode, wParam, lParam, ReturnVal
    
    'remember to pass on to the next filter in the chain
    ForeGroundIdleProc = CallNextHookEx(gStatusOfHooks(3).HS_HOOKID, nCode, wParam, lParam)
    
    Exit Function
    
ForeGroundIdleProc_Error:
    'do not fail without passing onto next hook filter
    ForeGroundIdleProc = CallNextHookEx(gStatusOfHooks(3).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Public Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim ReturnVal As Long
Dim tCopyData As COPYDATASTRUCT
Dim tActivateDets As CBTACTIVATESTRUCT
Dim aByte(7) As Byte

On Error GoTo WindowProc_Error

LogEvent fcWindowProc, "In Window proc, wMsg is: " & wMsg & "  wParam: " & wParam & "  lParam: " & lParam
'is the message our registered one
If wMsg = giNewMessage Then
    '....Perform required processing here

    'in this prototype, the WM_BATTMESSAGE is defined in the C++ DLL
    'this is what we will be looking for here.

    'wParam holds the action code, lParam contains handle to destroyed window
    If wParam = HCBT_DESTROYWND Then
        LogEvent fcWindowProc, "**SIGNALLING client of HCBT_DESTROYWND message passed by CBTProc: hWnd destroyed " & Hex(lParam)
        goHook.SignalEvent WH_CBT, HCBT_DESTROYWND, lParam, 0, ReturnVal
    End If

    'flag completion of message
    WindowProc = 0
    Exit Function
End If

'this is where you will process messages posted from the global hook procedure in the DLL
'generally (if I can get it to work) these will be WM_COPYDATA messages, but it could be anything
If wMsg = WM_COPYDATA Then
    LogEvent fcWindowProc, "**Received WM_COPYDATA msg... hWnd of sender:" & Hex(wParam) & "  ptr to data: " & Hex(lParam)
    'shall I try it???????????????
    CopyMemory tCopyData, ByVal lParam, Len(tCopyData)
    LogEvent fcWindowProc, "The dwData entry has: " & tCopyData.dwData
'    LogEvent fcWindowProc, "**I THINK we just copied the data!!"
'    LogEvent fcWindowProc, "Here goes, in dwData I should have the value 999, the copied data is: " & tCopyData.dwData
    LogEvent fcWindowProc, "If that is correct, then the remaining ones should be - cbData: " & tCopyData.cbData & _
    "  lpData: " & Hex(tCopyData.lpData)
    'good up to here, so now try to copy the CBTACTIVATESTRUCT
    CopyMemory tActivateDets, ByVal tCopyData.lpData, tCopyData.cbData
    With tActivateDets
    LogEvent fcWindowProc, "*******************, fMouse: " & Hex(.fMouse) & "  hWndActive: " & Hex(.hWndActive)
    End With
    
    'let's have a look at the actual data shall we??
    CopyMemory aByte(0), ByVal tCopyData.lpData, 8
    Dim a As Integer
    For a = 0 To 7
        LogEvent fcWindowProc, "Byte number " & a & " is: " & Hex(aByte(a))
    Next a
    WindowProc = 1
    Exit Function
End If

'pass all other messages on for default processing
WindowProc = CallWindowProc(gOldWndProc, hWnd, wMsg, wParam, lParam)

Exit Function

WindowProc_Error:
    'do not fail without passing onto old window procedure
    WindowProc = CallWindowProc(gOldWndProc, hWnd, wMsg, wParam, lParam)

End Function

Public Sub LogEvent(ByVal SrcFunc As FuncCodes, ByVal MsgToLog As String, Optional fLogCode As LogCodes = bcLogEventTypeInformation)

Dim sLogCode As String
Dim sOPMsg As String

Select Case fLogCode
    Case bcLogEventTypeError
        sLogCode = "ERROR: "
        
    Case bcLogEventTypeWarning
        sLogCode = "WARNING: "
        
    Case bcLogEventTypeInformation
        sLogCode = "INFORMATION: "
End Select

'compose o/p msg
sOPMsg = sLogCode & Time & " Func: " & ResolveFuncCodes(SrcFunc) & _
"   " & MsgToLog & vbNewLine

Print #gfFile, sOPMsg



End Sub

Public Function ResolveFuncCodes(ByVal SrcFunc As FuncCodes) As String

Select Case SrcFunc

Case fcClassInitialize
    ResolveFuncCodes = "Class_Initialize"
    
Case fcCloseUp
    ResolveFuncCodes = "CloseUp"
    
Case fcClearHookStatus
    ResolveFuncCodes = "ClearHookStatus"
    
Case fcFreeLibraryDLL
    ResolveFuncCodes = "FreeLibraryDLL"
    
Case fcInit
    ResolveFuncCodes = "Init"
    
Case fcInitializeHookStatus
    ResolveFuncCodes = "InitializeHookStatus"
    
Case fcLoadHookDLL
    ResolveFuncCodes = "LoadHookDLL"
    
Case fcRegisterNewMessage
    ResolveFuncCodes = "RegisterNewMessage"
    
Case fcReleaseHook
    ResolveFuncCodes = "ReleaseHook"
    
Case fcResolveError
    ResolveFuncCodes = "ResolveError"
    
Case fcResolveHookType
    ResolveFuncCodes = "ResolveHookType"
    
Case fcResolveScopeIdentifier
    ResolveFuncCodes = "ResolveScopeIdentifier"
    
Case fcRestoreWndProc
    ResolveFuncCodes = "RestoreWndProc"
    
Case fcSetHookType
    ResolveFuncCodes = "SetHookType"
    
Case fcSetUpHook
    ResolveFuncCodes = "SetUpHook"
    
Case fcSignalEvent
    ResolveFuncCodes = "SignalEvent"
    
Case fcSubClassWnd
    ResolveFuncCodes = "SubClassWnd"
    
Case fcUnHookAll
    ResolveFuncCodes = "UnHookAll"
    
Case fcCallWndProc
    ResolveFuncCodes = "CallWndProc"
    
Case fcCBTProc
    ResolveFuncCodes = "CBTProc"
    
Case fcDebugProc
    ResolveFuncCodes = "DebugProc"
    
Case fcDecodeKeyInfo
    ResolveFuncCodes = "DecodeKeyInfo"
    
Case fcForeGroundIdleProc
    ResolveFuncCodes = "ForeGroundIdleProc"
    
Case fcGetMsgProc
    ResolveFuncCodes = "GetMsgProc"
    
Case fcHIWORD
    ResolveFuncCodes = "HIWORD"
    
Case fcKeyBoardProc
    ResolveFuncCodes = "KeyBoardProc"
    
Case fcLOWORD
    ResolveFuncCodes = "LOWORD"
    
Case fcMessageProc
    ResolveFuncCodes = "MessageProc"
    
Case fcMouseProc
    ResolveFuncCodes = "MouseProc"
    
Case fcShellProc
    ResolveFuncCodes = "ShellProc"
    
Case fcWindowProc
    ResolveFuncCodes = "WindowProc"
    
End Select

End Function



Public Function JournalPlaybackProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo JournalPlaybackProc_Error
    
    Dim ReturnVal As Long
    Dim tJournalData As EVENTMSG
    
    If nCode < HC_ACTIONREQUIRED Then
        JournalPlaybackProc = CallNextHookEx(gStatusOfHooks(5).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If

    'when the user presses one of the set key combinations to stop journalling, the hooks filters are all removed
    'this is fine unless the client needs to know that this has occured.  Therefore see explanation below for how to
    'detect this
    
    'the playback function requires that a pointer to an EVENTMSG struct be supplied, the client provides this
    'via the JournalEventData property (note that this property also provides the client with the data when using the
    'JournalRecord hook
    
    'inform the client of what to do next
    goHook.SignalEvent WH_JOURNALPLAYBACK, nCode, wParam, lParam, ReturnVal
    
    'the returnval is used by the client to signal if there are no more events to play back
    If ReturnVal Then
        'true indicates that there are no more events to play back
        goHook.ReleaseHook WH_JOURNALPLAYBACK
        'remember, code after here is not executed, Exit Function just reminds you what happens next
        Exit Function
    End If
    
    'there are more events that the client wants to process, they should present this info on HC_SKIP and we process data
    'on the HC_GETNEXT
    If nCode = HC_GETNEXT Then
        tJournalData = goHook.JournalEventData
        lParam = VarPtr(tJournalData)
        JournalPlaybackProc = tJournalData.Time
        Exit Function
    End If
    
    'don't exit without calling any other filters in the chain
    JournalPlaybackProc = CallNextHookEx(gStatusOfHooks(5).HS_HOOKID, nCode, wParam, lParam)
   
    Exit Function
    
JournalPlaybackProc_Error:
    'do not fail without passing onto next hook filter
    JournalPlaybackProc = CallNextHookEx(gStatusOfHooks(5).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Function JournalRecordProc(ByVal nCode As HookCodes, ByVal wParam As Long, ByVal lParam As Long) As Long

    '**********************************************************************************************************
        'Because all messages are funneled through to a single message queue, if one application on the system
        'hangs, the whole system will grind to a halt because messages are serialised (they have to be if they
        'are routed through to a single message queue).  For this reason Microsoft allows CTRL-ESC or CTRL-ALT-DEL
        'key combinations to remove all filter functions from the chain of this hook.  When this happens, a
        'WM_CANCELJOURNAL message is sent with its hWnd member set to NULL.  This causes a problem because we cannot
        'use sub-classing to intercept the key combination (since hWnd is null, we have no window to sub-class).
        'Therefore, when using this hook, it is a good idea to also use the WH_GETMESSAGE hook and watch for the
        'WM_CANCELJOURNAL message.  When this is received, simply remove the hooks to stop journalling.
        'Microsoft also recommends that CTRL-BREAK be used as a general way of cancelling a hook, although the system
        'doesn't implement this key combination.  In this procedure we will test for this key combination in order
        'to allow the user to cancel journalling.
        'Note that because of the special role these key combinations perform in Journalling, it is not possible
        'to record these key combinations.
    '**********************************************************************************************************
        
    On Error GoTo JournalRecordProc_Error
    
    Dim heErr As HOOKERRORS
    Dim ReturnVal As Long
    Dim tJournalData As EVENTMSG
    
    If nCode < HC_ACTIONREQUIRED Then
        JournalRecordProc = CallNextHookEx(gStatusOfHooks(6).HS_HOOKID, nCode, wParam, lParam)
        Exit Function
    End If

    If nCode = HC_ACTION Then
        'grab the event details
        CopyMemory tJournalData, ByVal lParam, Len(tJournalData)
        goHook.JournalEventData = tJournalData
        
        'the first thing I want to do is check for the key combination to cancel the journalling
        If tJournalData.message = WindowsMessages.WM_KEYDOWN And LOWORD(tJournalData.paramL) = VK_CANCEL Then
            LogEvent fcJournalRecordProc, "** Detected cancel key combination, attempting to unhook..."
            heErr = goHook.ReleaseHook(WH_JOURNALRECORD)
        End If
        'Note - if the key combination is detected, then any other code after the ReleaseHook will be ignored.  This is
        'because the filter (where we are now) is removed from the hook chain, and therefore so is the code.
        'The Exit Function simply shows that this is what occurs, it never actually gets executed.
        With tJournalData
            LogEvent fcJournalRecordProc, "wMsg: " & .message & " paramL: " & .paramL & " paramH: " & .paramH
            LogEvent fcJournalRecordProc, "Time: " & .Time & " hWnd: " & .hWnd
        End With
        '.... signal the client
        goHook.SignalEvent WH_JOURNALRECORD, nCode, wParam, lParam, ReturnVal
        'ReturnVal is ignored
    End If
    'we ignore the HC_SYSMODALON and HC_SYSMODALOFF codes because they only caused problems in Win16
        
    JournalRecordProc = CallNextHookEx(gStatusOfHooks(6).HS_HOOKID, nCode, wParam, lParam)
        
    Exit Function
    
JournalRecordProc_Error:
    'do not fail without passing onto next hook filter
    JournalRecordProc = CallNextHookEx(gStatusOfHooks(6).HS_HOOKID, nCode, wParam, lParam)
    LogEvent fcCBTProc, "**An unexpected error has occured, the system reports the following:", bcLogEventTypeError
    LogEvent fcCBTProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function





