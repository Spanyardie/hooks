VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'private properties
Private mCWPData As CWPSTRUCT

Private mCBTActivateData As CBTACTIVATESTRUCT
Private mCBTMouseData As MOUSEHOOKSTRUCT
Private mCBTCreateData As CBTCREATEDATA
Private mCBTKeySkipped As KeyStrokeInfo
Private mCBTMinMaxData As CBTSizeWndCodes
Private mCBTMoveSizeData As RECT
Private mCBTSysCommandData As CBTSYSCOMMAND_DATA

Private mDebugData As DEBUGHOOKINFO

Private mGetMsgData As GETMSGDATA

Private mKeyboardData As KeyStrokeInfo

Private mMessageProcData As MessageData

Private mMouseData As MouseProcData

Private mJournalEventData As EVENTMSG

Private mShellGetMinRectData As RECT


'event
Public Event CallWndProcEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event CBTEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event DebugEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event ForeGroundIdleEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event GetMessageEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event JournalPlayBackEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event JournalRecordEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event KeyBoardEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event MouseEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event MsgFilterEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)
Public Event ShellEvent(ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnVal As Long)


Public Function SetUpHook(ByVal HookType As HookTypes, ByVal HookScope As HookIdentifier) As HOOKERRORS
    'set up the required hook
    Dim lErRet As HOOKERRORS
    Dim lRet As Long
    Dim sErrDesc As String
    
    LogEvent fcSetUpHook, "Entered function"
    
    'localised error trap
    On Error GoTo SetUpHook_Error
    
    '**********************************************************************************************
    'FOR FIRST VERSION OF THE DLL, ENFORCE APPLICATION (or should I say Thread) SPECIFIC HOOKS ONLY
        HookScope = HI_APPLICATION
    '**********************************************************************************************
        
    'attempt to set the hook type in the local array
    LogEvent fcSetUpHook, "  Attempting to set hook type in local array..."
    lErRet = SetHookType(HookType, HookScope)
    If lErRet <> ERROR_SUCCESS Then
        ResolveError lErRet, sErrDesc
        LogEvent fcSetUpHook, "  ...failed with error '" & sErrDesc & "'", bcLogEventTypeError
        SetUpHook = lErRet
        Exit Function
    End If
    LogEvent fcSetUpHook, "  ...successfully set hook type '" & ResolveHookType(HookType) & "' to " & ResolveScopeIdentifier(HookScope)
    
    Select Case HookType
    
    Case WH_CALLWNDPROC
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(0).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf CallWndProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(0).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_CALLWNDPROC_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_CALLWNDPROC)
            If lRet = 0 Then
                gStatusOfHooks(0).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_CALLWNDPROC_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        
        'successfully set the hook, record details
        With gStatusOfHooks(0)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    
    Case WH_CBT
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(1).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf CBTProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(1).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_CBTPROC_FAIL
                Exit Function
            End If
        Else
            LogEvent fcSetUpHook, "  Calling DLL with (hWnd:" & FHookSubClass.hWnd & ", Hook:" & ResolveHookType(HookType) & ")..."
            'set up a global hook
            lRet = SetUpHookGlobal(WH_CBT)
            If lRet = 0 Then
                gStatusOfHooks(1).HS_ACTIVE = False
                ResolveError lRet, sErrDesc
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_CBTPROC_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(1)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
        
    Case WH_DEBUG
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(2).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf DebugProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(2).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_DEBUG_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_DEBUG)
            If lRet = 0 Then
                gStatusOfHooks(2).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_DEBUG_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(2)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    Case WH_FOREGROUNDIDLE
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(3).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf ForeGroundIdleProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(3).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_FOREGROUNDIDLE_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_FOREGROUNDIDLE)
            If lRet = 0 Then
                gStatusOfHooks(3).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_FOREGROUNDIDLE_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(3)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    Case WH_GETMESSAGE
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(4).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf GetMsgProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(4).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_GETMESSAGE_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_GETMESSAGE)
            If lRet = 0 Then
                gStatusOfHooks(4).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_GETMESSAGE_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(4)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
        
    '**********************************************************************************************
        'The WH_JOURNALPLAYBACK and WH_JOURNALRECORD are defined as global scope only in the
        'Microsoft documentation, however because of the nature of these hooks (messages are all
        'funneled through a single message queue) it is possible to install it within this thread
        'and still receive all messages that the system generates.
    '**********************************************************************************************
        
    Case WH_JOURNALPLAYBACK
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(5).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        'set up a global hook
        lRet = SetWindowsHookEx(HookType, AddressOf JournalPlaybackProc, ghInstDLL, 0) 'note no threadID!!
        If lRet = 0 Then
            gStatusOfHooks(5).HS_ACTIVE = False
            LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
            SetUpHook = HOOK_SETHOOK_JOURNALPLAYBACK_FAIL
            Exit Function
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(5)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
        
    Case WH_JOURNALRECORD
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(6).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        'set up a global hook
        lRet = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf JournalRecordProc, ghInstDLL, 0)
        If lRet = 0 Then
            gStatusOfHooks(6).HS_ACTIVE = False
            LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
            SetUpHook = HOOK_SETHOOK_JOURNALRECORD_FAIL
            Exit Function
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(6)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
        
    Case WH_KEYBOARD
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(7).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf KeyboardProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(7).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_KEYBOARD_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_KEYBOARD)
            If lRet = 0 Then
                gStatusOfHooks(7).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_KEYBOARD_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(7)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    Case WH_MOUSE
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(8).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf MouseProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(8).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_MOUSE_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_MOUSE)
            If lRet = 0 Then
                gStatusOfHooks(8).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_MOUSE_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(8)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    Case WH_MSGFILTER
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(9).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf MessageProc, ghInstDLL, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(9).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_MSGFILTER_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_MSGFILTER)
            If lRet = 0 Then
                gStatusOfHooks(9).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_MSGFILTER_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(9)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
    Case WH_SHELL
        LogEvent fcSetUpHook, " Attempting to set " & _
        ResolveScopeIdentifier(gStatusOfHooks(10).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
        "  DLL address: " & ghInstDLL & _
        "  with thread ID: " & App.ThreadID & "..."
        If HookScope = HI_APPLICATION Then
            'attempt to set hook
            lRet = SetWindowsHookEx(HookType, AddressOf Shellproc, 0, App.ThreadID)
            If lRet = 0 Then
                gStatusOfHooks(10).HS_ACTIVE = False
                SetUpHook = HOOK_SETHOOK_SHELL_FAIL
                Exit Function
            End If
        Else
            'set up a global hook
            lRet = SetUpHookGlobal(WH_SHELL)
            If lRet = 0 Then
                gStatusOfHooks(10).HS_ACTIVE = False
                LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
                SetUpHook = HOOK_SETHOOK_SHELL_FAIL
                Exit Function
            End If
        End If
        
        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
        'successfully set the hook, record details
        With gStatusOfHooks(10)
            .HS_ACTIVE = True
            .HS_HOOKID = lRet
        End With
    
     'RESERVE UNTIL GLOBAL HOOK PORTION IS SORTED OUT
'    Case WH_SYSMSGFILTER
'        LogEvent fcSetUpHook, " Attempting to set " & _
'        ResolveScopeIdentifier(gStatusOfHooks(11).HS_IDENTIFIER) & " hook: " & ResolveHookType(HookType) & _
'        "  DLL address: " & ghInstDLL & _
'        "  with thread ID: " & App.ThreadID & "..."
'        'set up a global hook
'        lRet = SetUpHookGlobal(WH_SYSMSGFILTER)
'        If lRet = 0 Then
'            gStatusOfHooks(11).HS_ACTIVE = False
'            LogEvent fcSetUpHook, " ...failed!!", bcLogEventTypeError
'            SetUpHook = HOOK_SETHOOK_SYSMSGFILTER_FAIL
'            Exit Function
'        End If
'
'        LogEvent fcSetUpHook, " ...succeeded.  Hook handle: " & lRet
'        'successfully set the hook, record details
'        With gStatusOfHooks(11)
'            .HS_ACTIVE = True
'            .HS_HOOKID = lRet
'        End With
    
    
    End Select
        
    'flag success
    SetUpHook = ERROR_SUCCESS
    
    LogEvent fcSetUpHook, "Exiting normally"
    
    Exit Function

SetUpHook_Error:
    LogEvent fcSetUpHook, "The following error occured in 'CHooks.SetUpHook':", bcLogEventTypeError
    LogEvent fcSetUpHook, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    Err.Raise ResolveError(HOOK_UNKNOWNERROR, sErrDesc), "CHooks.SetUpHook", Err.Description
    
End Function

Public Function ReleaseHook(ByVal HookType As HookTypes) As HOOKERRORS

Dim lRet As Long
Dim sErrDesc As String

On Error GoTo ReleaseHook_Error

LogEvent fcReleaseHook, "Entering function"
LogEvent fcReleaseHook, "   Attempting to release hook " & ResolveHookType(HookType) & "..."
Select Case HookType


Case WH_CALLWNDPROC
    With gStatusOfHooks(0)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_CALLWNDPROC_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_CBT
    With gStatusOfHooks(1)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_CBTPROC_FAIL
                Exit Function
            Else
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
            End If
        End If
    End With
        
Case WH_DEBUG
    With gStatusOfHooks(2)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_DEBUG_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_FOREGROUNDIDLE
    With gStatusOfHooks(3)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_FOREGROUNDIDLE_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_GETMESSAGE
    With gStatusOfHooks(4)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_GETMESSAGE_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_JOURNALPLAYBACK
    With gStatusOfHooks(5)
        If .HS_ACTIVE = True Then
            lRet = UnhookWindowsHookEx(.HS_HOOKID)
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook (ID:" & .HS_HOOKID & ")!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_JOURNALPLAYBACK_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_JOURNALRECORD
    With gStatusOfHooks(6)
        If .HS_ACTIVE = True Then
            lRet = UnhookWindowsHookEx(.HS_HOOKID)
            If lRet = 0 Then
                lRet = GetLastError
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook (ID:" & .HS_HOOKID & ")!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with error(" & Err.LastDllError & ")!"
                ReleaseHook = HOOK_UNHOOK_JOURNALRECORD_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_KEYBOARD
    With gStatusOfHooks(7)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_KEYBOARD_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_MOUSE
    With gStatusOfHooks(8)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_MOUSE_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_MSGFILTER
    With gStatusOfHooks(9)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_MSGFILTER_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_SHELL
    With gStatusOfHooks(10)
        If .HS_ACTIVE = True Then
            If .HS_IDENTIFIER = HI_APPLICATION Then
                lRet = UnhookWindowsHookEx(.HS_HOOKID)
            Else
                lRet = UnHookGlobal(.HS_TYPE)
            End If
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_SHELL_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        
Case WH_SYSMSGFILTER
    With gStatusOfHooks(0)
        If .HS_ACTIVE = True Then
            lRet = UnHookGlobal(.HS_TYPE)
            If lRet = 0 Then
                LogEvent fcReleaseHook, "       **Error releasing " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!", bcLogEventTypeError
                LogEvent fcReleaseHook, "Exiting with errors!"
                ReleaseHook = HOOK_UNHOOK_SYSMSGFILTER_FAIL
                Exit Function
            Else
                .HS_ACTIVE = False
                .HS_HOOKID = 0
                .HS_IDENTIFIER = HI_APPLICATION
                LogEvent fcReleaseHook, "   ...successfully released " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " hook!"
            End If
        End If
    End With
        

End Select

LogEvent fcReleaseHook, "Exiting normally"

ReleaseHook = ERROR_SUCCESS

Exit Function

ReleaseHook_Error:
    LogEvent fcReleaseHook, "The following error occured in 'CHooks.ReleaseHook':", bcLogEventTypeError
    LogEvent fcReleaseHook, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    Err.Raise ResolveError(HOOK_UNKNOWNERROR, sErrDesc), "CHooks.ReleaseHook", Err.Description
    
End Function

Private Function InitialiseHookStatus() As HOOKERRORS
'********************************************************************************
'*  Name:           InitialiseHookStatus
'*
'*  Description:    Initialises status array for hooks
'*
'********************************************************************************

Dim hsHookItem As HookStatus

On Error GoTo InitialiseHookStatus_Error

With hsHookItem
    .HS_TYPE = WH_CALLWNDPROC
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(0) = hsHookItem
    
    .HS_TYPE = WH_CBT
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(1) = hsHookItem
    
    .HS_TYPE = WH_DEBUG
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(2) = hsHookItem

    .HS_TYPE = WH_FOREGROUNDIDLE
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(3) = hsHookItem
    
    .HS_TYPE = WH_GETMESSAGE
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(4) = hsHookItem
    
    .HS_TYPE = WH_JOURNALPLAYBACK
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(5) = hsHookItem
    
    .HS_TYPE = WH_JOURNALRECORD
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(6) = hsHookItem
    
    .HS_TYPE = WH_KEYBOARD
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(7) = hsHookItem
    
    .HS_TYPE = WH_MOUSE
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(8) = hsHookItem
    
    .HS_TYPE = WH_MSGFILTER
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(9) = hsHookItem
    
    .HS_TYPE = WH_SHELL
    .HS_IDENTIFIER = HI_APPLICATION
    .HS_ACTIVE = False
    
    gStatusOfHooks(10) = hsHookItem
    
    .HS_TYPE = WH_SYSMSGFILTER
    .HS_IDENTIFIER = HI_GLOBALONLY
    .HS_ACTIVE = False
    
    gStatusOfHooks(11) = hsHookItem
    
End With

'return an OK value
InitialiseHookStatus = ERROR_SUCCESS

Exit Function

InitialiseHookStatus_Error:
    'return error code
    InitialiseHookStatus = HOOK_UNKNOWNERROR
    LogEvent fcInitializeHookStatus, "**Unexpected error occured in InitializeHookStatus", bcLogEventTypeError
    LogEvent fcInitializeHookStatus, "Error Src: " & Err.Source & "  No: " & Err.Number & "  Desc: " & Err.Description, bcLogEventTypeError
End Function

Private Function SetHookType(ByVal htHook As HookTypes, ByVal hiIdentifier As HookIdentifier) As HOOKERRORS
'********************************************************************************
'*  Name:           SetHookType
'*
'*  Description:    Attempts to set a hook type
'*
'********************************************************************************
On Error GoTo SetHookType_Error

LogEvent fcSetHookType, "Entering function"

Select Case htHook

Case WH_CALLWNDPROC
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(0).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_CALLWNDPROC is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_CALLWNDPROC!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(0).HS_IDENTIFIER = hiIdentifier
    
Case WH_CBT
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(1).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_CBT is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_CBT!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(1).HS_IDENTIFIER = hiIdentifier
    LogEvent fcSetHookType, "   Identifier: " & ResolveScopeIdentifier(hiIdentifier)

Case WH_DEBUG
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(2).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_DEBUG is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_DEBUG!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(2).HS_IDENTIFIER = hiIdentifier

Case WH_FOREGROUNDIDLE
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(3).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_FOREGROUNDIDLE is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_FOREGROUNDIDLE!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(3).HS_IDENTIFIER = hiIdentifier

Case WH_GETMESSAGE
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(4).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_GETMESSAGE is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_GETMESSAGE!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(4).HS_IDENTIFIER = hiIdentifier

Case WH_JOURNALPLAYBACK
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(5).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_JOURNALPLAYBACK is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    'it doesnt really matter what the setting is for this hook
   
Case WH_JOURNALRECORD
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(6).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_JOURNALRECORD is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    'it doesnt really matter what the setting is for this hook
    
Case WH_KEYBOARD
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(7).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_KEYBOARD is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_KEYBOARD!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(7).HS_IDENTIFIER = hiIdentifier

Case WH_MOUSE
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(8).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_MOUSE is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_MOUSE!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(8).HS_IDENTIFIER = hiIdentifier

Case WH_MSGFILTER
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(9).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_MSGFILTER is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_MSGFILTER!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(9).HS_IDENTIFIER = hiIdentifier

Case WH_SHELL
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(10).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_SHELL is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier = HI_GLOBALONLY Then
        SetHookType = HOOK_CANT_SET_GLOBALONLY
        LogEvent fcSetHookType, "   can't set GLOBALONLY to WH_SHELL!", bcLogEventTypeWarning
        Exit Function
    End If
    
    gStatusOfHooks(10).HS_IDENTIFIER = hiIdentifier

Case WH_SYSMSGFILTER
    'Don't try to change the type if the hook is active
    If gStatusOfHooks(11).HS_ACTIVE Then
        SetHookType = HOOK_ACTIVE
        LogEvent fcSetHookType, "   WH_SYSMSGFILTER is already active!", bcLogEventTypeWarning
        Exit Function
    End If
    
    'Allow the change only if setting is not globalonly
    If hiIdentifier <> HI_GLOBALONLY Then
        SetHookType = HOOK_GLOBALONLY
        LogEvent fcSetHookType, "   can only set WH_SYSMSGFILTER to GLOBALONLY!", bcLogEventTypeWarning
        Exit Function
    End If
    
End Select

SetHookType = ERROR_SUCCESS

LogEvent fcSetHookType, "Exiting normally"

Exit Function

SetHookType_Error:
    SetHookType = HOOK_UNKNOWNERROR
    LogEvent fcSetHookType, "**An unexpected error occured in SetHookType - system reports the following details:", bcLogEventTypeError
    LogEvent fcSetHookType, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
End Function

Private Function LoadHookDLL() As HOOKERRORS
'attempt to load the DLL into my address space
Dim lRet As Long
Dim heError As HOOKERRORS

On Error GoTo LoadHookDLL_Error

LogEvent fcLoadHookDLL, "Entering LoadHookDLL at " & Time

'before installing the library, check if it has been mapped into my address space already
'because I am running a second instance of this program
LogEvent fcLoadHookDLL, "Retrieving DLL module instance handle..."
lRet = GetModuleHandle(gsDLLFile)
'if it is not loaded, then return will be zero
If lRet = 0 Then
    LogEvent fcLoadHookDLL, "Could not find library, attempting to load...", bcLogEventTypeWarning
    ghInstDLL = LoadLibrary(gsDLLFile)
    If ghInstDLL = 0 Then 'an error occured loading the DLL
        LogEvent fcLoadHookDLL, "Error occured loading the DLL!", bcLogEventTypeError
        LoadHookDLL = ERROR_LOADING_LIBRARY
        Exit Function
    End If
Else
    ghInstDLL = lRet
    LogEvent fcLoadHookDLL, "...successfully retrieved module handle " & ghInstDLL
End If

'now that we have got the module instance handle we can retrieve the proc addresses
LogEvent fcLoadHookDLL, "Attempting to retrieve DLL procedure addresses..."
glpCallWndProc = GetProcAddress(ghInstDLL, sCallWndProc)
If glpCallWndProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for CALLWNDPROC"
    LoadHookDLL = HOOK_PROCADDRESS_CALLWNDPROC_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for CALLWNDPROC: " & glpCallWndProc


glpCBTProc = GetProcAddress(ghInstDLL, sCBTProc)
If glpCBTProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for CBTPROC"
    LoadHookDLL = HOOK_PROCADDRESS_CBTPROC_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for CBTPROC: " & glpCBTProc

glpDebugProc = GetProcAddress(ghInstDLL, sDebugProc)
If glpDebugProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for DEBUGPROC"
    LoadHookDLL = HOOK_PROCADDRESS_DEBUG_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for DEBUGPROC: " & glpDebugProc

glpForeGroundIdleProc = GetProcAddress(ghInstDLL, sForeGroundIdleProc)
If glpForeGroundIdleProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for FOREGROUNDIDLEPROC"
    LoadHookDLL = HOOK_PROCADDRESS_FOREGROUNDIDLE_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for FOREGROUNDIDLEPROC: " & glpForeGroundIdleProc

glpGetMsgProc = GetProcAddress(ghInstDLL, sGetMsgProc)
If glpGetMsgProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for GETMESSAGEPROC"
    LoadHookDLL = HOOK_PROCADDRESS_GETMESSAGE_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for GETMESSAGEPROC: " & glpGetMsgProc

glpJournalPlaybackProc = GetProcAddress(ghInstDLL, sJournalPlaybackProc)
If glpJournalPlaybackProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for JOURNALPLAYBACKPROC"
    LoadHookDLL = HOOK_PROCADDRESS_JOURNALPLAYBACK_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for JOURNALPLAYBACKPROC: " & glpJournalPlaybackProc

glpJournalRecordProc = GetProcAddress(ghInstDLL, sJournalRecordProc)
If glpJournalRecordProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for JOURNALRECORDPROC"
    LoadHookDLL = HOOK_PROCADDRESS_JOURNALRECORD_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for JOURNALRECORDPROC: " & glpJournalRecordProc

glpKeyBoardProc = GetProcAddress(ghInstDLL, sKeyBoardProc)
If glpKeyBoardProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for KEYBOARDPROC"
    LoadHookDLL = HOOK_PROCADDRESS_KEYBOARD_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for KEYBOARDPROC: " & glpKeyBoardProc

glpMessageProc = GetProcAddress(ghInstDLL, sMessageProc)
If glpMessageProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for MSGFILTERPROC"
    LoadHookDLL = HOOK_PROCADDRESS_GETMESSAGE_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for MSGFILTERPROC: " & glpMessageProc

glpMouseProc = GetProcAddress(ghInstDLL, sMouseProc)
If glpMouseProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for MOUSEPROC"
    LoadHookDLL = HOOK_PROCADDRESS_MOUSE_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for MOUSEPROC: " & glpMouseProc

glpShellProc = GetProcAddress(ghInstDLL, sShellProc)
If glpShellProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for SHELLPROC"
    LoadHookDLL = HOOK_PROCADDRESS_SHELL_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for SHELLPROC: " & glpShellProc

glpSysMsgProc = GetProcAddress(ghInstDLL, sSysMsgProc)
If glpSysMsgProc = 0 Then
    LogEvent fcLoadHookDLL, "Failed to retrieve proc address for SYSMSGPROC"
    LoadHookDLL = HOOK_PROCADDRESS_SYSMSGFILTER_FAIL
    Exit Function
End If
LogEvent fcLoadHookDLL, "...retrieved proc address for SYSMSGPROC: " & glpSysMsgProc


LoadHookDLL = ERROR_SUCCESS

LogEvent fcLoadHookDLL, "Exiting LoadHookDLL normally"

Exit Function

LoadHookDLL_Error:
    LoadHookDLL = HOOK_UNKNOWNERROR
    If ghInstDLL Then
        heError = FreeLibraryDLL
        If heError <> ERROR_SUCCESS Then
            LoadHookDLL = ERROR_FREEING_LIBRARY
        End If
    End If
    LogEvent fcLoadHookDLL, "**An unexpected error occured in LoadHookDLL - system reports the following details:", bcLogEventTypeError
    LogEvent fcLoadHookDLL, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
End Function

Private Function FreeLibraryDLL() As HOOKERRORS

Dim lRet As Long

FreeLibraryDLL = ERROR_SUCCESS

If ghInstDLL <> 0 Then
    lRet = FreeLibrary(ghInstDLL)
    If lRet = 0 Then
        LogEvent fcFreeLibraryDLL, "**Error freeing hook library in FreeLibraryDLL!", bcLogEventTypeError
        FreeLibraryDLL = ERROR_FREEING_LIBRARY
    End If
End If


End Function

Private Function ClearHookStatus() As HOOKERRORS

'clear all the details up
Dim iCount As Integer

For iCount = 0 To 11
    With gStatusOfHooks(iCount)
        .HS_ACTIVE = False
        .HS_HOOKID = 0
        If .HS_IDENTIFIER <> HI_GLOBALONLY Then
            .HS_IDENTIFIER = HI_APPLICATION
        End If
    End With
Next iCount

ClearHookStatus = ERROR_SUCCESS

End Function

Public Function UnHookAll() As HOOKERRORS

'clear all the remaining active hooks
Dim iCount As Integer
Dim lRet As Long
Dim heErr As HOOKERRORS

On Error GoTo UnHookAll_Error

LogEvent fcUnHookAll, "Entering function"

For iCount = 0 To 11
    With gStatusOfHooks(iCount)
        If .HS_ACTIVE = True Then
        'since the first version is for application (or thread specific) hooks, then just remove it
'            If .HS_IDENTIFIER = HI_APPLICATION Then
'                LogEvent fcUnHookAll, " Found active " & ResolveScopeIdentifier(.HS_IDENTIFIER) & " " & ResolveHookType(.HS_TYPE) & " hook - attempting to unhook..."
'                lRet = UnhookWindowsHookEx(.HS_HOOKID)
'            Else
'                lRet = UnHookGlobal(.HS_TYPE)
'            End If
            lRet = UnhookWindowsHookEx(.HS_HOOKID)
            LogEvent fcUnHookAll, "Found active " & ResolveHookType(.HS_TYPE)
            If lRet = 0 Then
                LogEvent fcUnHookAll, "     ...failed!", bcLogEventTypeError
                UnHookAll = HOOK_UNHOOK_ALL_FAIL
                Exit Function
            End If
            LogEvent fcUnHookAll, "...released hook " & ResolveHookType(.HS_TYPE)
        End If
    End With
Next iCount

'just to make sure
LogEvent fcUnHookAll, " Clearing hook statuses..."
heErr = ClearHookStatus
If heErr <> ERROR_SUCCESS Then
    LogEvent fcUnHookAll, "    **Error occured!", bcLogEventTypeError
    LogEvent fcUnHookAll, "Exiting with errors!"
    UnHookAll = heErr
    Exit Function
End If

LogEvent fcUnHookAll, "Exiting normally"

UnHookAll = ERROR_SUCCESS

Exit Function


UnHookAll_Error:
    UnHookAll = HOOK_UNKNOWNERROR
    LogEvent fcUnHookAll, "**An unexpected error occured in UnHookAll - system reports the following details:", bcLogEventTypeError
    LogEvent fcUnHookAll, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError

End Function

Private Function RegisterNewMessage() As HOOKERRORS

Dim iRet As Integer

On Error GoTo RegisterNewMessage_Error

LogEvent fcRegisterNewMessage, "Entering function"
LogEvent fcRegisterNewMessage, "    Attempting to register new message..."
iRet = RegisterNewMessageGlobal
If iRet = 0 Then
    LogEvent fcRegisterNewMessage, "    ...failed to register new message!!", bcLogEventTypeError
    RegisterNewMessage = ERROR_REGISTERING_NEW_MESSAGE
    LogEvent fcRegisterNewMessage, "Exiting with errors!"
    Exit Function
End If

giNewMessage = iRet

If giNewMessage < 0 Then giNewMessage = giNewMessage + 65536

LogEvent fcRegisterNewMessage, "    ...successful, new message id: " & giNewMessage
RegisterNewMessage = ERROR_SUCCESS

LogEvent fcRegisterNewMessage, "Exiting normally"

Exit Function

RegisterNewMessage_Error:
    RegisterNewMessage = HOOK_UNKNOWNERROR
    LogEvent fcRegisterNewMessage, "**An unexpected error has occured in RegisterNewMessage - system reports the following details:", bcLogEventTypeError
    LogEvent fcRegisterNewMessage, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Private Sub Class_Initialize()

Dim heRet As HOOKERRORS
Dim sMsg As String
Dim lRet As Long

gfFile = FreeFile()

Open App.Path & "\HookEvt.log" For Output As #gfFile

On Error GoTo Initialise_Error

Load FHookSubClass

LogEvent fcClassInitialize, "  Attempting to sub-class hidden window..."
heRet = SubClassWnd
If heRet <> ERROR_SUCCESS Then
    ResolveError heRet, sMsg
    LogEvent fcClassInitialize, "** Error occured in CHooks.Initialize - '" & sMsg & "'!!"
    gbSUBCLASSED = False
    Exit Sub
Else
    gbSUBCLASSED = True
End If

LogEvent fcClassInitialize, "...sub-classed successfully.  Old window proc:" & Hex(gOldWndProc)

Set goHook = Me

LogEvent fcClassInitialize, "  Initializing DLL..."
lRet = InitDLL(FHookSubClass.hWnd)
If lRet = 1 Then
    LogEvent fcClassInitialize, "Error occured initialising DLL: 'Error occured opening DLL.log'", bcLogEventTypeError
End If
LogEvent fcClassInitialize, "  ...finished initialising!"
LogEvent fcClassInitialize, "Exiting CHooks.Initialize normally"

Exit Sub

Initialise_Error:
    'nothing
    
End Sub

Private Function SubClassWnd() As HOOKERRORS

On Error GoTo SubClassWnd_Error

LogEvent fcSubClassWnd, "Entering function"

LogEvent fcSubClassWnd, "   Attempting to subclass hWnd: " & Hex(FHookSubClass.hWnd) & "..."
gOldWndProc = SetWindowLong(FHookSubClass.hWnd, GWL_WNDPROC, AddressOf WindowProc)
If gOldWndProc = 0 Then
    LogEvent fcSubClassWnd, "   **Error occured!", bcLogEventTypeError
    SubClassWnd = ERROR_SUBCLASSING_WINDOW
Else
    LogEvent fcSubClassWnd, "   ...successful. Old window proc address: " & Hex(gOldWndProc)
    SubClassWnd = ERROR_SUCCESS
End If

LogEvent fcSubClassWnd, "Exiting normally"

Exit Function

SubClassWnd_Error:
    SubClassWnd = HOOK_UNKNOWNERROR
    LogEvent fcSubClassWnd, "**An unexpected error has occured in SubClassWnd - the system reports the following details:", bcLogEventTypeError
    LogEvent fcSubClassWnd, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Function RestoreWndProc() As HOOKERRORS

On Error GoTo RestoreWndProc_Error

LogEvent fcRestoreWndProc, "Entering function"

If gOldWndProc <> 0 Then
    LogEvent fcRestoreWndProc, "    Attempting to restore proc address of " & gOldWndProc & "..."
    gOldWndProc = SetWindowLong(FHookSubClass.hWnd, GWL_WNDPROC, gOldWndProc)
    If gOldWndProc = 0 Then
        LogEvent fcRestoreWndProc, "    **Error occured restoring window proc!", bcLogEventTypeError
        RestoreWndProc = ERROR_RELEASING_SUBCLASS
    Else
        LogEvent fcRestoreWndProc, "    ...successful!"
        RestoreWndProc = ERROR_SUCCESS
    End If
End If

LogEvent fcRestoreWndProc, "Exiting normally"

Exit Function

RestoreWndProc_Error:
    RestoreWndProc = HOOK_UNKNOWNERROR
    LogEvent fcRestoreWndProc, "**An unexpected error has occured in RestoreWndProc - the system reports the following details:", bcLogEventTypeError
    LogEvent fcRestoreWndProc, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description, bcLogEventTypeError
    
End Function

Public Sub Init()

Dim heErr As HOOKERRORS
Dim lErrNum As Long
Dim sErrDesc As String

On Error GoTo Init_Error

LogEvent fcInit, "Entered CHooks.Init"

'no point going any further if sub-classing of FHookSubClass failed, raise error
If Not gbSUBCLASSED Then
    LogEvent "**Error occured in CHooks.Init - sub-class isn't set!", bcLogEventTypeError
    Err.Raise ResolveError(ERROR_SUBCLASSING_WINDOW, sErrDesc), "CHooks.Init", sErrDesc
End If

'initialise statuses
heErr = InitialiseHookStatus
If heErr <> ERROR_SUCCESS Then
    LogEvent fcInit, "**Error occured in CHooks.Init - unable to initialise hook status!", bcLogEventTypeError
    Call UnHookAll
    Call RestoreWndProc
    Err.Raise ResolveError(heErr, sErrDesc), "CHooks.Init", sErrDesc
End If

'load the DLL into this address space
heErr = LoadHookDLL
If heErr <> ERROR_SUCCESS Then
    LogEvent fcInit, "**Error occured in CHooks.Init - unable to load hook DLL!", bcLogEventTypeError
    Call UnHookAll
    Call RestoreWndProc
    Err.Raise ResolveError(heErr, sErrDesc), "CHooks.Init", sErrDesc
End If

'register necessary message for global hook communication
heErr = RegisterNewMessage
If heErr <> ERROR_SUCCESS Then
    LogEvent fcInit, "**Error occured in CHooks.Init - unable to register WM_BATTMESSAGE!", bcLogEventTypeError
    Call UnHookAll
    Call RestoreWndProc
    Call FreeLibraryDLL
    Err.Raise ResolveError(heErr, sErrDesc), "CHooks.Init", sErrDesc
End If

LogEvent fcInit, "Exiting CHooks.Init normally"

Exit Sub

Init_Error:
    LogEvent fcInit, "**An unexpected error occured in CHooks.Init**", bcLogEventTypeError
    LogEvent fcInit, "No: " & Err.Number & "  Src: " & Err.Source & "  Desc: " & Err.Description
    Call UnHookAll
    Call RestoreWndProc
    Call FreeLibraryDLL
    With Err
        .Raise .Number, .Source, .Description
    End With
    
End Sub

Public Function ResolveError(ByVal HookError As HOOKERRORS, ByRef sDesc As String) As Long

Select Case HookError

Case ERROR_LOADING_LIBRARY
    sDesc = "An error occured attempting to load library 'WinHook.dll'. Ensure this file is in the application folder."
    
Case ERROR_FREEING_LIBRARY
    sDesc = "An error occured attempting to free the module 'WinHook.dll'."
    
Case ERROR_REGISTERING_NEW_MESSAGE
    sDesc = "An error occured registering new message for global hook!"
    
Case ERROR_RELEASING_SUBCLASS
    sDesc = "An error occured whilst attempting to restore the hook sub-class procedure address!"
    
Case ERROR_SUBCLASSING_WINDOW
    sDesc = "An error occured whilst attempting to sub-class the internal window!"
    
Case HOOK_ACTIVE
    sDesc = "The hook is currently active!"
    
Case HOOK_GLOBALONLY
    sDesc = "This hook is specified as a global hook only!"
    
Case HOOK_CANT_SET_GLOBALONLY
    sDesc = "You cannot set this hook to be Global Only!"
    
Case HOOK_UNKNOWNERROR
    sDesc = "An unknown or unexpected error has occured!"
    
Case HOOK_PROCADDRESS_CALLWNDPROC_FAIL
    sDesc = "An error occured retrieving the procedure address for CALLWNDPROC!"
    
Case HOOK_PROCADDRESS_CBTPROC_FAIL
    sDesc = "An error occured retrieving the procedure address for CBTPROC!"
    
Case HOOK_PROCADDRESS_DEBUG_FAIL
    sDesc = "An error occured retrieving the procedure address for DEBUGPROC!"
    
Case HOOK_PROCADDRESS_FOREGROUNDIDLE_FAIL
    sDesc = "An error occured retrieving the procedure address for FOREGROUNDIDLEPROC!"
    
Case HOOK_PROCADDRESS_GETMESSAGE_FAIL
    sDesc = "An error occured retrieving the procedure address for GETMESSAGEPROC!"
    
Case HOOK_PROCADDRESS_JOURNALPLAYBACK_FAIL
    sDesc = "An error occured retrieving the procedure address for JOURNALPLAYBACKPROC!"
    
Case HOOK_PROCADDRESS_JOURNALRECORD_FAIL
    sDesc = "An error occured retrieving the procedure address for JOURNALRECORDPROC!"
    
Case HOOK_PROCADDRESS_KEYBOARD_FAIL
    sDesc = "An error occured retrieving the procedure address for KEYBOARDPROC!"
    
Case HOOK_PROCADDRESS_MOUSE_FAIL
    sDesc = "An error occured retrieving the procedure address for MOUSEPROC!"
    
Case HOOK_PROCADDRESS_MSGFILTER_FAIL
    sDesc = "An error occured retrieving the procedure address for MSGFILTERPROC!"
    
Case HOOK_PROCADDRESS_SHELL_FAIL
    sDesc = "An error occured retrieving the procedure address for SHELLPROC!"
    
Case HOOK_PROCADDRESS_SYSMSGFILTER_FAIL
    sDesc = "An error occured retrieving the procedure address for SYSMSGFILTERPROC!"
    
Case HOOK_SETHOOK_CALLWNDPROC_FAIL
    sDesc = "An error occured setting the hook filter for WH_CALLWND!"
    
Case HOOK_SETHOOK_CBTPROC_FAIL
    sDesc = "An error occured setting the hook filter for WH_CBT!"
    
Case HOOK_SETHOOK_DEBUG_FAIL
    sDesc = "An error occured setting the hook filter for WH_DEBUG!"
    
Case HOOK_SETHOOK_FOREGROUNDIDLE_FAIL
    sDesc = "An error occured setting the hook filter for WH_FOREGROUNDIDLE!"
    
Case HOOK_SETHOOK_GETMESSAGE_FAIL
    sDesc = "An error occured setting the hook filter for WH_GETMESSAGE!"
    
Case HOOK_SETHOOK_JOURNALPLAYBACK_FAIL
    sDesc = "An error occured setting the hook filter for WH_JOURNALPLAYBACK!"
    
Case HOOK_SETHOOK_JOURNALRECORD_FAIL
    sDesc = "An error occured setting the hook filter for WH_JOURNALRECORD!"
    
Case HOOK_SETHOOK_KEYBOARD_FAIL
    sDesc = "An error occured setting the hook filter for WH_KEYBOARD!"
    
Case HOOK_SETHOOK_MOUSE_FAIL
    sDesc = "An error occured setting the hook filter for WH_MOUSE!"
    
Case HOOK_SETHOOK_MSGFILTER_FAIL
    sDesc = "An error occured setting the hook filter for WH_MSGFILTER!"
    
Case HOOK_SETHOOK_SHELL_FAIL
    sDesc = "An error occured setting the hook filter for WH_SHELL!"
    
Case HOOK_SETHOOK_SYSMSGFILTER_FAIL
    sDesc = "An error occured setting the hook filter for WH_SYSMSGFILTER!"
    
Case HOOK_UNHOOK_ALL_FAIL
    sDesc = "An error occured during call to UnHookAll!"
    
Case HOOK_UNHOOK_CALLWNDPROC_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_CALLWND!"
    
Case HOOK_UNHOOK_CBTPROC_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_CBT!"
    
Case HOOK_UNHOOK_DEBUG_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_DEBUG!"
    
Case HOOK_UNHOOK_FOREGROUNDIDLE_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_FOREGROUNDIDLE!"
    
Case HOOK_UNHOOK_GETMESSAGE_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_GETMESSAGE!"
    
Case HOOK_UNHOOK_JOURNALPLAYBACK_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_JOURNALPLAYBACK!"
    
Case HOOK_UNHOOK_JOURNALRECORD_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_JOURNALRECORD!"
    
Case HOOK_UNHOOK_KEYBOARD_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_KEYBOARD!"
    
Case HOOK_UNHOOK_MOUSE_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_MOUSE!"
    
Case HOOK_UNHOOK_MSGFILTER_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_MSGFILTER!"
    
Case HOOK_UNHOOK_SHELL_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_SHELL!"
    
Case HOOK_UNHOOK_SYSMSGFILTER_FAIL
    sDesc = "An error occured unhooking the hook filter for WH_SYSMSGFILTER!"
    
End Select

ResolveError = HookError

End Function

Friend Function SignalEvent(ByVal HookType As HookTypes, ByVal nCode As Integer, ByVal wParam As Long, ByVal lParam As Variant, ByRef ReturnVal As Long) As HOOKERRORS

LogEvent fcSignalEvent, "Entering function"


LogEvent fcSignalEvent, "   Raising event for " & ResolveHookType(HookType)
LogEvent fcSignalEvent, "   Params-  nCode: " & nCode & "  wParam: " & wParam & "  lParam: " & lParam

Select Case HookType

Case WH_CALLWNDPROC
    RaiseEvent CallWndProcEvent(nCode, wParam, lParam, ReturnVal)

Case WH_CBT
    RaiseEvent CBTEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_DEBUG
    RaiseEvent DebugEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_FOREGROUNDIDLE
    RaiseEvent ForeGroundIdleEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_GETMESSAGE
    RaiseEvent GetMessageEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_JOURNALPLAYBACK
    RaiseEvent JournalPlayBackEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_JOURNALRECORD
    RaiseEvent JournalRecordEvent(nCode, wParam, lParam, ReturnVal)

Case WH_KEYBOARD
    RaiseEvent KeyBoardEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_MOUSE
    RaiseEvent MouseEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_MSGFILTER
    RaiseEvent MsgFilterEvent(nCode, wParam, lParam, ReturnVal)
    
Case WH_SHELL
    RaiseEvent ShellEvent(nCode, wParam, lParam, ReturnVal)
    
End Select

LogEvent fcSignalEvent, "Exited normally"

End Function

Private Function ResolveHookType(ByVal HookType As HookTypes) As String

Select Case HookType

Case WH_CALLWNDPROC
    ResolveHookType = "WH_CALLWNDPROC"
    
Case WH_CBT
    ResolveHookType = "WH_CBT"
    
Case WH_DEBUG
    ResolveHookType = "WH_DEBUG"
    
Case WH_FOREGROUNDIDLE
    ResolveHookType = "WH_FOREGROUNDIDLE"
    
Case WH_JOURNALPLAYBACK
    ResolveHookType = "WH_JOURNALPLAYBACK"
    
Case WH_JOURNALRECORD
    ResolveHookType = "WH_JOURNALRECORD"

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
    
Case WH_SYSMSGFILTER
    ResolveHookType = "WH_SYSMSGFILTER"

End Select


End Function


Private Function ResolveScopeIdentifier(ByVal HookScope As HookIdentifier) As String

ResolveScopeIdentifier = ""

Select Case HookScope

Case HI_APPLICATION
    ResolveScopeIdentifier = "APPLICATION"
    
Case HI_GLOBAL
    ResolveScopeIdentifier = "GLOBAL"
    
Case HI_GLOBALONLY
    ResolveScopeIdentifier = "GLOBALONLY"
    
End Select

End Function

Public Function CloseUp() As HOOKERRORS

Dim heRet As HOOKERRORS
Dim sMsg As String

On Error Resume Next

CloseUp = ERROR_SUCCESS

LogEvent fcCloseUp, "Entering sub"

LogEvent fcCloseUp, "    Attempting to unhook all active hooks..."
heRet = UnHookAll
If heRet <> ERROR_SUCCESS Then
    ResolveError heRet, sMsg
    LogEvent fcCloseUp, "    ...error return from unhookall: " & sMsg, bcLogEventTypeError
    CloseUp = HOOK_UNHOOK_ALL_FAIL
Else
    LogEvent fcCloseUp, "    ...successful!"
End If

LogEvent fcCloseUp, "    Attempting to restore old window procedure..."
heRet = RestoreWndProc
If heRet <> ERROR_SUCCESS Then
    ResolveError heRet, sMsg
    LogEvent fcCloseUp, "    ...error return from RestoreWndProc: " & sMsg, bcLogEventTypeError
    CloseUp = ERROR_RELEASING_SUBCLASS
Else
    LogEvent fcCloseUp, "    ...successful!"
End If

LogEvent fcCloseUp, "    Closing down DLL..."
CloseDownDLL
LogEvent fcCloseUp, "    ...successful!"

LogEvent fcCloseUp, "    Attempting to free DLL..."
heRet = FreeLibraryDLL
If heRet <> ERROR_SUCCESS Then
    ResolveError heRet, sMsg
    LogEvent fcCloseUp, "    ...error return from FreeLibraryDLL: " & sMsg, bcLogEventTypeError
    CloseUp = ERROR_FREEING_LIBRARY
Else
    LogEvent fcCloseUp, "    ...successful!"
End If

LogEvent fcCloseUp, "    Closing sub-class hidden form..."
Unload FHookSubClass
LogEvent fcCloseUp, "    ...successful!"

LogEvent fcCloseUp, "    Setting global hook class reference to nothing..."
Set goHook = Nothing
LogEvent fcCloseUp, "    ...successful!"

LogEvent fcCloseUp, "    Closing log file now!!!"
Close #gfFile


End Function

Public Property Get CWPData() As CWPSTRUCT
    
    CWPData = mCWPData
    
End Property

Public Property Let CWPData(NewCWPStruct As CWPSTRUCT)
    
    mCWPData = NewCWPStruct
    
End Property

Public Property Get CBTActivateData() As CBTACTIVATESTRUCT

    CBTActivateData = mCBTActivateData
    
End Property

Public Property Let CBTActivateData(NewActivateStruct As CBTACTIVATESTRUCT)

    mCBTActivateData = NewActivateStruct
    
End Property

Public Property Get CBTMouseData() As MOUSEHOOKSTRUCT

    CBTMouseData = mCBTMouseData
    
End Property

Public Property Let CBTMouseData(NewMouseData As MOUSEHOOKSTRUCT)

    mCBTMouseData = NewMouseData
    
End Property
Public Property Get CBT_CreateData() As CBTCREATEDATA

    CBT_CreateData = mCBTCreateData
    
End Property

Public Property Let CBT_CreateData(NewCreateData As CBTCREATEDATA)

    mCBTCreateData = NewCreateData
    
End Property
Public Property Get CBTMinMaxData() As CBTSizeWndCodes

    CBTMinMaxData = mCBTMinMaxData
    
End Property

Public Property Let CBTMinMaxData(NewMinMaxData As CBTSizeWndCodes)

    mCBTMinMaxData = NewMinMaxData
    
End Property
Public Property Get CBTMoveSizeData() As RECT

    CBTMoveSizeData = mCBTMoveSizeData
    
End Property

Public Property Let CBTMoveSizeData(NewMoveSizeData As RECT)

    mCBTMoveSizeData = NewMoveSizeData
    
End Property

Public Property Get CBTSysCommandData() As CBTSYSCOMMAND_DATA

    CBTSysCommandData = mCBTSysCommandData
    
End Property

Public Property Let CBTSysCommandData(NewSysCommandData As CBTSYSCOMMAND_DATA)

    mCBTSysCommandData = NewSysCommandData
    
End Property
Public Property Get CBTKeySkipped() As KeyStrokeInfo

    CBTKeySkipped = mCBTKeySkipped
    
End Property

Public Property Let CBTKeySkipped(NewKeySkippedData As KeyStrokeInfo)

    mCBTKeySkipped = NewKeySkippedData
    
End Property
Public Property Get DebugData() As DEBUGHOOKINFO

    DebugData = mDebugData
    
End Property

Public Property Let DebugData(NewDebugData As DEBUGHOOKINFO)

    mDebugData = NewDebugData
    
End Property
Public Property Get KeyboardData() As KeyStrokeInfo

    KeyboardData = mKeyboardData
    
End Property

Public Property Let KeyboardData(NewKeyboardData As KeyStrokeInfo)

    mKeyboardData = NewKeyboardData
    
End Property
Public Property Get MessageProcData() As MessageData

    MessageProcData = mMessageProcData
    
End Property

Public Property Let MessageProcData(NewMessageProcData As MessageData)

    mMessageProcData = NewMessageProcData
    
End Property
Public Property Get MouseData() As MouseProcData

    MouseData = mMouseData
    
End Property

Public Property Let MouseData(NewMouseData As MouseProcData)

    mMouseData = NewMouseData
    
End Property
Public Property Get JournalEventData() As EVENTMSG

    JournalEventData = mJournalEventData
    
End Property

Public Property Let JournalEventData(NewJournalEventData As EVENTMSG)

    mJournalEventData = NewJournalEventData
    
End Property

Public Property Get Get_MsgData() As GETMSGDATA

    Get_MsgData = mGetMsgData

End Property

Public Property Let Get_MsgData(NewGetMsgData As GETMSGDATA)

    mGetMsgData = NewGetMsgData
    
End Property

Public Property Get ShellGetMinRectData() As RECT

    ShellGetMinRectData = mShellGetMinRectData
    
End Property

Public Property Let ShellGetMinRectData(NewGetMinRect As RECT)

    mShellGetMinRectData = NewGetMinRect
    
End Property
