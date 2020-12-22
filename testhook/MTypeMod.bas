Attribute VB_Name = "MTypeMod"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public glHWndTest As Long

Public gbDLL_LOADED As Boolean


Public Function HIWORD(ByVal lValue As Long) As Long

'retrieves the high-order word of a value
HIWORD = lValue And &HFFFF0000

End Function

Public Function LOWORD(ByVal lValue As Long) As Long

'retrieves the low-order word of a value
LOWORD = lValue And &HFFFF

End Function

Public Function ResolveSW(ByVal SWParam As Long) As String

Dim eParam As CBTSizeWndCodes

eParam = LOWORD(SWParam)

Select Case eParam

Case SW_ERASE
    ResolveSW = "ERASE"
    
Case SW_HIDE
    ResolveSW = "HIDE"
    
Case SW_INVALIDATE
    ResolveSW = "INVALIDATE"
    
Case SW_MAX
    ResolveSW = "MAX"
    
Case SW_MAXIMIZE
    ResolveSW = "MAXIMIZE"
    
Case SW_MINIMIZE
    ResolveSW = "MINIMIZE"
    
Case SW_NORMAL
    ResolveSW = "NORMAL"
    
Case SW_OTHERUNZOOM
    ResolveSW = "OTHERUNZOOM"
    
Case SW_OTHERZOOM
    ResolveSW = "OTHERZOOM"
    
Case SW_PARENTOPENING
    ResolveSW = "PARENTOPENING"
    
Case SW_PARENTCLOSING
    ResolveSW = "PARENTCLOSING"
    
Case SW_RESTORE
    ResolveSW = "RESTORE"
    
Case SW_CLOSECHILDREN
    ResolveSW = "CLOSECHILDREN"
    
Case SW_SHOW
    ResolveSW = "SHOW"
    
Case SW_SHOWDEFAULT
    ResolveSW = "SHOWDEFAULT"
    
Case SW_SHOWMAXIMIZED
    ResolveSW = "SHOWMAXIMIZED"
    
Case SW_SHOWMINIMIZED
    ResolveSW = "SHOWMINIMIZED"
    
Case SW_SHOWMINNOACTIVE
    ResolveSW = "SHOWMINNOACTIVATE"
    
Case SW_SHOWNA
    ResolveSW = "SHOWNA"
    
Case SW_SHOWNOACTIVATE
    ResolveSW = "SHOWNOACTIVATE"
    
Case SW_SHOWNORMAL
    ResolveSW = "SHOWNORMAL"
    
End Select

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


Public Function ResolveHookType(ByVal HookType As HookTypes) As String

Select Case HookType

Case WH_CALLWNDPROC
    ResolveHookType = "CALLWNDPROC"
    
Case WH_CBT
    ResolveHookType = "CBT"
    
Case WH_DEBUG
    ResolveHookType = "DEBUG"
    
Case WH_FOREGROUNDIDLE
    ResolveHookType = "FOREGROUNDIDLE"
    
Case WH_GETMESSAGE
    ResolveHookType = "GETMESSAGE"
    
Case WH_JOURNALPLAYBACK
    ResolveHookType = "JOURNALPLAYBACK"
    
Case WH_JOURNALRECORD
    ResolveHookType = "JOURNALRECORD"
    
Case WH_KEYBOARD
    ResolveHookType = "KEYBOARD"
    
Case WH_MOUSE
    ResolveHookType = "MOUSE"
    
Case WH_MSGFILTER
    ResolveHookType = "MSGFILTER"
    
Case WH_SHELL
    ResolveHookType = "SHELL"
    
Case WH_SYSMSGFILTER
    ResolveHookType = "SYSMSGFILTER"
    
End Select

End Function

Public Function ResolveFilterCode(ByVal FilterCode As MsgFilterCodes) As String

Select Case FilterCode

Case MSGF_DDEMGR
    ResolveFilterCode = "DDEMGR"
    
Case MSGF_DIALOGBOX
    ResolveFilterCode = "DIALOGBOX"

Case MSGF_MAINLOOP
    ResolveFilterCode = "MAINLOOP"

Case MSGF_MAX
    ResolveFilterCode = "MAX"

Case MSGF_MENU
    ResolveFilterCode = "MENU"

Case MSGF_MESSAGEBOX
    ResolveFilterCode = "MESSAGEBOX"

Case MSGF_MOVE
    ResolveFilterCode = "MOVE"

Case MSGF_NEXTWINDOW
    ResolveFilterCode = "NEXTWINDOW"

Case MSGF_SCROLLBAR
    ResolveFilterCode = "SCROLLBAR"

Case MSGF_SIZE
    ResolveFilterCode = "SIZE"

Case MSGF_USER
    ResolveFilterCode = "USER"

End Select

End Function



