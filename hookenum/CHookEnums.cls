VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHookEnums"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'log codes
Public Enum LogCodes
    bcLogEventTypeError = 1
    bcLogEventTypeWarning = 2
    bcLogEventTypeInformation = 4
End Enum

'enum for error logging
Public Enum FuncCodes
    fcClearHookStatus = 0
    fcFreeLibraryDLL
    fcInit
    fcInitializeHookStatus
    fcLoadHookDLL
    fcRegisterNewMessage
    fcReleaseHook
    fcResolveError
    fcResolveHookType
    fcResolveScopeIdentifier
    fcRestoreWndProc
    fcSetHookType
    fcSetUpHook
    fcSignalEvent
    fcSubClassWnd
    fcUnHookAll
    fcCallWndProc
    fcCBTProc
    fcDebugProc
    fcDecodeKeyInfo
    fcForeGroundIdleProc
    fcGetMsgProc
    fcHIWORD
    fcKeyBoardProc
    fcLOWORD
    fcMessageProc
    fcMouseProc
    fcShellProc
    fcWindowProc
    fcClassInitialize
    fcCloseUp
    fcJournalPlayBackProc
    fcJournalRecordProc
End Enum

'general hook codes
Public Enum HookCodes
    HC_ACTION = 0
    HC_ACTIONREQUIRED = HC_ACTION
    HC_GETNEXT = 1
    HC_NOREMOVE = 3
    HC_NOREM = HC_NOREMOVE
    HC_SKIP = 2
    HC_SYSMODALOFF = 5
    HC_SYSMODALON = 4
End Enum

'Hook types
Public Enum HookTypes
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_DEBUG = 9
    WH_FOREGROUNDIDLE = 11
    WH_GETMESSAGE = 3
    WH_JOURNALPLAYBACK = 1
    WH_JOURNALRECORD = 0
    WH_KEYBOARD = 2
    WH_MAX = 11
    WH_MIN = (-1)
    WH_MOUSE = 7
    WH_MSGFILTER = (-1)
    WH_SHELL = 10
    WH_SYSMSGFILTER = 6
End Enum

'WH_CBT Hook codes
Public Enum CBTHookCodes
    HCBT_ACTIVATE = 5
    HCBT_CREATEWND = 3
    HCBT_DESTROYWND = 4
    HCBT_MINMAX = 1
    HCBT_MOVESIZE = 0
    HCBT_ACTIONREQUIRED = HCBT_MOVESIZE
    HCBT_SYSCOMMAND = 8
    HCBT_CLICKSKIPPED = 6
    HCBT_KEYSKIPPED = 7
    HCBT_SETFOCUS = 9
    HCBT_QS = 2
End Enum
    
'WH_CBT min max codes
Public Enum CBTSizeWndCodes
    SW_ERASE = &H4
    SW_HIDE = 0
    SW_INVALIDATE = &H2
    SW_MAX = 10
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
    SW_OTHERUNZOOM = 4
    SW_OTHERZOOM = 2
    SW_PARENTOPENING = 1
    SW_PARENTCLOSING = 3
    SW_RESTORE = 9
    SW_CLOSECHILDREN = &H1
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWNORMAL = 1
End Enum

'WH_CBT system commands
Public Enum CBTSysCommands
    SC_ARRANGE = &HF110&
    SC_CLOSE = &HF060&
    SC_GROUP_IDENTIFIER = &H2B
    SC_HOTKEY = &HF150&
    SC_HSCROLL = &HF080&
    SC_KEYMENU = &HF100&
    SC_MANAGER_CONNECT = &H1
    SC_MANAGER_CREATE_SERVICE = &H2
    SC_MANAGER_ENUMERATE_SERVICE = &H4
    SC_MANAGER_LOCK = &H8
    SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
    SC_MANAGER_QUERY_LOCK_STATUS = &H10
    SC_MAXIMIZE = &HF030&
    SC_MINIMIZE = &HF020&
    SC_ICON = SC_MINIMIZE
    SC_MOUSEMENU = &HF090&
    SC_MOVE = &HF010&
    SC_NEXTWINDOW = &HF040&
    SC_PREVWINDOW = &HF050&
    SC_RESTORE = &HF120&
    SC_SCREENSAVE = &HF140&
    SC_SIZE = &HF000&
    SC_TASKLIST = &HF130&
    SC_VSCROLL = &HF070&
    SC_ZOOM = SC_MAXIMIZE
    STANDARD_RIGHTS_REQUIRED = &HF0000
    SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
End Enum

'Message filter codes
Public Enum MsgFilterCodes
    MSGF_DDEMGR = &H8001
    MSGF_DIALOGBOX = 0
    MSGF_ACTIONREQUIRED = MSGF_DIALOGBOX
    MSGF_MAINLOOP = 8
    MSGF_MAX = 8
    MSGF_MENU = 2
    MSGF_MESSAGEBOX = 1
    MSGF_MOVE = 3
    MSGF_NEXTWINDOW = 6
    MSGF_SCROLLBAR = 5
    MSGF_SIZE = 4
    MSGF_USER = 4096
End Enum

'Shell hook codes
Public Enum ShellHookCodes
    HSHELL_ACTIVATESHELLWINDOW = 3
    HSHELL_WINDOWCREATED = 1
    HSHELL_WINDOWDESTROYED = 2
    HSHELL_ACTIONREQUIRED = 0
   'HSHELL_ACCESSIBILITYSTATE=11    WINDOWS 2000
   'HSHELL_APPCOMMAND=12            WINDOWS 2000
    HSHELL_WINDOWACTIVATED = 4
    HSHELL_GETMINRECT = 5
    HSHELL_REDRAW = 6
    HSHELL_TASKMAN = 7
    HSHELL_LANGUAGE = 8
End Enum

'Hook Status codes
Public Enum HookIdentifier
    HI_APPLICATION = 0
    HI_GLOBAL = 1
    HI_GLOBALONLY = 3
End Enum

'removal options for WH_GETMESSAGE
Public Enum RemovalOptions
    PM_NOREMOVE = &H0
    PM_NOYIELD = &H2
    PM_REMOVE = &H1
End Enum

'enums for keystate struct
Public Enum ExtendedKey
    EK_EXTENDED = &H1
    EK_NOTEXTENDED = &H0
End Enum

Public Enum ContextCode
    CC_ALTDOWN = &H1
    CC_ALTNOTDOWN = &H0
End Enum

Public Enum PrevKeyState
    PK_KEYDOWN = &H1
    PK_KEYUP = &H0
End Enum

Public Enum TransState
    TS_KEYPRESSED = &H0
    TS_KEYRELEASED = &H1
End Enum

'Windows messages
Public Enum WindowsMessages
    WM_ACTIVATE = &H6
    WM_ACTIVATEAPP = &H1C
    WM_ASKCBFORMATNAME = &H30C
    WM_CANCELJOURNAL = &H4B
    WM_CANCELMODE = &H1F
    WM_CHANGECBCHAIN = &H30D
    WM_CHAR = &H102
    WM_CHARTOITEM = &H2F
    WM_CHILDACTIVATE = &H22
    WM_USER = &H400
    WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
    WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
    WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
    WM_CLEAR = &H303
    WM_CLOSE = &H10
    WM_COMMAND = &H111
    WM_COMMNOTIFY = &H44
    WM_COMPACTING = &H41
    WM_COMPAREITEM = &H39
    WM_CONVERTREQUESTEX = &H108
    WM_COPY = &H301
    WM_COPYDATA = &H4A
    WM_CREATE = &H1
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    WM_CUT = &H300
    WM_DDE_FIRST = &H3E0
    WM_DDE_ACK = (WM_DDE_FIRST + 4)
    WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
    WM_DDE_DATA = (WM_DDE_FIRST + 5)
    WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
    WM_DDE_INITIATE = (WM_DDE_FIRST)
    WM_DDE_LAST = (WM_DDE_FIRST + 8)
    WM_DDE_POKE = (WM_DDE_FIRST + 7)
    WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
    WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
    WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
    WM_DEADCHAR = &H103
    WM_DELETEITEM = &H2D
    WM_DESTROY = &H2
    WM_DESTROYCLIPBOARD = &H307
    WM_DEVMODECHANGE = &H1B
    WM_DRAWCLIPBOARD = &H308
    WM_DRAWITEM = &H2B
    WM_DROPFILES = &H233
    WM_ENABLE = &HA
    WM_ENDSESSION = &H16
    WM_ENTERIDLE = &H121
    WM_ENTERMENULOOP = &H211
    WM_ERASEBKGND = &H14
    WM_EXITMENULOOP = &H212
    WM_FONTCHANGE = &H1D
    WM_GETDLGCODE = &H87
    WM_GETFONT = &H31
    WM_GETHOTKEY = &H33
    WM_GETMINMAXINFO = &H24
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_HOTKEY = &H312
    WM_HSCROLL = &H114
    WM_HSCROLLCLIPBOARD = &H30E
    WM_ICONERASEBKGND = &H27
    WM_IME_CHAR = &H286
    WM_IME_COMPOSITION = &H10F
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_CONTROL = &H283
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYLAST = &H10F
    WM_IME_KEYUP = &H291
    WM_IME_NOTIFY = &H282
    WM_IME_SELECT = &H285
    WM_IME_SETCONTEXT = &H281
    WM_IME_STARTCOMPOSITION = &H10D
    WM_INITDIALOG = &H110
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_KEYDOWN = &H100
    WM_KEYFIRST = &H100
    WM_KEYLAST = &H108
    WM_KEYUP = &H101
    WM_KILLFOCUS = &H8
    WM_LBUTTONDBLCLK = &H203
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_MBUTTONDBLCLK = &H209
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MDIACTIVATE = &H222
    WM_MDICASCADE = &H227
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIGETACTIVE = &H229
    WM_MDIICONARRANGE = &H228
    WM_MDIMAXIMIZE = &H225
    WM_MDINEXT = &H224
    WM_MDIREFRESHMENU = &H234
    WM_MDIRESTORE = &H223
    WM_MDISETMENU = &H230
    WM_MDITILE = &H226
    WM_MEASUREITEM = &H2C
    WM_MENUCHAR = &H120
    WM_MENUSELECT = &H11F
    WM_MOUSEACTIVATE = &H21
    WM_MOUSEFIRST = &H200
    WM_MOUSELAST = &H209
    WM_MOUSEMOVE = &H200
    WM_MOVE = &H3
    WM_NCACTIVATE = &H86
    WM_NCCALCSIZE = &H83
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCHITTEST = &H84
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCMBUTTONDBLCLK = &HA9
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMOUSEMOVE = &HA0
    WM_NCPAINT = &H85
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NEXTDLGCTL = &H28
    WM_NULL = &H0
    WM_OTHERWINDOWCREATED = &H42
    WM_OTHERWINDOWDESTROYED = &H43
    WM_PAINT = &HF
    WM_PAINTCLIPBOARD = &H309
    WM_PAINTICON = &H26
    WM_PALETTECHANGED = &H311
    WM_PALETTEISCHANGING = &H310
    WM_PARENTNOTIFY = &H210
    WM_PASTE = &H302
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_POWER = &H48
    WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
    WM_PSD_FULLPAGERECT = (WM_USER + 1)
    WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
    WM_PSD_MARGINRECT = (WM_USER + 3)
    WM_PSD_MINMARGINRECT = (WM_USER + 2)
    WM_PSD_PAGESETUPDLG = (WM_USER)
    WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
    WM_QUERYDRAGICON = &H37
    WM_QUERYENDSESSION = &H11
    WM_QUERYNEWPALETTE = &H30F
    WM_QUERYOPEN = &H13
    WM_QUEUESYNC = &H23
    WM_QUIT = &H12
    WM_RBUTTONDBLCLK = &H206
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RENDERALLFORMATS = &H306
    WM_RENDERFORMAT = &H305
    WM_SETCURSOR = &H20
    WM_SETFOCUS = &H7
    WM_SETFONT = &H30
    WM_SETHOTKEY = &H32
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_SHOWWINDOW = &H18
    WM_SIZE = &H5
    WM_SIZECLIPBOARD = &H30B
    WM_SPOOLERSTATUS = &H2A
    WM_SYSCHAR = &H106
    WM_SYSCOLORCHANGE = &H15
    WM_SYSCOMMAND = &H112
    WM_SYSDEADCHAR = &H107
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_TIMECHANGE = &H1E
    WM_TIMER = &H113
    WM_UNDO = &H304
    WM_VKEYTOITEM = &H2E
    WM_VSCROLL = &H115
    WM_VSCROLLCLIPBOARD = &H30A
    WM_WINDOWPOSCHANGED = &H47
    WM_WINDOWPOSCHANGING = &H46
    WM_WININICHANGE = &H1A
End Enum


'WM_NCHITTEST values
Public Enum HitTestCodes
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTERROR = (-2)
    HTGROWBOX = 4
    HTHSCROLL = 6
    HTLEFT = 10
    HTMAXBUTTON = 9
    HTMENU = 5
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTREDUCE = HTMINBUTTON
    HTRIGHT = 11
    HTSIZE = HTGROWBOX
    HTSIZEFIRST = HTLEFT
    HTSIZELAST = HTBOTTOMRIGHT
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTTRANSPARENT = (-1)
    HTVSCROLL = 7
    HTZOOM = HTMAXBUTTON
End Enum

'Error codes for this module
Public Enum HOOKERRORS
    ERROR_SUCCESS = vbObjectError + 2000
    ERROR_LOADING_LIBRARY
    ERROR_FREEING_LIBRARY
    ERROR_REGISTERING_NEW_MESSAGE
    ERROR_RELEASING_SUBCLASS
    ERROR_SUBCLASSING_WINDOW
    HOOK_ACTIVE
    HOOK_GLOBALONLY
    HOOK_CANT_SET_GLOBALONLY
    HOOK_UNKNOWNERROR
    HOOK_PROCADDRESS_CALLWNDPROC_FAIL
    HOOK_PROCADDRESS_CBTPROC_FAIL
    HOOK_PROCADDRESS_DEBUG_FAIL
    HOOK_PROCADDRESS_FOREGROUNDIDLE_FAIL
    HOOK_PROCADDRESS_GETMESSAGE_FAIL
    HOOK_PROCADDRESS_JOURNALPLAYBACK_FAIL
    HOOK_PROCADDRESS_JOURNALRECORD_FAIL
    HOOK_PROCADDRESS_KEYBOARD_FAIL
    HOOK_PROCADDRESS_MOUSE_FAIL
    HOOK_PROCADDRESS_MSGFILTER_FAIL
    HOOK_PROCADDRESS_SHELL_FAIL
    HOOK_PROCADDRESS_SYSMSGFILTER_FAIL
    HOOK_SETHOOK_CALLWNDPROC_FAIL
    HOOK_SETHOOK_CBTPROC_FAIL
    HOOK_SETHOOK_DEBUG_FAIL
    HOOK_SETHOOK_FOREGROUNDIDLE_FAIL
    HOOK_SETHOOK_GETMESSAGE_FAIL
    HOOK_SETHOOK_JOURNALPLAYBACK_FAIL
    HOOK_SETHOOK_JOURNALRECORD_FAIL
    HOOK_SETHOOK_KEYBOARD_FAIL
    HOOK_SETHOOK_MOUSE_FAIL
    HOOK_SETHOOK_MSGFILTER_FAIL
    HOOK_SETHOOK_SHELL_FAIL
    HOOK_SETHOOK_SYSMSGFILTER_FAIL
    HOOK_UNHOOK_ALL_FAIL
    HOOK_UNHOOK_CALLWNDPROC_FAIL
    HOOK_UNHOOK_CBTPROC_FAIL
    HOOK_UNHOOK_DEBUG_FAIL
    HOOK_UNHOOK_FOREGROUNDIDLE_FAIL
    HOOK_UNHOOK_GETMESSAGE_FAIL
    HOOK_UNHOOK_JOURNALPLAYBACK_FAIL
    HOOK_UNHOOK_JOURNALRECORD_FAIL
    HOOK_UNHOOK_KEYBOARD_FAIL
    HOOK_UNHOOK_MOUSE_FAIL
    HOOK_UNHOOK_MSGFILTER_FAIL
    HOOK_UNHOOK_SHELL_FAIL
    HOOK_UNHOOK_SYSMSGFILTER_FAIL
End Enum


'WH_CALLWNDPROC struct
Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type

'keystroke info for WH_KEYBOARD
Public Type KeyStrokeInfo
    RepeatCount As Long
    ScanCode As Long
    ExtKey As Long
    Reserved As Integer
    ContCode As Long
    PreviousKeyState As Long
    TransitionState As Long
End Type

'WH_CBT Activate struct
Public Type CBTACTIVATESTRUCT
    fMouse As Long
    hWndActive As Long
End Type


'WH_CBT Create struct
Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

'WH_CBT CreateWnd struct
Public Type CBT_CREATEWND
    lpcs As CREATESTRUCT
    hWndInsertAfter As Long
End Type

'CBT create window data struct (my definition for property)
Public Type CBTCREATEDATA
    hWndInsertAfter As Long
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

'GetMsg data struct (my definition)
Public Type GETMSGDATA
    RemovalOption As Long
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

'message proc data struct (my definition)
Public Type MessageData
    MsgFilter As Long
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

    
'Rect struct for HCBT_MOVERESIZE
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'point struct for following struct
Public Type POINTAPI
    x As Long
    y As Long
End Type

'HCBT_CLICKSKIPPED mouse struct
Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

'mouse data (my definition)
Public Type MouseProcData
    mIdentifier As Long
    ptX As Long
    ptY As Long
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

'HCBT_SYSCOMMAND data structure (my definition)
Public Type CBTSYSCOMMAND_DATA
    scSysCommand As Long
    x As Long
    y As Long
End Type

'WH_DEBUG hook info struct
Public Type DEBUGHOOKINFO
    lThreadID As Long
    lThreadInstallerID As Long
    lParam As Long
    wParam As Long
    nCode As Long
End Type

'WH_GETMESSAGE struct
Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'WH_JOURNALRECORD - HC_ACTION Event message struct
Public Type EVENTMSG
    message As Long
    paramL As Long
    paramH As Long
    time As Long
    hWnd As Long
End Type

'COPY DATA STRUCTURE
Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Type HookStatus
    HS_TYPE As HookTypes
    HS_IDENTIFIER As HookIdentifier
    HS_ACTIVE As Boolean
    HS_HOOKID As Long
End Type

