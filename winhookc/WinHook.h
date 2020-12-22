
#include "Windows.h"
#include <stdio.h>

HRESULT InitDLL (HWND hCaller);
void CloseDownDLL (void);
LRESULT CALLBACK CallWndProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK CBTProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK DebugProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ForeGroundIdleProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK GetMsgProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK JournalPlaybackProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK JournalRecordProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK KeyboardProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MessageProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MouseProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ShellProc (int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK SysMsgProc (int nCode, WPARAM wParam, LPARAM lParam);

UINT RegisterNewMessageGlobal (void);

HHOOK SetUpHookGlobal (long hHookType);

BOOL UnHookGlobal (int hHookType);

struct HOOKDATA
{
	HHOOK hGlobalHooks;
	int iHookType;
	HOOKPROC lpFuncPtr;
};

HRESULT InitHookArray (void);
