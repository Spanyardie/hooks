// WinHook.cpp : Defines the entry point for the DLL application.
//

#include "WinHook.h"

#pragma data_seg(".SHARED")
HWND hCallingAppWnd = NULL;
UINT iNewMsg = 0;
#pragma data_seg()

HANDLE hInstDLL;
//FILE* fErrStream;

//array for hook handles
HOOKDATA hdHooks[11];

const long BM_CALLWNDPROC = 0;
const long BM_CBT = 1;
const long BM_DEBUG = 2;
const long BM_FOREGROUNDIDLE = 3;
const long BM_GETMESSAGE = 4;
const long BM_JOURNALPLAYBACK = 5;
const long BM_JOURNALRECORD = 6;
const long BM_KEYBOARD = 7;
const long BM_MOUSE = 8;
const long BM_MSGFILTER = 9;
const long BM_SHELL = 10;
const long BM_SYSMSGFILTER = 11;


HRESULT InitDLL (HWND hCaller)
{
	InitHookArray();

	//fErrStream = fopen( "bob.log", "w+" );


	//store the calling app hWnd for communication
	hCallingAppWnd = hCaller;

	return 0;
}

void CloseDownDLL (void)
{
	//fclose(fErrStream);
}

LRESULT CALLBACK CallWndProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_CALLWNDPROC].hGlobalHooks, nCode, wParam, lParam); 
}



LRESULT CALLBACK CBTProc (int nCode, WPARAM wParam, LPARAM lParam)
{

	CBTACTIVATESTRUCT* ActivateData=NULL;
	COPYDATASTRUCT DataToCopy;

	if (nCode < 0)  // do not process message, just pass it on 
        return CallNextHookEx(hdHooks[BM_CBT].hGlobalHooks, nCode, wParam, lParam); 
 
	switch(nCode)
	{	
		//is a window being activated?
		case HCBT_ACTIVATE:
			//wParam - hWnd of window about to be activated
			//lParam - pointer to CBTACTIVATESTRUCT
			//attempt to use WM_COPYDATA to pass the information to the VB sub-classed window
			//when using, the passed wParam should contain the hWnd of the window that is
			//passing the data, lParam a ptr to a COPYDATASTRUCT.  In here specify:
			//dwData - pass your own info value here
			//cbData - size in bytes of the block of data to send
			//lpData - pointer to block of data (in this case to a CBTACTIVATESTRUCT)

			//fprintf(fErrStream, "%s%d%s%d\n", "**Received HCBT_ACTIVATE - wParam: ", wParam, "  lParam: ", lParam);

			ActivateData = (CBTACTIVATESTRUCT*)lParam;
			//ActivateData->fMouse = ((CBTACTIVATESTRUCT*)lParam)->fMouse;
			//ActivateData->hWndActive = ((CBTACTIVATESTRUCT*)lParam)->hWndActive ;
			
			DataToCopy.dwData = (long)ActivateData->fMouse; //inform client of type of msg
			DataToCopy.cbData = sizeof(*ActivateData); //size of data to send
			DataToCopy.lpData = &ActivateData;

			//fprintf(fErrStream, "%s%d%s%d\n","fMouse: ", ActivateData->fMouse , "  hWndActive: ", ActivateData->hWndActive );

			SendMessage(hCallingAppWnd, WM_COPYDATA, (WPARAM)wParam, (LPARAM)&DataToCopy);


			break;

		//is the system creating a window?
		case HCBT_CREATEWND:
			//wParam contains handle to new window
			//lParam contains ptr to CBT_CREATEWND
			PostMessage(hCallingAppWnd, iNewMsg, wParam, lParam);
			break;
		
		case HCBT_DESTROYWND:
/*			fprintf( fErrStream, "%s\n", "Received a HCBT_DESTROYWND message in CBTProc");
			fprintf( fErrStream, "%s%d%s%d%s%d\n",
								"Posting message",
								iNewMsg, 
								" with hWnd ",
								wParam,
								" to window ",
								hCallingAppWnd);*/

			PostMessage(hCallingAppWnd, iNewMsg, HCBT_DESTROYWND, wParam);
			//fprintf( fErrStream, "%s\n", "Returned from posting message!");
			break;

		default:

			return CallNextHookEx (hdHooks[BM_CBT].hGlobalHooks, nCode, wParam, lParam);
			break;
	}

	return CallNextHookEx (hdHooks[BM_CBT].hGlobalHooks, nCode, wParam, lParam);
}

LRESULT CALLBACK DebugProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_DEBUG].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK ForeGroundIdleProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_FOREGROUNDIDLE].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK GetMsgProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_GETMESSAGE].hGlobalHooks, nCode, wParam, lParam); 
}


LRESULT CALLBACK JournalPlaybackProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_JOURNALPLAYBACK].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK JournalRecordProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_JOURNALRECORD].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK KeyboardProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_KEYBOARD].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK MessageProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_GETMESSAGE].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK MouseProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_MOUSE].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK ShellProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_SHELL].hGlobalHooks, nCode, wParam, lParam); 
}

LRESULT CALLBACK SysMsgProc (int nCode, WPARAM wParam, LPARAM lParam)
{
	return CallNextHookEx(hdHooks[BM_SYSMSGFILTER].hGlobalHooks, nCode, wParam, lParam); 
}


UINT RegisterNewMessageGlobal (void)
{
	UINT iRet;

/*	fprintf(fErrStream, "%s\n","Entered RegisterNewMessageGlobal");
	
	fprintf(fErrStream, "%s\n","Attempting to register a new message...");*/
	iRet = RegisterWindowMessage("WM_BATTHOOK");
//	fprintf(fErrStream, "%s%d\n","...returned from call: ", iRet);
	//store the new message number
	iNewMsg=iRet;
//	fprintf(fErrStream, "%s\n","Exited RegisterNewMessageGlobal normally!");
	return iRet;
}


HHOOK SetUpHookGlobal (long hHookType)
{
	HHOOK hRet;
	int iIndex, iSelIndex;
	HOOKPROC lHookFuncPtr=0;

	for (iIndex=0;iIndex<=11;iIndex++)
	{
		if (hdHooks[iIndex].iHookType == (int)hHookType)
		{
			lHookFuncPtr = hdHooks[iIndex].lpFuncPtr;
			iSelIndex = iIndex;
			break;
		}
	}

	hRet = SetWindowsHookEx((int)hHookType, lHookFuncPtr, (HINSTANCE)hInstDLL, 0);
	if (hRet==0)
		return 0;

	//ok to add hook handle to array
	hdHooks[iSelIndex].hGlobalHooks = hRet;

//return true otherwise
	return hRet;

}


BOOL UnHookGlobal (int hHookType)
{

	int iIndex;

	int iRet;

	for(iIndex=0;iIndex<=11;iIndex++)
	{
		if (hdHooks[iIndex].iHookType == hHookType)
		{
			iRet = UnhookWindowsHookEx(hdHooks[iIndex].hGlobalHooks);
			if (iRet==0)
			{
				return false;
				break;
			}
			hdHooks[iIndex].hGlobalHooks = 0;
			return true;
			break;
		}
	}
	return true;
}

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  fReason, 
                       LPVOID lpReserved
					 )
{
	hInstDLL = hModule;

    return TRUE;
}


HRESULT InitHookArray (void)
{
	hdHooks[BM_CALLWNDPROC].iHookType=4;
	hdHooks[BM_CALLWNDPROC].hGlobalHooks = 0;
	hdHooks[BM_CALLWNDPROC].lpFuncPtr = (HOOKPROC)CallWndProc;

	hdHooks[BM_CBT].iHookType = 5;
	hdHooks[BM_CBT].hGlobalHooks = 0;
	hdHooks[BM_CBT].lpFuncPtr = (HOOKPROC)CBTProc;

	hdHooks[BM_DEBUG].iHookType = 9;
	hdHooks[BM_DEBUG].hGlobalHooks = 0;
	hdHooks[BM_DEBUG].lpFuncPtr = (HOOKPROC)DebugProc;

	hdHooks[BM_FOREGROUNDIDLE].iHookType = 11;
	hdHooks[BM_FOREGROUNDIDLE].hGlobalHooks = 0;
	hdHooks[BM_FOREGROUNDIDLE].lpFuncPtr = (HOOKPROC)ForeGroundIdleProc;

	hdHooks[BM_GETMESSAGE].iHookType = 3;
	hdHooks[BM_GETMESSAGE].hGlobalHooks = 0;
	hdHooks[BM_GETMESSAGE].lpFuncPtr = (HOOKPROC)GetMsgProc;

	hdHooks[BM_JOURNALPLAYBACK].iHookType = 1;
	hdHooks[BM_JOURNALPLAYBACK].hGlobalHooks = 0;
	hdHooks[BM_JOURNALPLAYBACK].lpFuncPtr = (HOOKPROC)JournalPlaybackProc;

	hdHooks[BM_JOURNALRECORD].iHookType = 0;
	hdHooks[BM_JOURNALRECORD].hGlobalHooks =0;
	hdHooks[BM_JOURNALRECORD].lpFuncPtr = (HOOKPROC)JournalRecordProc;

	hdHooks[BM_KEYBOARD].iHookType = 2;
	hdHooks[BM_KEYBOARD].hGlobalHooks = 0;
	hdHooks[BM_KEYBOARD].lpFuncPtr = (HOOKPROC)KeyboardProc;

	hdHooks[BM_MOUSE].iHookType = 7;
	hdHooks[BM_MOUSE].hGlobalHooks =0;
	hdHooks[BM_MOUSE].lpFuncPtr = (HOOKPROC)MouseProc;

	hdHooks[BM_MSGFILTER].iHookType = -1;
	hdHooks[BM_MSGFILTER].hGlobalHooks = 0;
	hdHooks[BM_MSGFILTER].lpFuncPtr = (HOOKPROC)MessageProc;

	hdHooks[BM_SHELL].iHookType = 10;
	hdHooks[BM_SHELL].hGlobalHooks = 0;
	hdHooks[BM_SHELL].lpFuncPtr = (HOOKPROC)ShellProc;

	hdHooks[BM_SYSMSGFILTER].iHookType = 6;
	hdHooks[BM_SYSMSGFILTER].hGlobalHooks = 0;
	hdHooks[BM_SYSMSGFILTER].lpFuncPtr = (HOOKPROC)SysMsgProc;

	return 0;
}
// This is to prevent the CRT from loading, thus making this a smaller
//  and faster dll.
//extern "C" BOOL __stdcall _DllMainCRTStartup( HINSTANCE hinstStDLL, DWORD fdwReason, LPVOID lpvReserved) {
//    return DllMain( hinstStDLL, fdwReason, lpvReserved );
//}