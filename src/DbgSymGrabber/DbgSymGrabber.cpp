/*
This is the software license for DbgSymGrabber.
DbgSymGrabber has been written by quangnh89  <quangnh89(at)gmail(dot)com>

Copyright (c) 2015, quangnh89.
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
* Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.
* Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.
* Neither the name of the developer(s) nor the names of its
contributors may be used to endorse or promote products derived from this
software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
*/

#define DBGHELP_TRANSLATE_TCHAR
#include <tchar.h>
#include <Windows.h>
#include <commctrl.h>
#include "resource.h"
#include <Shlwapi.h>
#include <shlobj.h>
#include <Dbghelp.h>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")

/* enabling visual styles */
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define IS_FILE (0)
#define IS_DIR  (1)
#define SEARCH_PATH_SIZE (2 * MAX_PATH + 100)

typedef BOOL  (WINAPI *FP_SYMINITIALIZEW)( __in HANDLE hProcess, __in_opt PCWSTR UserSearchPath, __in BOOL fInvadeProcess );
typedef BOOL  (WINAPI *FP_SYMCLEANUP)( HANDLE hProcess );
typedef DWORD (WINAPI *FP_SYMGETOPTIONS)(void);
typedef DWORD (WINAPI *FP_SYMSETOPTIONS)( DWORD SymOptions );
typedef BOOL (WINAPI *FP_SYMGETSYMBOLFILEW)( HANDLE hProcess, PCWSTR SymPath,
									 PCWSTR ImageFile, DWORD Type,
									 PWSTR SymbolFile, size_t cSymbolFile,
									 PWSTR DbgFile, size_t cDbgFile );
typedef BOOL (WINAPI *FP_SYMSETPARENTWINDOW)( HWND hwnd );


typedef struct FILE_DATA 
{
	TCHAR szImage[MAX_PATH + 1];
	TCHAR szServer[SEARCH_PATH_SIZE];
}FILE_DATA, *LPFILE_DATA;

FP_SYMINITIALIZEW     fpSymInitializeW;
FP_SYMCLEANUP	      fpSymCleanup;
FP_SYMGETOPTIONS      fpSymGetOptions;
FP_SYMSETOPTIONS      fpSymSetOptions;
FP_SYMGETSYMBOLFILEW  fpSymGetSymbolFileW;
FP_SYMSETPARENTWINDOW fpSymSetParentWindow;

BOOL				  g_bReady = FALSE;
HWND				  g_hMain = NULL;
HWND				  g_hStatusBar = NULL;
HANDLE				  g_hThead = NULL;
HMODULE				  g_hDbgHelp = NULL;

INT_PTR CALLBACK DialogProc( __in HWND hwnd, __in UINT uMsg, __in WPARAM wParam, __in LPARAM lParam);

int WINAPI _tWinMain( __in HINSTANCE hInstance, __in_opt HINSTANCE hPrevInstance, __in LPTSTR lpCmdLine, __in int nShowCmd )
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	UNREFERENCED_PARAMETER(nShowCmd);

	InitCommonControls();
	if (FAILED(CoInitialize(NULL)))
		return 0;

	DialogBoxParam(hInstance, MAKEINTRESOURCE(IDD_DBGSYM), NULL, DialogProc, NULL);
	CoUninitialize();
	return 0;
}

/* drag and drop file or folder to text box */
LRESULT CALLBACK DropFilesEditProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
{
	UNREFERENCED_PARAMETER(uIdSubclass);

	if (uMsg == WM_DROPFILES)
	{
		TCHAR szFilePath[MAX_PATH + 1] = {};

		/* query file path */
		if (DragQueryFile((HDROP) wParam, 0, szFilePath, MAX_PATH)!= 0)
		{
			if(( dwRefData == IS_FILE && !PathIsDirectory(szFilePath)) || 
				(dwRefData == IS_DIR && PathIsDirectory(szFilePath)))
			SetWindowText(hWnd, szFilePath);
		}
		DragFinish((HDROP) wParam); 
		return 0;
	}
	return DefSubclassProc( hWnd, uMsg, wParam, lParam);
}

BOOL LoadDbgHelp( void )
{
	TCHAR szLibPath[MAX_PATH + 1];

	if (!GetModuleFileName(NULL, szLibPath, MAX_PATH))
		return FALSE;
	PathRemoveFileSpec(szLibPath);
	PathAppend(szLibPath, TEXT("Dbghelp.dll"));

	g_hDbgHelp = LoadLibrary(szLibPath);
	if (g_hDbgHelp == NULL)
		return FALSE;

	fpSymInitializeW    = (FP_SYMINITIALIZEW)GetProcAddress(g_hDbgHelp, "SymInitializeW");
	fpSymCleanup        = (FP_SYMCLEANUP)GetProcAddress(g_hDbgHelp, "SymCleanup");
	fpSymGetOptions     = (FP_SYMGETOPTIONS)GetProcAddress(g_hDbgHelp, "SymGetOptions");
	fpSymSetOptions     = (FP_SYMSETOPTIONS)GetProcAddress(g_hDbgHelp, "SymSetOptions");
	fpSymGetSymbolFileW = (FP_SYMGETSYMBOLFILEW)GetProcAddress(g_hDbgHelp, "SymGetSymbolFileW");
	fpSymSetParentWindow= (FP_SYMSETPARENTWINDOW)GetProcAddress(g_hDbgHelp, "SymSetParentWindow");
	if (!fpSymInitializeW || !fpSymCleanup || !fpSymGetOptions ||
		!fpSymSetOptions || !fpSymSetParentWindow || !fpSymGetSymbolFileW)
	{
		FreeLibrary(g_hDbgHelp);
		g_hDbgHelp = NULL;
		return FALSE;
	}

	return TRUE;
}

/* initialize the dialog box.  */
void OnInitDialog( __in HWND hwnd )
{
	/* load main icon */
	HICON hIcon = LoadIcon( GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_MAIN) );
	if( hIcon )
	{
		SendMessage( hwnd, WM_SETICON, ICON_BIG, (LPARAM)hIcon );
		SendMessage( hwnd, WM_SETICON, ICON_SMALL, (LPARAM)hIcon );
		DestroyIcon( hIcon );
	}

	/* Enable drag and drop file */
	SHAutoComplete(GetDlgItem(hwnd, IDC_TXT_IMAGE), 
		SHACF_AUTOSUGGEST_FORCE_ON | SHACF_FILESYSTEM | SHACF_USETAB);
	SetWindowSubclass(GetDlgItem(hwnd, IDC_TXT_IMAGE), DropFilesEditProc, 1, IS_FILE);
	DragAcceptFiles(GetDlgItem(hwnd, IDC_TXT_IMAGE), TRUE); 

	SHAutoComplete(GetDlgItem(hwnd, IDC_TXT_CACHE), 
		SHACF_AUTOSUGGEST_FORCE_ON | SHACF_FILESYS_DIRS | SHACF_USETAB);
	SetWindowSubclass(GetDlgItem(hwnd, IDC_TXT_CACHE), DropFilesEditProc, 1, IS_DIR);
	DragAcceptFiles(GetDlgItem(hwnd, IDC_TXT_CACHE), TRUE); 

	SHAutoComplete(GetDlgItem(hwnd, IDC_TXT_SERVER), 
		SHACF_AUTOSUGGEST_FORCE_ON | SHACF_URLALL | SHACF_USETAB);

	/* Draw status bar */
	RECT clientRect;
	GetClientRect(hwnd, &clientRect);
	g_hStatusBar = CreateWindowEx(0, STATUSCLASSNAME, TEXT(""), WS_VISIBLE | WS_CHILD, 
		0, clientRect.bottom - clientRect.top - 30,
		clientRect.right - clientRect.left, 30, 
		hwnd, NULL, GetModuleHandle(NULL), NULL);
	
	if (g_hStatusBar)
	{
		const int nStatusSize = clientRect.right - clientRect.left ;
		int status_parts[2] = { nStatusSize / 2 - 50, nStatusSize };
		SendMessage(g_hStatusBar, SB_SETPARTS, 2, (LPARAM)&status_parts);
		SendMessage(g_hStatusBar, SB_SETTEXT , 1, (LPARAM)TEXT("https://github.com/quangnh89/DbgSymGrabber"));
	}

	/* Initializing the Symbol Handler */
	g_bReady = LoadDbgHelp();
	if ( !g_bReady )
	{
		MessageBox(hwnd, TEXT("DbgHelp.dll not found."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONERROR);
		SendMessage(g_hStatusBar, SB_SETTEXT , 0, (LPARAM)TEXT("DbgHelp.dll not found."));
		return;
	}

	g_bReady = fpSymInitializeW(GetCurrentProcess(), NULL, FALSE);
 	if (g_bReady)
 	{
		fpSymSetParentWindow(hwnd);

		DWORD dwOption = fpSymGetOptions();
		dwOption |= SYMOPT_FAIL_CRITICAL_ERRORS; /* Do not display system dialog box when media failure */
		dwOption |= SYMOPT_FAVOR_COMPRESSED; /*good for slow connection */
		fpSymSetOptions(dwOption);
 	}
	if (g_hStatusBar)
	{
		SendMessage(g_hStatusBar, SB_SETTEXT , 0, (LPARAM)TEXT("Ready!"));
	}

	g_hMain = hwnd;
}

void OnClose( __in HWND hwnd )
{
	/* disable drag and drop file */
	 DragAcceptFiles(GetDlgItem(hwnd, IDC_TXT_IMAGE), FALSE); 
	 DragAcceptFiles(GetDlgItem(hwnd, IDC_TXT_CACHE), FALSE);
	 
	 if (g_hStatusBar)
		 DestroyWindow(g_hStatusBar);

	 if (g_bReady)
	 {
		 fpSymCleanup(GetCurrentProcess());
	 }
}

/* Display dialog box to select an image file */
void OnBtnImageClick( __in HWND hwnd )
{
	TCHAR szFile[MAX_PATH + 1] = {};
	TCHAR szSavedPath[MAX_PATH + 1];
	OPENFILENAME ofn = {};

	// get current path
	GetDlgItemText(hwnd, IDC_TXT_IMAGE, szSavedPath, MAX_PATH);
	PathRemoveFileSpec(szSavedPath);

	ZeroMemory( &ofn, sizeof( ofn));
	ofn.lStructSize    = sizeof ( ofn );
	ofn.hwndOwner      = hwnd;
	ofn.lpstrFile      = szFile;
	ofn.nMaxFile       = MAX_PATH;
	ofn.lpstrFilter    = TEXT("Executable Files(*.EXE;*.DLL;*.SYS;*.DRV;*.CPL)\0*.EXE;*.DLL;*.SYS;*.DRV;*.CPL\0All Files(*.*)\0*.*\0\0");
	ofn.nFilterIndex   = 1;
	ofn.lpstrFileTitle = NULL;
	ofn.nMaxFileTitle  = 0;
	ofn.lpstrInitialDir= szSavedPath;
	ofn.Flags          = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_LONGNAMES | OFN_ENABLESIZING | OFN_EXPLORER;

	if (GetOpenFileName(&ofn))
		SetDlgItemText(hwnd, IDC_TXT_IMAGE, szFile);
}

int CALLBACK BrowseCallbackProc( __in HWND hwnd, __in UINT uMsg, __in LPARAM lParam, __in LPARAM lpData)
{
	UNREFERENCED_PARAMETER(lParam);

	if(uMsg == BFFM_INITIALIZED)
		SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);

	return 0;
}

/* Display dialog box to select a directory */
void OnBtnCacheClick( __in HWND hwnd )
{
	TCHAR szPath[MAX_PATH + 1];
	TCHAR szSavedPath[MAX_PATH + 1];

	// get current path
	GetDlgItemText(hwnd, IDC_TXT_CACHE, szSavedPath, MAX_PATH);

	BROWSEINFO bi = { 0 };
	bi.lpszTitle  = TEXT("Please select a directory where symbol files can be saved.");
	bi.ulFlags    = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
	bi.lpfn       = BrowseCallbackProc;
	bi.lParam     = (LPARAM) szSavedPath;

	LPITEMIDLIST pidl = SHBrowseForFolder ( &bi );

	if ( pidl != NULL )
	{
		//get the name of the folder and put it in path
		SHGetPathFromIDList ( pidl, szPath );

		//free memory used
		IMalloc * imalloc = 0;
		if ( SUCCEEDED( SHGetMalloc ( &imalloc )) )
		{
			imalloc->Free( pidl );
			imalloc->Release( );
		}

		SetDlgItemText(hwnd, IDC_TXT_CACHE, szPath);
	}
}

/* set MS Symbol server */
void OnBtnMsServerClick( __in HWND hwnd )
{
	SetDlgItemText(hwnd, IDC_TXT_SERVER, TEXT("http://msdl.microsoft.com/download/symbols"));
}

/* Download thread */
DWORD WINAPI GrabFileThread(__in LPVOID lpParam)
{
	TCHAR szSymbolFile[MAX_PATH + 1];
	TCHAR szDbgFile[MAX_PATH + 1];
	LPFILE_DATA lpGrabData = (LPFILE_DATA)lpParam;
	
	if (lpGrabData == NULL)
	{
		goto QUIT;
	}

	/* Disable main button */
	EnableWindow(GetDlgItem(g_hMain, IDC_BTN_GRAB), FALSE);
	SetDlgItemText(g_hMain, IDC_BTN_GRAB, TEXT("Waiting..."));

	SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)TEXT("Start downloading..."));
	
	/* download symbol file */
	if (fpSymGetSymbolFileW(GetCurrentProcess(), 
		lpGrabData->szServer, lpGrabData->szImage, 
		sfPdb, 
		szSymbolFile, MAX_PATH, 
		szDbgFile, MAX_PATH))
	{
		SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)TEXT("Download successfully :)"));
	}
	else
	{
		SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)TEXT("Download unsuccessfully :("));
	}

	VirtualFree(lpGrabData, 0, MEM_RELEASE);

QUIT:
	CloseHandle(g_hThead);
	g_hThead = NULL;
	/* enable main button */
	SetDlgItemText(g_hMain, IDC_BTN_GRAB, TEXT("&Grab!"));
	EnableWindow(GetDlgItem(g_hMain, IDC_BTN_GRAB), TRUE);
	return 0;
}

void OnBtnGrabClick( __in HWND hwnd )
{
	TCHAR szCache[MAX_PATH + 1];	
	TCHAR szUrl[MAX_PATH + 1];

	LPFILE_DATA lpGrabData;

	if (g_bReady == FALSE)
	{
		MessageBox(hwnd, TEXT("DbgHelp.dll not found. Fix this error and restart program."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONERROR);
		return;
	}

	lpGrabData = (LPFILE_DATA)VirtualAlloc(NULL, sizeof(FILE_DATA), MEM_COMMIT, PAGE_READWRITE);
	if (lpGrabData == NULL)
	{
		MessageBox(hwnd, TEXT("Out of memory."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONWARNING);
		return;
	}

	if (GetDlgItemText(hwnd, IDC_TXT_CACHE, szCache, MAX_PATH) == 0)
	{
		MessageBox(hwnd, TEXT("Select cache directory."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONWARNING);
		VirtualFree(lpGrabData, 0, MEM_RELEASE);
		return;
	}
	if (GetDlgItemText(hwnd, IDC_TXT_IMAGE, lpGrabData->szImage, MAX_PATH) == 0)
	{
		MessageBox(hwnd, TEXT("Select image file."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONWARNING);
		VirtualFree(lpGrabData, 0, MEM_RELEASE);
		return;
	}
	if (GetDlgItemText(hwnd, IDC_TXT_SERVER, szUrl, MAX_PATH) < 0)
	{
		MessageBox(hwnd, TEXT("Select server."), TEXT("DbgSymGrabber"), MB_OK | MB_ICONWARNING);
		VirtualFree(lpGrabData, 0, MEM_RELEASE);
		return;
	}
	_stprintf_s(lpGrabData->szServer, SEARCH_PATH_SIZE, TEXT("SRV*%s*%s"), szCache, szUrl);
	g_hThead = CreateThread(NULL, 0, GrabFileThread, lpGrabData, 0, NULL);
	if (g_hThead == NULL)
	{
		//  CreateThread() failed
		VirtualFree(lpGrabData, 0, MEM_RELEASE);
		SendMessage(g_hStatusBar, SB_SETTEXT, 0, (LPARAM)TEXT("Download unsuccessfully."));
	}
}

/* application-defined callback function that processes messages */
BOOL CALLBACK DialogProc( __in HWND hwnd, __in UINT uMsg, __in WPARAM wParam, __in LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch(uMsg)
	{
	case WM_INITDIALOG:
		OnInitDialog(hwnd);
		break;

	case WM_CLOSE:
		OnClose(hwnd);
		EndDialog(hwnd, 0);
		break;

	case WM_COMMAND:
		{
			switch(LOWORD(wParam))
			{
			case IDC_BTN_IMAGE:
				OnBtnImageClick(hwnd);
				break;

			case IDC_BTN_CACHE:
				OnBtnCacheClick(hwnd);
				break;

			case IDC_BTN_MS_SERVER:
				OnBtnMsServerClick(hwnd);
				break;

			case IDC_BTN_GRAB:
				OnBtnGrabClick(hwnd);
				break;

			default:
				return FALSE;
			}
			break;
		}
	default:
		return FALSE;
	}

	return TRUE;
}
