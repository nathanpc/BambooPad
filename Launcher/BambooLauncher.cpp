/**
 * BambooLauncher.cpp
 * A simple launcher for the eVB program BambooPad.
 *
 * @author Nathan Campos <nathan@innoveworkshop.com>
 */

#include "stdafx.h"
#include "shellapi.h"

LPTSTR szAppName = _T("BambooPad.vb");
LPTSTR szStubName = _T("BambooLauncher.exe");

int WINAPI WinMain(HINSTANCE hInstance,
				   HINSTANCE hPrevInstance,
				   LPTSTR    lpCmdLine,
				   int       nCmdShow) {
	long retVal;
	TCHAR szPath[128];
	LPTSTR Instr;
	LPTSTR szVerb = _T("open");

	// Start the Visual Basic application and exit.
	SHELLEXECUTEINFO lpExecInfo;
	memset(&lpExecInfo, 0, sizeof(SHELLEXECUTEINFO));
	lpExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);

	// Get the path to the current directory.
	retVal = GetModuleFileName(hInstance, szPath, 128);
	if (retVal) {
		// Remove the stubs file name to get just the path.
		Instr = wcsstr(szPath, szStubName);

		// Add the target file to the resulting path.
		if (Instr != NULL)
			wcscpy(Instr, szAppName);

		//MessageBox(0, szPath, _T("This is the path I got"), 0);

		// Now use this to start the application.
		lpExecInfo.lpFile = szPath;
		lpExecInfo.nShow = SW_SHOWNORMAL;
		lpExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
		lpExecInfo.lpVerb = szVerb ;

		ShellExecuteEx(&lpExecInfo);
		return 0;
	}

	MessageBox(0, _T("An error occured while trying to launch the eVB application"),
		_T("Cannot launch eVB program"), MB_OK + MB_ICONERROR);
	return -1;
}

