
// ghosttrip.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CghosttripApp:
// See ghosttrip.cpp for the implementation of this class
//

class CghosttripApp : public CWinApp
{
public:
	CghosttripApp();

// Overrides
public:
	virtual BOOL InitInstance();
	virtual BOOL ExitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CghosttripApp theApp;