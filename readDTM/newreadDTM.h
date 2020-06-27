// newreadDTM.h : main header file for the NEWREADDTM application
//

#if !defined(AFX_NEWREADDTM_H__24C37FA4_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_)
#define AFX_NEWREADDTM_H__24C37FA4_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CNewreadDTMApp:
// See newreadDTM.cpp for the implementation of this class
//

class CNewreadDTMApp : public CWinApp
{
public:
	CNewreadDTMApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CNewreadDTMApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CNewreadDTMApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NEWREADDTM_H__24C37FA4_4B82_11D3_A3C5_CA4CD97BAB5E__INCLUDED_)
