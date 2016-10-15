// WordPRO.h : main header file for the WORDPRO DLL
//

#if !defined(AFX_WORDPRO_H__96F3D297_6322_415A_A085_71651CD79F0B__INCLUDED_)
#define AFX_WORDPRO_H__96F3D297_6322_415A_A085_71651CD79F0B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
#include "WordPRO_i.h"

/////////////////////////////////////////////////////////////////////////////
// CWordPROApp
// See WordPRO.cpp for the implementation of this class
//

class CWordPROApp : public CWinApp
{
public:
	CWordPROApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWordPROApp)
	public:
	virtual BOOL InitInstance();
		virtual int ExitInstance();
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CWordPROApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	BOOL InitATL();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WORDPRO_H__96F3D297_6322_415A_A085_71651CD79F0B__INCLUDED_)
