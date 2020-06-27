// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__24F82719_C5C2_48E5_8403_1CA5C88AD903__INCLUDED_)
#define AFX_STDAFX_H__24F82719_C5C2_48E5_8403_1CA5C88AD903__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

#include <stdio.h>


//dynamic array that contains DTM data
typedef struct tag_elevat
{
	double vert[3];
} ELEVAT;


// TODO: reference additional headers your program requires here

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__24F82719_C5C2_48E5_8403_1CA5C88AD903__INCLUDED_)
