//VB.CPP
#include "stdafx.h"


#include <Stdio.h>
#include <Stdlib.h>
#include <Assert.h>
#include <stdarg.h>
#include "VBString.h"


//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::InStr(int nStart,const char* lpszSub)
{
	unsigned long int i;
	nStart -= 1;

	i=this->Find(lpszSub,nStart);
	if(i>=0)
	{
		return (i+1);
	}
	else if(i==-1)
	{
		return 0;
	}

	return 0;
}
