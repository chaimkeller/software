//VBStringFunctions.CPP

#include "stdafx.h"

#include <Stdio.h>
#include <Stdlib.h>
#include <Assert.h>
#include <stdarg.h>
#include "VBString.h"

//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Compare(const char* lpsz) const
{
	return MyStrcmp((unsigned char*)lpsz,(unsigned char*)this->m_pBuffer);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::CompareNoCase(const char* lpsz) const
{
	CVBString s1=lpsz;
	CVBString s2=*this;
	s1.UCase();
	s2.UCase();
	return MyStrcmp((unsigned char*)s1.m_pBuffer,(unsigned char*)s2.m_pBuffer);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString CVBString::Mid(int nFirst, int nCount) const
{
	CVBString s="";
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	if((unsigned long int)nFirst>=strLen)
	{
		return s;
	}
	if((unsigned long int)(nFirst+nCount)>=strLen)
	{
		nCount=strLen-(unsigned long int)(nFirst);
	}

	s=*this;

	s.Delete(nFirst,nCount);

	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString CVBString::Mid(int nFirst) const
{
	CVBString s="";
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	unsigned long int nCount=strLen-(unsigned long int)nFirst;

	if((unsigned long int)nFirst>=strLen)
	{
		return s;
	}
	if((unsigned long int)(nFirst+nCount)>=strLen)
	{
		nCount=strLen-(unsigned long int)(nFirst);
	}

	s=*this;

	s.Delete(nFirst,nCount);

	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString CVBString::Left(int nCount)
{
	CVBString s;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	if((unsigned long int)nCount>strLen)
	{
		s="";
		return s;
	}

	strLen-=(unsigned long int)nCount;

	s.GetBuffer(strLen+1);

	MyStrcpy(s.m_pBuffer,&this->m_pBuffer[nCount]);

	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString CVBString::Right(int nCount)
{
	CVBString s;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	if((unsigned long int)nCount>strLen)
	{
		s=*this;
		return s;
	}

	strLen-=(unsigned long int)nCount;

	s.GetBuffer(strLen+1);

	MyStrncpy(s.m_pBuffer,this->m_pBuffer,strLen);
	s.m_pBuffer[strLen]=0;

	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::UCase()
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]>='a'&&this->m_pBuffer[i]<='z')
			this->m_pBuffer[i]+=('A'-'a');
	}
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::LCase()
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]>='A'&&this->m_pBuffer[i]<='Z')
			this->m_pBuffer[i]+=('a'-'A');
	}
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::Reverse()
{
	CVBString s=*this;
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=0;i<strLen;i++)
	{
		this->m_pBuffer[i]=s.m_pBuffer[(strLen-1)-i];
	}

}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::RTrim()
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=(strLen-1);i!=0;i--)
	{
		if(IsWhiteSpace(this->m_pBuffer[i])==false)
			break;
	}
	*this=this->Right((strLen-1)-i);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::LTrim()
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=0;i<strLen;i++)
	{
		if(IsWhiteSpace(this->m_pBuffer[i])==false)
			break;
	}
	*this=this->Left(i);

}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::RTrim(char chTarget)
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=(strLen-1);i!=0;i--)
	{
		if(this->m_pBuffer[i]!=chTarget)
			break;
	}
	*this=this->Right((strLen-1)-i);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::RTrim(const char* lpszTargets)
{
	if(MyStrcmp((unsigned char*)&this->m_pBuffer[MyStrlen(this->m_pBuffer)-MyStrlen(lpszTargets)],(unsigned char*)lpszTargets)==0)
	{
		this->Right(MyStrlen(this->m_pBuffer));
	}
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::LTrim(char chTarget)
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]!=chTarget)
			break;
	}
	*this=this->Left(i);

}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::LTrim(const char* lpszTargets)
{
	if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)lpszTargets)==0)
	{
		this->Left(MyStrlen(this->m_pBuffer));
	}
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Replace(char chOld, char chNew)
{
	unsigned long int i;
	unsigned long int iChangedChars=0;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]==chOld)
		{
			this->m_pBuffer[i]=chNew;
			iChangedChars++;
		}
	}

	return iChangedChars;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Replace(const char* lpszOld, const char* lpszNew)
{
	CVBString s;
	unsigned long int i;
	unsigned long int j;
	unsigned long int iBuffSize;
	unsigned long int iIndex=0;
	unsigned long int iReplacedStrings=0;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	while(1)
	{
		iIndex=this->Find(lpszOld,iIndex);
		if(iIndex==-1)
			break;

		iIndex+=MyStrlen(lpszOld);

		iReplacedStrings++;
	}

	if(!iReplacedStrings)
	{
		return 0;
	}

	iBuffSize = MyStrlen(this->m_pBuffer)+1;
	iBuffSize -= MyStrlen(lpszOld)*iReplacedStrings;
	iBuffSize += MyStrlen(lpszNew)*iReplacedStrings;
	s.GetBuffer(iBuffSize);

	iIndex=0;
	for(i=0,j=0;i<(strLen);i++,j++)
	{
		if(MyStrncmp(&this->m_pBuffer[i],(char*)lpszOld,MyStrlen(lpszOld))==0)
		{
			MyStrcat(&s.m_pBuffer[j],(char*)lpszNew);
			i+=(MyStrlen(lpszOld)-1);
			j+=(MyStrlen(lpszNew)-1);
		}
		else
		{
			s.m_pBuffer[j]=this->m_pBuffer[i];
			s.m_pBuffer[j+1]=0;
		}
	}

	*this = s;

	return iReplacedStrings;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Remove(char chRemove)
{
	CVBString s;
	unsigned long int i;
	unsigned long int j;
	unsigned long int iChangedChars=0;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	unsigned long int iNewStrLen=MyStrlen(this->m_pBuffer)+1;

	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]==chRemove)
		{
			iNewStrLen--;
			iChangedChars++;
		}
	}

	s.GetBuffer(iNewStrLen);

	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[i]!=chRemove)
		{
			s.m_pBuffer[j]=this->m_pBuffer[i];
			j++;
		}
	}
	*this = s;
	return iChangedChars;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Insert(int nIndex, char ch)
{
	CVBString s;
	unsigned long int i;
	unsigned long int j;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	s.GetBuffer(MyStrlen(this->m_pBuffer)+2);

	for(i=0,j=0;i<strLen;i++,j++)
	{
		if(i==(unsigned long int)nIndex)
		{
			s.m_pBuffer[i]=ch;
			j++;
		}
		else
		{
			s.m_pBuffer[j]=this->m_pBuffer[i];
		}
	}
	return MyStrlen(this->m_pBuffer);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Insert(int nIndex, const char* pstr)
{
	CVBString s;
	unsigned long int i;
	unsigned long int j;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);

	s.GetBuffer(MyStrlen(this->m_pBuffer)+1+MyStrlen(pstr));

	for(i=0,j=0;i<strLen;i++,j++)
	{
		if(MyStrcmp((unsigned char*)&this->m_pBuffer[i],(unsigned char*)pstr)==0)
		{
			MyStrcpy(&s.m_pBuffer[i],pstr);
			j+=MyStrlen(pstr);
		}
		else
		{
			s.m_pBuffer[j]=this->m_pBuffer[i];
		}
	}
	return MyStrlen(this->m_pBuffer);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Delete(int nIndex, int nCount)
{
	CVBString s;
	unsigned long int i;
	unsigned long int j;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	unsigned long int iNewstrLen=MyStrlen(this->m_pBuffer)-(unsigned long int)(nCount);
	iNewstrLen+=1;

	if((unsigned long int)(nIndex)>strLen)
	{
		return MyStrlen(this->m_pBuffer); 
	}
	if((unsigned long int)(nIndex+nCount)>strLen)
	{
		nCount=strLen-(unsigned long int)nIndex;
	}

	s.GetBuffer(iNewstrLen);

	for(i=0,j=0;i<strLen;i++,j++)
	{
		if(i==(unsigned long int)nIndex)
		{
			i+=nCount;
		}
		else
		{
			s.m_pBuffer[j]=this->m_pBuffer[i];
		}
	}
	return MyStrlen(this->m_pBuffer);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Find(char ch) const
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=0;i<strLen;i++)
	{
		if(this->m_pBuffer[ch]==ch)
			return i;
	}

	return -1;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::ReverseFind(char ch) const
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=(strLen-1);i!=strLen;i--)
	{
		if(this->m_pBuffer[ch]==ch)
			return i;
	}

	return -1;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Find(char ch, int nStart) const
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=nStart;i<strLen;i++)
	{
		if(this->m_pBuffer[ch]==ch)
			return i;
	}

	return -1;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Find(const char* lpszSub) const
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=0;i<strLen;i++)
	{
		if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)lpszSub)==0)
			return i;
	}

	return -1;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int CVBString::Find(const char* lpszSub, int nStart) const
{
	unsigned long int i;
	unsigned long int strLen=MyStrlen(this->m_pBuffer);
	for(i=nStart;i<strLen;i++)
	{
		if(MyStrncmp(&this->m_pBuffer[i],(char*)lpszSub,MyStrlen(lpszSub))==0)
			return i;
	}

	return -1;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void __cdecl CVBString::Format(const char* lpszFormat, ...)
{
	CVBString s;
	va_list ArgList;

	va_start(ArgList, lpszFormat);
	vsprintf(s.GetBuffer(4096),lpszFormat,ArgList);
	va_end(ArgList);

	*this=s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::FormatV(const char* lpszFormat, va_list argList)
{
	CVBString s;

	vsprintf(s.GetBuffer(4096),lpszFormat,argList);

	*this=s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
bool CVBString::IsWhiteSpace(char ch)
{
	if((ch>='\t'&&ch<='\r')||ch==' ')//0x9 - 0x13 , 0x20
		return true;

	return false;
}
