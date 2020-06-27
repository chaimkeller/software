//VBString.CPP

/******************************************************************
*
* Emulate VB String Functions For C++
* 
* 
*
******************************************************************/
#include "stdafx.h"

#include <Stdio.h>
#include <Stdlib.h>
#include <Assert.h>
#include <stdarg.h>
#include "VBString.h"

//	fgets(g_Buffer,256,f);
//	g_Buffer[strlen(g_Buffer)-1]=0;


/*	__try{
	throw("jhkh");

	}
	catch(char *str)
	{
		printf(str);
	}*///CString

//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
/*void main()
{
	CVBString s, s2, s3[3][4];

	s="Hello World ";
	short tmp = 0;
	s3[0][0] = "Hello ";
	s3[0][1] = "Very Intersting ";
	s3[0][2] = " Nice ";
    s3[0][3] = " Most Unusual ";
	s3[1][0] = "Hello You ";
	s3[1][1] = "Very Intersting You ";
	s3[1][2] = " Nice You";
    s3[1][3] = " Most Unusual You ";
	s3[2][0] = "Hello Them";
	s3[2][1] = "Very Intersting them";
	s3[2][2] = " Nice them";
    s3[2][3] = " Most Unusual them";

    
		
	s = s + s;
	printf("%s\n", s);
    tmp = s.InStr(8, "World" );
	printf("%d\n", tmp);
	s2 = s + "Hello Shalom Zvi";
	printf("%s\n", s2);
	if (s2.InStr(10, "Shalom" ) )
	{
		printf("found\n");
	}

	tmp =s2.InStr(10, "Shalom" );
	s3[2][2] = s3[2][2] + s3[0][1];
	printf( "%s\n", s3[2][2] );


}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void Testing()
{
	
}*/
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString::CVBString()
{
	m_pBuffer=0;
	m_nBufferLength=0;
	//------------------------------------------
	m_nBufferLength=1;
	m_pBuffer = (char*)malloc(m_nBufferLength);
	if(!m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	m_pBuffer[0]=0;
	//------------------------------------------
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString::~CVBString()
{
	if(m_pBuffer)
	{
		free(m_pBuffer);
		m_nBufferLength=0;
	}
	m_pBuffer=0;
	m_nBufferLength=0;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString::CVBString(const CVBString& stringSrc)
{
	m_pBuffer=0;
	m_nBufferLength=MyStrlen(stringSrc.m_pBuffer)+1;
	m_pBuffer = (char*)malloc(m_nBufferLength);
	if(!m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(m_pBuffer,stringSrc.m_pBuffer);

}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString::CVBString(const char* lpsz)
{
	m_pBuffer=0;
	m_nBufferLength=MyStrlen(lpsz)+1;
	m_pBuffer = (char*)malloc(m_nBufferLength);
	if(!m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(m_pBuffer,lpsz);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
const CVBString& CVBString::operator=(const CVBString& stringSrc)
{
	if(m_pBuffer)
	{
		free(m_pBuffer);
		m_nBufferLength=0;
	}

	//---------------------------
	if(stringSrc.m_pBuffer==0)//
	{
		m_nBufferLength=1;
		m_pBuffer = (char*)malloc(m_nBufferLength);
		if(!m_pBuffer)
		{
			assert(0);//Out Of Memory!
		}
		m_pBuffer[0]=0;
		return *this;
	}
	//---------------------------
	m_nBufferLength=MyStrlen(stringSrc.m_pBuffer)+1;
	m_pBuffer = (char*)malloc(m_nBufferLength);
	if(!m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(m_pBuffer,stringSrc.m_pBuffer);

	return *this;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
const CVBString& CVBString::operator=(const char *lpsz)
{
	if(m_pBuffer)
	{
		free(m_pBuffer);
		m_nBufferLength=0;
	}
	//---------------------------
	if(lpsz==0)//
	{
		m_nBufferLength=1;
		m_pBuffer = (char*)malloc(m_nBufferLength);
		if(!m_pBuffer)
		{
			assert(0);//Out Of Memory!
		}
		m_pBuffer[0]=0;
		return *this;
	}
	//---------------------------
	m_nBufferLength=MyStrlen(lpsz)+1;
	m_pBuffer = (char*)malloc(m_nBufferLength);
	if(!m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(m_pBuffer,lpsz);

	return *this;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char *MyStrcpy(char * dst,  const char *src)
{
	while( *dst++ = *src++ );
	return( dst );
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int MyStrlen(const char *str)
{
	int length = 0;

	while( *str++ )
		++length;

	return( length );
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char *MyStrcat(char * dst, char * src)
{
	char * cp = dst;
	while( *cp )
		++cp;// Find end of dst
	while( *cp++ = *src++ );
	return( dst );
}

//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString operator+(const CVBString& string1,const CVBString& string2)
{
	CVBString s;

	if(s.m_pBuffer != 0)
	{
		free(s.m_pBuffer);
		s.m_nBufferLength=0;
	}
	s.m_nBufferLength = MyStrlen(string1.m_pBuffer)+MyStrlen(string2.m_pBuffer)+1;
	s.m_pBuffer = (char*)malloc(s.m_nBufferLength);
	if(!s.m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(s.m_pBuffer,string1.m_pBuffer);
	MyStrcat(s.m_pBuffer,string2.m_pBuffer);
	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString operator+(const CVBString& string, char ch)
{
	CVBString s;

	if(s.m_pBuffer != 0)
	{
		free(s.m_pBuffer);
		s.m_nBufferLength=0;
	}
	s.m_nBufferLength = MyStrlen(string.m_pBuffer)+1+1;
	s.m_pBuffer = (char*)malloc(s.m_nBufferLength);//ch+'/0' = 1+1
	if(!s.m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(s.m_pBuffer,string.m_pBuffer);
	s.m_pBuffer[MyStrlen(s.m_pBuffer)]=ch;
	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString operator+(char ch, const CVBString& string)
{
	CVBString s;

	if(s.m_pBuffer != 0)
	{
		free(s.m_pBuffer);
		s.m_nBufferLength=0;
	}
	s.m_nBufferLength = MyStrlen(string.m_pBuffer)+1+1;
	s.m_pBuffer = (char*)malloc(s.m_nBufferLength);//ch+'/0' = 1+1
	if(!s.m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	s.m_pBuffer[0]=ch;
	s.m_pBuffer[1]=0;
	MyStrcat(s.m_pBuffer,string.m_pBuffer);
	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString operator+(const CVBString& string, const char* lpsz)
{
	CVBString s;

	if(s.m_pBuffer != 0)
	{
		free(s.m_pBuffer);
		s.m_nBufferLength=0;
	}
	s.m_nBufferLength = MyStrlen(string.m_pBuffer)+MyStrlen(lpsz)+1;
	s.m_pBuffer = (char*)malloc(s.m_nBufferLength);
	if(!s.m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(s.m_pBuffer,string.m_pBuffer);
	MyStrcat(s.m_pBuffer,(char*)lpsz);
	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
CVBString operator+(const char* lpsz, const CVBString& string)
{
	CVBString s;

	if(s.m_pBuffer != 0)
	{
		free(s.m_pBuffer);
		s.m_nBufferLength=0;
	}
	s.m_nBufferLength = MyStrlen(string.m_pBuffer)+MyStrlen(lpsz)+1;
	s.m_pBuffer = (char*)malloc(s.m_nBufferLength);
	if(!s.m_pBuffer)
	{
		assert(0);//Out Of Memory!
	}
	MyStrcpy(s.m_pBuffer,lpsz);
	MyStrcat(s.m_pBuffer,string.m_pBuffer);
	return s;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
_inline CVBString::operator const char*() const
{ 
	return m_pBuffer;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
const CVBString& CVBString::operator+=(const CVBString& string)
{
	*this = *this+string;
	return *this;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
const CVBString& CVBString::operator+=(char ch)
{
	*this = *this+ch;
	return *this;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
const CVBString& CVBString::operator+=(const char* lpsz)
{
	*this = *this+lpsz;
	return *this;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char CVBString::operator[](int nIndex) const
{
	if((unsigned long int)nIndex>=m_nBufferLength)
	{
		assert(0);//ERROR! you are accessing memory not allocated for the string.
	}
	return m_pBuffer[nIndex];
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char CVBString::GetAt(int nIndex) const
{
	if((unsigned long int)nIndex>=m_nBufferLength)
	{
		assert(0);//ERROR! you are accessing memory not allocated for the string.
	}
	return m_pBuffer[nIndex];
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
void CVBString::SetAt(int nIndex, char ch)
{
	if((unsigned long int)nIndex>=m_nBufferLength)
	{
		assert(0);//ERROR! you are accessing memory not allocated for the string.
	}
	m_pBuffer[nIndex]=ch;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int MyStrcmp ( unsigned char *src , unsigned char *dst )
{
	int ret = 0 ;
	while( ! (ret = *src - *dst) && *dst)
		++src, ++dst;
	if ( ret < 0 )
		ret = -1 ;
	else if ( ret > 0 )
		ret = 1 ;
	return( ret );
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char* CVBString::GetBuffer(int nMinBufLength)
{
	CVBString s;
	if((unsigned long int)nMinBufLength>this->m_nBufferLength)
	{
		s=*this;

		if(this->m_pBuffer)
		{
			free(this->m_pBuffer);
		}
		
		this->m_nBufferLength=nMinBufLength;
		this->m_pBuffer=(char*)malloc(nMinBufLength);
		if(!this->m_pBuffer)
		{
			assert(0);//Out Of Memory!
		}
		MyStrcpy(this->m_pBuffer,s.m_pBuffer);
	}
	return this->m_pBuffer;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
bool CVBString::operator==(const char* lpsz)
{
	if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)lpsz)==0)
		return true;
	else 
		return false;

	return false;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
bool CVBString::operator==(CVBString string)
{
	if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)string.m_pBuffer)==0)
		return true;
	else 
		return false;

	return false;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
bool CVBString::operator!=(const char* lpsz)
{
	if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)lpsz)==0)
		return false;
	else 
		return true;

	return false;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
bool CVBString::operator!=(CVBString string)
{
	if(MyStrcmp((unsigned char*)this->m_pBuffer,(unsigned char*)string.m_pBuffer)==0)
		return false;
	else 
		return true;

	return false;
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
char *MyStrncpy(char *dest,char *source,unsigned long int count)
{
	char *start = dest;

	while (count && (*dest++ = *source++))
		count--;
	if (count)
	while (--count)
		*dest++ = '\0';
	return(start);
}
//////////////////////////////////////////////////////////////////////////////////
//
//////////////////////////////////////////////////////////////////////////////////
int MyStrncmp (char *first,char *last,unsigned count)
{
	if (!count)
		return(0);
	while (--count && *first && *first == *last)
	{
		first++;
		last++;
	}
	return(*first - *last);
}
