//VBString.H



#ifndef __VBString_h_
#define __VBString_h_


//_LPCTSTR = const char*
//_LPCWSTR = const short int*

class CVBString
{
	char *m_pBuffer;
	unsigned long int m_nBufferLength;
public:

	//--------------------------------------------
	CVBString();
	~CVBString();
	CVBString(const CVBString& stringSrc);
	CVBString(const char* lpsz);
	//--------------------------------------------
	const CVBString& operator=(const CVBString& stringSrc);
	const CVBString& operator=(const char* lpsz);
	//--------------------------------------------
	bool operator==(const char* lpsz);
	bool operator==(CVBString string);
	//--------------------------------------------
	bool operator!=(const char* lpsz);
	bool operator!=(CVBString string);
	//--------------------------------------------
	const CVBString& operator+=(const CVBString& string);
	const CVBString& operator+=(char ch);
	const CVBString& operator+=(const char* lpsz);
	//--------------------------------------------
	friend CVBString operator+(const CVBString& string1,const CVBString& string2);
	friend CVBString operator+(const CVBString& string, char ch);
	friend CVBString operator+(char ch, const CVBString& string);

	friend CVBString operator+(const CVBString& string, const char* lpsz);
	friend CVBString operator+(const char* lpsz, const CVBString& string);
	//--------------------------------------------
	operator const char*() const;
	//--------------------------------------------
	char operator[](int nIndex) const;
	//--------------------------------------------
	char GetAt(int nIndex) const;
	void SetAt(int nIndex, char ch);
	//--------------------------------------------
	int Compare(const char* lpsz) const;
	int CompareNoCase(const char* lpsz) const;
	//--------------------------------------------
	CVBString Mid(int nFirst, int nCount) const;
	CVBString Mid(int nFirst) const;
	CVBString Left(int nCount);
	CVBString Right(int nCount);
	//--------------------------------------------
	void UCase();
	void LCase();
	void Reverse();
	//--------------------------------------------
	void RTrim();
	void LTrim();
	void RTrim(char chTarget);
	void RTrim(const char* lpszTargets);
	void LTrim(char chTarget);
	void LTrim(const char* lpszTargets);
	void Trim(){RTrim();LTrim();}
	//--------------------------------------------
	int Replace(char chOld, char chNew);
	int Replace(const char* lpszOld, const char* lpszNew);
	int Remove(char chRemove);
	int Insert(int nIndex, char ch);
	int Insert(int nIndex, const char* pstr);
	int Delete(int nIndex, int nCount = 1);
	//--------------------------------------------
	int Find(char ch) const;
	int ReverseFind(char ch) const;
	int Find(char ch, int nStart) const;
	int Find(const char* lpszSub) const;
	int Find(const char* lpszSub, int nStart) const;
	//--------------------------------------------
	void __cdecl Format(const char* lpszFormat, ...);
	void FormatV(const char* lpszFormat, va_list argList);
	//--------------------------------------------
	char* GetBuffer(int nMinBufLength);
	//--------------------------------------------
	bool IsWhiteSpace(char ch);
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	int InStr(int nStart,const char* lpszSub);
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

protected:
};


int MyStrlen(const char *str);
char *MyStrcpy(char * dst, const char *src);
char *MyStrcat(char * dst, char * src);
int MyStrcmp ( unsigned char *src , unsigned char *dst );
char *MyStrncpy(char *dest,char *source,unsigned long int count);
int MyStrncmp (char *first,char *last,unsigned count);

void Testing();



#endif