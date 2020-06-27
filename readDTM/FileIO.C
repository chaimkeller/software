/**********************************************************************************************
* File Name: FileIO.C
*
* Description: Implemets File IO.
*
* Programmer Name: Shalom Keller.
*
**********************************************************************************************/
//-------------------------------------------------------------------
//
//-------------------------------------------------------------------
#define _CRT_SECURE_NO_DEPRECATE	//For MS VC 2005
//-------------------------------------------------------------------
//
//-------------------------------------------------------------------

#include <Stdio.h>
#include "FileIO.h"

//-------------------------------------------------------------------
//
//-------------------------------------------------------------------
#ifdef __USE_C__

int File_Open(char *szFileName, int OpenMode)
{
	int ID;
	if(OpenMode == OPEN_MODE_NEW_FILE)
		ID = (int)fopen(szFileName,"w+");
	else
		ID = (int)fopen(szFileName,"r+");
	return ID;
}
int File_Read(int ID, char *pBuffer, int nBytes)
{
	int iRet = (int)fread(pBuffer,sizeof(char),nBytes,(FILE*)ID);
	if(iRet < nBytes)
	{
		return IO_FAIL;
	}

	return IO_SUCESS;
}
int File_Write(int ID, char *pBuffer, int nBytes)
{
	int iRet = (int)fwrite(pBuffer,sizeof(char),nBytes,(FILE*)ID);
	if(iRet < nBytes)
	{
		return IO_FAIL;
	}

	return IO_SUCESS;
}
int File_GetSize(int ID, int *pnBytes)
{
	int pos;
	int end;

	pos = ftell((FILE*)ID);
	if(fseek ((FILE*)ID, 0, SEEK_END)!=0)
	{
		printf("File_GetSize: fseek failled!\n");
		perror("Reason: ");
		return IO_FAIL;
	}
	end = ftell((FILE*)ID);
	if(fseek ((FILE*)ID, pos, SEEK_SET)!=0)
	{
		printf("File_GetSize: fseek failled!\n");
		perror("Reason: ");
		return IO_FAIL;
	}
	*pnBytes = end;

	return IO_SUCESS;
}
int File_Close(int ID)
{
	fclose((FILE*)ID);
	return IO_SUCESS;
}
int File_Seek(int ID, int bytesFromBeginningOfFile)
{
	int iRet = fseek((FILE*)ID,bytesFromBeginningOfFile,SEEK_SET );
	if(iRet!=0)
		return IO_FAIL;

	return IO_SUCESS;
}
#else

//Uses Windows API for File IO.
#include <Windows.h>

int File_Open(char *szFileName, int OpenMode)
{
	HANDLE hFile; 
	if(OpenMode == OPEN_MODE_NEW_FILE)
		hFile = CreateFileA(szFileName,GENERIC_READ|GENERIC_WRITE,FILE_SHARE_READ|FILE_SHARE_WRITE, 0,CREATE_ALWAYS,0,0 );
	else
		hFile = CreateFileA(szFileName,GENERIC_READ|GENERIC_WRITE,FILE_SHARE_READ|FILE_SHARE_WRITE, 0,OPEN_EXISTING,0,0 );
	if(hFile ==  INVALID_HANDLE_VALUE)
	{
		return 0;
	}
	return (int)hFile;
}
int File_Read(int ID, char *pBuffer, int nBytes)
{
	int iBytesRead = 0;
	DWORD error;
	if(!ReadFile((HANDLE)ID,pBuffer,nBytes,&iBytesRead,0))
		return IO_FAIL;

	if(iBytesRead != nBytes)
	{
		error = GetLastError();
		__asm __emit 0xCC
		File_GetSize(ID, &error);
		return IO_FAIL;
	}

	return IO_SUCESS;
}
int File_Write(int ID, char *pBuffer, int nBytes)
{
	int iBytesWrite = 0;
	if(!WriteFile((HANDLE)ID,pBuffer,nBytes,&iBytesWrite,0))
		return IO_FAIL;

	if(iBytesWrite != nBytes)
	{
		return IO_FAIL;
	}

	return IO_SUCESS;
}
int File_GetSize(int ID, int *pnBytes)
{
	int iSize = GetFileSize((HANDLE)ID,0);
	if(iSize == 0xFFFFFFFF)
	{
		return IO_FAIL;
	}
	*pnBytes = iSize;
	return IO_SUCESS;
}
int File_Close(int ID)
{
	CloseHandle((HANDLE)ID);
	return IO_SUCESS;
}
int File_Seek(int ID, int bytesFromBeginningOfFile)
{
	if(SetFilePointer((HANDLE)ID,bytesFromBeginningOfFile,0,FILE_BEGIN) == 0xFFFFFFFF)
		return IO_FAIL;
	return IO_SUCESS;
}
#endif