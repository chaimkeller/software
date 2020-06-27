/**********************************************************************************************
* File Name: FileIO.h
*
* Description: FileIO header.
*
* Programmer Name: Shalom Keller.
*
**********************************************************************************************/


#ifndef __FILE_IO_H__
#define __FILE_IO_H__

#define IO_SUCESS 1
#define IO_FAIL 0

#define OPEN_MODE_NEW_FILE 1
#define OPEN_MODE_OPEN_FILE 2


int File_Open(char *szFileName, int OpenMode);
int File_Read(int ID, char *pBuffer, int nBytes);
int File_Write(int ID, char *pBuffer, int nBytes);
int File_GetSize(int ID, int *pnBytes);
int File_Close(int ID);
int File_Seek(int ID, int bytesFromBeginningOfFile);


#endif
