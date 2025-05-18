// newreadDTMDlg.cpp : implementation file
//
//	Version 19 06112j0 - full implimatation of Wikipedia like terrestrial refraction model option
//
#define VC_EXTRALEAN

#include "stdafx.h"
#include "newreadDTM.h"
#include "newreadDTMDlg.h"
#include "math.h"
#include <sys/stat.h>
#include <io.h>
#include <stdlib.h>
#include <stdio.h>
#include <fcntl.h>
#include <conio.h>
#include <ctype.h>
#include <string.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CNewreadDTMDlg dialog

CNewreadDTMDlg::CNewreadDTMDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CNewreadDTMDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CNewreadDTMDlg)
	m_Label = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CNewreadDTMDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CNewreadDTMDlg)
	DDX_Control(pDX, IDC_STATIC_PROGRESS, m_NewLabel);
	DDX_Control(pDX, IDC_PROGRESS1, m_Progress);
	DDX_Text(pDX, IDC_STATIC_PROGRESS, m_Label);
	DDV_MaxChars(pDX, m_Label, 4);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CNewreadDTMDlg, CDialog)
	//{{AFX_MSG_MAP(CNewreadDTMDlg)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//////////////////////external modules///////////////////
short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] );
char *RemoveCRLF( char *str );
char *fgets_CR( char *str, short strsize, FILE *stream );
int InStr( short nstart, char * str, char * str2 );
int InStr( char * str, char * str2 );

/////////////////////////////////////////////////////////////////////////////
// CNewreadDTMDlg message handlers

///////////////////////////////PLEASE NOTE//////////////////////////////////////////////
//This program is called by Maps&More VB6 program.
//The program modules, flag files, and DTM directories must reside either of Windows Disk c or e
/////////////////////////////////////////////////////////////////////////////////////////////

char worlddtm[]="";  //drive letter where SRTM1 files are stored
char d3asdtm[]="";   //drive letter where SRTM3 - MERIT files are stored
char arosdtm[]="";	 //drive letter where ALOS bil files are stored
bool worlddtmcd;
char ramdrive[]="";
char usadrive[255]="";  //folder where SRTM1 tiles are stored
char tasdrive[266]="";  //folder where SRTM3 tiles are stored
char bildrive[255]="";  //folder where ALOS bil tiles are stored
char tmpfil[81];
//int DTMflag = 0; //default DTM = SRTM30
//bool IgnoreMissingTiles = false;
//short noVoid = 0; //used for removing radar shadow
//short calcProfile = 0; //=1 for calculating profile, = 0 for not calculating profile
//double AzimuthStep = 0.1; //step size in azimuth for horizon profiles


//declare function
int WhichJK();
int WhichJK_2();
int WhichJK_3();
//declare global variables used with the above functions
//WhereJK stores the eros.tm3 folder's letter number, WhereJK2 stores the folder's letter number for the mapcdinfo.sav file
//the numbers are from 0-13 corresponding to the testdrv letters below.
int WhereJK, WhereJK2, WhereJK3;
char testdrv[13] = "cdefghijklmn";
char drivlet[2] = "c";
char bufFileName[255] = "";
char doclin[255] = ""; //used for fget_CR reading of one line of file
char seps[]   = " ,\t\n"; //used for strtok parsing the above one line of the file


//int Profile( double kmyo,double kmxo, double hgt, double mang, double aprn,double nnetz,double modval, 
//		double lgo,double endlg, double beglt, double endlt);

////////////////////

BOOL CNewreadDTMDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	int nTimer, reply;

    //double L1, L2, hgt;
	//int ang = 45;
	//double aprn = 0.5;
    char st2[] = "g:\\E130S30.BIN";
	char s[81]="";
	FILE *stream;
	CString lpszText;
	short ier,MinT[12],AvgT[12],MaxT[12],iTK;
	char *token;

	//bool TimerSet = FALSE;


	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	m_Progress.SetRange (0,100);
	m_Progress.SetStep (1);
	m_progresspercent = 0;

   worldCD[1] = 1;
   worldCD[2] = 1;
   worldCD[3] = 1;
   worldCD[4] = 1;
   worldCD[5] = 3;
   worldCD[6] = 3;
   worldCD[7] = 3;
   worldCD[8] = 3;
   worldCD[9] = 3;
   worldCD[10] = 1;
   worldCD[11] = 1;
   worldCD[12] = 1;
   worldCD[13] = 2;
   worldCD[14] = 2;
   worldCD[15] = 2;
   worldCD[16] = 3;
   worldCD[17] = 3;
   worldCD[18] = 4;
   worldCD[19] = 4;
   worldCD[20] = 4;
   worldCD[21] = 2;
   worldCD[22] = 2;
   worldCD[23] = 2;
   worldCD[24] = 2;
   worldCD[25] = 4;
   worldCD[26] = 4;
   worldCD[27] = 4;
   worldCD[28] = 5;

   
   // read the worddtm file drive contained in the mapcdinfo.sav file
   //find out where that file resides and open it
   WhereJK2 = WhichJK_2();

   if (WhereJK2 != 0)
   {
	   drivlet[0] = testdrv[WhereJK2 - 1];
	   drivlet[1] = 0;
   }
   else
   {
		lpszText = "Can't find the file mapcdinfo.sav!";
		reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
		OnCancel();
   }

	sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\mapcdinfo.sav");
	if ( (stream = fopen( bufFileName, "r" )) != NULL )
	{
		fgets_CR(doclin, 255, stream);
		//parse the output
		token = strtok( doclin, seps );
		strcpy(worlddtm, token);
		token = strtok( NULL, seps );
		worlddtmcd = atoi(token);
		fgets_CR(doclin, 255, stream);
		fscanf( stream, "%s\n", &ramdrive );
		fclose( stream );
	}

	//read the /jk_c/mapSRTMinfo.sav file for location of the DEM-SRTM tiles
   WhereJK3 = WhichJK_3();

   if (WhereJK3 != 0)
   {
	   drivlet[0] = testdrv[WhereJK2 - 1];
	   drivlet[1] = 0;
   }
   else
   {
		lpszText = "Can't find the file mapSRTMinfo.sav!";
		reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
		OnCancel();
   }

	sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\mapSRTMinfo.sav");
	if ( (stream = fopen( bufFileName, "r" )) != NULL )
	{
		fgets_CR(doclin, 255, stream);
		//parse the output
		token = strtok( doclin, seps );
		strcpy(worlddtm, token);
		token = strtok( NULL, seps );
		worlddtmcd = atoi(token);
		fgets_CR(doclin, 255, stream); //USA directory for DTMflag = 1
		sprintf(usadrive, "%s:\\%s\\", worlddtm, doclin );
		fgets_CR(doclin, 255, stream); //3AS directory for DTMflag = 2
		token = strtok( doclin, seps );
		strcpy(d3asdtm, token);
		token = strtok( NULL, seps );
		sprintf(tasdrive, "%s:\\%s\\", d3asdtm,token );
		fgets_CR(doclin, 255, stream); //BIL directory for DTMflag = 3
		token = strtok( doclin, seps );
		strcpy(arosdtm, token);
		token = strtok( NULL, seps );
		sprintf(bildrive, "%s:\\%s\\", d3asdtm, token );
		fclose( stream );
	}

	
	//open eros.tm3 file and read beglat,endlat,etc.
w2000:

   WhereJK = WhichJK();
   if (WhereJK != 0)
   {
	   drivlet[0] = testdrv[WhereJK - 1];
	   drivlet[1] = 0;
	   sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\eros.tm3");
   }
   else
   {
		lpszText = "Can't find the file eros.tm3!";
		reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
		OnCancel();
   }

    stream = fopen( bufFileName, "r" );
 	fscanf( stream, "%s", st2);
	fscanf( stream, "%lg,%lg,%lg,%d,%lg,%d,%f\n", &lat0,&lon0,&hgt0,&ang,&aprn,&mode,&modeval );
	fscanf( stream, "%lg,%lg,%lg,%lg", &beglog,&endlog,&beglat,&endlat );
	fscanf( stream, "%d,%d\n", &noVoid, &calcProfile); //radar shadow interpolation
	//newer versions
	try
	{
		fscanf( stream, "%lg\n", &AzimuthStep);
		fscanf( stream, "%d\n", &IgnoreMissingTiles);
		fscanf( stream, "%d\n", &TemperatureModel);
		fscanf( stream, "%lg\n", &Tground); //degrees Celsius
		fscanf( stream, "%lg\n", &treehgt); //meterse

		if (Tground == 0) {
			ier = Temperatures(lat0,lon0,MinT,AvgT,MaxT);

			if (ier == 0) {

				//use average temperature for calculation of terrestrial refraction
				for (iTK = 0; iTK < 12; iTK++) {
					Tground += AvgT[iTK];
				}
				Tground /= 12.0; //take average over year

				if (mode >= 4) {
					//both horizons, average over both
					for (iTK = 0; iTK < 12; iTK++) {
						Tground += MinT[iTK];
						Tground += AvgT[iTK];
					}
					Tground /= 24.0; //take average over year

				}else if (mode >= 1) {
					//sunrise horizon, use minimum temperatures
					for (iTK = 0; iTK < 12; iTK++) {
						Tground += MinT[iTK];
					}
					Tground /= 12.0; //take average over year
				}else if (mode <= 0) {
					//sunset horizon, use average temperatures
					for (iTK = 0; iTK < 12; iTK++) {
						Tground += AvgT[iTK];
					}
					Tground /= 12.0; //take average over year
				}
				
				Tground += 273.15; //convert to deg. Kelvin

			}else{
				//failed to load temperature files, use default
				CString lpszText = "Can't find or read the WorldClim temperature files!\n\nUsing default mean temp. of 15 deg C instead.";
				reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );
				Tground = 288.15; //default ground temperature = 15 deg C
			}
		}else{
			Tground += 273.15; //convert to deg. Kelvin
		}

		//calculate conserved portion of terrain refraction based on Wikipedia's expression
		//lapse rate is set at -0.0065 K/m, Ground Pressure = 1013.25 mb
		TRpart = 8.15 * 1013.25 * 1000.0 * (0.0342 - 0.0065)/(Tground * Tground * 3600); //units of deg/km
	}
	catch( char *ErrorStr )
	{
		//do nothing
	}

	fclose( stream );
	
	/*
	if (calcProfile == 1) //when calculating profiles, ignore missing tiles
	{
		IgnoreMissingTiles = true;
	}
	*/

	// check for eros.tm4 file
	strncpy( tmpfil, (const char *)ramdrive, 1);
	strncpy( tmpfil + 1, ":\\eros.tm4", 11 );
	if ( (stream = fopen(tmpfil, "r" )) != 0 )
		{
		directx = TRUE;
		fscanf( stream, "%lg\n", &lat0 );
		fscanf( stream, "%lg\n", &lon0 );
		fscanf( stream, "%lg\n", &beglog );
		fscanf( stream, "%lg\n", &endlog );
		fscanf( stream, "%lg\n", &endlat );
		fscanf( stream, "%lg\n", &beglat );
		fscanf( stream, "%i\n", &landflag );
		fclose( stream );
		}
	else
		{
		directx = FALSE;
		}

 	// Check if eros.tm6 file exists
	// It contains the integer, DTMflag, the flag of DTM type, where
	// DTMflag =
	//         = 0 for GTOPO30 (1000 meter) -- default
	//         = 1 for SRTM-1 DEM (30 meter)
	//         = 2 for SRTM-2/3 MERIT DEM (90 meter)
	//		   = 3 for ALOS DEM (30 meter)
	// also read in the hard drive letter of the SRTM worlddtm
	strncpy( tmpfil, (const char *)ramdrive, 1);
	strncpy( tmpfil + 1, ":\\eros.tm6", 11 );
	if ( (stream = fopen(tmpfil, "r" )) != 0 )
		{
		fscanf( stream, "%1s, %d\n", &worlddtm, &DTMflag);
		fclose( stream );
		}
	else
	{
		reply = AfxMessageBox( "Can't find the file eros.tm6 in the jk_c folder!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );
		OnCancel();
	}

	if (DTMflag == 1) //SRTM1 DEM
	{
		//sprintf(usadrive, "%s%s%s%s", worlddtm, ":\\", doclin, "\\" );
		//worlddtm already defined
		;
	}
	else if (DTMflag == 2) //SRTM3-MERIT DEM
	{
		//sprintf(tasdrive, "%s%s%s%s", worlddtm, ":\\", doclin, "\\" );
		worlddtm[0] = 0;
		strcpy(worlddtm, d3asdtm );
	}
	else if (DTMflag == 3) //ALOS DEM
	{
		//sprintf(bildrive, "%s%s%s%s", worlddtm, ":\\", doclin, "\\" );
		worlddtm[0] = 0;
		strcpy(worlddtm, arosdtm );
	}

    //usadrive contains full path to the DEM-SRTM files which can now be any folder: EK 061322

	/****************************test mode*********/
	//strncpy( worlddtm, "d", 1 );
	//DTMflag = 2;
	/***************************************/

	// Check if eros.tm7 file exists
	// This flags the program to automatically use default tiles
	// whenever a missing tile is detected
	int skipflag;
	//IgnoreMissingTiles = false;
	strncpy( tmpfil, (const char *)ramdrive, 1);
	strncpy( tmpfil + 1, ":\\eros.tm7", 11 );
	if ( (stream = fopen(tmpfil, "r" )) != 0 )
		{
		fscanf( stream, "%d\n", &skipflag);
		fclose( stream );
		IgnoreMissingTiles = 1;
		}

	/* create timer to signal the program to read the DTM*/
	TimerSet = FALSE;
	if (TimerSet == FALSE)
		{
		nTimer = SetTimer(1, 100, NULL ); /* 1/10 of second*/
		}
	if (nTimer != 0)
		{
		TimerSet = TRUE; /*flag the program that a timer is set*/
		}

	return TRUE;  // return TRUE  unless you set the focus to a control
	
}	

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CNewreadDTMDlg::OnPaint() 
{
    if (directx)
	   CWnd::ShowWindow(SW_HIDE);

	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CNewreadDTMDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CNewreadDTMDlg::OnOK() 
{
	CDialog::OnOK();
}

///////////global arrays and variables///////////////////////

union unich //big endian short int heights read from DTM files
	{
	short int i5;
	unsigned char chi5[2]; //break up the short int into its two bytes
	} uni5_tmp[1];

union unich *uni5;

union i4switch //used to rotating little endian to big endian
	{
	short int i4s;
	unsigned char chi6[2]; //two bytes for a short integer
	} i4sw[1];

union ver
{
	float vert[3];
} ver_tmp[1];

union ver *vertices;

union fac
{
	int face[4];
} fac_tmp[1];

union fac *faces;

union hgt
{
	short int height;
} hgt_tmp[1];

union hgt *heights;

union profile
{
	double ver[3];
}prof_tmp[1];

union profile *prof;

union dtm
{
	double ver[3];
}dtm_tmp[1];

union dtm *dtms;
	


//delcare and allocate a minimum memory for other arrays
int *lg1chk = new int[1];
int *lt1chk = new int[1];
double *kmxb1 = new double[1];
double *kmxb2 = new double[1];
double *kmyb1 = new double[1];
double *kmyb2 = new double[1];
long *numstepkmyt = new long[1];
long *numstepkmxt = new long[1];
char (*filt1)[8] = new char[1][8];

//////////////////////////////////////////////////////////////

//bool manytiles, record, 
bool poles;
short numtiles;
short whichCDs[4];
double ntmp, tmp, nx;
int ny, lg1, lg11, lt1, lt11;

int totnumtiles;
int numtotrec;

int lg2, lt2, oldpercent;
double xdim,ydim;
int nrows, ncols;
int gototag, reply, rep;
long i,j,ii,jj,kk,ikmy,ikmx,itnum;
long iikmy,iikmx;
long maxrange,maxranget;
double kmx,kmy;
long pos, numrc, numrec, numCD, tmprc, numrctot;
long numrc2; //used for removing radar shadow
long numstepkmx, numstepkmy,numx,numy;
double kmxo,kmyo;
short int i4,i4min,i4max;
short int i41x,i42x,i41y,i42y; //used for removing radar shadow
long i41xt,i42xt,i41yt,i42yt; //used for removing radar shadow
double wtxd,wtyd,w4dtot; //used for removing radar shadow
double num1,num2,num3,num4; //,testx,testz;
double fracx,fracy;
long numface, nskip; //newsize[31],
long  charout, nverts, numfracx1, numfracx2, numfracy1, numfracy2;
char st2[16];
char quest[81];
char buffer[5];
char CDbuffer[1];
char filt[8];
char filtmp[8];
char filn[255] = "";  //EK 061322  instead of fixed length of 18, can now be any folder of any string length
char filnn[255] = "";
char binfil[18];
char DTMfile[255] = "";
char lg1ch[3];
char lt1ch[2];
char lch[2];
char EW[1];
char NS[1];

/////declare functions/////////

long Cint(double x);
int Nint(double x);
int WhichJK(); 
int WhichJK_2();

////////////////////////////////

 void CNewreadDTMDlg::OnTimer(UINT nIDEvent) 
{

	unsigned int bytesread,byteswritten;
	MSG Message;

	//local variables for profiles
    double pi = acos(-1.);
    double cd = pi / 180;
    int nkmx1, ikmx, ikmy;
    short minview = 10;
	double dstpointx;
	double dstpointy;
	double mang;
	double maxang = 0, maxang0 = 0;
	short nomentr = 0;
	double apprnr, nnetz;
	bool netz = false, found = false;
    short numskip = 0;

	double startkmx;
    double sofkmx;
    double lto;
	int numkmy, numhalfkmy, numdtm;
	//int numcol = 0,numrow = 0;
	short optang = 0;
	double kmyoo, kmxoo, lt, lg;
	double hgtobs = 1.8f;
	double anav = 0;
	double hgt = 0;
	double begkmx, endkmx, skipkmx, skipkmy;
	int nbegkmx, numkmx, i__, i__1, i__2, i__3, i__4, nk;
	double re, rekm, range, kmx, kmy, dkmy, testazi;
	float hgt2, fudge = 0, xentry;
	double lg2, lt2, x1, x2, y1, y2, z1, z2, d__1, d__2, d__3;
	double distd, re1, re2, deltd, dist1, dist2, angle, viewang;
	double d__, x1d, y1d, z1d, azicos, x1s, y1s, z1s, x1p, y1p, z1p, azisin, azi;
	double defm, defb;
	double avref, kmxp, kmyp, bne1, bne2;
	int istart1, ndela, ndelazi, ne, nstart2;
	double az2,vang2,kmy2,az1,vang1,kmy1,bne, vang, proftmp, delazi, dist;
	int invAzimuthStep = 10;
	double PATHLENGTH = 0.0; //Pathlength from observer to obstruction
	double RETH = 6356.766; //mean radius of Earth in kms
	double Exponent = 1.0;
	bool AzimuthOpt = false;


	FILE* f;

	invAzimuthStep = (int)(1/AzimuthStep);
	AzimuthStep = 1.f / (double)invAzimuthStep;

	xdim = 0.008333333333333; //GTOPO30 -- the default
	ydim = 0.008333333333333;

	if (DTMflag == 1 || DTMflag == 3) //SRTM-1 (1 arcsec/30 meter DTM), or ALOS DEM (1 arcsec/30 meters)
	{
		xdim = xdim/30.0;
		ydim = ydim/30.0;
	}
	else if (DTMflag == 2) //SRTM-2/3 or MERIT DEM (3 arcsec/100 meter DTM)
	{
		xdim = xdim * 0.1;
		ydim = ydim * 0.1;
	}

	#define TOKEN_NAME 1 //0x0001
	#define TOKEN_STRING 2 //0x0002
	#define TOKEN_INTEGER 3 //0x0003
	#define TOKEN_GUID 5 //0x00000005
	#define TOKEN_INTEGER_LIST 6 //0x0006
	#define TOKEN_REALNUM_LIST 7 //0x0007  
	#define TOKEN_OBRACE 10 //0x000A  //10
	#define TOKEN_CBRACE 11 //0x000B  //11
	#define TOKEN_OPAREN 12 //0x000C  //12
	#define TOKEN_CPAREN 13 //0x000D  //13
	#define TOKEN_OBRACKET 14 //0x000E  //14
	#define TOKEN_CBRACKET 15 //0x000F  //15
	#define TOKEN_OANGLE 16  //0x0010  //16
	#define TOKEN_CANGLE 17  //0x0011  //17
	#define TOKEN_DOT 18  //0x0012  //18 
	#define TOKEN_COMMA 19  //0x0013   //19
	#define TOKEN_SEMICOLON 20  //0x0014  //20
	#define TOKEN_TEMPLATE 31  //0x001F  //31
	#define TOKEN_WORD 40  //0x0028  //40
	#define TOKEN_DWORD 41  //0x0029  //41
	#define TOKEN_FLOAT 42  //0x002A //42
	#define TOKEN_DOUBLE 43  //0x002B  //43
	#define TOKEN_CHAR 44  //0x002C  //44
	#define TOKEN_UCHAR 45  //0x002D  //45
	#define TOKEN_SWORD 46  //0x002E  //46
	#define TOKEN_SDWORD 47  //0x002F  //47
	#define TOKEN_VOID 48  //0x0030  //48
	#define TOKEN_LPSTR 49  //0x0031  //49
	#define TOKEN_UNICODE 50  //0x0032  //50
	#define TOKEN_CSTRING 51  //0x0033  //51
	#define TOKEN_ARRAY 52  //0x0034  //52

	short sizeofit[31];
	char headbuffer[] = "xof 0302bin 0032";
	float materials[11];
	double lim1,lim2,lim3,lim4,dlim1,dlim2;
	FILE *stream3;//, *stream4;

	sizeofit[0] = TOKEN_NAME;
	sizeofit[1] = TOKEN_STRING;
	sizeofit[2] = TOKEN_INTEGER;
	sizeofit[3] = TOKEN_GUID;
	sizeofit[4] = TOKEN_INTEGER_LIST;
	sizeofit[5] = TOKEN_REALNUM_LIST;
	sizeofit[6] = TOKEN_OBRACE;
	sizeofit[7] = TOKEN_CBRACE;
	sizeofit[8] = TOKEN_OPAREN;
	sizeofit[9] = TOKEN_CPAREN;
	sizeofit[10] = TOKEN_OBRACKET;
	sizeofit[11] = TOKEN_CBRACKET;
	sizeofit[12] = TOKEN_OANGLE;
	sizeofit[13] = TOKEN_CANGLE;
	sizeofit[14] = TOKEN_DOT;
	sizeofit[15] = TOKEN_COMMA;
	sizeofit[16] = TOKEN_SEMICOLON;
	sizeofit[17] = TOKEN_TEMPLATE;
	sizeofit[18] = TOKEN_WORD;
	sizeofit[19] = TOKEN_DWORD;
	sizeofit[20] = TOKEN_FLOAT;
	sizeofit[21] = TOKEN_DOUBLE;
	sizeofit[22] = TOKEN_CHAR;
	sizeofit[23] = TOKEN_UCHAR;
	sizeofit[24] = TOKEN_SWORD;
	sizeofit[25] = TOKEN_SDWORD;
	sizeofit[26] = TOKEN_VOID;
	sizeofit[27] = TOKEN_LPSTR;
	sizeofit[28] = TOKEN_UNICODE;
	sizeofit[29] = TOKEN_CSTRING;
	sizeofit[30] = TOKEN_ARRAY;

    if (directx)
	   CWnd::ShowWindow(SW_HIDE);


	switch ( abs( mode ) ) // determine size of DirectX output file
		{
		case 0:  // largest sunset file, looking only east.
			fracx = 1.6;
			fracy = 1.6;
			break;
		case 1:  // largest sunrise file, no west included.
			fracx = 1.6;
			fracy = 1.6;
			break;
		case 2:  // medium sunset/sunrise files, no east included.
			fracx = 2.8;//3.2;
			fracy = 2.8;//3.2;
			break;
		case 3:  // smallest sunset/sunrise files, no west/east included. 
			fracx = 6.4;
			fracy = 6.4;
			break;
		case 4:  // all-directional file 
			fracx = 2.8;//1.2;//1.4;//2.8;//3.2;
			fracy = 2.8;//1.2;//1.4;//2.8;//3.2;
			break;
		}
	if ( modeval != 0 )  // then use this value for fracx,fracy
		{
		fracx = (double)modeval;
		fracy = (double)modeval;
		}

	if ( directx ) // then the eros.tm4 file overides the above values
				   // and it uses the original boundaries	
		{
		fracx = (double)modeval; //1.0; //(double)modeval; // 1.4;
		fracy = (double)modeval; //1.0; //(double)modeval; //1.4;
		mode = 4; // calculate for all directions
		}

	
	if (TimerSet == TRUE )
		{
		if (!directx)
			{
		    rep = ::SetWindowPos( CNewreadDTMDlg::m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE);
			}
		else
			{
		    rep = ::SetWindowPos( CNewreadDTMDlg::m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_HIDEWINDOW );
			}
		TimerSet = FALSE;
		KillTimer( nIDEvent );
		}

	//determine number of columns of rows to be extracted
	numstepkmy = Nint(fabs(beglat - endlat) / ydim) + 1;
	numstepkmx = Nint(fabs(endlog - beglog) / xdim) + 1;
	//will need to allocate maxrange = numstepkmx * numstepkmy * 4 byte integer memory
	//(but add five columns and rows as a cushion)
	maxrange = (numstepkmy + 5) * (numstepkmx + 5);
	numtotrec = numstepkmy * numstepkmx;

	//allocate memory for tile names and boundaries
    int lg1SRTM,lg2SRTM,lt1SRTM,lt2SRTM;
	if (DTMflag == 0) //GTOPO30
	{
		totnumtiles = 4;
	}
	else //SRTM
	{
	  lg1SRTM = (int)beglog;
	  if (beglog < 0 && (beglog < lg1SRTM) ) lg1SRTM -= 1;
	  lg2SRTM = (int)endlog;
	  if (endlog < 0 && (endlog < lg2SRTM) ) lg2SRTM -= 1;
	  lt1SRTM = (int)endlat;
	  if (endlat < 0 && (endlat < lt1SRTM) ) lt1SRTM -= 1;
	  lt2SRTM = (int)beglat;
	  if (beglat < 0 && (beglat < lt2SRTM) ) lt2SRTM -= 1;
	  
	  //number of tiles
	  totnumtiles = (abs(lg2SRTM - lg1SRTM) + 1) * (abs(lt2SRTM - lt1SRTM) + 1);

	}
	//deallocate memory for tile boundaries and names
	//and then allocate the required amount
	delete [] lg1chk;
	delete [] lt1chk;
	delete [] kmxb1;
	delete [] kmxb2;
	delete [] kmyb1;
	delete [] kmyb2;
	delete [] numstepkmyt;
	delete [] numstepkmxt;
	delete [] filt1;
	int *lg1chk = new int[totnumtiles];
	int *lt1chk = new int[totnumtiles];
	double *kmxb1 = new double[totnumtiles];
	double *kmxb2 = new double[totnumtiles];
	double *kmyb1 = new double[totnumtiles];
    double *kmyb2 = new double[totnumtiles];
	long *numstepkmyt = new long[totnumtiles];
	long *numstepkmxt = new long[totnumtiles];
	char (*filt1)[8] = new char[totnumtiles][8];


		/* now readin eros.tm3 and extract the DTM
 		as extracting DTM, advance the progress bar
		if user interrupts extraction with Cancel, then
		erase what was already extracted in the cancel button handler.*/

		// determine if need to read multiple tiles
		gototag = 1;
		goto findtile;
		
		// now extract the file name of the BIN file
		// to reside on the g: RAMDRIVE
ret1:   if (WhereJK != 0) 
		{
			drivlet[0] = testdrv[WhereJK - 1];
			drivlet[1] = 0;
			bufFileName[0] = 0; //zero string
			sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\eros.tm3" );
		}
		else
		{
			reply = AfxMessageBox( "Can't find the file eros.tm3 in the jk_c folder!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
			delete [] lg1chk;
			delete [] lt1chk;
			delete [] kmxb1;
			delete [] kmxb2;
			delete [] kmyb1;
			delete [] kmyb2;
			delete [] numstepkmyt;
			delete [] numstepkmxt;
			delete [] filt1;
			OnCancel();
		}
	    stream = fopen( bufFileName, "r" );
		fscanf( stream, "%s", st2);
		fclose( stream );

		//now open the .BIN file (first erase old version if exists)
		if (!(directx))
			{
			if ((stream2 = fopen( (const char*)st2, "rb" )) != NULL)
				{
				fclose( stream2 );
				DeleteFile( st2 );
				}

			stream2 = fopen( (const char *)st2, "wb" );
			if ( stream2 == NULL )
				{
				reply = AfxMessageBox( "Can't open the ramdrive:\\*.BIN file!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
				//deallocate memory
				delete [] lg1chk;
				delete [] lt1chk;
				delete [] kmxb1;
				delete [] kmxb2;
				delete [] kmyb1;
				delete [] kmyb2;
				delete [] numstepkmyt;
				delete [] numstepkmxt;
				delete [] filt1;
				OnCancel();
				}
			}
		
		//now open the .DEM file
		gototag = 1;
		goto dtmfiles;

ret2:	;
		
		if (directx) 
	    // open BI1 file which will contain height info for 3D Viewer file
		{
		strncpy( binfil, (const char *)st2, 10 );//strncpy( binfil, (const char *)DTMfile, 10 );
		strncpy( binfil + 10, ".BI1", 4 );
		strncpy( binfil + 14, "\0", 1 );

		if ((stream3 = fopen( (const char*)binfil, "rb" )) != NULL)
			{
			fclose( stream3 );
			DeleteFile( binfil );
			}

		stream3 = fopen( (const char *)binfil, "wb" );
		if ( stream3 == NULL )
			{
			reply = AfxMessageBox( "Can't open the ramdrive:\\*.BI1 file!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		
			//deallocate memory
			delete [] lg1chk;
			delete [] lt1chk;
			delete [] kmxb1;
			delete [] kmxb2;
			delete [] kmyb1;
			delete [] kmyb2;
			delete [] numstepkmyt;
			delete [] numstepkmxt;
			delete [] filt1;
			OnCancel();
			}
		}

		
		//begin writing to the BIN file

		numfracy2 = (int)(numstepkmy * 0.5 * (1 + 1.0/fracy));
		numfracy1 = (int)(numstepkmy * 0.5 * (1 - 1.0/fracy));

		if ( abs( mode ) == 4 )
			{
			numfracx2 = (int)(numstepkmx * 0.5 * (1 + 1.0/fracx));
			numfracx1 = (int)(numstepkmx * 0.5 * (1 - 1.0/fracx));
			}
		else
			{
			if ( mode >= 1 )  // sunrise
				{
				numfracx1 = (int)(0.1 / xdim);
				numfracx2 = numfracx1 + (int)(numstepkmx * (1.0/fracx));
				if ( numfracx2 > numstepkmx )
					{
					numfracx2 = numstepkmx;
					}
				}
			else if ( mode <= 0 )  // sunset
				{
				numfracx1 = (int)(numstepkmx * (1 - 1.0/fracx) - 0.1 / xdim);
				if (numfracx1 < 0 )
					{
					numfracx1 = 0;
					}
				numfracx2 = (int)(numstepkmx  - 0.1 / xdim);
				}
			}

		num1 =  (double)numstepkmy / ((double)(numfracy2 - numfracy1));
		num2 = 	0.002 / ((double)(numfracx2 - numfracx1 + numfracy2-numfracy1)); 
		num3 =	(double)numstepkmx / ((double)(numfracx2 - numfracx1));
		if ( mode >= 1 )
			{
			num4 =	(double)(beglog + xdim * numfracx1);
			}
		else if ( mode <= 0 )
			{
			num4 =	(double)(beglog + xdim * numfracx2);
			}
		
		nverts = (numfracy2 - numfracy1) * (numfracx2 - numfracx1);
		

	//allocate memory for vertices and faces arrays (add a little bit more to be save = 5)
	//these arrays are used to generate the landscape file 
	if (directx)
	//using these arrays, so redeclare them with a full allocation of memory
	{
		free( vertices );
        vertices = (union ver *) malloc((nverts + 5) * 3 * sizeof(union ver));
		free( faces );
        faces = (union fac *) malloc((nverts + 5) * 3 * sizeof(union fac));
		free( heights );
        heights = (union hgt *) malloc((nverts + 5) * 3 * sizeof(union hgt));
	}

	//allocate memory for main data array (add a little more just to be safe)
	//union unich *uni5;
	free( uni5 ); //free the memory allocated before for uni1, and reallocate
    uni5 = (union unich *) malloc(maxrange * sizeof(union unich));

	//check that memory was properly allocated to all the arrays
	if ( (uni5 == NULL) || ( (vertices == NULL || faces == NULL || heights == NULL) && (directx == TRUE) ) )
	{
		CString lpszText = "Can't allocate enough memory!";
		reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

		//deallocate memory
		delete [] lg1chk;
		delete [] lt1chk;
		delete [] kmxb1;
		delete [] kmxb2;
		delete [] kmyb1;
		delete [] kmyb2;
		delete [] numstepkmyt;
		delete [] numstepkmxt;
		delete [] filt1;
		OnCancel();
	}

	short iter;
	long nvert;
	numrctot = 0;
	oldpercent = 0;

    for ( iter = 0; iter < totnumtiles ; ++iter )
		{
		if (iter != 0) // open new tile
			{
			// check that this is first time to process this tile
			// need to open new tile, so close old tile
			fclose ( stream );
			if (DTMflag == 0) //GTOPO30 data
			{
				strncpy( filt, (const char *)filt1[iter], 7); //18 );
				//new string manipulation to allow for tiles existing in sub directories 02/06/24
				/*
				strncpy( filn, (const char *)worlddtm, 1);
				strncpy( filn + 1, ":\\", 2 );
				strncpy( filn + 3, (const char *)filt1[iter], 7 );
				strncpy( filn + 10, "\\", 1 );
				strncpy( filn + 11, (const char *)filt1[iter], 7 );
				*/
				sprintf(filn, "%s:\\:%s\\%s", worlddtm, filt1[iter], filt1[iter]);
				numCD = whichCDs[iter];
			}
			else if (DTMflag == 1) //SRTM1 data
			{
				filn[0] = 0;
				strncpy( filt, (const char *)filt1[iter], 7 );
				sprintf( filn, "%s%s", usadrive, filt );
				//strncpy( filn + 7, (const char *)filt1[iter], 7);
			}
			else if (DTMflag == 2) //SRTM3 data
			{
				filn[0] = 0;
				strncpy( filt, (const char *)filt1[iter], 7 );
				sprintf( filn, "%s%s", tasdrive, filt );
			}
			else if (DTMflag == 3) //ALOS data in bil form
			{
				filn[0] = 0;
				strncpy( filt, (const char *)filt1[iter], 7 );
				sprintf( filn, "%s%s", bildrive, filt );
			}

			gototag = 2;
			tmprc = numrc;
			goto dtmfiles;
ret4:		;	//manytiles = TRUE;
			}
			
		kmyo = -2000;
		kmxo = -2000;
		for ( ii = 1; ii <= numstepkmyt[iter]; ii++)
			{
			
			/* the following is the C++ version of VB's "Doevents" */
			/*  see p. 205 of D. Kruglinski's book */
			if (::PeekMessage(&Message, NULL, 0, 0, PM_REMOVE))
				{
				::TranslateMessage(&Message);
				::DispatchMessage(&Message);
				}

			kmy = kmyb1[iter] - ydim * (double)(ii - 1);
			i = Nint((endlat - kmy)/ydim) + 1;
			double kmyff;
			kmyff = kmy;
			poles = FALSE;
			if ( kmy < -90 )
				{
				kmyff = -180 - kmy;
				poles = TRUE;
				}
			else if ( kmy > 90 )
				{
				kmyff = 180 - kmy;
				poles = TRUE;
				}
			
			ikmy = Nint((lt1chk[iter] - kmyff) / ydim) + 1;

			{
				kmx = kmxb1[iter]; //beglog;  // for polar regions manytiles = TRUE
				//ikmx = (int)(floor((kmx - lg1) / xdim) + 1);
				ikmx = Nint((kmx - lg1chk[iter]) / xdim) + 1;
				numrec = (ikmy - 1) * ncols + ikmx - 1;
				//position DEM file at starting coordinate for reading
				pos = fseek( stream, numrec * 2L, SEEK_SET );
				//determine beginning record number of BIN file to write to
				iikmx = Nint((kmx - beglog)/xdim) + 1;
				iikmy = Nint((endlat - kmy)/ydim) + 1;
				numrc = (iikmy - 1) * numstepkmx + iikmx - 1;
				bytesread = fread( &uni5[numrc + 1].i5, numstepkmxt[iter] * 2L, 1, stream );

				numrctot += numstepkmxt[iter];
			}
			m_progresspercent = (int)((numrctot /  (float)maxranget) * 100);
			if ( m_progresspercent != oldpercent )
				{
				oldpercent = m_progresspercent;
				m_Progress.StepIt (); /*  step the progress bar */
				itoa(m_progresspercent, buffer, 10); /* convert m_prog... to char */
				strcat( buffer, "%");
				m_Label = buffer;    /* update the value of the progress bar */
				m_NewLabel.SetWindowText (m_Label);
				}
			}
		}
		fclose( stream );
//		iter -= 1;


		//write the vertices to the land.x file  and or
		//smooth radar shadow voids
		if ( directx || noVoid == 1 )
			{
			////////////////////////////////////////////////////////

      ;    /* notify the user of what is going on by the form's caption */
			m_Progress.SetPos (0);
			m_progresspercent = 0;
			m_Label = _T("");
			m_NewLabel.SetWindowText (m_Label);

			if (noVoid == 1 && DTMflag > 0) //change the caption
				{
				SetWindowText("Calculating vertices and/or filling voids");
				}


			long nver;
			nvert = 0;
			numy = 0;
			numx = 0;
			i4min = 20000;
			i4max = -20000;

			numrctot = 0;
			oldpercent = 0;
			for ( iter = 0; iter < totnumtiles ; ++iter )
				{
					
				kmyo = -2000;
				kmxo = -2000;

				for ( ii = 1; ii <= numstepkmyt[iter]; ii++)
					{
					
					/* the following is the C++ version of VB's "Doevents" */
					/*  see p. 205 of D. Kruglinski's book */
					if (::PeekMessage(&Message, NULL, 0, 0, PM_REMOVE))
						{
						::TranslateMessage(&Message);
						::DispatchMessage(&Message);
						}

					kmy = kmyb1[iter] - ydim * (double)(ii - 1);
					i = Nint((endlat - kmy)/ydim) + 1;
					double kmyff;
					kmyff = kmy;
					poles = FALSE;
					if ( kmy < -90 )
						{
						kmyff = -180 - kmy;
						poles = TRUE;
						}
					else if ( kmy > 90 )
						{
						kmyff = 180 - kmy;
						poles = TRUE;
						}
					
					ikmy = Nint((lt1chk[iter] - kmyff) / ydim) + 1;

					kmx = kmxb1[iter]; //beglog;  // for polar regions manytiles = TRUE
					ikmx = Nint((kmx - lg1chk[iter]) / xdim) + 1;
					numrec = (ikmy - 1) * ncols + ikmx - 1;
					//determine beginning record number of BIN file to write to
					iikmx = Nint((kmx - beglog)/xdim) + 1;
					iikmy = Nint((endlat - kmy)/ydim) + 1;
					numrc = (iikmy - 1) * numstepkmx + iikmx - 1;

					for ( jj = 1; jj <= numstepkmxt[iter]; jj++ )
						{
                		kmx = kmxb1[iter] + xdim * (double)(jj - 1);
						j = Nint((kmx - beglog)/xdim) + 1;
						if ((( i >= numfracy1 + 1) && ( i <= numfracy2 )) && (( j >= numfracx1 + 1) && (j <= numfracx2 )))
							{
							//vertices must be in form of rows of kmx vs kmy, so determine position in ordered array
							nver = (j - numfracx1 - 1) + (numfracx2 - numfracx1) * (i - numfracy1 - 1);

							if ( kmyo != kmy )
								{
								kmyo = kmy;
								numy++;
								}
							if ( kmx > kmxo )
								{
								kmxo = kmx;
								numx++;
								}

							// vertices require little endian representation
							i4 = (short int)(uni5[numrc + jj].chi5[0] << 8 | uni5[numrc + jj].chi5[1]);
							if ( DTMflag == 0 && i4 == -9999 ) // the ocean
								{
								i4 = 0; 
								}
							if ( DTMflag > 0 && (i4 == -32768  || i4 < -500)) // SRTM radar shadow/void
								{
								i4 = 0;  //if not removing SRTM void, set them to hgt = 0
								if (noVoid == 1) //remove SRTM void by linear interpolation
									{
									goto NoVoids;
unVoid:							    ;
								    //rotate back to big endian representation
								    //and write fixed value to uni5
								    i4sw[0].i4s = i4;
									uni5[numrc + jj].chi5[0] = i4sw[0].chi6[1];
									uni5[numrc + jj].chi5[1] = i4sw[0].chi6[0];
									}
								}

							if ( i4 < i4min )
								{
								i4min = i4;
								}
							if ( i4 > i4max )
								{
								i4max = i4;
								}

							if (directx) //record vertices
								{
								heights[nver].height = i4;
								vertices[nver].vert[1] = (float)(i4 * num2); //convert heights to normalized scale 
								if ( abs( mode ) == 4 )
									{
									vertices[nver].vert[2] = (float)((2 * (kmx - beglog)/(endlog - beglog) - 1) * num3); // convert to normalized z scale
									vertices[nver].vert[0] = (float)((2*(kmy - endlat)/(beglat - endlat) - 1) * num1); // convert to normalized x scale
									}
								else
									{
									vertices[nver].vert[2] = (float)fabs(kmx - num4); // convert to normalized z scale
									if ( mode >= 1 )
										{
										vertices[nver].vert[0] = (float)((2*(kmy - endlat)/(beglat - endlat) - 1) * num1); // convert to normalized x scale
										}
									else if ( mode <=0 )
										{
										vertices[nver].vert[0] = (float)((1 - 2*(kmy - endlat)/(beglat - endlat)) * num1); // convert to normalized x scale
										}
									}
								nvert++;
								}
							}
						}
					numrctot += numstepkmxt[iter];
					m_progresspercent = (int)((numrctot /  (float)maxranget) * 100);
					if ( m_progresspercent != oldpercent )
						{
						oldpercent = m_progresspercent;
						m_Progress.StepIt (); /*  step the progress bar */
						itoa(m_progresspercent, buffer, 10); /* convert m_prog... to char */
						strcat( buffer, "%");
						m_Label = buffer;    /* update the value of the progress bar */
						m_NewLabel.SetWindowText (m_Label);
						}
					}
				}



			////////////////////////////////////////////////////////
			}


		if (!(directx)) //calculate profile and write the big-endian heights to the BIN file if flagged
			{
			//int ier = Profile(lat0,lon0,hgt0,ang,aprn,mode,modeval,
	        //                  beglog,endlog,beglat,endlat );
			//allocate memory for prof array
			free( prof ); //free the memory allocated before for prof, and reallocate
			prof = (union profile *) malloc((2 * ang * invAzimuthStep + 10) * sizeof(union profile));

			//allocate memory for dtms array
			free( dtms ); //free the memory allocated before for dtms, and reallocate
			dtms = (union dtm *) malloc((28000) * sizeof(union dtm));

			//check that memory was properly allocated to all the arrays
			if ( prof == NULL || dtms == NULL )
			{
				CString lpszText = "Can't allocate enough memory!";
				reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

				//deallocate memory
				delete [] lg1chk;
				delete [] lt1chk;
				delete [] kmxb1;
				delete [] kmxb2;
				delete [] kmyb1;
				delete [] kmyb2;
				delete [] numstepkmyt;
				delete [] numstepkmxt;
				delete [] filt1;

				OnCancel();
			}

			if (calcProfile == 1) goto Profiles;
endProfiles:

			//address of the beginning of the data array is uni5[1]
            byteswritten = fwrite(&uni5[1].i5, numrctot * 2L, 1, stream2 );
		    fclose( stream2 );
			}
		else if (directx) //write heights to BI1 file
			{
			byteswritten = fwrite(&heights[0].height, nvert * sizeof(short), 1, stream3 );
			fclose( stream3 ); 
			}

		//deallocate memory
		free ( uni5 );
		free ( heights );
		delete [] lg1chk;
		delete [] lt1chk;
		delete [] kmxb1;
		delete [] kmxb2;
		delete [] kmyb1;
		delete [] kmyb2;
		delete [] numstepkmyt;
		delete [] numstepkmxt;
		delete [] filt1;

        if (!directx) goto skipdx; // don't write land.x (directx landscape) 
		
		strncpy( tmpfil, (const char *)ramdrive, 1);
		strncpy( tmpfil + 1, ":\\land.tm3", 11 );
		if( (stream3 = fopen(tmpfil, "w")) == NULL ) 
			{
			reply = AfxMessageBox( "Can't open the ramdrive:\\land.tm3 file!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

			OnCancel();
			}

		fprintf( stream3, "%s\n", binfil );//fprintf( stream3, "%s\n", filnn );
		fprintf( stream3, "%d,%d,%d,%d\n", numstepkmx, numstepkmy, i4min, i4max );
		fprintf( stream3, "%d,%d\n", numfracx2 - numfracx1, numfracy2 - numfracy1 );//,numx, numy );
		//actual x,y limits ( x = longiutde, y = latitude )
		fprintf( stream3, "%15.15f,%15.15f,%15.15f,%15.15f,%15.15f,%15.15f\n", beglog + xdim * numfracx1, beglog + xdim * (numfracx2 - 1), \
			   xdim, endlat - ydim * numfracy1, endlat - ydim * (numfracy2 - 1), ydim );
		//normalized x,y limits ( x = longiutde, y = latitude )
		if ( abs( mode ) == 4 )
			{
			kmx = beglog + xdim * numfracx1;
			lim1 = (2 * (kmx - beglog)/(endlog - beglog) - 1) * num3;
			kmx = beglog + xdim * (numfracx2 - 1);
			lim2 = (2 * (kmx - beglog)/(endlog - beglog) - 1) * num3;
			dlim1 = (2 * xdim * num3) /(endlog - beglog);
			kmy = endlat - ydim * numfracy1;
			lim3 = (2*(kmy - endlat)/(beglat - endlat) - 1) * num1;
			kmy = endlat - ydim * (numfracy2 - 1);
			lim4 = (2*(kmy - endlat)/(beglat - endlat) - 1) * num1;
			fprintf( stream3, "%15.15f,%15.15f\n", 0.5 * (endlog + beglog) , 0.5 * (beglat + endlat) );
			}
		else
			{
			kmx = beglog + xdim * numfracx1;
			lim1 = fabs(kmx - num4);
			kmx = beglog + xdim * (numfracx2 - 1);
			lim2 = fabs(kmx - num4);
			dlim1 = xdim;
			if ( mode >= 1 )
				{
				kmy = endlat - ydim * numfracy1;
				lim3 = (2*(kmy - endlat)/(beglat - endlat) - 1) * num1;
				kmy = endlat - ydim * (numfracy2 - 1);
				lim4 = (2*(kmy - endlat)/(beglat - endlat) - 1) * num1;
				}
			else if ( mode <=0 )
				{
				kmy = endlat - ydim * numfracy1;
				lim3 = (1 - 2*(kmy - endlat)/(beglat - endlat)) * num1;
				kmy = endlat - ydim * (numfracy2 - 1);
				lim4 = (1 - 2*(kmy - endlat)/(beglat - endlat)) * num1;
				}
			fprintf( stream3, "%15.15f,%15.15f\n", num4, 0.5 * (beglat + endlat) );
			}
		dlim2 = fabs((2 * ydim * num1) / (beglat - endlat));
		fprintf( stream3, "%15.15f,%15.15f,%15.15f,%15.15f,%15.15f,%15.15f,%d\n", lim1,lim2,dlim1,lim3,lim4,dlim2,nverts );//nvert );
		//scaling factors, x scale, y scale, z scale
		if ( abs( mode ) == 4 )
			{
			fprintf( stream3, "%15.15f,%15.15f,%15.15f", num3, num1, num2 * 1.0e+06 );
			}
		else
			{
			fprintf( stream3, "%15.15f,%15.15f,%15.15f", num4, num1, num2 * 1.0e+06 );
			}
		fclose( stream3 );

		//also write data to DirectX file "ramdrive:land.x"
		if (landflag == 1) //continuous motion land file
			{
			strncpy( tmpfil, (const char *)ramdrive, 1);
			strncpy( tmpfil + 1, ":\\land1.x", 10 );
			if( (stream3 = fopen(tmpfil, "wb")) == NULL ) 
				{
				reply = AfxMessageBox( "Can't open the ramdrive:\\land1.x file!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

				OnCancel();
				}			
			}
		else //land file as requested by Maps & More
			{
			strncpy( tmpfil, (const char *)ramdrive, 1);
			strncpy( tmpfil + 1, ":\\land.x", 9 );
			if( (stream3 = fopen(tmpfil, "wb")) == NULL ) 
				{
				reply = AfxMessageBox( "Can't open the ramdrive:\\land.x file!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

				OnCancel();
				}			
			}
		fwrite( &headbuffer, 16, 1, stream3 );   // xof 0302bin 0032
		fwrite( &sizeofit[0], 2, 1, stream3 ); //Token_Name
		charout = 6;  // size of name
		fwrite( &charout, 4, 1, stream3 );
		fwrite( (char *)"Header", 6, 1, stream3 ); // Token Name = "Header"
		fwrite( &sizeofit[6], 2, 1, stream3 ); // Token_Obrace
		fwrite( &sizeofit[4], 2, 1, stream3 ); // Token_Integer_List 
		charout = 3;  // Number of integers = 3
		fwrite( &charout, 4, 1, stream3 );
		charout = 1;  // Major version: 1
		fwrite( &charout, 4, 1, stream3 );
		charout = 0;  // Minor version: 0
		fwrite( &charout, 4, 1, stream3 );
		charout = 0;  // flags = 0 (binary)
		fwrite( &charout, 4, 1, stream3 );
		fwrite( &sizeofit[7], 2, 1, stream3 ); //Token_Cbrace
		fwrite( &sizeofit[0], 2, 1, stream3 );  //Token_Name
		charout = 4;  // Size of name
		fwrite( &charout, 4, 1, stream3 );
		fwrite( (char *)"Mesh", 4, 1, stream3 );  // Token Name = "Mesh"
		fwrite( &sizeofit[6], 2, 1, stream3 ); // Token_Obrace
		fwrite( &sizeofit[4], 2, 1, stream3 );  // Token_Integer_List
		charout = 1;  // Number of integers to follow = 1 integer of value  = number of vertices   
		fwrite( &charout, 4, 1, stream3 );
		charout = nvert;  // (Number of vertices ) //<--determine number of verteces !!!!!!!!!!!!!
		fwrite( &charout, 4, 1, stream3 );
		fwrite( &sizeofit[5], 2, 1, stream3 ); //Token_Realnum_list
		charout = 3 * nvert;  // (Number of coordinates of vertices )
		fwrite( &charout, 4, 1, stream3 );
		fwrite( vertices, 4, 3 * nvert, stream3 ); // coordinates of vertices

		//deallocate memory for vertices
		free (vertices );

		fwrite( &sizeofit[4], 2, 1, stream3 );  // Token_Integer_List
		//numface = 2 * (numstepkmx - 1) * (numstepkmy - 1);
		numface = 2 * (numfracy2 - numfracy1 - 1) * (numfracx2 - numfracx1 - 1);
		charout = 1 + 4 * numface;  // Number of integers to follow
		fwrite( &charout, 4, 1, stream3 );
		charout = numface;  // (Number of faces)
		fwrite( &charout, 4, 1, stream3 );

		// calculate faces
		numface = -1;
		for ( i = 1; i <= (numfracy2 - numfracy1 - 1); i++ ) // 1 to numstepkmy - 1
			{
			//numface = -1; 
			for( j = 1; j <= (numfracx2 - numfracx1 - 1); j++ )
				{
				numface++; // numface = numface + 1
				//nskip = (i - 1) * numstepkmx;
				nskip = (i - 1) * (numfracx2 - numfracx1);
				faces[numface].face[0] = 3;
				faces[numface].face[1] = j - 1 + nskip;
				faces[numface].face[2] = j + nskip;
				faces[numface].face[3] =  numfracx2 - numfracx1 + j  + nskip - 1;
				numface++; // numface = numface + 1
				faces[numface].face[0] = 3;
				faces[numface].face[1] = numfracx2 - numfracx1 + j  + nskip - 1;
				faces[numface].face[2] = j + nskip;
				faces[numface].face[3] = numfracx2 - numfracx1 + 1 + j + nskip - 1;
				}
			}

		fwrite( faces, 4, 4 * (numface + 1), stream3 ); // record faces
		fwrite( &sizeofit[0], 2, 1, stream3 );  //Token_Name
		charout = 16;  // Size of name
		fwrite( &charout, 4, 1, stream3 );
		fwrite( (char *)"MeshMaterialList", 16, 1, stream3 );  // Token Name = "Mesh"
		fwrite( &sizeofit[6], 2, 1, stream3 ); // Token_Obrace
		fwrite( &sizeofit[4], 2, 1, stream3 );  // Token_Integer_List
		charout = 3;  // Number of integers to follow = 1 integer of value 8 = number of vertices   
		fwrite( &charout, 4, 1, stream3 );
		charout = 1;
		fwrite( &charout, 4, 1, stream3 );
		charout = 1;
		fwrite( &charout, 4, 1, stream3 );
		charout = 0;
		fwrite( &charout, 4, 1, stream3 );
		fwrite( &sizeofit[0], 2, 1, stream3 );  //Token_Name
		charout = 8;  // Size of name
		fwrite( &charout, 4, 1, stream3 );
		fwrite( (char *)"Material", 8, 1, stream3 );  // Token Name = "Mesh"
		fwrite( &sizeofit[6], 2, 1, stream3 ); // Token_Obrace
		fwrite( &sizeofit[5], 2, 1, stream3 ); //Token_Realnum_list
		charout = 11;  // Number of integers to follow = 1 integer of value 8 = number of vertices   
		fwrite( &charout, 4, 1, stream3 );
		materials[0] = 1.0;
		materials[1] = 1.0;
		materials[2] = 1.0;
		materials[3] = 1.0;
		for (j = 4; j<12; j++ )
			{
			materials[j] = 0.0;
			}
		fwrite( &materials, 4, 11, stream3 );
		fwrite( &sizeofit[7], 2, 1, stream3 ); // Token_Cbrace
		fwrite( &sizeofit[7], 2, 1, stream3 ); // Token_Cbrace
		fwrite( &sizeofit[7], 2, 1, stream3 ); // Token_Cbrace

		fclose( stream3 );

		//deallocate memory for faces
		free( faces );

skipdx:  ;
		/* write flag to disk to tell VB program that computations
		   have ended successfully*/
		drivlet[0] = testdrv[WhereJK - 1];
        drivlet[1] = 0;
		bufFileName[0] = 0;
		sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\newreaddtm.end" );
	    stream2 = fopen( bufFileName, "w" );
		fprintf( stream2, "%d", 0 );
		fclose( stream2 );
		fcloseall();


		::DestroyWindow( CNewreadDTMDlg::m_hWnd );
		exit( 0 );


/**************findtile*****************/
findtile:
	double longitude, latitude;
	int itnum,itnum1,itnum2,uniquetilesfound;
	short uniquetile[3];
	
if (DTMflag == 0) //GTOPO30 DTM
{
	totnumtiles = 4; //maximum number of unique titles for analysis using GTOPO30
	uniquetilesfound = 0;

	for ( itnum = 0; itnum < 4; itnum++ ) // check the 4 corners
	{
	   switch (itnum)
		   {
			case 0:
			   longitude = beglog;
			   latitude = endlat;
			   break;
			case 1:
			   longitude = beglog;
			   latitude = beglat;
			   break;
			case 2:
			   longitude = endlog;
			   latitude = endlat;
			   break;
			case 3:
			   longitude = endlog;
			   latitude = beglat;
			   break;
		   }



	   if ( latitude > -60 ) // not Antartic
		{
			if ( latitude > 90 )
				{
				//manytiles = TRUE;
				latitude = 180 - latitude ;
				longitude = (endlog + beglog) - longitude;
				longitude += 180;
				if ( longitude > 180 )
					longitude -= 360;
				}

			if ( longitude > 180 )
				{
				//manytiles = TRUE;
				longitude -= 360;
				}
			else if ( longitude < -180 )
				{
				//manytiles = TRUE;
				longitude += 360;
				}

			tmp = (longitude + 180) * 0.025;
			nx = modf( tmp, &ntmp);
			lg1 = (int)(-180 + ntmp * 40);
			lg11 = abs(lg1);
			if (lg11 >= 100)
			{
				_itoa( lg11, lg1ch, 10);
			}
			else
			{
				strncpy(lg1ch, "0", 1);
				strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);
			}
			if (lg1 < 0) 
			{
				strncpy( EW, "W", 1 );
			}
			else
			{
				strncpy( EW, "E", 1 );
			}
			ny = (int)((90 - latitude) * 0.02);
			lt1 = 90 - 50 * ny;
			lt11 = abs(lt1);
			_itoa( lt11, lt1ch, 10);
			if (lt1 > 0) 
			{
				strncpy( NS, "N", 1 );
			}
			else
			{
				strncpy( NS , "S", 1 );
			}
			strncpy( filt1[itnum], (const char *)EW, 1 );
			strncpy( filt1[itnum] + 1, (const char *)lg1ch, 3 );
			strncpy( filt1[itnum] + 4, (const char *)NS, 1 );
			strncpy( filt1[itnum] + 5, (const char *)lt1ch, 2 );
			strncpy( filt1[itnum] + 7, "\0", 1);
			strncpy( filnn, (const char *)worlddtm, 1);
			strncpy( filnn + 1, ":\\", 2 );
			strncpy( filnn + 3, (const char *)filt1[itnum], 7 );
			strncpy( filnn + 10, "\\", 1 );
			strncpy( filnn + 11, (const char *)filt1[itnum], 7 );
			nrows = 6000;
			ncols = 4800;
			whichCDs[itnum] = worldCD[ny * 9 + (int)ntmp + 1];

			//determine kmx,kmy boundaries of analysis within this tile
			kmxb1[itnum] = __max(lg1, beglog);
			kmxb2[itnum] = __min(lg1 + 40, endlog);
			kmyb1[itnum] = __min(lt1 ,endlat);
			kmyb2[itnum] = __max(lt1 - 50, beglat);

		}
		else // Antartic CD #5
		{
			if ( latitude < -90 )
				{
				//manytiles = TRUE;
				latitude = -180 - latitude;
				longitude = (endlog + beglog) - longitude;
				longitude += 180;
				}

			if ( longitude > 180 )
				{
				longitude -= 360;
				}
			else if ( longitude < -180 )
				{
				longitude += 360;
				}

			
			tmp = (longitude + 180) / 60;
			nx = modf( tmp, &ntmp);
			lg1 = (int)(-180 + ntmp * 60);
			lg11 = abs(lg1);
			if (lg11 >= 100 )
				{
				_itoa( lg11, lg1ch, 10);
				}
			else if (( lg11 < 100 ) && ( lg11 != 0 ))
				{
				strncpy(lg1ch, "0", 1);
				strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);
				}
			else if (lg11 == 0 )
				{
				strncpy(lg1ch, "000", 3);
				}
			if (lg1 <= 0) 
				{
				strncpy( EW, "W", 1 );
				}
			else
				{
				strncpy( EW, "E", 1 );
				}
			strncpy( NS, "S", 1 );
			lt1 = -60;
			strncpy(lt1ch, "60", 2);
			strncat( filt1[itnum], EW, 1 );
			strncat( filt1[itnum], lg1ch, 3 );
			strncat( filt1[itnum], NS, 1 );
			strncat( filt1[itnum], lt1ch, 2 );
			strncpy( filt1[itnum] + 7, "\0", 1);
			strncpy( filnn, (const char *)worlddtm, 1);
			strncpy( filnn + 1, ":\\", 2 );
			strncpy( filnn + 3, (const char *)filt1[itnum], 7 );
			strncpy( filnn + 10, "\\", 1 );
			strncpy( filnn + 11, (const char *)filt1[itnum], 7 );
			nrows = 3600;
			ncols = 7200;
			whichCDs[itnum] = 5;

			//determine kmx,kmy boundaries of analysis within this tile
			kmxb1[itnum] = __max(lg1, beglog);
			kmxb2[itnum] = __min(lg1 + 60, endlog);
			kmyb1[itnum] = __min(lt1 ,endlat);
			kmyb2[itnum] = __max(lt1 - 30, beglat);
		
		}

		lt1chk[itnum] = lt1;
		lg1chk[itnum] = lg1;

        //determine number of steps in this tile   
		numstepkmyt[itnum] = Nint((fabs(kmyb1[itnum] - kmyb2[itnum])) / ydim) + 1;
		numstepkmxt[itnum] = Nint((fabs(kmxb1[itnum] - kmxb2[itnum])) / xdim) + 1;

		if ( itnum == 0 )
			{
			strncpy( filn, filnn, 18 );
			maxranget += numstepkmyt[itnum] * numstepkmxt[itnum];
			uniquetile[itnum] = 1;
		}
		else
	    //check if this is unique tile
			{
  			maxranget += numstepkmyt[itnum] * numstepkmxt[itnum];
		    uniquetile[itnum] = 1; //assume it is an unique tile name
			for (int itcheck = 0; itcheck < itnum; itcheck++)
				{
			    if (_strnicoll( (const char *)filt1[itcheck], (const char *)filt1[itnum], 7) == 0 )
					{
					totnumtiles -= 1; //this tile not unique
        			maxranget -= numstepkmyt[itnum] * numstepkmxt[itnum];
					//it is not an unique tile name
					uniquetile[itnum] = 0;
					break;
					}
				}
			}


		}
		//now shift the unique tile names in the stack downward
		//along with its other parameters
		for (int itcheck = 1; itcheck < 4; itcheck++)
			{
			if (uniquetile[itcheck] == 1)
				{
				uniquetilesfound++;
				strncpy(filt1[uniquetilesfound], filt1[itcheck], 8);
				numstepkmyt[uniquetilesfound] = numstepkmyt[itcheck];
				numstepkmxt[uniquetilesfound] = numstepkmxt[itcheck];
				kmxb1[uniquetilesfound] = kmxb1[itcheck];
				kmxb2[uniquetilesfound] = kmxb2[itcheck];
				kmyb1[uniquetilesfound] = kmyb1[itcheck];
				kmyb2[uniquetilesfound] = kmyb2[itcheck];
				}
			}


}

else if (DTMflag > 0) //SRTM or AROS data, determine number of tiles

{
	  //now determine the tile names
	  itnum = 0;
	  for ( itnum1 = 0; itnum1 <= abs(lg2SRTM - lg1SRTM); itnum1++ )
	  {
		  for ( itnum2 = 0; itnum2 <= abs(lt2SRTM - lt1SRTM); itnum2++ )
		  {
			lg1 = lg1SRTM + itnum1;
			lg11 = abs(lg1);

			if (lg11 >= 100)
			{
				_itoa( lg11, lg1ch, 10);
			}
			else if (lg11 >= 10 && lg11 < 100 )
			{
				strncpy(lg1ch, "0", 1);
				strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);
			}
			else if (lg11 < 10 )
			{
				strncpy(lg1ch, "00", 2);
				strncpy(lg1ch + 2, (const char *)_itoa( lg11, lch, 10), 1);
			}


			if (lg1 < 0) //<--check out that name convention for lg1=0
						 //is E000 and not W000 -- otherwise use "<="
			{
				strncpy( EW, "W", 1 );
			}
			else
			{
				strncpy( EW, "E", 1 );
			}
			
			lt1 = lt2SRTM + itnum2;
			lt11 = abs(lt1);

			if (lt11 >= 10 )
			{
			   _itoa( lt11, lt1ch, 10);
			}
			else
			{
        		strncpy(lt1ch, "0", 1);
				strncpy(lt1ch + 1, (const char *)_itoa( lt11, lch, 10), 1);
			}

			if (lt1 >= 0) 
			{
				strncpy( NS, "N", 1 );
			}
			else
			{
				strncpy( NS , "S", 1 );
			}

			strncpy( filt1[itnum], (const char *)NS, 1 );
			strncpy( filt1[itnum] + 1, (const char *)lt1ch, 2 );
			strncpy( filt1[itnum] + 3, (const char *)EW, 1 );
			strncpy( filt1[itnum] + 4, (const char *)lg1ch, 3 );
			strncpy( filt1[itnum] + 7, "\0", 1);
			strncpy( filnn, (const char *)worlddtm, 1);
			strncpy( filnn + 1, ":\\", 2 );
			
			if (DTMflag == 1) // 30 meter NED or SRTM1
			{
				nrows = 3601;
				ncols = 3601; 
				strcpy( filnn, usadrive );
				//strncpy( filnn + 3, "USA", 3 );
				//strncpy( filnn + 6, "\\", 1 );
				strncat( filnn, (const char *)filt1[itnum], 7 );
				//strncpy( filnn + 7, (const char *)filt1[itnum], 7 );
			}
			else if (DTMflag == 2) //90 meter SRTM3 or MERIT DEM
			{
				nrows = 1201;
				ncols = 1201;
				//strncpy( filnn + 3, "3AS", 3 );
				//strncpy( filnn + 6, "\\", 1 );
				strcpy( filnn, tasdrive );
				//strncpy( filnn + 7, (const char *)filt1[itnum], 7 );
				strncat( filnn, (const char *)filt1[itnum], 7 );
			}
			else if (DTMflag == 3) //ALOS 30 meter DEM BIL files
			{
				nrows = 3600;
				ncols = 3600;
				//strncpy( filnn + 3, "BIL", 3 );
				//strncpy( filnn + 6, "\\", 1 );
				strcpy( filnn, bildrive );
				//strncpy( filnn + 7, (const char *)filt1[itnum], 7 );
				strncat( filnn, (const char *)filt1[itnum], 7 );
			}

			lt1chk[itnum] = (int)(lt1 + 1.0); //SRTM DTM is named by SW corner
									   //unlike GTOPO30 which is named by NW corner

            lg1chk[itnum] = lg1;

			//determine kmx,kmy boundaries of analysis within this tile
			//and the number of steps in xdim,ydim within this tile
			kmxb1[itnum] = __max(lg1, beglog);
			kmxb2[itnum] = __min(lg1 + 1, endlog);
			kmyb1[itnum] = __min(lt1 + 1, endlat);
			kmyb2[itnum] = __max(lt1, beglat);

			numstepkmyt[itnum] = Nint((fabs(kmyb1[itnum] - kmyb2[itnum])) / ydim) + 1;
			numstepkmxt[itnum] = Nint((fabs(kmxb1[itnum] - kmxb2[itnum])) / xdim) + 1;

			//record tile name in filn buffer
			if ( itnum == 0 )
			{
				if (DTMflag != 0 ) // not GTOPO30
				{
					sprintf( filn, "%s", filnn);
				}
				else
				{
					strncpy( filn, filnn, 14 );
				}
			}

			maxranget += numstepkmyt[itnum] * numstepkmxt[itnum];
			itnum++;
		  }
	  }
}
    //current tile name
	strncpy( filt, filt1[0], 7 );
	numCD = whichCDs[0];

	if (gototag == 1 )
		{
		goto ret1;
		}
	else
		{
		}
/********************************************/


	
/****************dtmfiles*****************/
dtmfiles:
	{
		if (DTMflag == 0) //GTOPO30
		{
			strncpy( DTMfile, filn, 18 );
			strncpy( DTMfile + 18, ".DEM", 4 );
			strncpy( DTMfile + 22, "\0", 1);
		}
		else if (DTMflag == 1 || DTMflag == 2) //SRTM DTM data
		{
			DTMfile[0] = 0;
			sprintf( DTMfile, "%s%s", filn, ".hgt"); //EK 061322
			//strncpy( DTMfile, filn, 14 );
			//strncpy( DTMfile + 14, ".hgt", 4 );
			//strncpy( DTMfile + 18, "\0", 1);
		}
		else if (DTMflag == 3) //ALOS DEM data
		{
			DTMfile[0] = 0;
			sprintf( DTMfile, "%s%s", filn, ".bil"); //EK 061322
		}

retry:	if ( (stream = fopen( (const char *)DTMfile, "rb" )) == NULL)
		if ( stream == NULL )
			{
			if (DTMflag == 0)
				{
				// tell user which CD it is on
				strncpy( quest, "Please insert USGS EROS CD#", 27 );
				strncpy( quest + 27, (const char *)_itoa( numCD, CDbuffer, 10 ), 1 );
				strncpy( quest + 28, "\0", 1);
				reply = AfxMessageBox( (const char *)quest, MB_OKCANCEL | MB_ICONINFORMATION | MB_APPLMODAL );		
				if ( reply == 2 ) // Cancel button selected
					{

					//deallocate memory
					delete [] lg1chk;
					delete [] lt1chk;
					delete [] kmxb1;
					delete [] kmxb2;
					delete [] kmyb1;
					delete [] kmyb2;
					delete [] numstepkmyt;
					delete [] numstepkmxt;
					delete [] filt1;

					OnCancel(); 
					}
				else
					{
					goto retry;
					}
				}
			else if (DTMflag > 0) //warn user
				{ 
				if (IgnoreMissingTiles == 0)
					{
					//strncpy( quest, "Can't find the SRTM tile: ", 26 );
					//strncpy( quest + 26, (const char *)DTMfile, 18 );
					//strncpy( quest + 44, "\n", 1);
					//strncpy( quest + 45, "Do you wan't to use an empty tile instead?", 42);
					//strncpy( quest + 87, "\0", 1);
					sprintf( quest, "%s%s\n%s", "Can't find the DTM tile: ", DTMfile, "Do you wan't to use an empty tile instead?" );
					reply = AfxMessageBox( (const char *)quest, MB_YESNOCANCEL | MB_ICONINFORMATION | MB_APPLMODAL );		
					}
				else if (IgnoreMissingTiles == 1) //automatically use default empty tile
					{
					reply = 6; 
					}

				if (reply == 6) //Replace the missing srtm file with an empty tile
					{
					//goto retry;
						if (DTMflag == 1) //SRTM1 DEM
						{
							//strncpy( DTMfile, (const char *)worlddtm, 1);
    						//strncpy( DTMfile + 1, ":\\USA\\Z000000.hgt\0", 18);
							sprintf( DTMfile, "%s%s", usadrive, "Z000000.hgt" );
							goto retry;
						}
						else if (DTMflag == 2) //SRTM3 - MERIT DEM
						{
							//strncpy( DTMfile, (const char *)worlddtm, 1);
    						//strncpy( DTMfile + 1, ":\\3AS\\Z000000.hgt\0", 18);
							sprintf( DTMfile, "%s%s", tasdrive, "Z000000.hgt" );
							goto retry;
						}
						else if (DTMflag == 3) //ALOS DEM
						{
							//strncpy( DTMfile, (const char *)worlddtm, 1);
    						//strncpy( DTMfile + 1, ":\\BIL\\Z000000.bil\0", 18);
							sprintf( DTMfile, "%s%s", bildrive, "Z000000.hgt" );
							goto retry;
						}
					}
				else if (reply == 7) //the user wants to try again
					{
					goto retry;
					}
				else if (reply == 2) // Cancel button selected
					{

					//deallocate memory
					delete [] lg1chk;
					delete [] lt1chk;
					delete [] kmxb1;
					delete [] kmxb2;
					delete [] kmyb1;
					delete [] kmyb2;
					delete [] numstepkmyt;
					delete [] numstepkmxt;
					delete [] filt1;

					OnCancel(); 
					}
				}

			}
			


		if (gototag == 1 ) 
			{
			goto ret2;
			}
		if (gototag == 2)
			{
			goto ret4;
			}
	}

/********************************************/
NoVoids: //In-line reoutine used for removing
	     //SRTM DTM data voids by linear interpolation in 
	     //three dimensions: kmx, kmy, and height. 
	     //The result has lower spatial resolution then
         //the original DTM grid, but is better than 
	     //leaving a void.
    
	i41x = i4;
	i42x = i4;
	i41y = i4;
	i42y = i4;
	i41xt = 0;
	i42xt = 0;
	i41yt = 0;
	i42yt = 0;
	wtxd = 0;
	wtyd = 0;
	w4dtot = 0;

	//search in kmx

	for ( kk = jj + 1; kk <= numstepkmx - iikmx; kk++ ) //search the entire extraction row to the end
		{
		i41x = (short int)(uni5[numrc + kk].chi5[0] << 8 | uni5[numrc + kk].chi5[1]);
		if ( i41x >= -499 )
			{
			i41xt = kk;
			goto noVoids1; //non void found
			}
		}
noVoids1: 
	//if ( i41xt == 0 ) goto unVoid; //non void not found

	for ( kk = jj - 1; kk >= - iikmx + 1 ; kk-- ) //search the entire extraction row to the beginning
		{
		i42x = (short int)(uni5[numrc + kk].chi5[0] << 8 | uni5[numrc + kk].chi5[1]);
		if ( i42x >= -499 )
			{
			i42xt = kk;
			goto noVoids2; //non void found
			}
		}
noVoids2:
	//if ( i42xt == 0 ) goto unVoid; //non void not found

	//search in kmy
	for ( kk = iikmy + 1; kk <= numstepkmy; kk++) //search the entire extraction column to the end
		{
		numrc2 = numrc + jj + (kk - iikmy) * numstepkmx;
 		i41y = (short int)(uni5[numrc2].chi5[0] << 8 | uni5[numrc2].chi5[1]);
		if ( i41y >= -499 )
			{
			i41yt = kk;
			goto noVoids3; //non void found
			}
		}
noVoids3:
	//if ( i41yt == 0 ) goto unVoid; //non void not found

	for ( kk = iikmy - 1; kk >= 1; kk--) //search the entire extraction column until the beginning
		{
		numrc2 = numrc + jj + (kk - iikmy) * numstepkmx;
		i42y = (short int)(uni5[numrc2].chi5[0] << 8 | uni5[numrc2].chi5[1]);
		if ( i42y >= -499 )
			{
			i42yt = kk;
			goto noVoids4; //non void found
			}
		}
noVoids4:
	//if ( i42yt == 0 ) goto unVoid; //non void not found

    //If got here, then interpolation succeeded.
	//If not on extraction boundary, then calculate
	//weighted average over three dimensions kmx,kmy,height.
	//If on extraction boundary, then interpolate in kmx, or kmy vs height.
    
	if (i41xt != 0 && i42xt !=0 && i41yt !=0 && i42yt != 0)
	{
	w4dtot = (double)(abs(i41xt - i42xt) + abs(i41yt - i42yt));
	wtxd = abs(i41xt - i42xt)/w4dtot;
	wtyd = abs(i41yt - i42yt)/w4dtot;
	i4 = (short)(((jj - i41xt)*(i42x - i41x)/(i42xt - i41xt) + i41x) * wtxd + 
		         ((iikmy - i41yt)*(i42y - i41y)/(i42yt - i41yt) + i41y) * wtyd);
	}
	else if (( (i41xt != 0) && (i42xt != 0) ) && (i41yt == 0 || i42yt == 0))
	{
		i4 = (short)((jj - i41xt)*(i42x - i41x)/(i42xt - i41xt) + i41x);
	}
	else if ((i41xt == 0 || i42xt == 0) && (i41yt != 0 && i42yt != 0))
	{
		i4 = (short)((ii - i41yt)*(i42y - i41y)/(i42yt - i41yt) + i41y);
	}
 
   goto unVoid;

/********************************************/

/*end of timer(calculation) routine*/

/////////////////Profiles inline routine///////////////////
//int ier = Profile(lat0,lon0,hgt0,ang,aprn,mode,modeval,
//                  beglog,endlog,beglat,endlat );

//int Profile( double kmyo,double kmxo, double hgt, double mang, double aprn,double nnetz,double modval, 
//		double lgo,double endlg, double beglt, double endlt)

Profiles:

	dstpointx = xdim;
	dstpointy = ydim;
	kmyo = lat0;
	kmxo = lon0;
	hgt = hgt0;
	mang = ang;
	nkmx1 = 1;
	startkmx = beglog;
    sofkmx = endlog;
    lto = endlat;

    //numcol = (int) ((endlog - beglog) / dstpointx) + 1;
    //numrow = (int) ((lto - beglat) / dstpointy) + 1;

    if (mang != 0.) {
	maxang = mang;
    }
    if (maxang > 90.) {  //EK 05/27/24 used to be 180, but that can't be since it is other horizon, rathe limit is 90 degrees
	maxang = 90.;
    }
    if (maxang < 30.) {
	maxang = 30.;
    }
    maxang0 = maxang;
    nomentr = (short) (2 * maxang * invAzimuthStep + 1);
    apprnr = 0.;
    if (aprn != 0.) {
	apprnr = aprn * .0083333333 / cos(kmyo * cd);
    }
	nnetz = mode;
    if (nnetz <= 0) {
	   netz = false;
    }
    if (nnetz >= 1) {
	   netz = true;
    }

	// now set optimization range 
    optang = (short) (maxang - 12);
    numskip = 1;

    kmyoo = kmyo;
    kmxoo = kmxo;
    lt = kmyoo;
    lg = kmxoo;
    hgt += hgtobs;


	if (hgt >= 0.) {
	    anav = sqrt(hgt) * -.0348f;
	}
	if (hgt < 0.) {
	    anav = -1.f;
	}

	double MinViewAngle = (float)anav - minview;

	for ( i__ = 0; i__ < nomentr; ++i__) {
		prof[i__].ver[0] = MinViewAngle;}
 
	//old code -- doesn't handle the apprnr right, needs to be checked for each coordinate in loop
	/*
	if (netz) {
		endkmx = sofkmx;
		if (startkmx < kmxo + apprnr) {
			begkmx = kmxo + apprnr;
		} else {
			begkmx = startkmx;
		}
    } else {
		begkmx = startkmx;
		if (sofkmx > kmxo - apprnr) {
			endkmx = kmxo - apprnr;
		} else {
			endkmx = sofkmx;
		}
    }
	*/

	if (netz) {
		endkmx = sofkmx;
		begkmx = startkmx;
    } else {
		begkmx = startkmx;
		endkmx = sofkmx;
    }

    // notify the user of what is going on by the form's caption */
	m_Progress.SetPos (0);
	m_progresspercent = 0;
	m_Label = _T("");
	m_NewLabel.SetWindowText (m_Label);
	
	numrctot = 0;
	oldpercent = 0;


	SetWindowText("Calculating the Profile");


    nbegkmx = Nint(begkmx / dstpointx);
    begkmx = nbegkmx * dstpointx;
    skipkmx = dstpointx;
    skipkmy = dstpointy;
    numkmx = Nint(fabs(endkmx - begkmx)/skipkmx) + 1;

	//////////diagnostics////////////////////////////////////
	double distclose = 0;
	double deltakmx;
	double deltakmy;
	double mindiffx = 9999;
	double mindiffy = 9999;
	double mindist = 9999;
	double maxcalcrange = -9999;
	char * fout;
	fout = "c://jk//testazi.txt";
	bool testopened = false;
	bool diagnostics = false;
	FILE * streamout;
	if (diagnostics) {
		if ( streamout = fopen(fout, "w" ) )
		{
			testopened = true;
		}
	}
	/////////////////////////////////////////////////////////////////

	if (maxang < 50) AzimuthOpt = true; //optimize calculations for the continental USA

	/////////////////////////////////////////////////////

    re = 6371315.;  //earth radius in Clark geoid in meters  //N.b. According to Widipedia Geographic Distances should be 6371009
    rekm = 6371.315f;  //Earth radius in Clark geoid in kms
    i__1 = numkmx;
    i__2 = numskip;
    for (i__ = nkmx1; i__2 < 0 ? i__ >= i__1 : i__ <= i__1; i__ += i__2) {

		kmx = begkmx + (i__ - 1) * skipkmx;
		ikmx = Nint((kmx - beglog) / skipkmx) + 1;

		if (ikmx == 0) {
			goto L550;
		}
/*      now determine optimized kmy angular range */
/*      first convert longitude range into kilometers to a first approximation */
		//it is Earth Radius * Delta-Angle (radians) along radius
		//Delta-Angle = fabs(kmx - kmxo) * cd -- cd converts from degrees to radians
		//Radius along the latitude is rekm * cos(kmyo * cd) for small changes in latitdue
		range = fabs(kmx - kmxo) * cd * rekm * cos(kmyo * cd);
		if (diagnostics)
		{
			if (range > maxcalcrange)
			{
				maxcalcrange = range;
			}
		}
		if (nnetz == 1) { //sunrise
			if (kmx <= kmxo + apprnr) goto L550;
/*          if range >= 20 km, then narrow azimuthal range to MAXANG=OptAng */
			if (range >= 20. && AzimuthOpt && maxang > (double) optang) {
				maxang = (double) optang; }
/*          if range >= 40 km and SRTM-1, then reduce step size to ~ 60m */
			if (range >= 40. && DTMflag != 2) {
				numskip = 2;
			if (range >= 40. && range < 60) {
				numskip = 2; }
			else if (range >= 60 && range < 80) {
				numskip = 3; }
			else if (range >= 80 && range < 120) {
				numskip = 4; }
			else if (range >= 120 && range < 160) {
				numskip = 5; }
			else if (range >= 160 && range < 200) {
				numskip = 6; }
			else if (range >= 200 && range < 220) {
				numskip = 8; }
			else if (range >= 220 && range < 240) {
				numskip = 9; }
			else if (range >= 240 && range < 260) {
				numskip = 10; }
			else if (range >= 260 && range < 280) {
				numskip = 11; }
			else if (range >= 280 && range < 300) {
				numskip = 12;}
			else if (range >= 300) {
				numskip = 13;}			
			
			}
		} else if (nnetz == 0) { //sunset
			if (kmx >= kmxo - apprnr) goto L550;
/*          if range >= 20 km, then narrow azimuthal range to MAXANG=OptAng */
			if (range > 20. && AzimuthOpt && maxang > (double) optang) {
				maxang = (double) optang; }
/*          if range >= 40 km and SRTM-1, then increase step size to ~ 60 m */
			if (range > 40. && DTMflag != 2) {
				//numskip = 2;
			
				if (range >= 40. && range < 60) {
					numskip = 2; }
				else if (range >= 60 && range < 80) {
					numskip = 3; }
				else if (range >= 80 && range < 120) {
					numskip = 4; }
				else if (range >= 120 && range < 160) {
					numskip = 5; }
				else if (range >= 160 && range < 200) {
					numskip = 6; }
				else if (range >= 200 && range < 220) {
					numskip = 8; }
				else if (range >= 220 && range < 240) {
					numskip = 9; }
				else if (range >= 240 && range < 260) {
					numskip = 10; }
				else if (range >= 260 && range < 280) {
					numskip = 11; }
				else if (range >= 280 && range < 300) {
					numskip = 12;}
				else if (range >= 300) {
					numskip = 13;}			
			
			}
/*          if range <= 20 then azimuthal range is maximized = maxang0 */
			if (range <= 20.) {
				maxang = maxang0; }
/*          if distance <= 40 km and SRTM-1, then reduce stepsize */
			if (range <= 40. && DTMflag != 2) {
				numskip = 1;}
		}
/*      optimized half angular range in latitude (degrees) = dkmy */
		if (maxang <= 43) {
			dkmy = range / (rekm * cd) * tan((maxang + 2) * cd);
		}
		else
		{
			dkmy = fabs(endlat - beglat); //range / (rekm * cd)
		}
		//dkmy = fabs(endlat - beglat) * 0.5 + 1;

		d__1 = fabs(dkmy/skipkmy);
		numkmy = 2 * Nint(d__1) + 1;
		numhalfkmy = Nint(d__1) + 1;
		numdtm = (int) ((maxang0 - maxang) * 10);
		i__3 = numkmy;
		i__4 = numskip;
		for (j = 1; i__4 < 0 ? j >= i__3 : j <= i__3; j += i__4) {

			//enable message reception during loop
			if (::PeekMessage(&Message, NULL, 0, 0, PM_REMOVE))
				{
				::TranslateMessage(&Message);
				::DispatchMessage(&Message);
				}

			kmy = kmyo - (numhalfkmy - j + 1) * skipkmy;
			if (kmy < beglat || kmy > endlat) {
			goto L120;  }

			//check the point is beyond the nearest approach allowed
			//use simplest calculation for the distance valid for small distances

			if (diagnostics)
			{
				deltakmx = kmx - kmxo;
				deltakmy = kmy - kmyo;
				distclose = rekm * cd * sqrt(pow(deltakmy,2.0) + pow(deltakmx * cos(kmyo*cd),2.0));
				if (fabs(deltakmx) < mindiffx)
				{
					mindiffx = fabs(deltakmx);
				}
				if (fabs(deltakmy) < mindiffy)
				{
					mindiffy = fabs(deltakmy);
				}
				if (distclose < mindist)
				{
					mindist = distclose;
				}
				if (distclose < aprn )
				{
					goto L120;
				}
			}
			//further optimization: use crude approximation of azimuth to
			//determine if data will lie within the required azimuth range
			if (kmx != lg && AzimuthOpt)
			{
				testazi = (kmy - lt)/(kmx - lg);
				testazi = atan(testazi)/cd;
				if (testazi > mang + 10.0 || testazi < -mang - 10.0 ) goto L120;
			}

			ikmy = Nint((lto - kmy) / skipkmy) + 1;
			if (ikmy == 0) {
				goto L120;}

/*      due to some subtle shift, all the data was written to */
/*      starting from an empty data point, so avoid it */
//	    if (ikmy == 1 && ikmx == 1) {
//		goto L120; }

	/*      this subroutine reads the SRTM DEM at any (lt,lg) and returns the altitude */
			numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1;
			if (numrc >= 0 && numrc <= numtotrec ) 
			{
				//hgt2 = (float)(uni5[numrc].chi5[0] << 8 | uni5[numrc].chi5[1]);

				if (DTMflag == 1 || DTMflag == 2) //SRTM hgt Motorola Big Endian
				{
					//rotate bits to little-endian
					i4sw[0].i4s = uni5[numrc].i5;
					uni5[numrc].chi5[0] = i4sw[0].chi6[1];
					uni5[numrc].chi5[1] = i4sw[0].chi6[0];
				}
				hgt2 = uni5[numrc].i5;
				if (fabs(hgt2) == 9999. || fabs(hgt2) == 32768. || fabs(hgt2) > 8848.) {
					hgt2 = 0.;} //ocean and radar voids or glitches
			}
			else
			{
				hgt2 = 0;
			}

/*          calculate the view angle and azimuth, without */
/*          atmospheric refraction for the selected observation point */
/*          at longitude kmx,latitude kmy, height hgt2 */
/* L45: */
			fudge = treehgt;
			if (hgt2 == -9999. || hgt2 == -32768. || fabs(hgt2) > 8848.) {
			hgt2 = 0.f;
			}
			hgt2 += fudge;
			lg2 = kmx;
			lt2 = kmy;
			x1 = cos(lt * cd) * cos(-lg * cd);
			x2 = cos(lt2 * cd) * cos(-lg2 * cd);
			y1 = cos(lt * cd) * sin(-lg * cd);
			y2 = cos(lt2 * cd) * sin(-lg2 * cd);
			z1 = sin(lt * cd);
			z2 = sin(lt2 * cd);
			/* Computing 2nd power */
			d__1 = x1 - x2;
			/* Computing 2nd power */
			d__2 = y1 - y2;
			/* Computing 2nd power */
			d__3 = z1 - z2;
			distd = rekm * sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
			re1 = hgt + re;
			re2 = hgt2 + re;
			deltd = hgt - hgt2;
			x1 = re1 * x1;
			y1 = re1 * y1;
			z1 = re1 * z1;
			x2 = re2 * x2;
			y2 = re2 * y2;
			z2 = re2 * z2;
			dist1 = re1;
			dist2 = re2;
			angle = acos((x1 * x2 + y1 * y2 + z1 * z2) / (dist1 * dist2));
/*          view angle in radians */
			viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
			d__ = (dist1 - dist2 * cos(angle)) / dist1;
			x1d = x1 * (1 - d__) - x2;
			y1d = y1 * (1 - d__) - y2;
			z1d = z1 * (1 - d__) - z2;
			x1p = -sin(-lg * cd);
			y1p = cos(-lg * cd);
			z1p = 0.;
			azicos = x1p * x1d + y1p * y1d;
			x1s = -cos(-lg * cd) * sin(lt * cd);
			y1s = -sin(-lg * cd) * sin(lt * cd);
			z1s = cos(lt * cd);
			azisin = x1s * x1d + y1s * y1d + z1s * z1d;
			azi = atan(azisin / azicos);
	/*      azimuth in degrees */
			azi /= cd;
			if (fabs(azi) > 80) {
				int cc;
				cc = 1;
			}

			if (TemperatureModel == -1) {
	/*      add old contribution of atmospheric refraction to view angle */

				/*  deprecated old model */
				if (deltd <= 0.) {
				defm = 7.82e-4f - deltd * 3.11e-7;
				defb = deltd * 3.4e-5f - .0141f;
				} else if (deltd > 0.) {
				defm = deltd * 3.09e-7 + 7.64e-4f;
				defb = -.00915f - deltd * 2.69e-5f;
				}
				avref = defm * distd + defb;
				if (avref < 0.) {
				avref = 0.;
				}

			}else if (TemperatureModel >= 1) {

				if (TemperatureModel == 2 && TemperatureModel == 4) 
				{
					//don't add any terrestrial refraction
					avref = 0.0;
				}
				else
				{
					//use terrestrial refraction formula from Wikipedia, suitably modified to fit van der Werf ray tracing
					PATHLENGTH = sqrt(distd*distd + pow(fabs(deltd)*0.001 - 0.5*(distd*distd/RETH),2.0));
					if (fabs(deltd) > 1000.0) {
						Exponent = 0.99;
					}else{
						Exponent = 0.9965;
					}

					avref = TRpart * pow(PATHLENGTH, Exponent); //degrees
					
				}

			}else if (TemperatureModel == 0) {
				//no added terrestrial refraction
				avref = 0.0;
			}

			++numdtm;
			dtms[numdtm].ver[0] = azi;
			dtms[numdtm].ver[1] = viewang / cd + avref;
			dtms[numdtm].ver[2] = kmy;

			////////////////diagnostics///////////////////
			if (testopened && diagnostics) 
			{
				fprintf(streamout, "%10.6f,%10.6f,%10.6f,%10.6f,%10.6f,%10.6f,%d,%10.6f\n", kmx,kmy,deltakmx,deltakmy,distclose,range,numskip,azi);
			}
			/////////////////////////////////////////

L120:
			;
		}
/*      interpolate to a grid of 0.1 degrees in azimuth */
/*      the interpolated points reside in the array prof() */

		istart1 = (int) ((maxang0 - maxang) * invAzimuthStep + 1);
L150:
		if (netz) {
			az2 = dtms[istart1].ver[0];
			vang2 = dtms[istart1].ver[1];
			kmy2 = dtms[istart1].ver[2];
		} else {
			az1 = dtms[istart1].ver[0];
			vang1 = dtms[istart1].ver[1];
			kmy1 = dtms[istart1].ver[2];
		}
		nstart2 = istart1 + 1;
L160:
		if (netz) {
			az1 = dtms[nstart2].ver[0];
			vang1 = dtms[nstart2].ver[1];
			kmy1 = dtms[nstart2].ver[2];
		} else {
			az2 = dtms[nstart2].ver[0];
			vang2 = dtms[nstart2].ver[1];
			kmy2 = dtms[nstart2].ver[2];
		}

		if (az2 >= -maxang0 && az1 <= maxang0 + AzimuthStep) {
			bne1 = (int)(az1 * invAzimuthStep) / (double)invAzimuthStep - AzimuthStep;
			bne2 = (int)(az2 * invAzimuthStep) / (double)invAzimuthStep + AzimuthStep;
			found = false;
			ndela = (int) ((bne2 - bne1) / AzimuthStep + 1);
			i__4 = ndela;
			for (ndelazi = 1; ndelazi <= i__4; ++ndelazi) {
				delazi = bne1 + (ndelazi - 1) * AzimuthStep;
				if (delazi < -maxang0) {
					goto L165;}
				if (delazi > maxang0 + AzimuthStep) {
					goto L170;}
				if (az1 <= delazi && az2 >= delazi) {
					bne = (delazi + maxang0) * (double)invAzimuthStep + 1.f;
					ne = Nint(bne);
					if (ne >= 0 && ne <= nomentr) {
						vang = (vang2 - vang1) / (az2 - az1) * (delazi - az1) + vang1;
						proftmp = prof[ne].ver[0];
						if (vang > proftmp) {
							prof[ne].ver[0] = (float)vang; //viewangle
							prof[ne].ver[1] = (float)(begkmx + (i__ - 1) * skipkmx); //x coordinate
							prof[ne].ver[2] = (float)((kmy2 - kmy1) / (az2 - az1) * (delazi - az1) + kmy1); //y coordinate

							//fprintf(f,"%10.4f  %10.4f  %10.4f\n",prof[ne * 3 - 2],prof[ne * 3 - 1],prof[ne * 3 - 3]);
							}
						found = true; }
				} else if (az1 < delazi && az2 < delazi) {
					goto L170;}
L165:
				;
			}
L170:
			if (found) {
				goto L180; }
			++nstart2;
			if (nstart2 <= numdtm) {
				goto L160;
			} else {
				goto L180; }
		}
L180:
		++istart1;
		if (istart1 < numdtm) {
	    goto L150;	}

		numrctot ++;
		m_progresspercent = (int)((numrctot /  (float)numkmx) * 100);
		if ( m_progresspercent != oldpercent )
			{
			oldpercent = m_progresspercent;
			m_Progress.StepIt (); /*  step the progress bar */
			itoa(m_progresspercent, buffer, 10); /* convert m_prog... to char */
			strcat( buffer, "%");
			m_Label = buffer;    /* update the value of the progress bar */
			m_NewLabel.SetWindowText (m_Label);
			}

		//fprintf( stdout, "%s,%d,%lg,%lg\n", " completed interpreting: ",i__, kmx, kmy );
L550:
		;
    }

	////////////////diagnostics///////////////////
	if (testopened && diagnostics) 
	{
		fclose(streamout);
	}
	/////////////////////////////////////////


	char Header[255] = "";

	if (WhereJK != 0)
	{
		drivlet[0] = testdrv[WhereJK - 1];
		drivlet[1] = 0;
		bufFileName[0] = 0;
		sprintf(bufFileName, "%s%s", drivlet, ":\\jk_c\\eros.tmp"); 

		if ( f =fopen(bufFileName, "w" ) ) 
		{
			if (TemperatureModel == 1 || TemperatureModel == 3) { //new Wikipedia version of TR added, record the temprature used to calculate TR
				sprintf( Header, "%s,%6.2f,%6.2f","Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy", Tground, 0.0 );
				fprintf( f, "%s\n", Header );
				//fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR,1" );
			}else if (TemperatureModel == -1) { //old version
				fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR" );
			}else if (TemperatureModel == 0 || TemperatureModel == 2 || TemperatureModel == 4) { //no added terrestrial refraction, or TR removed
				fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR,0" );
			}
			//fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR" );

			////////////////////////////edit 101520/////////////////////////////////////////////////
			//record the height used for the calculation = hgt, and not the original ground height hgt0
			////////////////////////////////////////////////////////////////////////////////////////////
			fprintf( f, "%f,%f,%lg,%lg,%lg,%lg,%lg,%lg\n", lat0, -lon0, hgt, -startkmx, -sofkmx, dstpointx, dstpointy, apprnr);
			///////////////////////////////////end edit//////////////////////////////////////////////////////
		}
		else
		{
			reply = AfxMessageBox( "Can't open the profile file: c:/jk_c/eros.tmp!", MB_OK | MB_ICONSTOP | MB_APPLMODAL );		

		}
	}

/* 260     FORMAT(F9.4,2X,F9.4,2X,F8.2,4(3X,F8.3),3X,F5.3) */
    i__2 = nomentr;
    for (nk = 1; nk <= i__2; ++nk) {
		xentry = (float)(-maxang0 + (nk - 1) / 10.f);
	/*  find distances and heights to calculated horizon profile */
	/*  as a function of azimuth, and write to the output file */
	/*  called EROS.TMP */
		kmxp = prof[nk].ver[1];
		kmyp = prof[nk].ver[2];
	/*  IKMX = NINT((kmxp - (lgo+skipkmx*.5d0)) / skipkmx) + 1 */
	/*  IKMY = NINT(((lto -skipkmy*.5d0) - kmyp) / skipkmy) + 1 */
		ikmx = Nint((kmxp - beglog) / skipkmx) + 1;
		ikmy = Nint((endlat - kmyp) / skipkmy) + 1;
		numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1;
		hgt2 = 0;
		if (numrc >= 0 && numrc <= numtotrec ) 
		{
			hgt2 = (float)uni5[numrc].i5; //.chi5[0] << 8 | uni5[numrc].chi5[1]);
			if (fabs(hgt2) == 9999. || fabs(hgt2) == 32768. ) {
				hgt2 = 0.;} //ocean and radar voids
			else if ( fabs(hgt2) > 8848.) {
				//try rotating back
				i4sw[0].i4s = uni5[numrc].i5;
				uni5[numrc].chi5[0] = i4sw[0].chi6[1];
				uni5[numrc].chi5[1] = i4sw[0].chi6[0];
				hgt2 = uni5[numrc].i5;
				if (fabs(hgt2) == 9999. || fabs(hgt2) == 32768. || fabs(hgt2) > 8848.) {
					hgt2 = 0.;} //ocean and radar voids or glitches
			}
		}
		lg2 = kmxp;
		lt2 = kmyp;
		x1 = cos(lt * cd) * cos(-lg * cd);
		x2 = cos(lt2 * cd) * cos(-lg2 * cd);
		y1 = cos(lt * cd) * sin(-lg * cd);
		y2 = cos(lt2 * cd) * sin(-lg2 * cd);
		z1 = sin(lt * cd);
		z2 = sin(lt2 * cd);
		re = 6371.315;
		re1 = 1.;
		re2 = 1.;
		x1 = re1 * x1;
		y1 = re1 * y1;
		z1 = re1 * z1;
		x2 = re2 * x2;
		y2 = re2 * y2;
		z2 = re2 * z2;
	/* Computing 2nd power */
		d__1 = x1 - x2;
	/* Computing 2nd power */
		d__2 = y1 - y2;
	/* Computing 2nd power */
		d__3 = z1 - z2;
		dist = sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
		dist *= re;

		//now if flagged, remove the new type of terrestrial refraction addition
		/* fixed bug on 051825//////////////////////////////////////////////////////
		////////////////also -- now never adds TR if TemperatureModel == 2 || TemperatureModel == 4
		//so don't need to subtract later on
		///////////////////////////////////////////////////////////////////////
		if (TemperatureModel == 2 || TemperatureModel == 4) {
			//remove the terrestrial refraction addition
			deltd = hgt0 - hgt2;
			PATHLENGTH = sqrt(dist*dist + pow(fabs(deltd)*0.001 - 0.5*(distd*distd/RETH),2.0));
			if (fabs(deltd) > 1000.0) {
				Exponent = 0.99;
			}else{
				Exponent = 0.9965;
			}

			avref = TRpart * pow(PATHLENGTH, Exponent); //degrees
			prof[nk].ver[0] -= avref;
		}
		*/

		if (prof[nk].ver[0] > MinViewAngle + 0.1) //don't print out beyond range of azimuths that were found
		{
		   fprintf( f, "%lg,%lg,%f,%lf,%lg,%lg\n", xentry, prof[nk].ver[0], prof[nk].ver[1], prof[nk].ver[2], dist, hgt2);
		}

/* L300: */
    }
	fclose( f );

	if ( f = fopen( "c:/jk_c/erosend.tmp", "w" ))
	{
		fprintf(f, "%s\n", "finished" );
		fclose(f);
	}


   goto endProfiles;
}


void CNewreadDTMDlg::OnCancel() 
	{

	//deallocate memory
	free ( uni5);
	free ( vertices );
	free ( faces );
	free ( heights );

	//now write ending file for Maps and More
	char st2[] = "c:\\E130S30.BIN";
	int numclosed;

	numclosed = fcloseall(); //closed all opened files
	//look for jk or jk_c by looking for eros.tm3

	drivlet[0] = testdrv[WhereJK - 1];
	drivlet[1] = 0;

	char *jkdir,*enddir;
	strcpy(jkdir, drivlet);
	strcat(jkdir, ":\\jk_c\\eros.tm3");
	strcpy(enddir, drivlet);
	strcat(enddir, ":\\jk_c\\newreaddtm.end");

    if ( (stream = fopen( jkdir, "r" )) == NULL )
		{
		}
	else
		{
		fscanf( stream, "%s", st2);
		fclose( stream );
		//DeleteFile( "c:\\jk\\eros.tm3"); //erase the temporary eros.tm3 file
		}

    if ( (stream = fopen( (const char *)st2, "r" )) == NULL )
		{
		}
	else
		{
		fclose( stream );
		DeleteFile( st2 ); //kill the aborted .BIN file
		}

opca100: ;

	CDialog::OnCancel();
	// signal that data extraction was not successful
	stream2 = fopen( enddir, "w" );
	fprintf( stream2, "%d", 1 );
	fclose( stream2 );

	exit( 1 );
	}




////////////////////////////////////////////////
//    emulation of VB function Cint(x,y)      // 
////////////////////////////////////////////////
long Cint(double x)
{
	long nc1,nc2;
	nc1 = long(ceil(x));
	nc2 = long(floor(x));
	if ( fabs(x - nc1) <= fabs(x - nc2) )
		{
		return nc1;
		}
	else
		{
		return nc2;
		}
}
//////////////////////////////////////////////////

//////////////////////////////////////////////////
//   emulation of Fortran 95 Nint(double x)
//////////////////////////////////////////////////////////////
//source: http://www.lahey.com/docs/lfprohelp/F95ARNINTFn.htm
///////////////////////////////////////////////////////////////
int Nint(double x)
{
	if (x > 0)
	{
		return (int)(x + 0.5);
	}
	else if (x <= 0)
	{
		return (int)(x - 0.5);
	}
}

//////////////////////////////////////////////////
//function determine which drive contains /jk_c/eros.tm3
int WhichJK()
{
	FILE *stream;
	char drivlet[2] = "c";
	char jkdir[255] = "";
	int i;
	for (i=1; i<= strlen(testdrv); i++)
	{
		drivlet[0] = testdrv[i - 1];
		drivlet[1] = 0;
		strcpy(jkdir, (const char *)drivlet);
		strcat(jkdir, ":\\jk_c\\eros.tm3" );
		if ( (stream = fopen( jkdir, "r" )) != NULL )
		{
		   fclose(stream);
		   return i;
		}
	}
	/*
	char *jkdir;
	jkdir = "c:\\jk_c\\eros.tm3";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 1;
	jkdir = "e:\\jk_c\\eros.tm3";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 2;
	*/
	return 0; //return 0 if didn't find anything
}
////////////////////////////////////////////////////////////

//determine which drive contains /jk_c/mapcdinfo.sav
int WhichJK_2()
{
	FILE *stream;
	char testdrv[13] = "cdefghijklmn";
	char drivlet[2] = "c";
	char jkdir[255] = "";
	int i;
	for (i=1; i<= strlen(testdrv); i++)
	{
		drivlet[0] = testdrv[i - 1];
		drivlet[1] = 0;
		strcpy(jkdir, (const char *)drivlet);
		strcat(jkdir, ":\\jk_c\\mapcdinfo.sav" );
		if ( (stream = fopen( jkdir, "r" )) != NULL )
		{
		   fclose(stream);
		   return i;
		}
	}
	/*
	char *jkdir;
	jkdir = "c:\\jk_c\\mapcdinfo.sav";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 1;
	jkdir = "e:\\jk_c\\mapcdinfo.sav";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 2;
	*/
	return 0; //didn't find anything
}
//////////////////////////////////////////////////////////////EK 061322

//determine which drive contains /jk_c/mapSRTMinfo.sav
int WhichJK_3()
{
	FILE *stream;
	char testdrv[13] = "cdefghijklmn";
	char drivlet[2] = "c";
	char jkdir[255] = "";
	int i;
	for (i=1; i<= strlen(testdrv); i++)
	{
		drivlet[0] = testdrv[i - 1];
		drivlet[1] = 0;
		strcpy(jkdir, (const char *)drivlet);
		strcat(jkdir, ":\\jk_c\\mapSRTMinfo.sav" );
		if ( (stream = fopen( jkdir, "r" )) != NULL )
		{
		   fclose(stream);
		   return i;
		}
	}
	/*
	char *jkdir;
	jkdir = "c:\\jk_c\\mapcdinfo.sav";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 1;
	jkdir = "e:\\jk_c\\mapcdinfo.sav";
	if ( (stream = fopen( jkdir, "r" )) != NULL )
	   return 2;
	*/
	return 0; //didn't find anything
}
//////////////////////////////////////////////////////////////

short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] )
{

	//extract the WorldClim averaged minimum and average temperature for months 1-12 for this lat,lon
	//constants of the bil files

	int NROWS = 21600; // 'number of rows of the bil files
	int NCOLS = 43200; // 'number of columns of the bil files
	double XDIM = 8.33333333333333e-03; // 'longitude steps of bil files in degrees
	double YDIM = 8.33333333333333e-03; // 'latitude steps of bil files in degrees
	int NODATA = -9999; // 'no temp data flag of bil files
	double ULXMAP = -179.995833333333; // 'top left corner longitude of bil files
	double ULYMAP = 89.9958333333333; // 'top left corner latitude of bil files

	char FilePathBil[255] = "";
	char FileNameBil[255] = "";

    int tncols, IKMY, IKMX, numrec,pos, i;
	short IO = 0, Tempmode = 0;
	double X, Y;
	//short MinTemp[12], AvgTemp[12];

	FILE *stream;

	strcpy(FilePathBil,"c:\\DevStudio\\Vb\\WorldClim_bil");

	/*
	DIR* dir = opendir("FilePathBil");
	if (dir) {
		// Directory exists.
		closedir(dir);
	} else if (ENOENT == errno) {
		// Directory does not exist. 
		//return -1;
	} else {
		// opendir() failed for some other reason.
		return -1;
	}
	*/

	//first extract minimum temperatures, then average temperatures

T50:

	if (Tempmode == 0) { // read the minimum temperatures for this lat,lon
		strcat(FilePathBil,"//min_");
	}
	else if ( Tempmode == 1) { // read the average temperatures for this lat, lon
		strcpy(FilePathBil,"c:\\DevStudio\\Vb\\WorldClim_bil");
		strcat(FilePathBil,"//avg_");
	}
	else if ( Tempmode == 2) { // read the maximum temperatures for this lat, lon
		strcpy(FilePathBil,"c:\\DevStudio\\Vb\\WorldClim_bil");
		strcat(FilePathBil,"//max_");
	}

	for( i = 1; i<=12; i++ ) {
        
		strcpy(FileNameBil,FilePathBil);

		switch (i)
		{
			case 1:
			  strcat(FileNameBil, "Jan");
			  break;
			case 2:
			  strcat(FileNameBil, "Feb");
			  break;
			case 3:
			  strcat(FileNameBil, "Mar");
			  break;
			case 4:
			  strcat(FileNameBil, "Apr");
			  break;
			case 5:
			  strcat(FileNameBil, "May");
			  break;
			case 6:
			  strcat(FileNameBil, "Jun");
			  break;
			case 7:
			  strcat(FileNameBil, "Jul");
			  break;
			case 8:
			  strcat(FileNameBil, "Aug");
			  break;
			case 9:
			  strcat(FileNameBil, "Sep");
			  break;
			case 10:
			  strcat(FileNameBil, "Oct");
			  break;
			case 11:
			  strcat(FileNameBil, "Nov");
			  break;
			case 12:
			  strcat(FileNameBil, "Dec");
			  break;
		}

		strcat(FileNameBil,".bil");

		if ((stream = fopen( (const char*)FileNameBil, "rb" )) != NULL)
		{

			Y = lt;
			X = lg;
    
			IKMY = (int)((ULYMAP - Y) / YDIM) + 1;
			IKMX = (int)((X - ULXMAP) / XDIM) + 1;
			tncols = NCOLS;
			numrec = (IKMY - 1) * tncols + IKMX - 1;
			pos = fseek(stream,numrec * 2L,SEEK_SET);
			fread(&IO, 2L, 1, stream);

			if (IO == NODATA) IO = 0;
			if ( Tempmode == 0) {
				MinTemp[i - 1] = IO;
			} else if(Tempmode == 1) {
				AvgTemp[i - 1] = IO;
			} else if(Tempmode == 2) {
				MaxTemp[i - 1] = IO;
			}
        
			fclose(stream);

		}else{
			return -1;
		}
     
	}

  //now go back and calculate the AvgTemps
	if ( Tempmode == 0) {
       Tempmode = 1;
       goto T50;
	} else if ( Tempmode == 1) {
	   Tempmode = 2;
	   goto T50;
	}

	return 0;
}

///////////////RemoveCRLF/////////////////
char *RemoveCRLF( char *str )
///////////////////////////////////////
//remove the carriage return and line feed that typically end DOS files
///////////////////////////////////////////////
{
	short pos = 0;
	pos = InStr( str, "\r" );
	if (pos)
	{
	     str[pos - 1] = 0; //terminate the string at "\r"
	}

	pos = InStr( str, "\n" );
	if (pos)
	{
	     str[pos - 1] = 0; //terminate the string at "\r"
	}

	return str;
}

//////////////fgets_CR////////////////////////
char *fgets_CR( char *str, short strsize, FILE *stream )
///////////////////////////////////////////////
//fgets_CR for reading MS text lines (removes the CR)
///////////////////////////////////////////
{
	fgets(str, strsize, stream); //read in line of text
	RemoveCRLF( str );
	//Replace( str, '\r', '\0' ); //replace the CR with a character termination
	//RemoveCRLF( str ); //remove CR and LF at end of MS line
	return str;
}

///////////////////emulation of InStr(3 parameters)////////////////////
int InStr( short nstart, char * str, char * str2 )
{
	//emulates VB, i.e., substrings begin at position 1 and NOT at 0
	//so if found at beginning of string, returns 1 instead of 0

	char *pch;

	pch = strstr( str + nstart - 1, str2 ); //first occurance of str2 after nstart characters

	if (pch != NULL )
	{
		return pch-str+1; //returns value of position according to 1 = beginning
	}
	else
	{
		return 0; //return 0 for not being found
	}
}

//////////////////emulation of InStr(2 parameters) /////////////////////
int InStr( char * str, char * str2 )
{
	//works like VB, i.e., substrings begin at position 1 and NOT at 0
	return InStr( 1, str, str2 );
}

