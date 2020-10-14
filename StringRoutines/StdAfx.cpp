#ifdef _WIN32
	#define _WIN32_WINNT 0x0400 //<--comment out when not debugging
	#include <Windows.h>	//<--comment out for Linux
#endif

#include <math.h>
#include <stdio.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <string.h>
#include <stdlib.h>
#include <algorithm>
#include <iostream>
#include "dirent.h"
#include <ctype.h>
#include <errno.h>

/////////////////define macros, source: homeplanet///////////////////////
#define PI 3.14159265358979323846
#define fixangle(a) ((a) - 360.0 * (floor((a) / 360.0)))  /* Fix angle for degrees   */
#define fixangr(a)  ((a) - (PI*2) * (floor((a) / (PI*2))))  /* Fix angle in radians*/
#define dtr(x) ((x) * (PI / 180.0))                       /* Degree->Radian */
#define rtd(x) ((x) / (PI / 180.0))                       /* Radian->Degree */

/*  //<--uncomment for UNIX
#define __min(a,b) (((a)<(b))?(a):(b)) 
#define __max(a,b) (((a)>(b))?(a):(b))
*/ //<--uncomment for UNIX


//////////////////////global array bounds///////////////////////////////////////
#define MAXANGRANGE 80 //maximum degrees assuming 0.1 degrees steps
#define MAXCGIVARS 30 //maximum number of cgi variables
#define MAXDBLSIZE 100 //maximum character size of data elements to be read using ParseString
#define WIDTHFIELD 0 //0 for false - use predeter__mind html cell widths, = 1 for letting program deter__min it
#define INPUT_TYPE 1 //0 for post, 1 for get
#define MAX_NUM_TILES 15 //maximum possible number of tiles due to webmaster's memory limitiation for this thread
#define MAX_HEIGHT_DISCREP 1000 //maximum acceptable height discrepency between inputed and DTM's height at chosen coords.
#define BEG_SUN_APPROX 1950 //first year to use simplified ephem of sun
#define END_SUN_APPROX 2050 //last year to use simpliefied ephem of sun
#define MinGoldenDist 1 //minimum distance to Golden Horizon point to enable calculations
#define MaxGoldenViewAng 50 //maximum view angle to allow Golden Horizon calculations

//#define MAX_ANGULAR_RANGE 45 //maximum azimuth range from -MAX_ANGULAR_RANGE to MAX_ANGULAR_RANGE

//**********declare functions*****************/

////////////functions that emulate MS VB 6.0 and VC 6.0 functions//////////
//int isspace ( char c );
char *LTrim(char * p);
char *RTrim(char * p);
char *Trim(char * str);
char *Mid(char *str, short nstart, short nsize, char buff[]);
char *Mid(char *str, short nstart, short nsize);
int InStr( short nstart, char * str, char * str2 );
int InStr( char * str, char * str2 );
char *Replace( char *str, char Old, char New);
char *LCase(char *str); //, char buff[]);
char *UCase(char *str);
double Fix( double x );
short Cint( double x );
double Sgn(double a);
char *itoa ( int value, char *str, int base );
char *strlwr(char *str);
char *strupr(char *str);


///////////other functions/////////////////////
void hebnum(short k, char *hebnumch );
char *hebweek( short dayweek, char ch[] );
short RdHebStrings();
short RdHebErrorMessages();
short LoadParshiotNames();
void ParseSedra(short PMode);
void HebCal(short yrheb);
void InsertHolidays(char *caldayHeb, char *caldayEng, char *calday,
					char *calhebweek, short *i, short *j, short *k, short *yrn,
					short yl, short *ntotshabos, short *dayweek, short *fshabos,
					short *nshabos,	short *addmon, short iheb );
char *hebyear(char *yrcal, short yr);
short WriteTables(char TitleZman[], short numsort, short numzman,
				  double *avekmxzman, double *avekmyzman, double *avehgtzman,
				  short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[]);
void heb12monthHTML(short setflag);
void heb13monthHTML(short setflag);
void timestring( short ii, short k, short iii);
void parseit( char outdoc[], bool *closerow, short linnum, short mode);
short CheckInputs(char *strlat, char *strlon, char *strhgt);
short Caldirectories();
short SunriseSunset(short types, short ExtTemp[]);
short CreateHeaders(short types);
short ReadHebHeaders();
short PrintSingleTable(short setflag, short types);
short calnearsearch( char *batfile, char *sun);
short netzskiy( char *batfile, char *sun);
short WriteUserLog( int UserNumber, short zmanyes, short typezman, char *errorstr );
void ErrorHandler( short mode, short ier);
short YearLength( short myear);
short int height(long numrc);

short iswhite(char c);
short find_first_nonwhite_character(char *str, short start_pos);
short find_last_nonwhite_character(char *str, short start_pos);
short ParseData( char *doclin, char *sep, double arr[], short NUMENTRIES, short MAXARRsize );
short ParseString( char *doclin, char *sep, char strarr[][MAXDBLSIZE], short NUMENTRIES );
char *fgets_CR( char *str, short strsize, FILE *stream );
char *RemoveCRLF( char *str );

char *IsoToUnicode( char *str );
short LoadConstants(short zmanyes, short typezman, char filzman[] );
short LoadTitlesSponsors();
short PrintMultiColTable(char filzman[], short ExtTemp[] );
char *round( double tim, short stepps, short accurr, short sp, char t3subb[] );
char *EngCalDate(short ndy, short EngYear, short yl, char *caldate );
void PST_iterate( short *k, short *j, short *i, short *ntotshabos, short *dayweek
				 , short *yrn, short *fshabos, short *nshabos, short *addmon
				 , short setflag, short sp, short types, short yl, short iheb );
short WriteHtmlHeader();
short WriteHtmlClosing();
short decl_(double *dy1, double *mp, double *mc, double *ap, double *ac
	      , double *ms, double *aas, double *es, double *ob, double *d__);
short casgeo_(double *gn, double *ge, double *l1, double *l2);
short geocasc_(double *l1, double *l2, double *g1, double *g2);
double fnp_(double *p);
double fnm_(double *p);

//////////////////atmospheric refraction routines///////////////////////////////
//based on Menat type ray tracing
short refract_(double *al1, double *p, double *t, double *hgt
			  ,short *nweather, double *dy, double *refrac1, short *ns1, short *ns2);
//based on VDW ray tracing
short refvdw_(double *hgt, double *x, double *fref);
//////////////////////////////////////////////////////////////////////////////

void InitAstConst(short *nyr, short *nyd, double *ac, double *mc,
							   double *ec, double *e2c, double *ap, double *mp,
							   double *ob, short *nyf, double *td);
void AstronRefract( double *hgt, double *lt, short *nweather, short *PartOfDay,
				    double *dy, short *ns1, short *ns2,
					double *air, double *a1, short mint[], short avgt[], short maxt[],
					double *vdwsf, double *vdwrefr, double *vbweps, short *nyr);
void InitWeatherRegions( double *lt, double *lg, 
						 short *nweather, short *ns1, short *ns2,
						 short *ns3, short *ns4, double *WinTemp,
						 short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[]);

short ReadProfile( char *fileon, double *hgt, double *p, short *maxang, double *meantemp );

////////////////////////netzski6///////////////////////////////////////
short netzski6(char *fileon, double lg, double lt, double hgt, double aprn,
			               float geotd, short nsetflag, short nfil,
						   short *nweather, double *meantemp, double *p,
     		               short *ns1, short *ns2, short *ns3, short *ns4, double *WinTemp,
						   short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[]);

//////////////////////AstroNoon///////////////////////////////////////////////
void AstroNoon( double *jdn, double *td, double *dy1, double *mp, double *mc, double *ap,
		   double *ac, double *ms, double *aas, double *es, double *ob,
		   double *d__, double *tf, short *nyear, short *mday, short *mon, int *nyl, double *t6);

///////////Astronomical/Mishor sunset///////////////////////////////////////////////////////////
void AstronSunset(short *ndirec, double *dy, double *lr, double *lt, double *t6,
			 double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
			 double *es, double *ob, double *d__, double *air, double *sr1,
			 short *nweather, short *ns3, short *ns4, bool *adhocset, double *t3,
			 bool *astro, double *azi1, double *jdn, short *mday, short *mon,
			 short *nyear, int *nyl, double *td, short mint[], double *WinTemp,
			 bool *FindWinter, bool *EnableSunsetInv);
///////////Astronomical/Mishor sunrise///////////////////////////////////////////////////////////
void AstronSunrise(short *ndirec, double *dy, double *lr, double *lt, double *t6,
			 double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
			 double *es, double *ob, double *d__, double *air, double *sr1,
			 short *nweather, short *ns3, short *ns4, bool *adhocrise, double *t3,
			 bool *astro, double *azi1, double *jdn, short *mday, short *mon,
			 short *nyear, int *nyl, double *td, short mint[], double *WinTemp,
			 bool *FindWinter, bool *EnableSunriseInv);
////////////////////////VisRiseSet///////////////////////////////////////////////
short VisRiseSet(short *nloop, short *nlpend, short *ndirec, double *dy, double *t3,
				   double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
                   double *es, double *ob, double *lr, short *nsetflag, short *maxang,
				   double *stepsec, short *nyrstp, short *niter, short *ndystp,
				   short *nfil, double *sr1, double *p, double *t, double *hgt,
				   short *nweather, short *ns1, short *ns2, double *a1,
				   double *jdn, short *nyear, short *mday, short *mon, int * nyl, double *td);
/////////////////////////newzemanim//////////////////////////////////////////////////////
short newzemanim(short *yr, short *jday, double *air, double *lr, double *tf, double *dyo, double *hgto,
				 short *yro, double *td, double *dy, char *caldayHeb, short *dayweek,
				 double *mp, double *mc, double *ap, double *ac,  double *ms, double *aas,
				 double *es, double *ob, double *fdy, double *ec, double *e2c,
				 double *avekmxzman, double *avekmyzman, double *avehgtzman,
		         short *ns1, short *ns2, short *ns3, short *ns4, short *nweather,
				 short *numzman, short *numsort, double *t6, double *t3, double *t3sub99,
 				 double *jdn, int *nyl, short *mday, short *mon,
				 short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[]);
////////////////////////////cal (for zemanim)/////////////////////////////////////////////////////////
short cal(double *ZA, double *air, double *lr, double *tf, double *dyo, double *hgto, short *yro, double *td,
		  short *yr, double *dy, double *mp, double *mc, double *ap, double *ac, double *ms,
		  double *aas, double *es, double *ob, double *fdy, double *ec, double *e2c, short *ns1,
		  short *ns2, short *ns3, short *ns4, short *nweather, short *PartOfDay, 
		  double *avekmxzman, double *avekmyzman, double *avehgtzman, double *t6,
		  double *t3, double *t3sub, double *t3sub99, double *jdn, int *nyl, short *mday, short *mon, 
		  short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[]);

/////////////////ephem functions//////////////////////////////
double Decl3(double *, double *, double *, double *, short *, short *,
			 short *, int *, double *, double *, double *, double *,
			 double *, int);

////////////////WorldClim temperature files///////////////////////////////
short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] );

short DaysInYear(short *nyr);
double MonthFromDay(short *nyl, double *dy);

double obliqeq(double *jd);

double kepler(double *, double *);

static double J2000 = 2451545.0;
static double JulianCentury = 36525.0;  // Days in Julian century


/////////////////////////cgi functions////////////////////////////////////////////
short cgiinput(short *NumInputs);
char * getpair(short *NumInputs, char * name);
short RedirectToGeofinder();
////////////////html header diagnostics////////////////////////
short checkenviron();
////////////////html5 plot/////////////////
void PlotGraphHTML5(short nomentr, double vang_max, double vang_min, bool NearObstructions, short optang);
/////////////////////////////////////////////

/************file folder paths/weather modeling****************/
//eventually these will be stored in a file with a php interface
//with values for Windows and values for Linux
char drivjk[255] = ""; //path to hebrew strings
char drivweb[255] = ""; //path to web files
char drivfordtm[255] = ""; //path to fordtm (temp directory)
char drivcities[255] = ""; //path to cities directory
char drivdtm[255] = ""; //path to 30 m dtm directories
char drivtas[255] = ""; //path to 90 m dtm directory
char drivTemp[255] = "";//path to WorldClim 2.1 temperature bil files

//////////weather modeling////////////////
float latzone[5]; //four weather latitude zones excluding Eretz Yisroel
                  //used both for Northern Hemisphere, and mirror image for Southern Hemisphere
                  //latzone[0] = maximum latitude for tropical atmosphere with year long summer refraction
                  //latzone[1] = Mediterranian climate between latitudes latzone[0] and latzone[1]
                  //latzone[2] = mid northern climate between latitudes latzone[1] and latzone [2]
                  //latzone[3] = northern climate between latitudes latzone[3] to latzone[4]
				  //latzone[4] = far northern climate for latitudes larger than latzone[4]
short sumwinzone[5][4]; //weather zone winter summer ranges
                        //sumwinzones[0-4] correspond to latzones[0-4]
                        //sumwinzones[5] is for Eretz Yisroel
                        //sumwinzone[0] = end of winter refractions
                        //sumwinzone[1] = beginning of winter refraction
                        //sunwinzome[2] = end of ad hoc fix for winter refraction
                        //sumwinzone[3] = beginning of ad hoc fix for winter refraction
short nacurrFix;  //ad hoc shift of winter and summer time in seconds

//based on output from program Menat.for for winter and summer Northern Hemisphere mean-latitude atmospheres
//for heights of 1 - 999 meters.
double winrefo, sumrefo;
double sumref[8];
double winref[8];
double rearth = 6356.766f;
double vdweps[5];
double vdwref[6];
double vdwconst[4];
//short mint[12],avgt[12],maxt[12];

///////////////////////////////temperature model for vdw refraction///////////////////
//for VDW_REF = true, set the flag for which ground temperature is used for vdw refraction
short TempModel;  ; //1 -- minimum 30 arcsec WorldClim ground temperatures
					// 2 -- average 30 arcsec WorldClim ground temperatures
					// 3 -- maximum 30 arcsec WorldClim ground temperatures
////////////////////////////////////////////////////////////////////////////////////////

double WinterPercent; //filter out inversion layers during months that have a minimum temperature
					  //larger than this fraction of the average yearly minimum temperature
					 // this is an added filter to filter out negative view angle regions during hot parts of the yeaR
					 // or for hot climates

double HorizonInv;  //only add inversion adhoc correction if calculated view angle of sunrise/sunset is <= then HorizonInv 

bool VDW_REF = true;  //flag to use vand der Werf refraction instead of Menat atmospheres

double Pressure0;	//ground pressure

/////////astronomical constansts/////////////////
double ac0, mc0, ec0, e2c0, ob10, ap0, mp0, ob0;
short RefYear, ValidUntil;
//default values of astronomical constants are defined in routine LoadConstants

////////////////time cushions////////////////////
short cushion[5]; //EY, GTOPO30, SRTM-1, SRTM-2, astronomical calculations

/************************************************************/

//////////////multi-dimen.arrayrs//////////
//variables pertaining to netzski6
double dobs[2][2][366];
double tims[2][7][366];
double azils[2][2][366];

char heb1[85][255]; //[sunrisesunset] NOTICE: largest hebrew string is 99 characters long
char heb2[25][255];//[newhebcalfm]
char heb3[33][255];//[zmanlistfm]
char heb4[8][100]; //[hebweek] NOTICE: largest hebrew string is 20 characters long
char heb5[40][100];//[holidays] NOTICE largest hebrew string is 20 characters long
char heb6[10][100];//[candle_ligthing] NOTICE largest hebrew string is 20 characters long
char heb7[30][100];//Hebrew alphabet in Unicode
char heb8[25][255];//Hebrew error messages
char heb9[16][255]; //PlotGraphHTML5 Hebrew strings
char holidays[2][12][255]; //NOTICE: largest holiday string is 40 characters long
char arrStrParshiot[2][63][255]; //NOTICE: largest parshia name is 40 characters long
char arrStrSedra[2][63][255]; //NOTICE: Sedra length must equla Parshiot name
char stortim[2][13][31][40]; //actually stors times, parshios, days, etc.
                              //NOTICE: maximum length = length of Parshios
short yrend[2]; //the last day number of the civil years within the Hebrew Year
short yrstrt[2]; //the first day number of the civil year within the Hebrew Year
short FirstSecularYr = 0;  //first civil year within the Hebrew Year
short SecondSecularYr = 0; //second civil year within the Hebrew Year
short StartYr = 0;         //to be used to start the printed output at a certain civil year >= FirstSecularYr
char storheader[2][9][500]; //stores the header and footer strings
char monthh[2][15][50];  //Heberw month strings

/////////////////////zemanim arrays/////////////////////////////////
//they have an array size = maxzemanim
short maxzemanim = 50; //maximum number of zemanim that can be calculated
////////////////////////////////////////////////////////////////////

char zmannames[50][8][200]; //parsed elements of zemanim parameters
char zmantitles[50][200]; //sorted zman legend names
char c_zmantimes[50][9];  //contains sorted zemanim as time strings
double zmantimes[50];     //unsorted zemanim as fractional hours
short zmannumber[2][50];


//dynamic array that contains DTM data
//union elevat
//{
//	double vert[3];
//} elevat_tmp[1];

//union elevat *elev;

typedef struct _structElev
{
  double vert[4];
} structElev;
structElev elev[20 * MAXANGRANGE + 2]; //<--array size = MAXANGRANGE

/////////buffer of places to calculate times////////
short nplac[2] = {0, 0};
char fileo[2][200][255];
double latbat[2][200];
double lonbat[2][200];
double hgtbat[2][200];
double avekmx[2];
double avekmy[2];
double avehgt[2];

////////strings///////////////
char TitleLine[1255] = "";
char SponsorLine[8][1255];
char eroscity[255] = "";
char citnamp[255] = "";
char title[1255] = "";
char currentdir[255] = "";
char eroscountry[255] = "";
char erosareabat[255] = "";
char astname[255] = "";
char engcityname[255] = "";
char hebcityname[255] = "";
char EYhebNeigh[255] = "";
char MetroArea[255] = "";
char country[255] = "";
char TableType[5] = "";
char DomainName[255] = "";

/////////////////////global static string/buffers////////////////////////////
char doclin[1600] = "";
char buff[10000] = "";
char buff1[1600] = "";
char buff2[10000] = "";
char buff3[10000] = "";
char filroot[255] = "";
char myfile[255] = "";

//////global non string variables///////
short sortzman[50];
short datavernum = 0;
short progvernum = 0;
short yr = 5767;
float distlim = 0;
double cushion_factor = 1.0;
float distlims[4];
short parshiotflag = 0; //==0 for no parshiot
                        //==1 for Eretz Yisroel sedra (VB: parshiotEY = true)
                        //==2 for diaspora sedra (VB: parshiotdiaspora = true)
double eroslatitude = 0;
double eroslongitude = 0;
double eroshgt = 0;
double erosaprn = 0;
double erosdiflat = 1.f;
double erosdiflog = 1.f;
float searchradius = 3;
double pi, cd, dcmrad, pi2, ch, hr; //3.14..., conv deg to rad, mrad to degrees, 2*pi, hours/rad, min/hour
short fshabos0 = 0; //keeps tracks of first Shabbos
short rhday = 0; //keeps track of day of Rosh Hashono
short endyr = 12;

float geotz = 2.0;
int UserNumber = 0;

//////////boolean flags//////////////
short sunsyes = 0; //!= 0 prints individual sunrise/sunset tables
bool ast = false; //false for mishor, true for astronomical
bool astronplace = false; //= true for non eros non EY astronomical times
bool optionheb = false; // = true for Hebrew tables
bool eros = false; // = true for GTOPO/SRTM/EY neighborhoods
bool geo = false; // = true for using geo coordinates of latitude and longitude
bool Israel_neigh = false; // = true for Eretz Yisroel neighborhoods (needed to invert coordinates due to convention)
bool eroscityflag = false;
bool NearWarning[2]; //flag for near obstruction warning
bool InsufficientAzimuth[2]; //flag for insufficient Azimuth Range warning
short types = 0; //see below
short orig_types = 0; //records original value of types
short SRTMflag = 0; //0 for GTOPO30/SRTM30, 1 for SRTM-2, 2 for SRTM-1, 9 for EY neighborhoods
short PrintSeparate = false; //signals whether to print sunrise and sunsets as separate tables
                             // (if = true) or as column in zemanim output (if = false)
bool exactcoordinates = false; //= true to use the eros place's coordinates to calculate
                               //  the zemanim (rather than the average of the data points' coords)
bool g_IgnoreMissingTiles = false; //= true to ignore any missing tiles and use a default tile instead
								   //= false to return error code if any required tile is missing
short steps[2]; //round to steps seconds
short accur[2]; //cushions
short RoundSeconds = 0;
bool twentyfourhourclock = false; //use 24 hour clock ( if false, then use 12 hour clock)
bool abortprog = true; //= true to abortprog program if requested files not found/ = false not to abortprog rather to use defaults
bool wroteheader = false; //= true after successful exit of WriteHtmlHeader (used to avoid writing header twice)
bool wrotecloser = false; //= true after successful exit of WriteHtmlClosing (used to avoid writing close twice))
bool HebrewInterface = false; //= true when handling Hebrew versions of the webpages

bool Geofinder = false; //flag to redirect to chai_geofinder.php

bool IgnoreInsuffAzimuthRange = true; //true to return "*" if past the profile's azimuth range limit

bool AllowShavinginEY = false; //true to allow unlimited shaving for EY
short GlobalWarmingCorrection = 1; //shift latitude when determining  ground temmperature to reflect
									 //Global Warming

/////////////First Light flag -- new flag for Version 20 initiated on 091920/////////////////
short GoldenLight = 0; //0 for no GoldenLight calculation
					   //1 for GoldenLight calculation, first pass
					   //2 for GoldenLight calculation second pass

typedef struct _strucGold
{
  short nkk; //index of elev array corresponding to highest viewangle in the horizon profile
  double vert[6];
	//.vert]0] = azimuth
	// vert]1] = viewangle
	// vert[2] = longitude
	// vert[3] = latitude
	// vert[4] = distance
	// vert[5] = height
} structGold;
structGold Gold; 

///////////////////////////////////////////added 101420//////////////////////////////////////////
bool FixZmanHgts = false; //set to true to open each profile file to find the height instead ofusing
						   //the height recorded in the bat file, which may be in error by 1.8 meters
						   //This is typically uncessary if the zemanim are rounded to minutes
						   //The netz/skiya tables already querry the hgt recorded in the profile file
						   //independent of this flag
/////////////////////////////////////////////////////////////////////////////////////////////////

/////////////Hebrew Year/////////////////////
short g_yrheb = 0; //Hebrew Year in numbers
char g_yrcal[20] = ""; //Hebrew Year in Hebrew characters

/////output of HebYr calculations//////////
bool g_hebleapyear = 0;
short g_yeartype = 0;
short g_dayRoshHashono = 0;
short g_mmdate[2][13];  //contains civil day numbers corresponding to 1st and last day of the Hebrew month
					  //g_mmdate[0][i] = day number of 1st day of the Hebrew Month, i
					  //g_mmdate[1][i] = day number of last day of the Hebrew Month, i

char g_mdates[3][14][13]; //contains civil dates corresponding to 1st and last days of Hebrew months

////global constants for Hebrew Calendar///////////////
short g_n3mon = 793; // monthly change in molad after 1 lunation
short g_n2mon = 12;
short g_n1mon = 1;
short g_n3ylp = 589; // change in molad after 13 lunations of leap year
short g_n2ylp = 21;
short g_n1ylp = 5;
short g_n3yreg = 876; //change in molad after 12 lunations of reg. year
short g_n2yreg = 8;
short g_n1yreg = 4;
short g_n33 = 0;
short g_n22 = 0;
short g_n11 = 0;
short g_n3rh = 0;
short g_n2rh = 0;
short g_n1rh = 0;
short g_n1rhooo = 0;
short g_n1rho = 0;
short g_ncal3 = 0;
short g_ncal2 = 0;
short g_ncal1 = 0;
short g_n1rhoo = 0;
short g_nt3 = 0;
short g_nt2 = 0;
short g_nt1 = 0;
short g_rh2 = 0;
short g_leapyear = 0;
short g_yl = 0;
short g_dyy = 0;
short g_Hourr = 0;
short g_Min = 0;
short g_hel = 0;
char g_dates[12] = "";
char *g_gm = " PM afternoon";
short g_stdyy = 0;
short g_leapyr2 = 0;
short g_leapyr = 0;
char g_hdryr[] = "-2007";
short g_styr = 0;
short g_yrr = 0;

/************************Conversions from ITM to WGS84, etc*********************/
//*****************************************************************************************************
//*                                                                                                   *
//*  This code is free software; you can redistribute it and/or modify it at your will.               *
//*  It is my hope however that if you improve it in any way you will find a way to share it too.     *
//*                                                                                                   *
//*  If you have any comments, questions, suggestions etc. mail me (jgray77@gmail.com)                *
//*                                                                                                   *
//*  This program is distributed AS-IS in the hope that it will be useful, but WITHOUT ANY WARRANTY;  *
//*  without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.        *
//*                                                                                                   *
//*****************************************************************************************************
//
//
//===================================================================================================
//	Israel Local Grids <==> WGS84 conversions
//===================================================================================================
//
// The Israel New Grid (ITM) is a Transverse Mercator projection of the GRS80 ellipsoid.
// The Israel Old Grid (ICS) is a Cassini-Soldner projection of the modified Clark 1880 ellipsoid.
//
// To convert from a local grid to WGS84 you first do a "UTM to Lat/Lon" conversion using the
// known formulas but with the local grid data (Central Meridian, Scale Factor and False
// Easting and Northing). This results in Lat/Long in the local ellipsoid coordinate system.
// Afterwards you do a Molodensky transformation from this ellipsoid to WGS84.
//
// To convert from WGS84 to a local grid you first do a Molodensky transformation from WGS84
// to the local ellipsoid, after which you do a Lat/Lon to UTM conversion, again with the data of
// the local grid instead of the UTM data.
//
// The UTM to Lat/Lon and Lat/Lon to UTM conversion formulas were taken as-is from the
// excellent article by Prof.Steven Dutch of the University of Wisconsin at Green Bay:
//		http://www.uwgb.edu/dutchs/UsefulData/UTMFormulas.htm
//
// The [abridged] Molodensky transformations were taken from
//		http://home.hiwaay.net/~taylorc/bookshelf/math-science/geodesy/datum/transform/molodensky/
// and can be found in many sources on the net.
//
// Additional sources:
// ===================
// 1. dX,dY,dZ values:  http://www.geo.hunter.cuny.edu/gis/docs/geographic_transformations.pdf
//
// 2. ITM data:  http://www.mapi.gov.il/geodesy/itm_ftp.txt
//    for the meridional arc false northing, the value is given at
//    http://www.mapi.gov.il/reg_inst/dir2b.doc
//    (this doc also gives a different formula for Lat/lon -> ITM, but not the reverse)
//
// 3. ICS data:  http://www.mapi.gov.il/geodesy/ics_ftp.txt
//    for the meridional arc false northing, the value is given at several places as the
//    correction value for Garmin GPS sets, the origin is unknown.
//    e.g. http://www.idobartana.com/etrexkb/etrexisr.htm
//
// Notes:
// ======
// 1. The conversions between ICS and ITM are
//			ITM Lat = ICS Lat - 500000
//			ITM Lon = ICS Lon + 50000
//	  e.g. ITM 678000,230000 <--> ICS 1178000 180000
//
//	  Since the formulas for ITM->WGS84 and ICS->WGS84 are different, the results will differ.
//    For the above coordinates we get the following results (WGS84)
//		ITM->WGS84 32.11'43.945" 35.18'58.782"
//		ICS->WGS84 32.11'43.873" 35.18'58.200"
//      Difference    ~3m            ~15m
//
// 2. If you have, or have seen, formulas that contain the term Sin(1"), I recommend you read
//    Prof.Dutch's enlightening explanation about it in his link above.
//
//===================================================================================================
//double pi() { return 3.141592653589793; }
double sin2(double x) { return sin(x)*sin(x); }
double cos2(double x) { return cos(x)*cos(x); }
double tan2(double x) { return tan(x)*tan(x); }
double tan4(double x) { return tan2(x)*tan2(x); }

/////declare functions/////////
int Nint(double x);
int WhichJK();
int WhichJK_2();

/////declare global arrays//////////////

enum eDatum {
	eWGS84=0,
	eGRS80=1,
	eCLARK80M=2
};

typedef struct _DTM {
	double a;	// a  Equatorial earth radius
	double b;	// b  Polar earth radius
	double f;	// f= (a-b)/a  Flatenning
	double esq;	// esq = 1-(b*b)/(a*a)  Eccentricity Squared
	double e;	// sqrt(esq)  Eccentricity
	// deltas to WGS84
	double dX;
	double dY;
	double dZ;
} DATUM,*PDATUM;

static DATUM Datum[3] = {

	// WGS84 data
	{
		6378137.0,				// a
		6356752.3142,			// b
		0.00335281066474748,  	// f = 1/298.257223563
		0.006694380004260807,	// esq
		0.0818191909289062, 	// e
		// deltas to WGS84
		0,
		0,
		0
	},

	// GRS80 data
	{
		6378137.0,				// a
		6356752.3141,			// b
		0.0033528106811823,		// f = 1/298.257222101
		0.00669438002290272,	// esq
		0.0818191910428276,		// e
		// deltas to WGS84
		-48,
		55,
		52
	},

	// Clark 1880 Modified data
	{
		6378300.789,			// a
		6356566.4116309,		// b
		0.003407549767264,		// f = 1/293.466
		0.006803488139112318,	// esq
		0.08248325975076590,	// e
		// deltas to WGS84
		-235,
		-85,
		264
	}
};

enum gGrid {
	gICS=0,
	gITM=1
};

typedef struct _GRD {
	double lon0;
	double lat0;
	double k0;
	double false_e;
	double false_n;
} GRID,*PGRID;

static GRID Grid[2] = {

	// ICS data
	{
		0.6145667421719,			// lon0 = central meridian in radians of 35.12'43.490"
		0.55386447682762762,		// lat0 = central latitude in radians of 31.44'02.749"
		1.00000,					// k0 = scale factor
		170251.555,					// false_easting
		2385259.0					// false_northing
	},

	// ITM data
	{
		0.61443473225468920,		// lon0 = central meridian in radians 35.12'16.261"
		0.55386965463774187,		// lat0 = central latitude in radians 31.44'03.817"
		1.0000067,					// k0 = scale factor
		219529.584,					// false_easting
		2885516.9488				// false_northing = 3512424.3388-626907.390
									// MAPI says the false northing is 626907.390, and in another place
									// that the meridional arc at the central latitude is 3512424.3388
	}
};

// prototypes of local Grid <=> Lat/Lon conversions
void Grid2LatLon(int N,int E,double& lat,double& lon,gGrid from,eDatum to);
void LatLon2Grid(double lat,double lon,int& N,int& E,eDatum from,gGrid to);
// prototype of Moldensky transformation function
void Molodensky(double ilat,double ilon,double& olat,double& olon,eDatum from,eDatum to);
void itm2wgs84(int N,int E,double& lat,double &lon);
void wgs842itm(double lat,double lon,int& N,int& E);
void ics2wgs84(int N,int E,double *lat, double *lon);
void wgs842ics(double lat,double lon,int& N,int& E);
short readDTM( double *kmxo, double *kmyo, double *khgt, short *kmaxang,
			   double *kaprn, double *p, double *meantemp, short nsetflag );
/************************End Conversions section********************************/

/************************global variables used by Profile routine readDTM*************/
///////////global arrays and variables///////////////////////

struct unich //big endian short int heights read from DTM files
	{
	short int i5;
	//unsigned char chi5[2]; //break up the short int into its two bytes
	} uni5_tmp[1];

struct unich *uni5;

union i4switch //used to rotating little endian to big endian
	{
	short int i4s;
	unsigned char chi6[2]; //two bytes for a short integer
	} i4sw[1];

/*
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
*/

/*
union hgt
{
	short int height;
} hgt_tmp[1];

union hgt *heights;
*/

union profile
{
	float ver[3];
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
char (*filt1)[255] = new char[1][255];

double MaxWinLen;  //minimum length of day for adding inversion fix


/*************************************************************************************/

/********************************
This global is for CGI -- Yedidia
********************************/
char *vars[2][MAXCGIVARS]; //<--up to 2*MAXCGIVARS + 2 of inputs
bool CGI; //flag that deter__mins if using CGI input and output
bool Windows = false; //debugging flag, if Windows always write out output file when generating Astro profiles
/////////////////////////////////////////////////////////////////////

//////////////version number for 30m DTM tables/////////////
float vernum = 5.0f;
/////////////////////////////////////////////////////////////

int main(int argc, char* argv[])
{
	///////local variables for main program///////////////////////////////////
	short foundhebcity = 0;
	short zmanyes;
	short typezman;
	short numUSAcities1 = 0;
	short numUSAcities2 = 0;
	float searchradiusin = 0;
	short yesmetro = 0;
	short cnum = 0;
	char *eroslatitudestr = "";
	char *eroslongitudestr = "";
	char *eroshgtstr = "";
	char *erosaprnstr = "";
	char *erosdiflatstr = "";
	char *erosdiflogstr = "";
	double eroslatitudein = 0;
	double eroslongitudein = 0;
	double eroshgtin = 0;
    short RoundSecondsin = 0;
	char eroscityin[50] = "";
	double lat, lon, hgt;
    	//bool aveusa = true; //invert coordinates in SunriseSunset due to
                        //silly nonstandard coordinate format
	bool viseros = false;
	static char filzman[255] = "";
	short ier = 0;
	short optionhebnum = 0;
	short NumCgiInputs = 0;
	short foundcity = 0;

	////////////////////if first entry is not = -9999, then using external temperatures inputed by user/////
	////////////////////to calculat the vdw scaling factors//////////////////////////////////////
	short ExtTemp[12] = {-9999,0,0,0,0,0,0,0,0,0,0,0};
	///////////////////////////////////////////////////////////////////////////////////////////

	int N, E;

    FILE *stream2;

////////////////////////debug handle///////////////////////////////////////
/*//<--comment out in order to debug CGI
#ifdef _DEBUG
char szMessage [256];
wsprintf (szMessage, "Please attach a debugger to the process %d (%s) and click OK",
      GetCurrentProcessId(), argv[0]);
MessageBox(NULL, szMessage, "CGI Debug Time!",
      MB_OK|MB_SERVICE_NOTIFICATION);
__asm{
	__emit 0xCC
	__emit 0xCC
}
#endif
*///<--comment out in order to debug CGI

//================================================================================
	////////////////process inputed parameters for php pages///////////////
	if (argc != 26) // && argc != 1 ) //<--comment out the "&& argc != 1) for cgi, uncomment to debug
	{
		/* working as CGI  */

		if (cgiinput( &NumCgiInputs )) //something wrong with cgi
		{
			strcpy( drivweb, "/home/chkeller/domains/chaitables.com/public_html/cgi-bin/" ); //default path to log file directory
			ErrorHandler(0,0);
			return -1;
		}

		try
		{
			strcpy( TableType, Trim(getpair(&NumCgiInputs, "cgi_TableType")) ); //Type of Table, "Astr", or "BY", "Chai"
			strcpy( country, Trim(getpair(&NumCgiInputs, "cgi_country")) ); //country name
			strcpy( MetroArea, Trim(getpair(&NumCgiInputs, "cgi_MetroArea")) ); //Metro area name
			ier = 1;
		}
        catch( short ier )
        {
            //now give error warning that place wasn't picked prooperly
            ErrorHandler(0, ier);
            exit(-1);
        }

		UserNumber = atoi(getpair(&NumCgiInputs, "cgi_UserNumber")); //usernumber
		g_yrheb = atoi(getpair(&NumCgiInputs, "cgi_yrheb")); //Hebrew Year

		typezman = atoi(getpair(&NumCgiInputs, "cgi_typezman")); // acc. to which opinion
		zmanyes = 0;
		if (typezman != -1) zmanyes = 1;
		optionhebnum = atoi(getpair(&NumCgiInputs, "cgi_optionheb")); //flag for Hebrew;
		optionheb = true;
		if (optionhebnum) optionheb = false;
		RoundSecondsin = atoi(getpair(&NumCgiInputs, "cgi_RoundSecond"));

		///unique to Chai tables
		numUSAcities1 = atoi(getpair(&NumCgiInputs, "cgi_USAcities1")); //Metro area
		numUSAcities2 = atoi(getpair(&NumCgiInputs, "cgi_USAcities2")); //city area
		if (numUSAcities2 == 0)
		{
			yesmetro = 1; //will be entering coordinates
		}
		else
		{
			yesmetro = 0; //found city area
		}
		
		//Place Naee, test lat,lon, hgt inputs
		strcpy ( eroscityin, Trim(getpair(&NumCgiInputs, "cgi_Placename")) ); //Place Name
		//remove the added underlines
		Replace(eroscityin, '_', ' ' );

		eroslatitudestr = getpair(&NumCgiInputs, "cgi_eroslatitude"); //as inputed from cgi as strings
        eroslongitudestr = getpair(&NumCgiInputs, "cgi_eroslongitude"); //"21.89521";
		eroshgtstr = getpair(&NumCgiInputs, "cgi_eroshgt"); //"83.2";
		// ***************TO DO ***************
		erosaprnstr = getpair(&NumCgiInputs, "cgi_erosaprn");
		erosdiflatstr = getpair(&NumCgiInputs, "cgi_erosdiflat");
		erosdiflogstr = getpair(&NumCgiInputs, "cgi_erosdiflon");
        // ***************TO DO ***************

		if (!strstr(Trim(getpair(&NumCgiInputs, "cgi_searchradius")), "Not found") )
		{
    		searchradiusin = atof(getpair(&NumCgiInputs, "cgi_searchradius")); //serach radius;
        }

		geotz = atof(getpair(&NumCgiInputs, "cgi_geotz"));//-8;

		////////////table type (see below for key)///////
		types = atoi(getpair(&NumCgiInputs, "cgi_types")); //8;

		sunsyes = 0; //default = don't print individual sunrises or sunsets
		if (types != -1) sunsyes = 1; //print individual sunrise or sunset tables
		////////////////////////////////////////////////

		twentyfourhourclock = true;
		if (strstr(Trim(getpair(&NumCgiInputs, "cgi_24hr")), "Not found") )
		{
           //checkbox is OFF so no cgi input returned for cgi_zmanyes
		   //Therefore getpair returns "Error - Not found"
		   twentyfourhourclock = false;
		}

		exactcoordinates = true;
		if (strstr(Trim(getpair(&NumCgiInputs, "cgi_exactcoord") ), "Not found") || strstr(Trim(getpair(&NumCgiInputs, "cgi_exactcoord")), "OFF" ) )
		{
						//= false, use averaged data points' coordinates
						//= true - use the eroslatitude, eroslongitude, eroshgt
			exactcoordinates = false;
		}

		g_IgnoreMissingTiles = true;
		if (strstr(Trim(getpair(&NumCgiInputs, "cgi_ignoretiles") ), "Not found") || strstr(Trim(getpair(&NumCgiInputs, "cgi_ignoretiles")), "OFF" ) )
		{
						//= false, use averaged data points' coordinates
						//= true - use the eroslatitude, eroslongitude, eroshgt
			g_IgnoreMissingTiles = false;
		}

		AllowShavinginEY = true;
		if (strstr(Trim(getpair(&NumCgiInputs, "cgi_AllowShaving") ), "Not found") || strstr(Trim(getpair(&NumCgiInputs, "cgi_AllowShaving")), "OFF" ) )
		{
						//= false, use averaged data points' coordinates
						//= true - use the eroslatitude, eroslongitude, eroshgt
			AllowShavinginEY = false;
		}


		if ( strstr(Trim(getpair(&NumCgiInputs, "cgi_Language")), "Hebrew") )
           HebrewInterface = true; //Hebrew Interface flagged

		CGI = 1;

   }
/*//<-- leave only "/*" for CGI, comment out the "/*" -> "///*" for debugging
	else if (argc == 1) //not using console
	{
		strcpy( TableType, "BY"); //"Astr"); //"BY");//"Chai" ); //"BY" ); //"Chai" );//"Astr" );//"Chai" )//"BY" );
		yesmetro = 0;
		strcpy( MetroArea, "jerusalem");//"beit-shemes"); //"jerusalem"); //"London"); //"jerusalem"); //"Kfar Pinas");//"Mexico");//"jerusalem"); //"almah"); //"jerusalem" ); //"telz_stone_ravshulman"; //"???_????";
		strcpy( country, "Israel");//"England"); //"Israel");//"Mexico");//"Israel" ); //"USA" ); //"Reykjavik, Iceland" ); //"USA" );//"Israel";
		UserNumber = 302343;
		g_yrheb = 5781; //5779;//5776; //5775;
		zmanyes = 1; //0; //1;//1; //0; //1; //0 = no zemanim, 1 = zemanim
		typezman = 8; // acc. to which opinion
		optionheb = true;//false;
		RoundSecondsin = 1;//5;

		///unique to Chai tables
		numUSAcities1 = 3;//1; //2;//11; //2; //Metro area
		numUSAcities2 = 2; //7;//0; //75; //city area
		searchradius = 1;//4;//1; //search area
		yesmetro = 1; //0;  //yesmetro = 0 if found city, = 1 if didn't find city and providing coordinates
		GoldenLight = 1;

		//test lat,lon, hgt inputs 32.027585
		eroslatitudestr = "31.706938"; //"30.960207"; //"37.683140"; //"30.92408"; //"31.858246"; //"31.789";//"51.472"; //"31.820158";//"31.815051"; //"31.8424065";//"52.33324";//"32.027585";//"32.483266";//"19.434381";//"32.08900"; //"40.09553"; //"32.08900"; //"31.046051"; //31.864977"; //31.805789354"; //"31.806467"; //"64.135338"; //as inputed from cgi as strings
		eroslongitudestr = "-35.123139"; //"-35.195719"; //"119.170837"; //"-34.96542"; //"-35.158354"; //"-35.217";//"0.2074"; //"-35.201774";//"-35.217059"; //"-35.247203";//"1.11254";//"-34.841867";//"-35.004566";//"99.201608";//"-34.83177";//"74.222"; //"-34.83177"; //"-34.851612"; //-35.164481"; //-35.239226491"; //"-35.239291"; //"21.89521";
		eroshgtstr = "787.23"; //"29.70"; //"2843.73"; //"495.0"; //"644.26"; //"800";//"0";//"684.3";//"164.1";//"26.93";//"68.69";//"2274.5"; //"63.5";//"22.72"; //"63.5";//"788.19"; //832.3"; //"832"; //"83.2";
		erosaprnstr = "0.5";//"1.5";//"0.5";
		erosdiflatstr = "1";
		erosdiflogstr = "1"; //"2";


		searchradiusin = 1;//4; //0.9;
		geotz = 2; //-8; //2;//0;//-6; //2;//-5; //2; //0;//-8;
		sunsyes = 1;//0; //1;//0; //= 1 print individual sunrise/sunset tables
					 //= 0 don't print individual sunrise/sunset tables

		///////////////////astronomical type//////////////////////
		ast = true;//false; //= false for mishor, = true for astronomical
		//////////////////////////////////////////////////////////

		////////////table type (see below for key)///////
		types = 0; //11; //13; //1;//10; //0; //10;//0; //5;//0;//2; //3; //10;//10; //0;//10;//2; //10;//10;//0;
		SRTMflag = 0; //11;//13; //0;//10; //0; //10; //10;//10; //0;//10;//0; //10; //10;
		////////////////////////////////////////////////

    	exactcoordinates = true; //false;  //= false, use averaged data points' coordinates
								   //= true - use the eroslatitude, eroslongitude, eroshgt
        twentyfourhourclock = false; //true; //=true - use 24 hour clock
		  						 //false,- use 12 hour clock)

		g_IgnoreMissingTiles = true; //= true to ignore missing tiles
									  // = false to return error message if tile is missing

		HebrewInterface = false;// true;//false; //=true for Hebrew webpages
	}
//*/
   else if ( argc == 27 ) //console
	{

        strcpy( TableType, Trim(argv[1])); //Type of Table, Astr, or BY, or Chai
		strcpy( MetroArea, argv[2] ); //English city name of Israeli city, e.g.: "afulah"
		UserNumber = atoi(argv[3]);
		g_yrheb = atoi(argv[4]);
		//zmanyes = atoi(argv[5]);
		typezman = atoi(argv[5]);
		zmanyes = 0; //default = don't print zemanim tables
		if (typezman != -1) zmanyes = 1; //print zemanim tables

		/////////Hebrew or English table/////////////////
		short optionhebs = atoi(argv[6]);
		optionheb = true;
		if (optionhebs == 0 ) optionheb = false;
		/////////////////////////////////////////////////

        RoundSecondsin = atoi(argv[7]);

		///unique to Chai tables
		numUSAcities1 = atoi(argv[8]); //Metro area
		numUSAcities2 = atoi(argv[9]); //city area
/*
		yesmetro = atoi(argv[10]);  //yesmetro = 0 if found city, = 1 if didn't find city and providing coordinates
*/

		//Placename, test lat,lon, hgt inputs
		strcpy( eroscityin, Trim(argv[10]) );
		eroslatitudestr = argv[11]; //as inputed from cgi as strings
		eroslongitudestr = argv[12];
		eroshgtstr = argv[13];

		searchradiusin = atoi(argv[14]);
		UserNumber = atoi(argv[15]);
		geotz = atoi(argv[16]);
	    /*
		sunsyes = atoi(argv[19]); //= 1 print individual sunrise/sunset tables
					 //= 0 don't print individual sunrise/sunset tables
      	 */

		///////////////////astronomical type//////////////////////
		short asts = atoi(argv[17]); //= false for mishor, = true for astronomical
		ast = true;
		if (asts == 0) ast = false;
		//////////////////////////////////////////////////////////

		////////////table type (see below for key)///////
		types = atoi(argv[18]);
		sunsyes = 0; //default: don't print individual sunrise/sunset tables
		if (types != -1) sunsyes = 1; //print individual sunrise/sunset tables
		////////////////////////////////////////////////

	    exactcoordinates = false;
		if (atoi(argv[19])) exactcoordinates = true;
		                           //= false, use averaged data points' coordinates
								   //= true - use the eroslatitude, eroslongitude, eroshgt
	    twentyfourhourclock = true;
		if (!atoi(argv[20])) twentyfourhourclock = false;
		                           //=true - use 24 hour clock
								   //=false,- use 12 hour clock)
		g_IgnoreMissingTiles = false;
		if (!atoi(argv[21])) g_IgnoreMissingTiles = true;
								   //=true - ignore missing tiles
								   //=false - return error code if tile is missing	

        if ( strstr( Trim(argv[21]), "Hebrew") )
           HebrewInterface = true; //Hebrew Interface flagged

	}
///////////////////////////////end of input section////////////////////////////////////////

	//////////eros, astro type//////////////////////////
	if (!strcmp(TableType, "Chai")) eros = true;
	if (!strcmp(TableType, "Astr")) astronplace = true;
	/////////////////////////////////////////////

    /////////////////LoadConstants////////////////////////////////
	ier = LoadConstants(zmanyes, typezman, filzman); //load constants and paths
	if (ier) //missing essential file -- severe error flagged
	{
		ErrorHandler(1, ier );
		return -1;
	}

	////////////////////////////////////////////////////////
	PrintSeparate = true; //not implemented yet -- will be used to plot table
	                      //of sunrise/sunset/zemanim over a few days only
	///////////////////////////////////////////////////////

	if (HebrewInterface) //read and store Hebrew error messages
	{
		if (RdHebErrorMessages())
		{
			HebrewInterface = false; //error means that can't load Hebrew, so use English defaults
			optionheb = false;
		}
	}

	/////////////////////finish loading strings////////////////////////////////
	if (LoadTitlesSponsors() ) //load Titles and Sponsor information
	{
		ErrorHandler(13, 1); //SponsorsLogo.txt not found
		return -1;
	}
	//////////////////////////////////////////////////////////////////////

    //////////////check if user selected a table to calculate///////////

    if (sunsyes == 0 && zmanyes == 0)
    {
        ErrorHandler(0,1); //no tables chosen
        return -1;
    }

    orig_types = types; //record original values of types before processing

/*
	switch (types)
	{
	case 0:
	  nsetflag = 0; //visible sunrise/astr. sunrise/chazos
	  break;
	case 1:
	  nsetflag = 1; //visible sunset/astr. sunset/chazos
	  break;
	case 2:
	  nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
	  break;
	case 3:
	  nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
	  break;
	case 4:
	  nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
	  break;
	case 5:
	  nsetflag = 5; //visible sunset/ mishor sunset/ chazos
	  break;
	case 6:
	  nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
	  break;
	case 7:
	  nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
	  break;
	case 8:
	  nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
	  break;
	case 10:
      nsetflag = 10; // visible sunrise calculated using 30 m DTM
      break;
    case 11:
      nsetflag = 11; //visible sunset calculated using 30 m DTM
      break;
	default:
	  {
	  ErrorHandler( 1, 2, stream2);; //uncrecogniable "types"
	  return -1;
	  }
	}

	nsetflag = types;

/*       Nsetflag = 0 for astr sunrise/visible sunrise/chazos */
/*       Nsetflag = 1 for astr sunset/visible sunset/chazos */
/*       Nsetflag = 2 for mishor sunrise if hgt = 0/chazos or astro. sunrise/chazos */
/*       Nsetflag = 3 for mishor sunset if hgt = 0/chazos or astro. sunset/chazos */
/*       Nsetflag = 4 for mishor sunrise/vis. sunrise/chazos */
/*       Nsetflag = 5 for mishor sunset/vis. sunset/chazos */
/*       Nsetflag = 10 for visible sunrise calculated using 30 m DTM */
/*       Nsetflag = 11 for visible sunset calculated using 30 m DTM */

	if (CGI) //use special definitions for types to define astronomical times
			//corresponding to form information sent from the html page
	{
		switch (types)
		{
			case -1: //no individual sunrise/sunset tables chosen
				break;
			case 0: //visible sunrise
			case 1: //visible sunset
				break;
			case 2: //mishor sunrise/chazos
				ast = false;
				break;
			case 3: //astronomical sunrise/chazos
				ast = true;
				types = 2;
				break;
			case 4: //mishor sunset/chazos
				ast = false;
				types = 3;
				break;
			case 5: //astronomical sunset/chazos
				ast = true;
				types = 3;
				break;
			case 6:
			case 7:
			case 8:
				break;
		    case 10: //visible sunrise calculated from 30m DTM
                ast = true; //use inputed heights for calculation
                SRTMflag = 10; //30 m DTM flag, i.e., JK's DTM in Israel, NED in the US
                break;
            case 11: //visible sunset calculated from 30m DTM
                ast = true; //use inputed heights for calculation
                SRTMflag = 11; //30 m DTM flag, i.e., JK's DTM in Israel, NED in the US
                break;
		    case 12: //visible sunrise/sunset and GoldenLight calculated from 30m DTM
			case 13:
				//first calculate the sunset profile to find the highest point on the Western horizon
                ast = true; //use inputed heights for calculation
                SRTMflag = types; //30 m DTM flag, i.e., JK's DTM in Israel, NED in the US
                break;
			default:
				{
				ErrorHandler( 1, 2 ); //uncrecogniable "types"
				return -1;
				}
		}
	}

	//find version number of netzski6 calculations//////////////
	sprintf( filroot, "%s%s", drivcities, "version.num" );

	if ( (stream2 = fopen( filroot, "r" ) ) )
	{
		fscanf( stream2, "%d\n", &datavernum );
        fclose( stream2 );
	}
	else //can't find version file, use default
	{
		datavernum = 15;
	}

///////////////////////////////end of input parameters/////////////////////////

	//check the inputs for errors --- inputs will be strings that need to be converted, etc
	//before converting them, check that they are really numbers and that the numbers
	//fall between certain values
	if ( ( yesmetro == 1 && InStr(TableType,"Chai") ) || InStr(TableType,"Astr") ) //inputing coordinates so check them
	{

		if (CheckInputs( eroslatitudestr, eroslongitudestr, eroshgtstr))
		{
            ErrorHandler(1, 2); //something wrong with inputs <<<<<<<<<<<<<<
			return -1;
		}

		//now convert from the input string to the values
		eroslatitudein = atof( eroslatitudestr);
		eroslongitudein = atof( eroslongitudestr);
		eroshgtin = atof(eroshgtstr);
		erosaprn = atof(erosaprnstr);
		erosdiflat = atof(erosdiflatstr);
		erosdiflog = atof(erosdiflogstr);

		//set limits for erosdiflog, erosdiflat
		erosdiflat = __max(1.0, erosdiflat );
		erosdiflat = __min(2.0, erosdiflat );

		erosdiflog = __max(1.0, erosdiflog );
		erosdiflog = __min(3.0, erosdiflog );

    }

	RoundSeconds = RoundSecondsin;

	if ( ( (strstr(TableType, "Chai") != NULL) || (strstr(TableType, "Astr") != NULL) ) && (RoundSecondsin == -1) )
	{
		RoundSeconds = 5; //default RoundSeconds for Chai and Astr tables
	}

	steps[0] = RoundSeconds;
	steps[1] = RoundSeconds;

///////////////////deter__min which table will be generated///////////////////////

	//short langMetro = 0;

	if ( InStr( TableType, "BY") ) //Eretz Yisroel tables based on 25 m DTM
	{
/*
		//if ( (char)Mid( MetroArea, 1, 1, buff) < (char)(128) ) //English table
		if (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", Mid( MetroArea, 1, 1, buff)) ) //English table
		{
			langMetro = 0;
		}
		else // if ( (char)Mid( MetroArea, 1, 1, buff) >= (char)(128) )  //Hebrew table
		{
			langMetro = 1;
		}
*/
		///////////////////////////////////////////////////////////////
		//find English and Hebrew equivalents of directory name
		sprintf( filroot, "%s%s", drivcities, "citynams.txt");

		if ( !( stream2 = fopen( filroot, "r")) )
		{
			ErrorHandler( 1, 3); //report can't open citynams.txt file
			return -1;
		}
		else
		{
			while (!feof(stream2) )
			{
			fgets_CR(doclin, 255, stream2); //read in line of text
			strcpy( engcityname, doclin );

			fgets_CR(doclin, 255, stream2); //read in line of text
			strcpy( hebcityname, doclin );

			if ( ( !HebrewInterface && !strcmp(MetroArea, engcityname) ) ||
				  (  HebrewInterface && !strcmp(MetroArea, hebcityname) )    )
				{
				sprintf( currentdir, "%s%s%s", drivcities, engcityname, "/" );
				foundhebcity = 1;
				break;
				}
			}

			fclose( stream2 );
			if (!foundhebcity ) //Hebrew city equivalent not found
			{
                strcpy( citnamp, MetroArea ); //record city for log file
				ErrorHandler(1, 4);
				return -1;
			}

		}

	    strcpy( eroscountry, "Israel" );

		eros = false;
		geo = false;
		geotz = 2;

		//remove the underlines in the city names
		Replace( engcityname, '_', ' ');
		Replace( hebcityname, '_', ' ');

		//captilize first letter
		strupr( engcityname ); //capitalize first letter of English name
		strlwr( engcityname + 1); //lower case for rest of string

		//if name has space, then capatalize letter after space
		char * pch = NULL;
		char str[2] = " ";

		pch = strstr( engcityname, " ");
		if (pch != NULL )
		{
		  strncpy( &str[0], pch + 1, 1);
		  strupr( &str[0] );
		  strncpy( pch + 1, &str[0], 1);
		}

	    strcpy( citnamp, engcityname );

        parshiotflag = 1; //Eretz Yisroel parshiot

	}
	else if ( InStr( TableType, "Chai") ) //World-wide tables based on SRTM
	{

        parshiotflag = 2; //diaspora parshiot
		strcpy( eroscountry, country );
		eros = true;
		geo = true;
		eroscityflag = true;
		//now read in parameters
		eroscity[0] = '\0';
		eroslatitude = eroslatitudein;
		eroslongitude = eroslongitudein;
		eroshgt = eroshgtin;
		searchradius = searchradiusin;

		if (InStr( eroscountry, "Israel") && HebrewInterface ) //Israel neighborhoods in Hebrew
		{
			sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities6.dir");
			//Israelcities6.dir is the list of Israeli metro areas sorted according to the order of Israelcities1.dir
		}
		else
		{
			sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities1.dir");
		}

		/* write closing clauses on html document */
		if ( !( stream2 = fopen( filroot, "r")) )
		{
			ErrorHandler( 1, 5);; //can't open cities1 file
			return -1;
		}
		else //find English Metro name laving list number: numUSAcities1
		{
			cnum = 0;
			while ( !feof(stream2) )
			{
				fgets_CR(doclin, 255, stream2); //read in line of text
				cnum++;
				if (cnum == numUSAcities1) break; //reached the metro area, so break
			}

			fclose( stream2 );

			if (cnum < numUSAcities1)
			{
				ErrorHandler( 1, 6); //never reached requested metro area
				return -1;
			}
		}

		//this will define the eroscity name unless flagged by
		//yesmetro = 1 to use the inputed placename and coordinates
		strcpy( erosareabat, doclin );

		//remove _area_state_country to form the english Metro area name
		//engcityname = portion of ersareabat before "_area"
		Mid( erosareabat, 1, InStr(erosareabat, "_area") - 1, engcityname );
		Replace(engcityname, '_', ' ' ); //remove any "_" in the engcityname

		if (Geofinder) //allready collected all the info needed to redirect to chai_geofinder.php
		{
			ier = RedirectToGeofinder();
		}

		if ( yesmetro == 0 ) // Then 'found the city in the city list, use it
		{

			//now find eroslongitude, eroslatitude

			sprintf( myfile, "%s%s%s%s%s", "[", erosareabat, "_", eroscountry, "]" ); //header of metro area in cities5/8.dir

			if (InStr( eroscountry, "Israel") && HebrewInterface ) //Israel neighborhoods in Hebrew
			{
				sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities8.dir");
				//Israelcities6.dir is list of Israeli neigh. areas sorted according to the order of Israelcities5.dir
			}
			else
			{
				sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities5.dir");
			}
			/* write closing clauses on html document */
			if ( !(stream2 = fopen( filroot, "r")) )
			{
				ErrorHandler( 1, 7);; //can't open cities5 file
				return -1;
			}
			else //find eroscity's coordinates
			{
				foundcity = 0;
				while (!feof(stream2) )
				{
					fgets_CR( doclin, 255, stream2 );

					if ( strcmp(myfile, doclin) == 0  )
					{
						//read numUSAcities2 lines to find neighborhood/city
						cnum = 0;
						while ( cnum < numUSAcities2 )
						{
							fgets_CR(doclin, 255, stream2); //read in line of text

							//check that haven't passed header for next Metro area
							if ( (strstr( doclin, "[") != NULL) && (strstr( doclin, "]") != NULL ) )
							{
								//already reached next metro area
								ErrorHandler( 1, 9);; //never found requested city/neighborhood
								return -1;
							}
							cnum++;
						}

						//now extract city name and coordinates
						const short NUMENTRIES = 4;
						char strarr[NUMENTRIES][MAXDBLSIZE];

						if ( ParseString( doclin, ",", strarr, NUMENTRIES ) )
						{
							ErrorHandler( 1, 8);; //can't read line of text in cities5 file
							return -1;
						}
						else //extract text
						{
							short citylen = strlen(&strarr[0][0]);
							strncpy( eroscity, &strarr[0][0], citylen - 1);
							eroscity[citylen] = 0;
							eroslatitude = atof( &strarr[1][0] );
							eroslongitude = atof( &strarr[2][0] );
							eroshgt = atof( &strarr[3][0] );
							foundcity = 1;

							if (InStr( eroscountry, "Israel"))
							{
								//remove any underlines: "_" from city name
								Replace( eroscity, '_', ' ');
								Israel_neigh = true; //flag to reverse coordinates for zemanim calc
							}

							break;
						}
					}
				}

				fclose(stream2); //close cities5

				if (cnum < numUSAcities2)
				{
					ErrorHandler( 1, 9);; //never reached requested city/neighborhood
					return -1;
				}

				if ( foundcity == 0 ) //didn't find appropriate city and therefore didn't find its coordinates
				{
					ErrorHandler( 1, 10); //can't find right city
					return -1;
				}

				if (InStr( eroscountry, "Israel")) //Israel neighborhoods
				{
					geo = false; //treat coordinates as ITM and not as geographical coordinates of longitude and latitude
					parshiotflag = 1; //Eretz Yisroel parshiot

					///////////find equivalent Hebrew Name of Neighborhood///////////////////////

					//create the English city string that is listed in Israelhebneigh.dir
					sprintf(buff1, "%s%s%s%s", erosareabat, "/", eroscity, "\0" );
					Replace(buff1, ' ', '_' );  //buff1 is written in format of English city name in Israelhebneigh.dir

					//open Israelhebneigh.dir and search for Hebrew city name corresponding to the
					//English city name = buff1
					sprintf( filroot, "%s%s%s", drivcities, "eros/", "Israelhebneigh.dir");
					if ( !(stream2 = fopen( filroot, "r")) )
					{
						strcpy( EYhebNeigh, eroscity ); //can't open Isrealhebneigh.dir so use English name
					}
					else //find eroscity's coordinates
					{
						foundcity = 0;
						while (!feof(stream2) )
						{
							fgets_CR( doclin, 255, stream2 );
							if (strstr(doclin, buff1))
							{
								//found it
								foundcity = 1;
								//read Hebrew name
								fgets_CR( doclin, 255, stream2 );
								strcpy( EYhebNeigh, doclin );
								break;
							}
						}
						fclose(stream2);
						if (foundcity == 0) //didn't find it, so use English name as default
						{
							strcpy( EYhebNeigh, eroscity );
						}
					}
					/////////////////////////////////////////////////////////////////////////////
				}

			}

		}
		else //use inputed coordinates through cgi
		{
			strcpy( eroscity, eroscityin );
			if (optionheb) strcpy( EYhebNeigh, eroscity );
			eroslatitude = eroslatitudein;
			eroslongitude = eroslongitudein;
			eroshgt = eroshgtin;

			if (InStr( eroscountry, "Israel")) //Israel neighborhoods
			{
				//convert geo coordinates to ITM in order to search for vantage points
				double ltgeo, lggeo;//, kmyg, kmxg;
//				ltgeo = eroslatitude - 0.0013; //add approx conversion from WGS84 to Clark 1888 geographic coordinates
//				lggeo = -eroslongitude - 0.0013; //sign convention for longitude, approx conversion from WGS84 to Clark 1888
				ltgeo = eroslatitude; //WGS84 latitude in degrees
				lggeo = -eroslongitude; //sign convention for WGS84 longitude

                //convert from WGS84 geo coordinates to Old ITM kmx,kmy coordinates
                wgs842ics(ltgeo,lggeo,N,E);
				eroslongitude = Fix(0.5 + E) * 1e-3; //convert to km convention
				eroslatitude = Fix(0.5 + N - 1000000) * 1e-3;

				Israel_neigh = true; //flag to reverse coordinates for zemanim calc  <<<<<<<<<<<<<<<changes 010415

				//also include approximate fudge factor to convert to Clark 1888 geoid geographic coordinates

//			    if (!(geocasc_(&ltgeo, &lggeo, &kmyg, &kmxg)) )
//				{
//					eroslongitude = Fix(0.5 + kmxg) * 1e-3; //sign convention for longitude
//					eroslatitude = Fix(0.5 + kmyg) * 1e-3;
//				}
//				else
//				{
//					//geocas didn't converge
//					ErrorHandler( 1, 11);
//					return -1;
//				}

				geo = false; //treat coordinates as ITM and not as geographical coordinates of longitude and latitude
				parshiotflag = 1; //Eretz Yisroel parshiot
			}

		}

		switch (types)
		{
			case 0:
			case 1:
			case 4:
			case 5:
			case 6:
			case 7:
				viseros = true;
				break;
			case 2:
			case 3:
			case 8:
				viseros = false;
				break;
			default:
				viseros = false;
				break;
		}

		strcat( erosareabat, "_" );
		strcat( erosareabat, eroscountry );
		sprintf( currentdir, "%s%s%s%s", drivcities, "eros/", &erosareabat, "/" );

	}
	else if ( InStr( TableType, "Astr") ) //Astronomical calculations
	{
		//strcpy( country, "??÷??" );
		strcpy( astname, eroscityin );
		strcpy( hebcityname, eroscityin );
		strcpy( engcityname, eroscityin );
		geo = true;
		astronplace = true;
		lat = eroslatitudein;
		lon = eroslongitudein;

		if (ast || types > 9) //ast = true, astronomical times
		{
			hgt = eroshgtin;
		}
		else
		{
			hgt = 0; //ast = false (default), mishor times
		}

		steps[0] = RoundSeconds;
		steps[1] = RoundSeconds;

		distlim = 0;

		//sprintf( userloglin, "%s%s", "Astronomical name: ", astname );
		strcpy( eroscountry, "Astron" );
		strcpy( eroscity, astname );
		eroslatitude = lat;
		eroslongitude = lon;
		eroshgt = hgt;

		strcpy( citnamp, astname );

		if ( lon < -34.21 &&  lon> -35.8333 && lat > 29.55 && lat < 33.3417 && geotz == 2 )
		{
			parshiotflag = 1;  //'inside the borders of Eretz Yisroel
			cushion_factor = 1.0;

			if (SRTMflag > 9)
			{
				//give error message that can only shave up to a techum Shabbos in Eretz Eretz acc. to the psak of Rav Elyashiv z"l,
				//and anyone who has issues with this should send the request by Feedback.
				if (erosaprn > 1.5  && !AllowShavinginEY)
				{
					ier = -3;
					ErrorHandler(16, 1);
					return -1;
				}
			}
		}
        else
		{
			parshiotflag = 2; //somewhere in the diaspora
			/*
			if (types == 10 || types == 11) //not yet supported outside Eretz Yisroel
			{
				ErrorHandler( 1, 12);
			}
			*/

			//at different latitudes, spacing between 1 arcsec data points in longitude is about half of 30 m
			//so deter__min the distance corresponding to 1 arcsec at the relevant latitude
			//and adjust the cushions accordingly.
			//The calculation is as follows:
			//2*pi*cos(lat) * Re is distance corresponding to 2*pi radians
			//i.e. cos(lat) * Re is distance corresponding to 1 radian
			//1 arcsecond = 1/3600 degrees = 1/3600 * cd radians = 4.84814 x 10-6 radians
			//so distance of 1 arcsec = 4.84814 x 10-6 * cos(lat) * Re(m) 
			//so cushion will be above expression/30
			//rather it is 30m * cos(lat)/cos(reflat).  take reflat = 31 degrees of EY.
			//cushion_factor = cos(eroslatitude * cd) * 1.02963; 

			//however, the spacing in latitude is approx constant, and that is really what matters, so
			cushion_factor = 1.0; 

			//however, latitude accuracy in England is less since DTM based on 50 m spacing
			//if (eroslongitude <= 11 && eroslongitude >= -2) 
			if (eroslongitude <= 6 && eroslongitude >= -1 && eroslatitude >= 50.0 && eroslatitude <= 61.0 )
			{
				cushion_factor = 1.67;
			}

		}


		switch (types)
		{
			case -1: //<--use inputed coordinates for calculating zemanim
			case 0: //visible sunrise calculated from 30m DTM
			case 1: //visible sunset calculated from 30m DTM
            case 2:
 				nplac[0] = 1;
		  		nplac[1] = 0;
				avekmx[0] = lon; //parameters used for z'manim tables
				avekmy[0] = lat;
				avehgt[0] = hgt;
				break;
			case 3:
     	  		nplac[0] = 0;
		  		nplac[1] = 1;
				avekmx[1] = lon;
				avekmy[1] = lat;
				avehgt[1] = hgt;
		  		break;
			case 8:
				nplac[0] = 1;
				nplac[1] = 1;
				avekmx[0] = lon; //parameters used for z'manim tables
				avekmy[0] = lat;
				avehgt[0] = hgt;
				avekmx[1] = lon;
				avekmy[1] = lat;
				avehgt[1] = hgt;
		  		break;
		    case 10: //visible sunrise calculated from 30m DTM
 				nplac[0] = 1;
		  		nplac[1] = 0;
				avekmx[0] = lon; //parameters used for z'manim tables
				avekmy[0] = lat;
				avehgt[0] = hgt;
				distlims[3] = distlims[3] * cushion_factor;
				distlim = distlims[3];
				cushion[3] = __max(15, Cint(cushion[3] * cushion_factor));
                break;
            case 11: //visible sunset calculated from 30m DTM
     	  		nplac[0] = 0;
		  		nplac[1] = 1;
				avekmx[1] = lon;
				avekmy[1] = lat;
				avehgt[1] = hgt;
				distlims[3] = distlims[3] * cushion_factor;
				distlim = distlims[3];
				cushion[3] = __max(15, Cint(cushion[3] * cushion_factor));
		  		break;
		    case 12: //visible GokldenLight sunrise/sunset calculated from 30m DTM
			case 13:
				//first calculate the sunset horizon to find the highest point
				//then use those coordinates to find the sunrise from that point
				avekmx[0] = lon; //parameters used for z'manim tables
				avekmy[0] = lat;
				avehgt[0] = hgt;
				distlims[3] = distlims[3] * cushion_factor;
				distlim = distlims[3];
				cushion[3] = __max(15, Cint(cushion[3] * cushion_factor));
                break;
		}

      //aveusa = true; //invert coordinates in SunriseSunset due to
      //silly nonstandard coordinate format

      sprintf( currentdir, "%s%s%s%s", drivcities, "ast/", Trim(astname), "/" );
	}

    //write the parameters to the userlog.log file
	ier = WriteUserLog( UserNumber, zmanyes, typezman, "" );
	{
		if (ier == -1)
		{
			ErrorHandler(4, 1);
			return -1;
		}
		else if (ier == -2)
		{
			ErrorHandler(4, 2);
			return -1;
		}
   }



////////////////////////<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>

	if (Caldirectories()) return -1;

	if (optionheb)
	{
		if (RdHebStrings()) return -1 ; //Hebrew constants
		hebyear( g_yrcal, g_yrheb); //Hebrew Year in Hebrew characters
	}

	if (LoadParshiotNames()) return -1;
	

	if (sunsyes) //print sunrise/sunset tables
	{
		if (SunriseSunset(types, ExtTemp)) return -1;
	}

	if (zmanyes) //print zemanim tables
	{
		if (PrintMultiColTable(filzman, ExtTemp)) return -1;
	}

	if (WriteHtmlClosing()) return -1; //write the html closing info and close the stream

	return 0;
}

/////////////////////////hebnum//////////////////////////////////////
void hebnum(short k, char *ch)
///////////////////////////////////////////////////////////////////////
{
//PURPOSE: to generate Hebrew calendar date letters k=1,2,3 makes ?,?,?,?, etc.
//Revised Last: 04/30/2007
//Programmer: Chaim Keller
// no error code returns
////////////////////////////////////////////////////////////////////////////

	if (k <= 10)
	{
		sprintf( ch, "%s%s", &heb7[k][0], "\0" );
	}
	else if (k > 10 && k < 20 && k != 15 && k != 16 )
	{
		sprintf( ch, "%s%s%s", &heb7[10][0], &heb7[k - 10][0], "\0" );
	}
	else if (k == 15) //special way of writing 15,16, i.e., טו, טז
	{
		sprintf( ch, "%s%s%s", &heb7[9][0], &heb7[6][0], "\0" );
	}
	else if (k == 16)
	{
		sprintf( ch, "%s%s%s", &heb7[9][0], &heb7[7][0], "\0" );
	}
	else if ( k == 20 )
	{
		sprintf( ch, "%s%s", &heb7[11][0], "\0" );
	}
	else if (k > 20 && k < 30)
	{
		sprintf( ch, "%s%s%s", &heb7[11][0], &heb7[k - 20][0], "\0" );
	}
	else if (k == 30)
	{
		sprintf( ch, "%s%s", &heb7[12][0], "\0" );
	}

}

///////////////////hebweek//////////////////////////
char *hebweek( short dayweek, char ch[] )
//////////////////////////////////////////////////
//returns Hebrew day of the weeks
////////////////////////////////////
{
   switch (dayweek)
   {
	case 1:
		return &heb4[1][0];
		break;
	case 2:
		return &heb4[2][0];
		break;
	case 3:
		return &heb4[3][0];
		break;
	case 4:
		return &heb4[4][0];
		break;
	case 5:
		return &heb4[5][0];
		break;
	case 6:
		return &heb4[6][0];
		break;
	case 0:
	case 7:
		return &heb4[7][0];
		break;
   }
   return 0;
}

//////////////////////////RdHebStrings//////////////////////////////
short RdHebStrings()
///////////////////////////////////////////////////////////////////
{

//PURPOSE: to load up Hebrew strings used in calendars//////
//Revised Last: 04/30/2007
//Programmer: Chaim Keller
// returns error codes: -1 = can't find "calhebrew.txt"
//////////////////////////////////////////////////////////////////////

	short flgheb = 0;
	short iheb = 0;
	FILE *stream;

	if (optionheb) //read in hebrew strings from calhebrew.txt
	{
		sprintf( myfile, "%s%s%s", drivjk, "calhebrew_dtm.txt", "\0" );

		if ( !( stream = fopen( myfile, "r")) )
		{
			//can't find hebrew strings, so use English strings as default
			optionheb = false;
			ErrorHandler(2, 1); //exit with error message
			return -1;
		}
		else
		{
			iheb = 0;

			while( !feof( stream ) )
			{
				doclin[0] = '\0'; //clear the buffer

				fgets_CR(doclin, 255, stream); //read in line of text

				if (strstr( doclin,"[sunrisesunset]"))
				{
					flgheb = 0;
				}
				else if (strstr( doclin,"[newhebcalfm]"))
				{
					flgheb = 1;
					iheb = 0;
				}
				else if (strstr( doclin,"[zmanlistfm]"))
				{
					flgheb = 2;
					iheb = 0;
				}
				else if (strstr( doclin,"[hebweek]" ))
				{
					flgheb = 3;
					iheb = 0;
				}
				else if (strstr( doclin,"[holiday_dates]" ))
				{
					flgheb = 4;
					iheb = 0;
				}
				else if (strstr( doclin,"[candle_lighting]" ))
				{
					flgheb = 5;
					iheb = 0;
				}
				else if (strstr( doclin,"[hebrew]" ))
				{
					flgheb = 6;
					iheb = 0;
				}
				else if (strstr( doclin, "[PlotGraphHTML5]" ))
				{
					flgheb = 7;
					iheb = 0;
				}

				if (strlen(Trim(doclin)) > 0 ) //ignore blank lines containing just "\n"
				{
					switch (flgheb)
					{
						case 0:
							strcpy( &heb1[iheb][0], doclin );
							break;
						case 1:
							strcpy( &heb2[iheb][0], doclin );
							break;
						case 2:
							strcpy( &heb3[iheb][0], doclin );
							break;
						case 3:
							strcpy( &heb4[iheb][0], doclin );
							break;
						case 4:
							strcpy( &heb5[iheb][0], doclin );
							break;
						case 5:
							strcpy( &heb6[iheb][0], doclin );
							break;
						case 6: //leave heb7[0] as NULL
							if (iheb != 0) strcpy( &heb7[iheb][0],  doclin );
							break;
						case 7:
							strcpy( &heb9[iheb][0], doclin );
							break;
					}

					iheb++;
				}
			}
			fclose(stream);
		}
	}

//////////Hebrew Months in Hebrew/English////////////
	sprintf( doclin, "%s%s%s", drivjk, "calhebmonths.txt", "\0" );

	if ( stream = fopen( doclin, "r" ) )
	{
		for( short i = 0; i < 14; i++ )
		{
			fgets_CR( doclin, 255, stream );
			strcpy( &monthh[0][i][0], doclin );
		}
		fclose( stream );
	}
	else //use defaults
	{
		if (abortprog)
		{
		        // issue warning to log file that "Calhebmonths.txt" can't be found
			//can't find hebrew strings, so use English strings as default
			optionheb = false;
			ErrorHandler(2, 2); //exit with error message
			return -2;
		}
		else //just write error message to log fileonAstroNoon(
		{
			WriteUserLog( 0, 0, 0, "Can't find or open Calhebmonths.txt in routine LoadConstants");
		}

		strcpy( &monthh[0][0][0], "תשרי");
		strcpy( &monthh[0][1][0], "מרחשון");
		strcpy( &monthh[0][2][0], "כסלו");
		strcpy( &monthh[0][3][0], "טבת");
		strcpy( &monthh[0][4][0], "שבט");
		strcpy( &monthh[0][5][0], "אדר א");
		strcpy( &monthh[0][6][0], "אדר ב");
		strcpy( &monthh[0][7][0], "ניסן");
		strcpy( &monthh[0][8][0], "אייר");
		strcpy( &monthh[0][9][0], "סיון");
		strcpy( &monthh[0][10][0], "תמוז");
		strcpy( &monthh[0][11][0], "אב");
		strcpy( &monthh[0][12][0], "אלול");
		strcpy( &monthh[0][13][0], "אדר");
/*
		//convert DOS Hebrew to ISO
		for ( i = 0; i < 14; i++ )
		{
			IsoToUnicode( &monthh[0][i][0], doclin );
			strcpy( &monthh[0][i][0], doclin );
		}
*/
	}

	return 0;
}

//////////////////RdHebErrorMessages///////////////////////////////////
short RdHebErrorMessages()
/////////////////////////////////////////////////////////////////////
//read Hebrew Error messages, and store themAstroNoon(
////////////////////////////////////////////////////////////////////
{
	short iheb = 0;
	FILE *stream;

	sprintf( myfile, "%s%s", drivjk, "HebErrors.txt" );

	if ( !( stream = fopen( myfile, "r")) )
	{
		return -1;
	}
	else
	{
		while( !feof( stream ) ) //read in the Hebrew Error strings
		{
			doclin[0] = '\0'; //clear the buffer

			fgets_CR(doclin, 255, stream); //read in line of text

			strcpy( &heb8[iheb][0], doclin );

			iheb++;
		}

		fclose(stream);
		return 0;

	}
}

//////////////////////////LoadParshiotNames//////////////////////////
short LoadParshiotNames()
////////////////////////////////////////////////////////////////////
{
//PURPOSE: to load up names of Parshios and Yom Tovim//////
//Revised Last: 04/30/2007
//Programmer: Chaim Keller
// returns error codes: 1 = can't find calparshios.txt"
//						2 = can't find correct sedra
//                      3 = can't find calparshios_data.txt
//////////////////////////////////////////////////////////////////////

	short nfound = 0;
	short numhol = 0;
	short nn = 0;
	short PMode = 0;
	char Dparsha[255] = ""; // Torah reading for Diaspora   //<<<<<<<<<<<changes 01/03/15
	char Iparsha[255] = ""; // Torah reading for Eretz Yisroel
	FILE *stream;

	//deter__min cycle name
	switch (g_hebleapyear)
	{
		case false: //Hebrew calendar non-leapyear
			strcpy( &Iparsha[0], "[0_" );
			break;
		case true: //Hebrew calendar leapyear
			strcpy( &Iparsha[0], "[1_" );
			break;
	}
	switch (g_dayRoshHashono) //day of the week of RoshHashono
	{
		case 2: //Monday
			strncpy( Iparsha + 3, "2_", 2);
			break;
		case 3: //Tuesday
			strncpy( Iparsha + 3, "3_", 2);
			break;
		case 5: //Thursday
			strncpy( Iparsha + 3, "5_", 2);
			break;
		case 7: //Shabbos
			strncpy( Iparsha + 3, "7_", 2);
			break;
	}
	switch (g_yeartype) //chaser, kesidrah, shalem
	{
		case 1: //chaser
			strncpy( Iparsha + 5, "1", 1);
			break;
		case 2: //kesidrah
			strncpy( Iparsha + 5, "2", 1);
			break;
		case 3: //shalem
			strncpy( Iparsha + 5, "3", 1);
			break;
	}

	strcpy( Dparsha, Iparsha);

	if ((strstr(Iparsha,"[0_2_1") ) ||
	   (strstr(Iparsha,"[0_5_3") ) ||
	   (strstr(Iparsha,"[0_7_1") ) ||
	   (strstr(Iparsha,"[0_7_3") ) ||
	   (strstr(Iparsha,"[1_5_1") ) ||
	   (strstr(Iparsha,"[1_5_3") ) ||
	   (strstr(Iparsha,"[1_7_1") ))
	{
		//alike for Eretz Yisroel and diaspora
		strcat( Iparsha, "]");
		strcat( Dparsha, "]");
	}
	else //different for Eretz Yisroel and the diaspora
	{
		strcat( Iparsha, "_israel]");
		strcat( Dparsha, "_diaspora]");
	}


	//now read parshios and holiday information

	sprintf( myfile, "%s%s%s", drivjk, "calparshiot.txt", "\0" );

    	if (( stream = fopen( myfile, "rt")) )
	{

		PMode = -1;
		while ( !feof(stream) )
		{
			doclin[0] = '\0'; //clear the buffer

			fgets_CR(doclin, 255, stream); //read in line of text
			short lendoc = strlen(doclin);

			if (lendoc > 1)
			{

				if	(strstr( doclin,"[hebparshiot]")) //Hebrew parshiot names
				{
					PMode = 0;
					nn = 0;
				}
				else if (strstr( doclin,"[engparshiot]")) //English parshiot names
				{
					PMode = 1;
					nn = 0;
				}
				else if (strstr( doclin,"[hebholidays]")) //Hebrew Yom Tovim names
				{
					PMode = 2;
					numhol = 0;
				}
				else if (strstr( doclin,"[engholidays]")) //English Yom Tovim names
				{
					PMode = 3;
					numhol = 0;
				}

				else if (strlen(doclin) > 1) //ignore blank lines
				{
					switch (PMode)
					{
						case 0:
						case 1:
							sprintf( &arrStrParshiot[PMode][nn][0], "%s%s", doclin, "\0");
							nn++;
							break;
						case 2:
						case 3:
							sprintf( &holidays[PMode - 2][numhol][0], doclin, "\0");
							numhol++;
							break;
						default:
							//keep on looping
						break;
					}
				}
			}
		}
		fclose(stream);
	}
	else
	{
		ErrorHandler(3, 1); //return with error flag set
		return -1;
	}


//////////load up the sedra for Eretz Yisroel and the diaspora/////////////////
	sprintf( myfile, "%s%s%s", drivjk, "calparshiot_data.txt", "\0" );

    	if ( (stream = fopen( myfile, "r")) )
	{
		PMode = -1;
		nfound = 0;
		while ( !feof(stream) )
		{
			doclin[0] = '\0'; //clear the buffer

			fgets_CR(doclin, 255, stream); //read in line of text

			switch (PMode)
			{
				case -1:
					if ((strstr( doclin, Iparsha)) && (strstr( doclin, Dparsha))) //parse out both sedra
					{
						PMode = 2;
						nfound += 2;
					}
					else if ((strstr( doclin, Iparsha)) && (strstr( doclin, Dparsha) == NULL)) //parse out Eretz Yisroel sedra
					{
						PMode = 0;
						nfound++;
					}
					else if ((strstr( doclin, Iparsha) == NULL) &&
						(strstr( doclin, Dparsha))) //parse out Eretz Yisroel sedra
					{
						PMode = 1;
						nfound++;
					}
					break;
				case 0:
				case 1:
				case 2: //parse out EY, diaspora sedra
					//ParseSedra(doclin, PMode);
					ParseSedra(PMode);
					if (nfound == 1) //look for sedra of Eretz Yisroel
					{
						PMode = -1;
					}
					if (nfound == 2) //found everything
					{
						fclose(stream);
						return 0;
					}
					break;
			}
		}
		fclose(stream);
		if (PMode == -1) //couldn't find parshiot
		{
			fclose(stream);
			ErrorHandler(3, 3);
			return -1;
		}
	}
	else
	{
		ErrorHandler(3, 2);
		return -1;
	}

	return 0;

}

///////////////////////////ParseSedra////////////////////////////////////
//void ParseSedra(char doclins[255], short PMode)
void ParseSedra( short PMode ) //doclin is now global so doesn't need to be passed
/////////////////////////////////////////////////////////////////////////
//parses the parsha data to stuff the Sedra arrays with the Shabbosim and holidays
//////////////////////////////////////////////////////////////////////////////
{
	//static char myfile[255] = "";
	short mParsha = 0;
	short i = 0;
	short mParshaD = 0;
	short mParshaI = 0;
	short hc = 0;
	char ch[2] = " ";
	bool doubleparsha = false;
	static char parsnum[255] = "";

	if (optionheb)
	{
		hc = 0;
	}
	else
	{
		hc = 1;
	}
	mParshaI = 0;
	mParshaD = 0;
	for (i = 0; i <= strlen(doclin)-1; i++) //doclin is a global variable loaded by the calling routine
	{
		strncpy( ch, (doclin + i), 1); //extract ith character in doclin

		if ((strstr( ch, ",")) || (strstr( ch, "}")) )
		{
			if (PMode == 0 || PMode == 1)
			{
				if (PMode == 0)
				{
					mParsha = mParshaI;
				}
				if (PMode == 1)
				{
					mParsha = mParshaD;
				}

				if (! doubleparsha) //single parsha
				{
					strcpy( &arrStrSedra[PMode][mParsha][0], &arrStrParshiot[hc][atoi(parsnum)][0]);
				}
				else //double parsha
				{
		                        sprintf( &arrStrSedra[PMode][mParsha][0],"%s%s%s",&arrStrParshiot[hc][atoi(parsnum)][0],
                                                "-",&arrStrParshiot[hc][atoi(parsnum) + 1][0]);
						doubleparsha = false; //reset double parsha flag
				}

				if (PMode == 0)
				{
					mParshaI++;
				}
				if (PMode == 1)
				{
					mParshaD++;
				}
				parsnum[0] = '\0';

			}
			else if (PMode == 2) //both EY and diaspora have same parsha
			{

				if (! doubleparsha) //single parsha
				{
					strcpy( &arrStrSedra[0][mParshaI][0], &arrStrParshiot[hc][atoi(parsnum)][0]);
					strcpy( &arrStrSedra[1][mParshaD][0], &arrStrSedra[0][mParshaI][0]);
					mParshaI++; //increment the parsha number
					mParshaD++; //increment the parsha number
				}
				else //double parsha
				{
					sprintf( &arrStrSedra[0][mParshaI][0],"%s%s%s",&arrStrParshiot[hc][atoi(parsnum)][0],
					         "-",&arrStrParshiot[hc][atoi(parsnum) + 1][0]);

					strcpy( &arrStrSedra[1][mParshaD][0], &arrStrSedra[0][mParshaI][0]);

					doubleparsha = false; //reset double parsha flag

					mParshaI++; //increment the parsha number
					mParshaD++; //increment the parsha number

				}
			}

			parsnum[0] = '\0'; //reset the sedra designation number

		}
        	else if (strstr( ch, "D"))
		{
			parsnum[0] = '\0';
			doubleparsha = true;
		}
//        	else if (strstr( ch, " "))
//		{
//			// do nothing
//		}
		else if ((strstr(ch, "0")) ||
                         (strstr(ch, "1")) ||
                         (strstr(ch, "2")) ||
                         (strstr(ch, "3")) ||
                         (strstr(ch, "4")) ||
                         (strstr(ch, "5")) ||
                         (strstr(ch, "6")) ||
                         (strstr(ch, "7")) ||
                         (strstr(ch, "8")) ||
                         (strstr(ch, "9")))
		{
			strcpy( &parsnum[strlen(parsnum)], ch);
			parsnum[strlen(parsnum) + 1] = 0;
		}
	}
}

///////////////////////InsertHolidays//////////////////////////////
void InsertHolidays(char *caldayHeb, char *caldayEng, char *calday, char *calhebweek,
		short *i, short *j, short *k, short *yrn, short yl, short *ntotshabos,
		short *dayweek, short *fshabos,	short *nshabos,	short *addmon, short iheb )
//////////////////////////////////////////////////////////////////
//returns the current hebrew date, civil date, and day description
//used to provide the above information for the WriteTable routine
//caldayHeb = Hebrew date
//caldayEng = civil date
//calday = Holiday and Day
//store
////////////////////////////////////////////////////////////////
{
	bool RemoveUnderline = 1;
	char CalParsha[255] = "";
	char caldate[255] = "";
	char ch[25] = "";
	char EYSedra[255] = "";
	char DiSedra[255] = "";
	char *pch;

	//passed variables are:
	//(1) the day field = calday$
	//(2) the increment variables i,k used to deter__min
	//    (a) the date field = CalDate$
	//    (b) the Shabbos Parsha field = CalParsha$
	//(3) variables that keep track of year, year length, day of week and shabbos:
	//    ntotshabos (total number of shabbosim), dayweek, fshabos,
	//    nshabos, addmon, yrn, yl, iheb
	//returned ier = -5 for error

    ///////////////day/date calculations if needed/////////////////////////
	if (*fshabos + *nshabos * 7 == *j) //this is shabbos, add parsha info
	{
		*nshabos += 1; //<<<2--->>>
		*ntotshabos += 1;
		//add sedra info
		sprintf( EYSedra, "%s%s", &arrStrSedra[0][*ntotshabos - 1][0], "\0" );
		sprintf( DiSedra, "%s%s", &arrStrSedra[1][*ntotshabos - 1][0], "\0" );

		//check for end of year
		if (*fshabos + *nshabos * 7 > yrend[0])
		{
			*fshabos = 7 - (yrend[0] - (*fshabos + (*nshabos - 1) * 7));
			*nshabos = 0;
		}
	}
	else //weekday, no parsha info
	{
		strcpy( EYSedra, "-----" );
		strcpy( DiSedra, "-----" );
	}

	//store the civil date
	EngCalDate( *j, FirstSecularYr + *yrn - 1, yl, caldate );
	sprintf( caldayEng, "%s%s", caldate, "\0" );

	if (optionheb == true)
	{
		hebnum(*k, ch);
	}
	else
	{
		itoa(*k, ch, 10);
	}

	//store Hebrew date
	sprintf( caldayHeb, "%s%s%s", ch, "-",  &monthh[iheb][*i + *addmon - 1][0] );

	if (endyr == 12 && *i == 6)
	{
		sprintf( caldayHeb, "%s%s%s", ch, "-", &monthh[iheb][13][0] );
		*addmon = 1;
	}

	Trim(caldayHeb);

	*dayweek += 1;
	if (*dayweek == 8)
	{
		*dayweek = 1;
	}

   //add weekdays////////////////////////////////////////

	//Hebrew weekdays
	if (optionheb)
	{
		sprintf( calhebweek, "%s%s", hebweek(*dayweek, ch), "\0" );
		sprintf( calday, "%s", calhebweek );
	}

	if ( !optionheb )
	//convert Hebrew weekdays to English weekdays
	{
		switch (*dayweek)
		{
		case 1:
			sprintf( calday, "%s", "Sunday");
			goto ih50;
			break;
		case 2:
			sprintf( calday, "%s", "Monday");
			goto ih50;
			break;
		case 3:
			sprintf( calday, "%s", "Tuesday");
			goto ih50;
			break;
		case 4:
			sprintf( calday, "%s", "Wednesday");
			goto ih50;
			break;
		case 5:
			sprintf( calday, "%s", "Thursday");
			goto ih50;
			break;
		case 6:
			sprintf( calday, "%s", "Friday");
			goto ih50;
			break;
		case 7:
		case 0:
			sprintf( calday, "%s", "Shabbos");
			goto ih50;
			break;
		}
	}

ih50:
	//================================================
        //add parshiot
	if (parshiotflag == 1) //Eretz Yisroel parshiot
	{
        if ( !strstr( EYSedra, "-----") )
 		{
			if (optionheb)
			{
            if (strstr( EYSedra, &heb4[6][0] ) )
				{
					sprintf( calday, "%s", EYSedra);
				}
				else
				{
            	sprintf( calday,"%s%s",&heb5[37][0], EYSedra);
				}
			}
			else if (! optionheb)
			{
				if (strstr( EYSedra, "Shabbos" ) )
				{
					sprintf( calday, "%s", EYSedra );
				}
				else
				{
					sprintf( calday, "%s%s", "Shabbos_", EYSedra);
				}
			}
		}
	}
	else if (parshiotflag == 2) //diaspora parshiot
	{
		if ( !strstr( DiSedra, "-----") )
 		{
			if (optionheb)
			{
				if (strstr( DiSedra, &heb4[6][0] )  )
				{
					sprintf( calday, "%s", DiSedra);
				}
				else
				{
					sprintf( calday,"%s%s", &heb5[37][0], DiSedra);
				}
			}
			else if (! optionheb)
			{
				if (strstr( DiSedra, "Shabbos" ) )
				{
					sprintf( calday, "%s", DiSedra );
				}
				else
				{
					sprintf( calday,"%s%s","Shabbos_", DiSedra );
				}
			}
		}
	}


	//=========================================================
	//now add holidays to Shabbosim

	if (parshiotflag == 1) //Eretz Yisroel parshiot
	{
		sprintf( CalParsha, "%s", EYSedra);
	}
	else if (parshiotflag == 2) //diaspora parshiot
	{
		sprintf( CalParsha, "%s", DiSedra);
	}

	if ( !strstr( CalParsha, "-----") )
	{
		//add in Shabbos Chanukah and Shabbos Shushan Purim when appropriate
		if (optionheb)
		{
			if ( !strcmp(caldayHeb, &heb5[13][0]) ||
			     !strcmp(caldayHeb, &heb5[14][0]) ||
			     !strcmp(caldayHeb, &heb5[15][0]) ||
			     !strcmp(caldayHeb, &heb5[16][0]) ||
			     !strcmp(caldayHeb, &heb5[17][0]) ||
			     !strcmp(caldayHeb, &heb5[18][0]) ||
			     !strcmp(caldayHeb, &heb5[19][0]) ||
			     !strcmp(caldayHeb, &heb5[20][0]) ||
			     !strcmp(caldayHeb, &heb5[21][0]) )//Chanukah
			{
				if ((g_yeartype == 2 || g_yeartype == 3) && !strcmp(caldayHeb, &heb5[21][0]) )
				{
					goto remU;
				}

				sprintf( buff1,"%s%s%s",&heb5[36][0], &holidays[0][6][0], "\0");
				strcat( calday, buff1);
				goto remU;
			}
			else if ( !strcmp(caldayHeb, &heb5[24][0]) || !strcmp(caldayHeb, &heb5[25][0]) ) //Shushan Purim
			{
				//sprintf( buff1,"%s%s%s",&heb5[36][0], &holidays[0][8][0]), "\0";
				//strcat( calday, buff1);
				strcat( calday, &heb5[36][0] );
				strcat( calday, &holidays[0][8][0]);
				goto remU;
			}
			else
			{
				goto remU;
			}
		}
		else
		{
			if ( !strcmp(caldayHeb, "25-Kislev")  ||
			     !strcmp(caldayHeb, "26-Kislev")  ||
			     !strcmp(caldayHeb, "27-Kislev")  ||
			     !strcmp(caldayHeb, "28-Kislev")  ||
			     !strcmp(caldayHeb, "29-Kislev")  ||
			     !strcmp(caldayHeb, "30-Kislev")  ||
			     !strcmp(caldayHeb, "1-Teves")  ||
			     !strcmp(caldayHeb, "2-Teves")  ||
			     !strcmp(caldayHeb, "3-Teves")  )
			{
				if ((g_yeartype == 2 || g_yeartype == 3) && !strcmp( caldayHeb,"3-Teves") )
				{
					goto remU;
				}
				sprintf( buff1, "%s%s%s", &holidays[1][6][0], ",_", calday );
				sprintf( calday, "%s", buff1 );
				goto remU;
			}
			else if (!strcmp(caldayHeb, "15-Adar") || !strcmp(caldayHeb, "15-Adar II") ) //Shushan Purim
			{
				sprintf( buff1, "%s%s%s", &holidays[1][8][0], ",_", calday );
				sprintf( calday, "%s", buff1 );
				goto remU;
			}
			else
			{
				goto remU;
			}
		}
	}

	//=======================================================
	//add in week day holidays
	if (optionheb)
	{
		if ( !strcmp(caldayHeb, &heb5[1][0]) || !strcmp(caldayHeb, &heb5[2][0]) )
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][0][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[3][0]))
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][1][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[4][0]))
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][2][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[5][0]) && parshiotflag == 1) //EY parshiot
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][3][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[5][0]) && parshiotflag == 2) //diaspora parshiot
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][2][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[6][0]) ||
		          !strcmp(caldayHeb, &heb5[7][0]) ||
		          !strcmp(caldayHeb, &heb5[8][0]) ||
		          !strcmp(caldayHeb, &heb5[9][0]) ||
		          !strcmp(caldayHeb, &heb5[10][0]) )
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][3][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[11][0]) )
		{
                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][4][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[12][0]) && parshiotflag == 2)
		{
                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][5][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[13][0]) ||
		          !strcmp(caldayHeb, &heb5[14][0]) ||
		          !strcmp(caldayHeb, &heb5[15][0]) ||
		          !strcmp(caldayHeb, &heb5[16][0]) ||
		          !strcmp(caldayHeb, &heb5[17][0]) ||
		          !strcmp(caldayHeb, &heb5[18][0]) ||
		          !strcmp(caldayHeb, &heb5[19][0]) ||
		          !strcmp(caldayHeb, &heb5[20][0]) ||
		          !strcmp(caldayHeb, &heb5[21][0]) ) //Chanukah
		{
			if ((g_yeartype == 2 || g_yeartype == 3) && !strcmp( caldayHeb, &heb5[21][0]) )
			{
				goto remU;
			}
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][6][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[22][0]) ||
		          !strcmp(caldayHeb, &heb5[23][0]) ) //Purim
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][7][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[24][0]) ||
		          !strcmp(caldayHeb, &heb5[25][0]) ) //Shushan Purim
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][8][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[26][0]) ) //Pesach
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][9][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[27][0]) && parshiotflag == 1 ) //Pesach 2nd day, EY parshiot
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][10][0]);
		}
		else if (!strcmp(caldayHeb, &heb5[27][0]) && parshiotflag == 2) //Pesach 2nd day, diaspora parshiot
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][9][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[28][0]) ||
		          !strcmp(caldayHeb, &heb5[29][0]) ||
		          !strcmp(caldayHeb, &heb5[30][0]) ||
		          !strcmp(caldayHeb, &heb5[31][0]) ) //Chol Hamoed Succos
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][10][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[32][0]) ) //Smini Azerat
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][9][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[33][0]) && parshiotflag == 2) //Simchas Torah diaspora parshiot
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][9][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[34][0]) ) //Simachas Torah
		{
	                sprintf( buff1,"%s%s%s",calday, &heb5[36][0], &holidays[0][11][0]);
		}
		else if ( !strcmp(caldayHeb, &heb5[35][0]) && parshiotflag == 2) //Simchas Torah diaspora parshiot
		{
	                sprintf( buff1,"%s%s%s", calday, &heb5[36][0], &holidays[0][11][0]);
		}
		else
		{
			//regular week day
			sprintf( buff1, "%s", calhebweek );
		}

		sprintf( calday, "%s", buff1 );
	}
	else
	{
		if ( !strcmp(caldayHeb, "1-Tishrey") ||
	             !strcmp(caldayHeb, "2-Tishrey")) //Rosh Hoshono
		{
				sprintf( buff1, "%s%s%s", &holidays[1][0][0], ",_", calday );
				strcpy( calday, buff1 );
		}
		else if ( !strcmp( caldayHeb, "10-Tishrey") ) //Yom Hakipurim
		{
				sprintf( buff1, "%s%s%s", &holidays[1][1][0], ",_", calday );
				strcpy( calday, buff1 );
		}
		else if ( !strcmp( caldayHeb, "15-Tishrey") ) //Succos
		{
				sprintf( buff1, "%s%s%s", &holidays[1][2][0], ",_", calday );
				strcpy( calday, buff1 );
		}
		else if ( !strcmp( caldayHeb, "16-Tishrey") ) //Succos
		{
			if (parshiotflag == 1) //Chol Hamoed Succos (EY parshiot)
			{
			sprintf( buff1, "%s%s%s", &holidays[1][3][0], ",_", calday );
			strcpy( calday, buff1 );
			}
			else if (parshiotflag == 2) //Succos (diaspora parshiot)
			{
			sprintf( buff1, "%s%s%s", &holidays[1][2][0], ",_", calday );
			strcpy( calday, buff1 );
			}
		}
		else if ( !strcmp(caldayHeb, "17-Tishrey") ||
                          !strcmp(caldayHeb, "18-Tishrey") ||
                          !strcmp(caldayHeb, "19-Tishrey") ||
                          !strcmp(caldayHeb, "20-Tishrey") ||
                          !strcmp(caldayHeb, "21-Tishrey") ) //Chol Hamoed Succos
		{
			sprintf( buff1, "%s%s%s", &holidays[1][3][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "22-Tishrey") ) //Shmini Azeres
		{
			sprintf( buff1, "%s%s%s", &holidays[1][4][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "23-Tishrey") ) //Simchas Torah
		{
			if (parshiotflag == 2) //diaspora parshiot
			{
			sprintf( buff1, "%s%s%s", &holidays[1][5][0], ",_", calday );
			strcpy( calday, buff1 );
			}
		}
		else if ( !strcmp(caldayHeb, "25-Kislev") ||
                          !strcmp(caldayHeb, "26-Kislev") ||
                          !strcmp(caldayHeb, "27-Kislev") ||
                          !strcmp(caldayHeb, "28-Kislev") ||
                          !strcmp(caldayHeb, "29-Kislev") ||
                          !strcmp(caldayHeb, "30-Kislev") ||
                          !strcmp(caldayHeb, "1-Teves") ||
                          !strcmp(caldayHeb, "2-Teves") ||
                          !strcmp(caldayHeb, "3-Teves") )
		{
			if ((g_yeartype == 2 || g_yeartype == 3) && !strcmp( caldayHeb,"3-Teves") )
			{
				goto remU;
			}
			sprintf( buff1, "%s%s%s", &holidays[1][6][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "14-Adar") || !strcmp(caldayHeb, "14-Adar II") ) //Purim
		{
			sprintf( buff1, "%s%s%s", &holidays[1][7][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "15-Adar") || !strcmp(caldayHeb, "15-Adar II") ) //Shushan Purim
		{
			sprintf( buff1, "%s%s%s", &holidays[1][8][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "15-Nisan")) //Pesach
		{
			sprintf( buff1, "%s%s%s", &holidays[1][9][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "16-Nisan")) //Second day Pesach for diaspora
		{
			if (parshiotflag == 1) //EY parshiot
			{
			sprintf( buff1, "%s%s%s", &holidays[1][10][0], ",_", calday );
			strcpy( calday, buff1 );
			}
			else if (parshiotflag == 2) //diaspora parshiot
			{
			sprintf( buff1, "%s%s%s", &holidays[1][9][0], ",_", calday );
			strcpy( calday, buff1 );
			}
		}
		else if ( !strcmp(caldayHeb, "17-Nisan") ||
                          !strcmp(caldayHeb, "18-Nisan") ||
                          !strcmp(caldayHeb, "19-Nisan") ||
                          !strcmp(caldayHeb, "20-Nisan") )
		{
			sprintf( buff1, "%s%s%s", &holidays[1][10][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "21-Nisan")) //last day of Pesach
		{
			sprintf( buff1, "%s%s%s", &holidays[1][9][0], ",_", calday );
			strcpy( calday, buff1 );
			if (parshiotflag == 1) //EY parshios
			{
				//remove the Chol_Hamoed
				short pos = InStr( calday, "Chol_Hamoed_Pesach" );
				if (pos != 0)
				{
					strncpy( calday + pos, "Pesach", 5);
				}
			}
		}
		else if ( !strcmp(caldayHeb, "22-Nisan")) //Second day last day of Pesach for diaspora
		{
			if (parshiotflag == 2) //diaspora parshiot
			{
			sprintf( buff1, "%s%s%s", &holidays[1][9][0], ",_", calday );
			strcpy( calday, buff1 );
			}
		}
		else if ( !strcmp(caldayHeb, "6-Sivan") ) //Shavuos
		{
			sprintf( buff1, "%s%s%s", &holidays[1][11][0], ",_", calday );
			strcpy( calday, buff1 );
		}
		else if ( !strcmp(caldayHeb, "7-Sivan") ) //Second day of Shavuos for diaspora
		{
			if (parshiotflag == 2) //diaspora parshiot
			{
			sprintf( buff1, "%s%s%s", &holidays[1][11][0], ",_", calday );
			strcpy( calday, buff1 );
			}
		}
		else //don't change anything
		{
			//just add in the week day
			sprintf( buff1, "%s", calhebweek );
		}
	}

	//========================================================
	remU: //final tidy ups of Hebrew calendar day

	if (optionheb)
	{
		//remove double Shabboses
		sprintf( buff1, "%s%s%s", &heb4[7][0], "_", &heb4[7][0] );

		pch = strstr( calday, buff1);
		if (pch != NULL)
		{
			strncpy( pch, "       ", 7);
			Trim(calday);
		}
	}

	if ( !optionheb && !strcmp(caldayHeb, "21-Nisan") ) //last day of Pesach of EY sedra
	{
		//remove the Chol_Hamoed
		short pos = InStr( calday, "Chol_Hamoed_Pesach" );
		if (pos != 0)
		{
			sprintf( calday + pos - 1, "%s%s", "Pesach", "\0" );
		}
	}

	if (RemoveUnderline && optionheb) //remove the "_"'s from Hebrew calday
	{
		Replace( calday, '_', ' ' );
	}


}


///////////////////hebyear/////////////////////////////////////
char *hebyear(char *yrcal, short yr)
/////////////////////////////////////////////////////////////
//convert hebrew year in numbers to hebrew characters
// inpurt: hebrew year, yr, in numbers, e.g., 5767
// outpout: pointer to Hebrew string containing the Hebrew Year
/////////////////////////////////////////////////////////////////
{
	short hund = 0;
	short tens = 0;
	short ones = 0;

	//hunds, tenss, ones are the index of heb7 that corresponds to the desired hebrew letter
	short hunds[4] = { 0,0,0,0 }; //if hunds[i] == 0 then blank
	short tenss = 0;

	//deter__min the hundreds, tens, ones that compose the Hebrew Year
	hund = (int)Fix((yr - 5000) * 0.01);
	tens = (int)Fix((yr - 5000 - hund * 100) * 0.1);
	ones = (int)Fix(yr - 5000 - hund * 100 - tens * 10);

	if ( hund == 0 )
	{
		hunds[0] = 0;
	}
	else if ( hund == 1 ) // ק
	{
		hunds[0] = 1;
		hunds[1] = 19;
	}
	else if ( hund == 2 )  //ר
	{
		hunds[0] = 1;
		hunds[1] = 20;
	}
	else if ( hund == 3 ) //ש
	{
		hunds[0] = 1;
		hunds[1] = 21;
	}
	else if ( hund == 4 ) //ת
	{
		hunds[0] = 1;
		hunds[1] = 22;
	}
	else if ( hund == 5 ) //תק
	{
		hunds[0] = 2;
		hunds[1] = 22;
		hunds[2] = 19;
	}
	else if ( hund == 6 ) //תר
	{
		hunds[0] = 2;
		hunds[1] = 22;
		hunds[2] = 20;
	}
	else if ( hund == 7 ) //תש
	{
		hunds[0] = 2;
		hunds[1] = 22;
		hunds[2] = 21;
	}
	else if ( hund == 8 ) //תת
	{
		hunds[0] = 2;
		hunds[1] = 22;
		hunds[2] = 22;
	}
	else if ( hund == 9 ) //תתק
	{
		hunds[0] = 3;
		hunds[1] = 22;
		hunds[2] = 22;
		hunds[3] = 19;
	}
	else if ( hund == 10 ) //תתר
	{
		hunds[0] = 3;
		hunds[1] = 22;
		hunds[2] = 22;
		hunds[3] = 20;
	}

	if ( tens == 1 ) //י
	{
		tenss = 10;
	}
	else if ( tens == 2 ) //כ
	{
		tenss = 11;
	}
	else if  ( tens == 3 ) //ל
	{
		tenss = 12;
	}
	else if  ( tens == 4 ) //מ
	{
		tenss = 13;
	}
	else if  ( tens == 5 ) //נ
	{
		tenss = 14;
	}
	else if  ( tens == 6 ) //ס
	{
		tenss = 15;
	}
	else if  ( tens == 7 ) //ע
	{
		tenss = 16;
	}
	else if  ( tens == 8 ) //פ
	{
		tenss = 17;
	}
	else if  ( tens == 9 ) //צ
	{
		tenss = 18;
	}


	//write the Hebrew Year in Unicode Hebrew.  The Hebrew Year can be a maximum of 5 characters
	//however depending on the year it can be even only one character long.
	//So need to keep this is mind in order to place the quotation mark in the right places
	//as is down in the following:

	if (ones != 0 ) //א,ב,ג,ד,ה,ו,ז,ח,ט
	{
		sprintf( yrcal, "%s%s%s%s%s%s", &heb7[hunds[1]][0], &heb7[hunds[2]][0], &heb7[hunds[3]][0], &heb7[tenss][0], "\"", &heb7[ones][0] );
	}
	else if (ones == 0)
	{
		if ( tenss != 0 )
		{
			sprintf( yrcal, "%s%s%s%s%s", &heb7[hunds[1]][0], &heb7[hunds[2]][0], &heb7[hunds[3]][0], "\"", &heb7[tenss][0]  );
		}
		else if (tenss == 0 )
		{
			if (hunds[0] == 3 )
			{
			sprintf( yrcal, "%s%s%s%s", &heb7[hunds[1]][0], &heb7[hunds[2]][0], "\"", &heb7[hunds[3]][0]  );
			}
			else if (hunds[0] == 2 )
			{
			sprintf( yrcal, "%s%s%s", &heb7[hunds[1]][0], "\"", &heb7[hunds[2]][0]);
			}
			else if (hunds[0] == 1 )
			{
			sprintf( yrcal, "%s", &heb7[hunds[1]][0] );
			}
		}
	}

	return yrcal;


}


//////////////////////////////////1////////////////////////////////
short WriteTables(char TitleZman[], short numsort, short numzman,
				  double *avekmxzman, double *avekmyzman, double *avehgtzman,
				  short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[])
///////////////////////////////////////////////////////////////////////
//this routine writes the zemanim html file
//the console or internet modes
//filroot$ = output file name with directory
//root$ = output file name without directory
//ext$ = the file extension either: "htm" (not supported now: "zip", "csv", or "xml")

/////////////////////////////////////////////////////////////////////////
{
	short j = 0, k = 0, i = 0, numday = 0, linnum = 0, m = 0, nn = 0;
	char hdr[255] = "";
	char address[255] = "";
	char l_progvernum[4] = "";
	char l_datavernum[3] = "";
	char l_yrheb[5] = "";
	bool closerow = false;

	char caldayHeb[50] = ""; //Hebrew date
	char caldayEng[20] = ""; //civil date
	char calday[100] = "";  //day and description
	char calhebweek[20] = ""; //Hebrew day of the week

	char ch[6] = "";
	char ch1[6] = "";
	char ch2[6] = "";
	char ch3[6] = "";
	//static char buff[255] = "";

	///////////////////////////////////////////////////////////////////////////
	short nweather = 3; //vary the percentage of summer/winter acc. to latitude
	///////////////////////////////////////////////////////////////////////////

	////////variables used for calculating zemanim//////////////

    short nyr, ns1, ns2, ns3, ns4, yro = 0;
	double ac, mc, ec, e2c, ap, mp, ob, td, lr, t6, t3;
	double dy, ms, aas, es, tf, air, fdy;
	double t3sub99 = 0;

	//dyo, hgt, are flags used for avoiding redundant astronomical constant
	//calculations, i.e., they keep track of the day number and height
	//in order to make sure that only one ast. constant initialization is done per
	//each unique day, dy, and height, hgt
	double dyo = 0, hgto = 0;

	////////////////////////////////////////////////////////////

	////////variables to keep track of day of week////////////

	short yl = 0;
	short addmon = 0;
	short dayweek = 0;
	short yrn = 1;
	short ntotshabos = 0;
	short nshabos = 0;
	short myear = StartYr;
	short fshabos = 0;
	short iheb = 0;

	//////////////variables from newzemanim/cal//////////
	double jdn;
	int nyl;
	short mday, mon;


	nshabos = 0;
	dayweek = rhday - 1;
	addmon = 0;
	ntotshabos = 0;
	fshabos = fshabos0;
 	if (!(optionheb)) iheb = 1;

	///////////////write html,version of sorted zmanim table////////////////////

	//write html header if not already written before writing the individual sunrise/sunset tables
	if (sunsyes == 0) WriteHtmlHeader();

	yl = YearLength( myear ); //length of starting year in days

	//convert Hebrew Year in English numbers to string
	itoa(g_yrheb, l_yrheb, 10);

	if ( !optionheb )
	{
		if ( strlen(Trim(eroscity)) <= 1 )
		{
			//strupr( citnamp ); //capitalize first letter of English name
			//strlwr( citnamp + 1); //lower case for rest of string
			sprintf( buff1, "%s%s%s%s%s%s%s", "<h2><center>", TitleLine, " Tables of ", citnamp, " for the Year ", l_yrheb, "</center></h2>" );
		}
		else
		{
			if ( strlen(Trim(citnamp)) <= 1 ) //probably Israeli City
			{
  				//strupr( eroscity ); //capitalize first letter of English name
 				//strlwr( eroscity + 1); //lower case for rest of string
   				sprintf( buff1, "%s%s%s%s%s%s%s", "<h2><center>", TitleLine, " Tables of ", eroscity, " for the Year ", l_yrheb, "</center></h2>" );
			}
			else
			{
				sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", TitleLine, " Tables of ", eroscity, " (", citnamp, ")", " for the Year ", l_yrheb, "</center></h2>" );
			}
		}
	}
	else
	{
		if ( strlen(Trim(&eroscity[0])) <= 1)
		{
			//sprintf( buff1, "%s%s%s%s%s%s", "<h2><center>", &heb3[1][0], hebcityname, &heb3[2][0], g_yrcal, "</h2>");
			sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb1[1][0], " \"", TitleLine, "\" ", &heb7[12][0], hebcityname, &heb3[2][0], g_yrcal, "</h2>");
		}
		else
		{
			//sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb3[12][0], " ", hebcityname, " (", eroscity, ") ", &heb3[2][0], g_yrcal, "</center></h2>");
            if (strstr(eroscountry, "Israel")) //use Hebrew neighborhood name
            {
                sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb1[1][0], " \"", TitleLine, "\" ", &heb3[29][0], " ", hebcityname, " (", EYhebNeigh, ") ", &heb3[2][0], g_yrcal, "</center></h2>");
            }
            else
            {
                sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb1[1][0], " \"", TitleLine, "\" ", &heb3[29][0], " ", hebcityname, " (", eroscity, ") ", &heb3[2][0], g_yrcal, "</center></h2>");
            }
		}
	}
	fprintf( stdout, "%s\n", &buff1[0]);

	//version number
	itoa( datavernum, l_datavernum, 10); //netzski6 version
	itoa( progvernum, l_progvernum, 10); //batfile version (version of profile collection)
	if (! optionheb )
	{
        sprintf( address, "%s%s%s%s%s", "&#169;", " Luchos Chai, info@chaitables.com; Tel/Fax: +972-2-5713765. Version: ", l_datavernum, ".", l_progvernum );
	}
	else
	{
		//sprintf( address, "%s%s%s%s%s", " ֲ©", &heb2[14][0], l_datavernum, ".", l_progvernum );
		sprintf( address, "%s%s%s%s%s", "&nbsp;&#169;", &heb2[14][0], l_datavernum, ".", l_progvernum );
	}


	//now write description of table
	strcpy( hdr, TitleZman );


	//Sponsor info
	sprintf(buff1, "%s%s%s", "<h4><center>", hdr, "</center></h4>");
	fprintf( stdout, "%s\n", buff1);
	sprintf(buff1, "%s%s%s", "<h5><center>", address, "</center></h5>");
	fprintf( stdout, "%s\n", buff1);

	if (Trim(&SponsorLine[0][0]) != "" && optionheb )
	{
        sprintf(buff1, "%s%s%s", "<h5><center>", &SponsorLine[0][0], "</center></h5>" );
		fprintf( stdout, buff1);
	}
	else if (Trim(&SponsorLine[1][0]) != "" && !optionheb )
	{
        sprintf(buff1, "%s%s%s", "<h5><center>", &SponsorLine[1][0], "</center></h5>" );
		fprintf( stdout, buff1);
	}

	//generate table column legend

	doclin[0] = '\0'; //initialize doclin
	nn = -1;
	for (m = numsort - 1; m >= 0; m--)
	{

		//attach letters or numbers to labels
		if (optionheb)
		{
			hebnum(m + 4, ch);
			strcpy( buff1, doclin );
			sprintf( doclin, "%s%s%s%s%s", ch, " = ", &zmantitles[m][0], " | ", buff1 );
		}
		else
		{
			nn++;
			itoa(nn + 4, ch, 10);
			strcpy( buff1, doclin );
			sprintf( doclin, "%s%s%s%s%s", buff1, " | ", ch, " = ", &zmantitles[nn][0] );
		}
	}

	//now add the the dates, and days legends
	if (optionheb)
	{
		hebnum(1, ch1);
		hebnum(2, ch2);
		hebnum(3, ch3);
		strcpy( buff1, doclin );
		sprintf( doclin, "%s%s%s%s%s%s%s%s%s%s%s%s%s%s", " | ", ch1, " = ", &heb3[11][0], " | ",
			     ch2, " = ", &heb3[10][0], " | ", ch3, " = ", &heb3[9][0], " | ", buff1 );
	}
	else
	{
		strcpy( buff1, doclin );
		sprintf( doclin, "%s%s","| 1 = hebrew date | 2 = day | 3 = civil date ", buff1 );
	}

	fprintf( stdout, "%s\n", "<hr/>"); //spacers between the headers and column headers
	fprintf( stdout, "%s\n", "<br/>");

	if (optionheb)
	{
		fprintf( stdout, "%s%s%s\n", "<font size = '3' dir='rtl'>", doclin, "</font>"); //print the legend for the csv, htm, and zip files
	}
	else
	{
		fprintf( stdout, "%s%s%s\n", "<font size = '3'>", doclin, "</font>"); //print the legend for the csv, htm, and zip files
	}

	fprintf( stdout, "%s\n", "<br/>"); //spacers between columnheaders and time
	fprintf( stdout, "%s\n", "<br/>");

	if (optionheb)
	{
		fprintf( stdout, "%s\n", "<table border='1' cellpadding='1' cellspacing='1' align='center' dir='rtl'>");
	}
	else
	{
		fprintf( stdout, "%s\n", "<table border='1' cellpadding='1' cellspacing='1' align='center'>");
	}


	//create table's column headers
	linnum = 1; //column headers are first row in the htm document

	doclin[0] = '\0'; //initialzie "doclin"

	//add new line tag for this first line
	fprintf( stdout, "%s\n", "    <tr>");

	nn = 0;
	for (m = 1; m <= numsort + 3; m++)
	{
		if (optionheb)
		{
			hebnum(m, ch);
		}
		else
		{
			nn = nn + 1;
			itoa( nn, ch, 10);
		}

		strcpy( doclin, ch);

		if (m == numsort + 3)
		{
			closerow = true;
		}
        	parseit( doclin, &closerow, linnum, 0);
		//doclin$ = doclin$ & comma$ & spacerb$ & ch$ & spacera$
	}

	//now create the actual table
	numday = -1;
	for (i = 1; i <= endyr; i++)
	{
		if (g_mmdate[1][i - 1] > g_mmdate[0][i - 1])
		{

			if (i > 1 && g_mmdate[0][i - 1] == 1)
			{
				//Rosh Chodesh is on January 1, so increment the year
				//(this happens in the year 2006, 2036)
				yrn++; //next year
				myear++;
				yl = YearLength( myear );

			}

			k = 0;
			//strcpy( totdoc, "\0");

			for (j = g_mmdate[0][i - 1]; j <= g_mmdate[1][i - 1]; j++)
			{
				numday++;
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf( stdout, "%s\n", "    <tr>");

				InsertHolidays(caldayHeb, caldayEng, calday, calhebweek,
					&i, &j, &k, &yrn, yl, &ntotshabos, &dayweek,
					&fshabos, &nshabos, &addmon, iheb );

				sprintf( doclin, "%s%s", caldayHeb, "\0" );
		        parseit( doclin, &closerow, linnum, 1);
				sprintf( doclin, "%s%s", calday, "\0");
		        parseit( doclin, &closerow, linnum, 2);
				sprintf( doclin, "%s%s", caldayEng, "\0" );
 		        parseit( doclin, &closerow, linnum, 3);

				nyr = atoi( Mid(caldayEng, 8, 4, buff) );
				dy = (double)j;

				if ( !(newzemanim( &nyr, &j, &air, &lr, &tf, &dyo, &hgto, &yro, &td,
					&dy, caldayHeb, &dayweek, &mp, &mc, &ap, &ac,  &ms, &aas, &es, &ob,
					&fdy,&ec, &e2c, avekmxzman, avekmyzman, avehgtzman, &ns1, &ns2,
					&ns3, &ns4, &nweather, &numzman, &numsort, &t6, &t3, &t3sub99,
					&jdn, &nyl, &mday, &mon, MinTemp, AvgTemp, MaxTemp, ExtTemp)) )
				{
					//exectued without error, all the zemanim strings are in array c_zmantimes
					//add them to table using parseit
					for (m = 0; m < numsort; m++)
					{
						sprintf( doclin, "%s%s", Trim(&c_zmantimes[sortzman[m]][0]), "\0" );
						if (m == numsort - 1)
						{
							closerow = true;
						}
						parseit( doclin, &closerow, linnum, 0);
					}

				}
				else
				{
					return -1; //error detected
				}


			}
		}
		else if (g_mmdate[1][i - 1] < g_mmdate[0][i - 1])
		{
			k = 0;
			//strcpy( totdoc, "\0" );

			for (j = g_mmdate[0][i - 1]; j <= yrend[0]; j++) //yl1%
			{
				numday++;
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf( stdout, "%s\n", "    <tr>");

				InsertHolidays(caldayHeb, caldayEng, calday, calhebweek,
					&i, &j, &k, &yrn, yl, &ntotshabos, &dayweek,
					&fshabos, &nshabos,	&addmon, iheb );
				sprintf( doclin, "%s%s", caldayHeb, "\0" );
		        parseit( doclin, &closerow, linnum, 1);
				sprintf( doclin, "%s%s", calday, "\0");
		        parseit( doclin, &closerow, linnum, 2);
				sprintf( doclin, "%s%s", caldayEng, "\0" );
		        parseit( doclin, &closerow, linnum, 3);

				nyr = atoi( Mid(caldayEng, 8, 4, buff) );
				dy = (double)j;

				if ( !(newzemanim( &nyr, &j, &air, &lr, &tf, &dyo, &hgto, &yro, &td,
					&dy, caldayHeb, &dayweek, &mp, &mc, &ap, &ac,  &ms, &aas, &es, &ob,
					&fdy,&ec, &e2c, avekmxzman, avekmyzman, avehgtzman, &ns1, &ns2,
					&ns3, &ns4, &nweather, &numzman, &numsort, &t6, &t3, &t3sub99,
					&jdn, &nyl, &mday, &mon, MinTemp, AvgTemp, MaxTemp, ExtTemp)) )
				{
					//exectued without error, all the zemanim strings are in array c_zmantimes
					//add them to table using parseit
					for (m = 0; m < numsort; m++)
					{
						sprintf( doclin, "%s%s", Trim(&c_zmantimes[sortzman[m]][0]), "\0" );
						if (m == numsort - 1)
						{
							closerow = true;
						}
						parseit( doclin, &closerow, linnum, 0);
					}

				}
				else
				{
					return -1; //error detected
				}

			}

			yrn++; //next year
			myear++;
			yl = YearLength( myear );

			for (j = 1; j <= g_mmdate[1][i - 1]; j++)
			{
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf( stdout, "%s\n", "    <tr>");

				numday++;

				InsertHolidays(caldayHeb, caldayEng, calday, calhebweek,
					&i, &j, &k, &yrn, yl, &ntotshabos, &dayweek,
					&fshabos, &nshabos,	&addmon, iheb );
				sprintf( doclin, "%s%s", caldayHeb, "\0" );
				parseit( doclin, &closerow, linnum, 1);
				sprintf( doclin, "%s%s", calday, "\0");
				parseit( doclin, &closerow, linnum, 2);
				sprintf( doclin, "%s%s", caldayEng, "\0" );
				parseit( doclin, &closerow, linnum, 3);

				nyr = atoi( Mid(caldayEng, 8, 4, buff));
				dy = (double)j;

				if ( !(newzemanim( &nyr, &j, &air, &lr, &tf, &dyo, &hgto, &yro, &td,
					&dy, caldayHeb, &dayweek, &mp, &mc, &ap, &ac,  &ms, &aas, &es, &ob,
					&fdy,&ec, &e2c, avekmxzman, avekmyzman, avehgtzman, &ns1, &ns2,
					&ns3, &ns4, &nweather, &numzman, &numsort, &t6, &t3, &t3sub99,
					&jdn, &nyl, &mday, &mon, MinTemp, AvgTemp, MaxTemp, ExtTemp)) )
				{
					//exectued without error, all the zemanim strings are in array c_zmantimes
					//add them to table using parseit
					for (m = 0; m < numsort; m++)
					{
						sprintf( doclin, "%s%s", Trim(&c_zmantimes[sortzman[m]][0]), "\0" );
						if (m == numsort - 1)
						{
							closerow = true;
						}
						parseit( doclin, &closerow, linnum, 0);
					}

				}
				else
				{
					return -1; //error detected
				}

			}
		}
	}
	fprintf( stdout, "%s\n", "          </table>");

	eroscity[0] = '\0'; //initialize eroscity

	return 0;
}


////////////////parseit//////////////////////////////////////////////
void parseit( char outdoc[], bool *closerow, short linnum, short mode)
{
	//parseit: adds tags to html, and accumlates text for the csv file

	char widthtim[4] = "";
	static char timlet[255] = "";
	//static char l_buff[255] = "";

	strcpy( timlet, Trim(&outdoc[0]) );

	if (WIDTHFIELD)
	{
		//program will independently deter__min witdth field
		itoa((int)(strlen(timlet) * 4.5), widthtim, 10);
	}
	else
	{
		//use set values for the different fields

		switch (mode)
		{
		case -1:
			itoa((int)(strlen(timlet) * 4.5), widthtim, 10);
			break;
		case 0: //time string, example 4:32, 17:32
			strcpy( widthtim, "30" );
			break;
		case 1: //Hebrew date
		    if (optionheb)
		    {
      			strcpy( widthtim, "70" );
            }
            else
            {
      			strcpy( widthtim, "90" );
            }
			break;
		case 2: //Day plus holidays
		    if (optionheb)
		    {
      			strcpy( widthtim, "120" );
            }
            else
            {
      			strcpy( widthtim, "200" );
            }
			break;
		case 3: //Civil date
			strcpy( widthtim, "80" );
			break;
		}
	}

	//parse lines looking for spaces

	if (linnum % 2 == 0)
	{

		sprintf( buff1, "%s%s%s", "        <td width=\"", widthtim, "\" bgcolor=\"#ffff00\" >" );
		fprintf(stdout, "%s\n", buff1 );
	}
	else
	{
		sprintf( buff1, "%s%s%s", "        <td width=\"", widthtim, "\" >" );
		fprintf(stdout, "%s\n", buff1 );
	}

	sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"2\" >", Trim(timlet), "</font></p>" );
	fprintf(stdout, "%s\n", buff1);
	fprintf(stdout, "%s\n", "        </td>");


	if (*closerow) //end of row signaled, so print end of row tags
	{
		fprintf(stdout, "%s\n", "    </tr>");
		*closerow = false;
	}


}


void heb12monthHTML(short setflag)
{
	char ch[6] = "";
	short j = 0;
	short i = 0;
	short iii = 0;
	short iheb = 0;
	short ii = 0;
	//static char l_buff[1255] = "";

	if (setflag == 0)
	{
		ii = 0;
	}
	else
	{
		ii = 1;
	}

	iheb = 0;
	if (!optionheb )
	{
		iheb = 1;
	}

	WriteHtmlHeader(); //initialize header information for stdout html output

	if (SRTMflag > 9)
	{
		fprintf(stdout, "%s\n", "<section>" );
		fprintf(stdout, "%s\n", " <style scoped>" );
		//fprintf(stdout, "%s\n", "  table { border-collapse: collapse; border: solid thin; }" );
		//fprintf(stdout, "%s\n", "  colgroup, tbody { }" );
		//fprintf(stdout, "%s\n", "  td { height: 0em; width: 30em; text-align: center; padding: 0; line-height:0px; }" );
		fprintf(stdout, "%s\n", "  *{ text-align: center; padding: 0; margin: auto; line-height: 1.3; }" );
		fprintf(stdout, "%s\n", " </style>" );
	}

	fprintf(stdout, "%s\n", "<table width=\"670\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\" >");
	fprintf(stdout, "%s\n", "    <col width=\"670\">");
//	if (SRTMflag != 10 && SRTMflag != 11)
//	{
		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"4\" style=\"font-size: 16pt\">", &storheader[ii][0][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");

		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\">", &storheader[ii][1][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");

		if (GoldenLight) //add information on azimuth and coordinates of first/last golden light
		{
			fprintf(stdout, "%s\n", "    <tr>");
			fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
			buff1[0] = '\0'; //clear buffer
			sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\">", &storheader[ii][8][0], "</font></p>" );
			fprintf(stdout, "%s\n", buff1 );
			fprintf(stdout, "%s\n", "        </td>");
			fprintf(stdout, "%s\n", "    </tr>");
		}

		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\">", &storheader[ii][3][0], "</font></p>" );// &storheader[ii][3][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");
//	}
//	else
//	{
//		fprintf(stdout, "%s\n", "    <tr>");
//		fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
//		buff1[0] = '\0'; //clear buffer
//		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\">", &storheader[ii][1][0], "</font></p>" );
//		fprintf(stdout, "%s\n", buff1 );
//		fprintf(stdout, "%s\n", "        </td>");
//		fprintf(stdout, "%s\n", "    </tr>");
//	}

	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"670\" border=\"1\" cellpadding=\"4\" cellspacing=\"0\" align=\"center\">");

	fprintf(stdout, "%s\n", "    <col width=\"15\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"41\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"46\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"16\">");

	fprintf(stdout, "%s\n", "    <tr>");
	fprintf(stdout, "%s\n", "        <td width=\"16\" valign=\"top\">");
	fprintf(stdout, "%s\n", "            <p align=\"left\"><br/>");
	fprintf(stdout, "%s\n", "            </p>");
	fprintf(stdout, "%s\n", "        </td>");

	if (optionheb)
	{
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][0][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"41\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][1][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][2][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][3][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][4][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][13][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][7][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][8][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][9][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][10][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"46\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][11][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][12][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"15\" valign=\"BOTTOM\">");
		fprintf(stdout, "%s\n", "            <p align=\"center\"><br/>");
		fprintf(stdout, "%s\n", "            </p>");
		fprintf(stdout, "%s\n", "        </td>");



	}
	else
	{
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][0][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"46\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][1][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][2][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][3][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][4][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][13][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][7][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][8][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][9][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][10][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"41\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][11][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\" valign=\"top\">");
    	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][12][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"15\" valign=\"BOTTOM\">");
		fprintf(stdout, "%s\n", "            <p align=\"center\"><br/>");
		fprintf(stdout, "%s\n", "            </p>");
		fprintf(stdout, "%s\n", "        </td>");
	}
	fprintf(stdout, "%s\n", "    </tr>");

	if ( !optionheb )
	{

		iii = -1;
		for (i = 1; i <= 5; i++) //5 rows of groups of 3 coloured, and 3 uncoloured
		{
			//check for shabboses (_) and for near mountains (*)
			for (j = 1; j <= 3; j++)
			{
				iii = iii + 1;
				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);;
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"16\" bgcolor=\"#ffff00\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"46\" bgcolor=\"#ffff00\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\" bgcolor=\"#ffff00\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 11, iii );

				fprintf(stdout, "%s\n", "        <td width=\"16\" bgcolor=\"#ffff00\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
			for (j = 1; j <= 3; j++)
			{
				iii = iii + 1;

				if (optionheb == true)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}
				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"16\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\" >");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"46\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"16\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
		}
	}
	else //direction is right to left
	{
		iii = -1;
		for (i = 1; i <= 5; i++) //5 rows of groups of 3 coloured, and 3 uncoloured
		{
			//check for shabboses (_) and for near mountains (*)
			for (j = 1; j <= 3; j++)
			{
				iii = iii + 1;
				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);;
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"16\" bgcolor=\"#ffff00\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\" bgcolor=\"#ffff00\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"46\" bgcolor=\"#ffff00\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 0, iii );

				fprintf(stdout, "%s\n", "        <td width=\"16\" bgcolor=\"#ffff00\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
			for (j = 1; j <= 3; j++)
			{
				iii = iii + 1;
				if (optionheb == true)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}
				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"16\">");
        		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\" >");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"46\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"16\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
		}
	}


	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"670\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"670\">");
	fprintf(stdout, "%s\n", "    <tr>");
	fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
   	buff1[0] = '\0'; //clear buffer
	sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][4][0], "</b></font></p>" );
	fprintf(stdout, "%s\n", buff1 );
	fprintf(stdout, "%s\n", "        </td>");
	fprintf(stdout, "%s\n", "    </tr>");
	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"670\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"670\">");
	fprintf(stdout, "%s\n", "    <tr>");
	fprintf(stdout, "%s\n", "        <td width=\"670\" valign=\"top\">");
	if (NearWarning[setflag]) //print near obstructions warning
	{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][5][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
	}
	if (InsufficientAzimuth[setflag]) //print insufficient azimuth range warning
	{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][6][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
	}
	if (GoldenLight)
	{
		if (SRTMflag == 12)
		{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font color=\"red\" size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[0][7][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		}
		else if (SRTMflag == 13)
		{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font color=\"red\" size=\"1\" style=\"font-size: 8pt\"\"><b>", &storheader[1][7][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		}
	}
	fprintf(stdout, "%s\n", "        </td>");
	fprintf(stdout, "%s\n", "    </tr>");
	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<p><br/><br/>");
	fprintf(stdout, "%s\n", "</p>");

	if (SRTMflag > 9)
	{
		fprintf( stdout, "%s\n", "</section>" );
	}

}

////////////////////timestring////////////////////////////////
void timestring( short ii, short k, short iii )
////////////////////////////////////////////////////////////////
//prints HTML time strings for heb12monthHTML, and heb13monthHTML routines
///////////////////////////////////////////////////////////////////
{

	static char ntim[255] = "";
	short positunder = 0;
	short positnear = 0;

	if (optionheb)
	{
		strcpy( ntim, &stortim[ii][endyr - k - 1][iii][0] );
	}
	else
	{
		strcpy( ntim, &stortim[ii][k][iii][0] );
	}

	positunder = InStr( ntim, "_" ); //Shabbosim
	if (positunder != 0) Mid( ntim, 1, positunder - 1); //remove the "_"

	positnear = InStr( ntim, "*" );
	if (positnear != 0) Mid( ntim, 1, positnear - 1 ); //remove the "*"

	if (positunder != 0)
	{
		sprintf( buff1, "%s%s%s", "<u>", ntim, "</u>" );
		strcpy( ntim, buff1 );
	}

	if (positnear != 0)
	{
		sprintf( buff1, "%s%s%s", "<font color=\"#33cc66\">", ntim, "</font>" );
		strcpy( ntim, buff1 );
	}

	sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font SIZE=\"1\" STYLE=\"font-size: 8pt\">", ntim, "</font></p>" );
	fprintf(stdout, "%s\n", buff1 );
	fprintf(stdout, "%s\n", "        </td>");

}

void heb13monthHTML(short setflag)
{
	char ch[6] = "";
	short j = 0;
	short i = 0;
	short iii = 0;
	short iheb = 0;
	short ii = 0;
	//static char l_buff[1255] = "";

	if (setflag == 0)
	{
		ii = 0;
	}
	else
	{
		ii = 1;
	}

	iheb = 0;
	if (!optionheb )
	{
		iheb = 1;
	}

	WriteHtmlHeader();

	if (SRTMflag > 9)
	{
		/*
		fprintf(stdout, "%s\n", "<section>" );
		fprintf(stdout, "%s\n", " <style scoped>" );
		fprintf(stdout, "%s\n", "  table { border-collapse: collapse; border: solid thin; }" );
		fprintf(stdout, "%s\n", "  colgroup, tbody { }" );
		fprintf(stdout, "%s\n", "  td { height: 0em; width: 30em; text-align: center; padding: 0; line-height:0px; }" );
		fprintf(stdout, "%s\n", " </style>" );
		*/
		fprintf(stdout, "%s\n", "<section>" );
		fprintf(stdout, "%s\n", " <style scoped>" );
		//fprintf(stdout, "%s\n", "  table { border-collapse: collapse; border: solid thin; }" );
		//fprintf(stdout, "%s\n", "  colgroup, tbody { }" );
		//fprintf(stdout, "%s\n", "  td { height: 0em; width: 30em; text-align: center; padding: 0; line-height:0px; }" );
		fprintf(stdout, "%s\n", "  *{ text-align: center; padding: 0; margin: auto; line-height: 1.3; }" );
		fprintf(stdout, "%s\n", " </style>" );
	}

	fprintf(stdout, "%s\n", "<table width=\"696\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"696\">");
//	if (SRTMflag != 10 && SRTMflag != 11)
//	{
		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\" bgcolor=\"#ffffff\">");
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"4\" style=\"font-size: 16pt\">", &storheader[ii][0][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");

		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\">");
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"4\" style=\"font-size: 8pt\">", &storheader[ii][1][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");

		fprintf(stdout, "%s\n", "    <tr>");
		fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\">");
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"4\" style=\"font-size: 8pt\">", &storheader[ii][3][0], "</font></p>" ); //&storheader[ii][3][0], "</font></p>" );
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "    </tr>");
//	}
//	else
//	{
//		fprintf(stdout, "%s\n", "    <tr>");
//		fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\">");
// 		buff1[0] = '\0'; //clear buffer
//		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"4\" style=\"font-size: 10pt\">", &storheader[ii][1][0], "</font></p>" );
//		fprintf(stdout, "%s\n", buff1 );
//		fprintf(stdout, "%s\n", "        </td>");
//		fprintf(stdout, "%s\n", "    </tr>");
//	}

	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"696\" border=\"1\" cellpadding=\"4\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"15\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"42\">");
	fprintf(stdout, "%s\n", "    <col width=\"45\">");
	fprintf(stdout, "%s\n", "    <col width=\"44\">");
	fprintf(stdout, "%s\n", "    <col width=\"44\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"41\">");
	fprintf(stdout, "%s\n", "    <col width=\"43\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"39\">");
	fprintf(stdout, "%s\n", "    <col width=\"51\">");
	fprintf(stdout, "%s\n", "    <col width=\"40\">");
	fprintf(stdout, "%s\n", "    <col width=\"14\">");
	fprintf(stdout, "%s\n", "    <tr valign=\"top\">");
	fprintf(stdout, "%s\n", "        <td width=\"14\">");
	fprintf(stdout, "%s\n", "            <p align=\"center\"><br/>");
	fprintf(stdout, "%s\n", "            </p>");
	fprintf(stdout, "%s\n", "        </td>");

	if (optionheb)
	{
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][0][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"51\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][1][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][2][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][3][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][4][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][5][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"43\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][6][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"41\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][7][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"44\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][8][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"44\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][9][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"45\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][10][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"42\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][11][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][12][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"15\">");
		fprintf(stdout, "%s\n", "            <p align=\"center\"><br/>");
		fprintf(stdout, "%s\n", "            </p>");
		fprintf(stdout, "%s\n", "        </td>");
	}
	else
	{
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][0][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"51\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][1][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][2][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][3][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][4][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"40\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][5][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"43\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][6][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"41\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][7][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"44\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][8][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"44\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][9][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"45\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][10][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"42\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][11][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"39\">");
	   	buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font face=\"Arial, sans-serif\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &monthh[iheb][12][0], "</b></font></font></p>");
		fprintf(stdout, "%s\n", buff1 );
		fprintf(stdout, "%s\n", "        </td>");
		fprintf(stdout, "%s\n", "        <td width=\"15\">");
		fprintf(stdout, "%s\n", "            <p align=\"center\"><br/>");
		fprintf(stdout, "%s\n", "            </p>");
		fprintf(stdout, "%s\n", "        </td>");
	}
	fprintf(stdout, "%s\n", "    </tr>");


	if ( !optionheb )
	{
		//**********finish here*******
		iii = -1;
		for (i = 1; i <= 5; i++) //5 rows of groups of 3 coloured, and 3 uncoloured
		{
			//check for shabboses (_) and for near mountains (*)
			for (j = 1; j <= 3; j++)
			{
				iii++;

				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"14\" bgcolor=\"#ffff00\">");
		   		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"51\" bgcolor=\"#ffff00\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"43\" bgcolor=\"#ffff00\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\" bgcolor=\"#ffff00\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\" bgcolor=\"#ffff00\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\" bgcolor=\"#ffff00\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"45\" bgcolor=\"#ffff00\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"42\" bgcolor=\"#ffff00\">");
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 12, iii );
				fprintf(stdout, "%s\n", "        <td width=\"14\" bgcolor=\"#ffff00\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
			for (j = 1; j <= 3; j++)
			{
				iii +=  1;

				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"14\">");
		   		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"51\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"43\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"45\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"42\">");
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 12, iii );
				fprintf(stdout, "%s\n", "        <td width=\"14\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
		}
	}
	else
	{
		//**********finish here*******
		iii = -1;
		for (i = 1; i <= 5; i++) //5 rows of groups of 3 coloured, and 3 uncoloured
		{
			//check for shabboses (_) and for near mountains (*)
			for (j = 1; j <= 3; j++)
			{
				iii++;

				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"14\" bgcolor=\"#ffff00\">");
		   		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 12, iii );
				fprintf(stdout, "%s\n", "        <td width=\"42\" bgcolor=\"#ffff00\">");
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"45\" bgcolor=\"#ffff00\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\" bgcolor=\"#ffff00\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\" bgcolor=\"#ffff00\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\" bgcolor=\"#ffff00\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"43\" bgcolor=\"#ffff00\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\" bgcolor=\"#ffff00\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"51\" bgcolor=\"#ffff00\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\" bgcolor=\"#ffff00\">");
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"14\" bgcolor=\"#ffff00\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
			for (j = 1; j <= 3; j++)
			{
				iii +=  1;

				if (optionheb)
				{
					hebnum(iii + 1, ch);
				}
				else
				{
					itoa( iii + 1, ch, 10);
				}

				fprintf(stdout, "%s\n", "    <tr valign=\"BOTTOM\">");
				fprintf(stdout, "%s\n", "        <td width=\"14\">");
		   		buff1[0] = '\0'; //clear buffer
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				//now deter__min if this time is underlined or shabbos and add appropriate HTML ******
				timestring( ii, 12, iii );
				fprintf(stdout, "%s\n", "        <td width=\"42\">");
				timestring( ii, 11, iii );
				fprintf(stdout, "%s\n", "        <td width=\"45\">");
				timestring( ii, 10, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\">");
				timestring( ii, 9, iii );
				fprintf(stdout, "%s\n", "        <td width=\"44\">");
				timestring( ii, 8, iii );
				fprintf(stdout, "%s\n", "        <td width=\"41\">");
				timestring( ii, 7, iii );
				fprintf(stdout, "%s\n", "        <td width=\"43\">");
				timestring( ii, 6, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 5, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 4, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 3, iii );
				fprintf(stdout, "%s\n", "        <td width=\"39\">");
				timestring( ii, 2, iii );
				fprintf(stdout, "%s\n", "        <td width=\"51\">");
				timestring( ii, 1, iii );
				fprintf(stdout, "%s\n", "        <td width=\"40\">");
				timestring( ii, 0, iii );
				fprintf(stdout, "%s\n", "        <td width=\"14\">");
				sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 10pt\"><b>", ch, "</b></font></p>" );
				fprintf(stdout, "%s\n", buff1 );
				fprintf(stdout, "%s\n", "        </td>");
				fprintf(stdout, "%s\n", "    </tr>");
			}
		}
	}

	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"696\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"696\">");
	fprintf(stdout, "%s\n", "    <tr>");
	fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\">");
   	buff1[0] = '\0'; //clear buffer
	sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][4][0], "</b></font></p>" );
	fprintf(stdout, "%s\n", buff1 );
	fprintf(stdout, "%s\n", "        </td>");
	fprintf(stdout, "%s\n", "    </tr>");
	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<table width=\"696\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" align=\"center\">");
	fprintf(stdout, "%s\n", "    <col width=\"696\">");
	fprintf(stdout, "%s\n", "    <tr>");
	fprintf(stdout, "%s\n", "        <td width=\"696\" valign=\"top\">");
	if (NearWarning[setflag]) //print near obstruction warning
	{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][5][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
	}
	if (InsufficientAzimuth[setflag]) //print insufficient azimuth range warning
	{
   		buff1[0] = '\0'; //clear buffer
		sprintf( buff1, "%s%s%s", "            <p align=\"center\"><font size=\"1\" style=\"font-size: 8pt\"><b>", &storheader[ii][6][0], "</b></font></p>" );
		fprintf(stdout, "%s\n", buff1 );
	}	
	fprintf(stdout, "%s\n", "        </td>");
	fprintf(stdout, "%s\n", "    </tr>");
	fprintf(stdout, "%s\n", "</table>");
	fprintf(stdout, "%s\n", "<p><br/><br/>");
	fprintf(stdout, "%s\n", "</p>");

	if (SRTMflag > 9)
	{
		fprintf( stdout, "%s\n", "</section>" );
	}


}

 ////////////////////////////////////////////////////////////////////////////////////////////
short WriteUserLog( int UserNumber, short zmanyes, short typezman, char *errorstr )
///////////////////////////////////////////////////////////////////////////////////////////
//WriteUserLog writes the user's inputed parameters to the userlog.log file
///////////////////////////////////////////////////////////////////////////////////////////
{

	//write userlog info
	char userloglin[255] = "";
	FILE *stream;

	sprintf( filroot, "%s%s", drivweb, "userlog.log");

	/* write closing clauses on html document */
	if ( !( stream = fopen( filroot, "a")) )
	{
	//ErrorHandler(4, 1); //can't open log file
	return -1;
	}

	if ( strlen(errorstr) > 2 ) //begin error report
	{
		fprintf( stream, "%s\n", "*******************Error/Warning***************" );
		fprintf( stream, "%s\n", errorstr );
	}

	//else log user information
	if (!strcmp(TableType, "BY" ))
	{
		sprintf( userloglin, "%s%s%s%s%s%s", "Israel-cities", ",", ",", ",", citnamp, "\0");
	}
	else
	{
		sprintf( userloglin, "%s%s%s%s%s%s", eroscountry, ",", erosareabat, ",", eroscity, "\0");
	}

	strcpy( buff1, userloglin );
	sprintf( userloglin, "%ld%s%s%s%f%s%f%s%g%s%d", UserNumber, ",", buff1, ",",
				eroslatitude, ",", eroslongitude, ",", eroshgt, ",", g_yrheb  );
	switch (orig_types)
	{
		case -1:
			strcat( userloglin, ",no sunrise or sunset tables" );
			break;
		case 0:
			strcat( userloglin, ",visible sunrise" );
			break;
		case 1:
			strcat( userloglin, ",visible sunset" );
			break;
		case 2:
			strcat( userloglin, ",mishor sunrise" );
			break;
		case 3:
			strcat( userloglin, ",astronomical sunrise" );
			break;
		case 4:
			strcat( userloglin, ",mishor sunset" );
			break;
		case 5:
			strcat( userloglin, ",astronomical sunset" );
			break;
		case 6:
			strcat( userloglin, ",visible sunrise and sunset" );
			break;
		case 7:
			strcat( userloglin, ",mishor sunrise and sunset" );
			break;
		case 8:
			strcat( userloglin, ",astron. sunrise and sunset" );
			break;
		case 10:
			strcat( userloglin, ",visible sunrise, 30m DTM" );
            break;
        case 11:
			strcat( userloglin, ",visible sunset, 30m DTM" );
            break;
		case 12:
			strcat( userloglin, ",Golden Light Sunrise, 30m DTM" );
            break;
        case 13:
			strcat( userloglin, ",Golden Light Sunset, 30m DTM" );
            break;
		default:
		{
		//ErrorHandler(4, 2); //error, unrecognizable types
		return -2;
		}
	  }

  	  if (zmanyes == 1)
	  {
		  strcat( userloglin, ",Yes Zemanim" );
	  }
	  else
	  {
		  strcat( userloglin, ",No Zemanim" );
	  }
  	  if (typezman != 0)
	  {
		  strcat( userloglin, ",Zman type: ");
		  strcat( userloglin, itoa(typezman, buff1, 10) );
	  }

	  if (optionheb)
	  {
		  strcat( userloglin, ",Hebrew" );
	  }
	  else
	  {
		  strcat( userloglin, ",English" );
	  }

	  fprintf( stream, "%s\n", userloglin );

	if ( strlen(errorstr) > 2 ) //complete error report
	{
		fprintf( stream, "%s\n", "************************************************" );
	}

	  fclose( stream );

	  return 0;
}

//////////////////////Error handler////////////////////////////
void ErrorHandler( short mode, short ier )
///////////////////////////////////////////////////////////////
//reads error code (ier) and prints appropriate message to output
// mode - identifies module reporting error
// ier - error code that identifies the specific error
////////////////////////////////////////////////////////////////////
{
	bool InternalError = false;
	short errornum = 0;
	static char errorstr[255] = "";

	//write Html header -- error message will be written instead of tables
	if (!HebrewInterface) optionheb = false; //write the html left to right if English interface
	if (HebrewInterface) optionheb = true; //write the html right to left
	WriteHtmlHeader(); //initialize header information for stdout html output

	//begin error report with standard header
	if (!HebrewInterface)
	{
		fprintf(stdout,"%s\n", "<tr><td><h2>The Chai Tables - Error Report</h2></td></tr>" );
	}
	else
	{
		fprintf(stdout,"%s%s%s\n", "<tr><td><h2>", &heb8[0][0], "</h2></td></tr>" );
	}

	fprintf(stdout,"%s\n", "<table border=\"0\">");
	fprintf(stdout,"%s\n", "<hr/>");

	//fprintf(stdout, "%s%d\n", "mode = ", mode );
	//fprintf(stdout, "%s%d\n", "ier = ", ier );

	switch (mode)
	{
		case 0: //used for diagnostics
			if (ier == -1) //stdout opening
			{
				fprintf( stdout, "%s\n", "<html>" );
				fprintf( stdout, "%s\n", "<body>" );
			}
			else if (ier == 0 ) //problem with cgi
			{
				//can't open "citynams.txt"
				InternalError = true;
				errornum = 0;
				strcpy (errorstr, "Cgi error: can't read html header, called in routine main" );
			}
			else if (ier == 1 ) //no table picked at all
			{
				if (!HebrewInterface)
				{
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>You haven't chosen any table to calculate!</strong></font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"%s\n", "</table>");
					fprintf(stdout,"%s\n", "<hr/>");
					fprintf(stdout,"%s\n", "<p>" );
					fprintf(stdout,"%s\n", "<font size=\"4\" color=\"black\">(To try again, use the return button of your browser)</font>" );
					fprintf(stdout,"%s\n", "</p>" );
				}
				else //Hebrew Interface
				{
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[14][0], "</strong></font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"%s\n", "</table>");
					fprintf(stdout,"%s\n", "<hr/>");
					fprintf(stdout,"%s\n", "<p>" );
					fprintf(stdout,"%s%s%s\n", "<font size=\"4\" color=\"black\">", &heb8[1][0], "</font>" );
					fprintf(stdout,"%s\n", "</p>" );
				}
			}
			break;
		case 1: //main routine errors
			if (ier == - 1)
			{
				//"PathsZones.txt" not found
				InternalError = true;
				errornum = 1;
				strcpy( drivweb, "/home/chkeller/public_html/new.chaitables.com/cgi-bin/" ); //default path to web files needed for writing error to log file
				strcpy(errorstr, "Can't open or find PathsZones_dtm.txt in routine LoadConstants" );
			}
			else if (ier == -2)
			{
				//"calhebmonths.txt" not found
				InternalError = true;
				errornum = 2;
				strcpy(errorstr, "Can't open or find calhebmonths.txt in routine LoadConstants" );
			}
			else if (ier == 2)
			{
				//error returned from CheckInputs, coordinates, and time zone
				//not numeric or not within range
		      if ( strstr(country, "Israel" ) )
				{
					if (!HebrewInterface)
					{
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>A problem has been detected with the coordinates you have entered</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>All coordinates must correspond to places within Eretz Yisroel (see below)</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The \"latitude\" must be greater than 29 and less than 34</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The \"longitude\" must be greater than -34 and less than -31 </strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>It is also possible that you haven't chosen your area and place properly in Step 1.</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
					}
					else //Hebrew InterfaceAstroNoonAstroNoon
					{
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[2][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[3][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[4][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[5][0], "</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[17][0], "</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
					}
				}
				else
				{
					if (!HebrewInterface)
					{
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>A problem has been detected with the coordinates or timezone you have entered</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The latitude must be greater than -66 degrees and less than 66 degrees</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The longitude must be between -180 and 180 degrees</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The timezone must be (+)/(-) for (-)/(+) (East/West) longitudes</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
					}
					else //Hebrew Interface
					{
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[1][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[6][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[7][0], "</strong></font></td>");
						fprintf(stdout,"\t%s\n", "</tr>");
						fprintf(stdout,"\t%s\n", "<tr>");
						fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[8][0], "</strong></font></td>");
						fprintf(stdout,"%s\n", "</tr>");
					}
				}

				fprintf(stdout,"%s\n", "</table>");
				fprintf(stdout,"%s\n", "<hr/>");
				fprintf(stdout,"%s\n", "<p></p>");

				if (!HebrewInterface)
				{
					fprintf(stdout,"%s\n", "<font size=\"4\" color=\"black\">(To try again, use the return button of your browser)</font>" );
				}
				else //Hebrew Interface
				{
					fprintf(stdout,"%s%s%s\n", "<font size=\"4\" color=\"black\">", &heb8[1][0], "</font>" );
				}
			}
			else if (ier == 3)
			{
				//can't open "citynams.txt"
				InternalError = true;
				errornum = 4;
				strcpy (errorstr, "Can't open or find citynams.txt in routine main" );
			}
			else if (ier == 4)
			{
				//Hebrew city equivalent to English city not found
				InternalError = true;
				errornum = 5;
				strcpy (errorstr, "Can't find Hebrew city equivalent to English city in routine main" );
			}
			else if (ier == 5)
			{
				//can't open cities1 file
				InternalError = true;
				errornum = 6;
				strcpy (errorstr, "Can't open or find cities1 file in routine main" );
			}
			else if (ier == 6)
			{
				//can't open cities5 file
				InternalError = true;
				errornum = 7;
				strcpy (errorstr, "Can't find the requested metro area in routine main" );
			}
			else if (ier == 7)
			{
				//can't open cities5 file
				InternalError = true;
				errornum = 8;
				strcpy (errorstr, "Can't open cities5.dir file in routine main" );
			}
			else if (ier == 8)
			{
				//can't open cities5 file
				InternalError = true;
				errornum = 9;
				strcpy (errorstr, "Error in reading line of text in cities5.dir file in routine main" );
			}
			else if (ier == 9)
			{
				//somehow the metro and city areas of the eros cities don't match
				//can't open cities3 file
				InternalError = true;
				errornum = 10;
				strcpy (errorstr, "Never found requested city or neighborhood--failure of javascript" );
			}
			else if (ier == 10)
			{
				//can't find city's coordinates
				InternalError = true;
				errornum = 11;
				strcpy (errorstr, "Can't find city's coordinates in routine main" );
			}
			else if (ier == 11)
			{
				//geocas didn't converge
				InternalError = true;
				errornum = 12;
				strcpy (errorstr, "Geographic to ITM coordinate conversion failed" );
			}
			else if (ier == 12)
			{
				//geocas didn't converge
				InternalError = true;
				errornum = 13;
				strcpy (errorstr, "At this time, DTM tables are only supported for places in Eretz Yisroel!" );
			}
			break;
		case 2: //RdHebStrings errors
			if (ier == 1)
			{
			//can't find Hebrew strings,
			InternalError = true;
			errornum = 13;
			sprintf( errorstr, "%s%s%s%s", "Can't open or find ", drivjk, "calhebrew_dtm.txt", " in routine RdHebStrings" );
			//strcpy (errorstr, "Can't open or find calhebrew_dtm.txt in routine RdHebStrings" );
			}
			else if (ier == 2)
			{
			//can't find or open calhebmonths.txt
			//can't find Hebrew strings,
			InternalError = true;
			errornum = 14;
			strcpy (errorstr, "Can't open or find calhebmonths.txt in routine RdHebStrings" );
			}
			break;
		case 3: //LoadParshiotNames errors
			if (ier == 1)
			{
				//can't open "calparshiot.txt" file
				InternalError = true;
				errornum = 15;
				strcpy (errorstr, "Can't open or find calparshiot.txt in routine LoadParshiotNames" );
			}
			else if (ier == 2)
			{
				//can't open "calparshiot_data.txt" file
				InternalError = true;
				errornum = 16;
				strcpy (errorstr, "Can't open or find calparshiot_data.txt in routine LoadParshiotNames" );
			}
			else if (ier == 3)
			{
				//can't find parshia
				InternalError = true;
				errornum = 17;
				strcpy (errorstr, "Can't open or find parshia in routine LoadParshiotNames" );
			}
			break;
		case 4: //WriteUserLog errors
			if (ier == 1)
			{
				//can't open log file
				InternalError = true;
				errornum = 18;
				strcpy (errorstr, "Can't open log file in routine WriteUserLog" );
			}
			else if (ier == 2)
			{
				//unrecognizable table type
				InternalError = true;
				errornum = 19;
				strcpy (errorstr, "Unrecognizable table type in routine WriteUserLog" );
			}
			break;
		case 5: //Caldirectories errors
			if (ier == 1)
			{
				//can't open eros city's netz folder
				InternalError = true;
				errornum = 20;
				/*
				if (!strcmp(TableType, "Astr"))
				{
                    sprintf (errorstr, "%s%s\n", "Can't open or find eros city's netz folder in routine Caldirectories, Astr = true ", country );
				}
				else
				{
                    sprintf (errorstr, "%s%s\n", "Can't open or find eros city's netz folder in routine Caldirectories, Astr = false ", country );
				}
				*/
                strcpy (errorstr, "Can't open or find eros city's netz folder in routine Caldirectories" );
			}
			else if (ier == 2)
			{
				//can't open eros city's skiy folder
				InternalError = true;
				errornum = 21;
				strcpy (errorstr, "Can't open or find eros city's skiy folder in routine Caldirectories" );
			}
			break;
		case 6: //calnearsearch
			if (ier == 1)
			{
				//can't open netz/skiy bat file of eros city
				InternalError = true;
				errornum = 22;
				strcpy (errorstr, "Can't open or find eros city's bat file in routine calnearsearch" );
			}
			else if (ier == 2)
			{
				//didn't find vantage point in eros city within selected search radius

				fprintf(stdout,"\t%s\n", "<tr>");

				if (!HebrewInterface)
				{
					fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>Couldn't find a vantage point within the chosen search radius</font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>You can increase the search radius and try again</strong></font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>Or click <a href=\"http://", "162.253.153.219","/chai_dtm_eng-BING.php\">here</a> to calculate visible zemanim for a specific place.</strong></font></td>");
				}
				else
				{
					fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[9][0], "</font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[10][0], "</font></td>");
					fprintf(stdout,"\t%s\n", "</tr>");
					fprintf(stdout,"\t%s\n", "<tr>");
					fprintf(stdout,"\t\t%s%s%s%s%s%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[21][0], " ","<a href=\"http://", "162.253.153.219","/chai_dtm_heb-BING.php\">", &heb8[22][0], "</a></font></td>");
				}

				fprintf(stdout,"\t%s\n", "</tr>");
				fprintf(stdout,"%s\n", "</table>");

				fprintf(stdout,"%s\n", "<hr/>");
				fprintf(stdout,"%s\n", "<p></p>" );

				if (!HebrewInterface)
				{
					fprintf(stdout,"%s\n", "<font size=\"4\" color=\"black\">(To try again, use the return button of your browser)</font>" );
				}
				else //Hebrew Interface
				{
					fprintf(stdout,"%s%s%s\n", "<font size=\"4\" color=\"black\">", &heb8[1][0], "</font>" );
				}

			}
			else if ( ier == 3)
			{
				//can't parse the eros city's bat file''
				InternalError = true;
				errornum = 23;
				strcpy (errorstr, "Can't parse the eros city's bat file in routine calnearsearch" );
			}
			break;
		case 7: //netzskiy
			if (ier == 1)
			{
				//can't open netz/skiy bat file of Israel neighborhood
				InternalError = true;
				errornum = 24;
				strcpy (errorstr, "Can't open or find the Israel neighborhood bat file in routine netzskiy" );
			}
			else if (ier == 2)
			{
				//didn't find vantage point in Israel neighborhood within search radius
			}
			break;
		case 8: //CreateHeaders
			if (ier == 1)
			{
				//astronomical constants need to update
				InternalError = true;
				errornum = 25;
				strcpy (errorstr, "Astronomical constants need to be updated in routine CreateHeaders" );
			}
			break;
		case 9: //ReadHebHeaders
			if (ier == 1)
			{
				//can't open EY city's netz/skiy directory
				InternalError = true;
				errornum = 26;
				strcpy (errorstr, "Can't open EY city's netz or skiy directory in routine ReadHebHeaders" );
			}
			else if (ier == 2)
			{
				//can't open "_port.sav" file -->just write warining on log
				strcpy (errorstr, "Can't find or open EY city's _port.sav file in routine ReadHebHeaders" );
				WriteUserLog(0, 0, 0, errorstr );
				return;
			}
			break;
		case 10: //netzski6 errors
			if (ier == -1)
			{
				//insufficient azimuth range
				InternalError = true;
				errornum = 27;
				strcpy (errorstr, "Insufficient azimuth range in routine netzski6" );
			}
			else if (ier == -2)
			{
				//can't open profile file
				InternalError = true;
				errornum = 28;
				strcpy (errorstr, "Can't open profile file in routine netzski6" );
			}
			else if (ier == -3)
			{
				//can't allocate memory for elev array
				InternalError = true;
				errornum = 29;
				strcpy (errorstr, "Can't allocate memory for elev array in routine netzski6" );
			}
			else if (ier == -4)
			{
                 //can't find one of the DTM tiles for IgnoreMissingTiles = false
                 InternalError = true;
                 errornum = 34;
                 strcpy(errorstr, "Detected missing DTM data; either (1) check 'Ignore Missing Tiles' or (2) decrease calculation range" );
            }
			else if (ier == -5)
			{
                 //can't allocate memory for uni5 array
                 InternalError = true;
                 errornum = 35;
                 strcpy(errorstr, "Can't allocate memory for calculations, please contact the webmaster" );
            }
			else if (ier == -6)
			{
                 //can't allocate memory for uni5 array
                 InternalError = true;
                 errornum = 36;
                 strcpy(errorstr, "Exceeded the maximum allowed number of DTM tiles, please decrease your lat. hw or log. width" );
            }
			else if (ier == -7)
			{
                 //can't allocate memory for uni5 array
                 InternalError = true;
                 errornum = 37;
                 sprintf(errorstr, "%s%i%s", "Detected a ", MAX_HEIGHT_DISCREP, "discrepency between the inputed and 30m DTM heights!" );
            }
			else if (ier == -8)
			{
                 //can't allocate memory for uni5 array
                 InternalError = true;
                 errornum = 38;
                 strcpy(errorstr, "You have chosen coordinates outside the supported georgraphic region" );
            }
			else if (ier == -9)
			{
                 //can't allocate memory for uni5 array or missing tiles
                 InternalError = true;
                 errornum = 38;
//				 if (!HebrewInterface)
//				 {

					sprintf( filroot, "%s%s%s%s%s%s%s%s%s\n", "Missing DTM tiles detected!",
                                                       "<br>",
                                                       "(a) Your chosen coordinates may be outside supported geographic regions (see <a href=\"http://www.chaitables.com/Introduction_to_DTM_calculations.html\">documentation</a>).",
                                                       "<br>", "(b) The chosen region for the calculation extends out to the ocean (there are no DTM tiles for oceans).",
                                                       "<br>",
                                                       "For case (b), check the \"Ignore Missing Tiles\" check box.",
                                                       "<br>",
                                                       "(If tiles are trully missing, report it using the link below.)" );
	                 strcpy(errorstr, filroot );
//				 }
//				 else //Hebrew interface
//				 {
//					 strcpy(errorstr, &heb8[18][0] );
//				 }
                 //strcpy(errorstr, "Missing DTM tiles detected.  Your chosen coordinates may be (a) outside supported geographic regions or (b) your calculation extends out to the ocean (see <a href=\"http://www.chaitables.com/Introduction_to_DTM_calculations.html\">documentation</a>).  For case (b), check the \"Ignore Missing Tiles\" check box" );
            }
			else if (ier == -10)
			{
				//Golden Light sunset error
				//signifies that the eastern horizon is too close
                 InternalError = true;
                 errornum = 60;

					sprintf( filroot, "%s%d%s%d%s%s%s%s%s%s%s%s\n", "The highest point in the eastern horizon is closer than ",MinGoldenDist, "km, or its view angle is larger than ",  MaxGoldenViewAng,

                                                       "<br>",
                                                       "Two possible options are:",
                                                       "<br>", "(a) Choose a different place to calculate the Golden Light sunset",
                                                       "<br>",
                                                       "(b) Choose to calculate the visible sunset instead of the Golden Light sunset",
                                                       "<br>",
                                                       "(N.b., Golden Light tables should not be used without consultation with a halochic authority)" );
	                 strcpy(errorstr, filroot );
			}
			else if (ier == -11)
			{
				//Golden Light sunrise error
				//signifies that the Western horizon is too close
                 InternalError = true;
                 errornum = 61;

					sprintf( filroot, "%s%d%s%d%s%s%s%s%s%s%s%s\n", "The highest point in the western horizon is closer than ",MinGoldenDist, "km, or has a view angle is larger than ",  MaxGoldenViewAng,
                                                       "<br>",
                                                       "Two possible options are:",
                                                       "<br>", "(a) Choose a different place to calculate the Golden Light surniset",
                                                       "<br>",
                                                       "(b) Choose to calculate the visible sunrise instead of the Golden Light sunrise",
                                                       "<br>",
                                                       "(N.b., Golden Light tables should not be used without consultation with a halochic authority)" );
	                 strcpy(errorstr, filroot );
			}
			break;
		case 11: //WriteHtmlHeader
			if (ier == 1)
			{
				//can't open html file
				InternalError = true;
				errornum = 30;
				strcpy (errorstr, "Can't open html file in routine WriteHtmlHeader" );
			}
			break;
		case 12: //PrintMultiColTable
			if (ier == 1)
			{
				//can't find zma table'
				InternalError = true;
				errornum = 31;
				strcpy (errorstr, "Can't open the zma file in routine PrintMultiColTable" );
			}
			break;
		case 13: //LoadTitlesSponsors
			if (ier == 1)
			{
				//"SponsorsLogo.txt" not found
				InternalError = true;
				errornum = 32;
				strcpy(errorstr, "Can't open or find SponsorLogo.txt in routine LoadTitlesSponsors" );
			}
			break;
		case 14: //WriteTables
			if (ier == 1)
			{
				InternalError = true;
				errornum = 33;
				strcpy (errorstr, "Incorrect parameters for WriteTalbes, reported in routine main" );
			}
			break;
		case 15: //WriteHtmlClosing
			if (ier == 1)
			{
				//stdout can't be closed, just notify log file'
				strcpy (errorstr, "Can't close html file in routine WriteHtmlClosing" );
				WriteUserLog(0, 0, 0, errorstr );
				return;
			}
			break;
		case 16: //overshaving error for Eretz Yisroel
			if (ier == 1)
			{
				InternalError = true;
				errornum = 50;
				//if (!HebrewInterface)
				//{
					sprintf( filroot, "%s%s%s%s%s%s%s\n", "In Eretz Yisroel, shaving is permitted only up to a techum Shabbos!*",
                                                       "<br>",
                                                       "Please limit the shaving to 1.5 km. and try again.",
													   "<br>",
													   "If you require a special exception, use the Feedback link below.",
                                                       "<br>", "(*P'sak of Rav Elyashiv z\"l).");
					strcpy(errorstr, filroot);
				//}
				//else //Hebrew
				//{
				//	strcpy(errorstr, &heb8[19][0]);
				//}

			}
			break;
		case 17: //errors in Temperatures
			if (ier > 0) 
			{
				//can't open one of the bil file
				InternalError = true;
				errornum = 60;
				sprintf(errorstr, "%s%d%s","Function Temperatures returned ier = ",ier,"can't open bil file min,(1-12), avg (13-24), max(25-36)");
			}
			break;
		default:
			break;
	}

	if (InternalError)
	{
		if (!HebrewInterface)
		{
			fprintf(stdout,"\t%s\n", "<tr>");
			sprintf( buff, "%s%d%s", "<td><font size=\"4\" color=\"red\"><strong>Your table can't be calculated due to internal error: #", errornum, "</strong></font></td>" );
			fprintf(stdout,"\t\t%s\n", buff );
			fprintf(stdout,"\t%s\n", "</tr>");
			fprintf(stdout,"\t%s\n", "<tr>");
			fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"red\"><strong>The error description follows:</strong></font></td>" );
		}
		else
		{
			fprintf(stdout,"\t%s\n", "<tr>");
			sprintf( buff, "%s%s%s%d%s", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[11][0], " ", errornum, "</strong></font></td>" );
			fprintf(stdout,"\t\t%s\n", buff );
			fprintf(stdout,"\t%s\n", "</tr>");
			fprintf(stdout,"\t%s\n", "<tr>");
			fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"red\"><strong>", &heb8[12][0], "</strong></font></td>" );
		}

		fprintf(stdout,"\t%s\n", "</tr>");
		fprintf(stdout,"\t%s\n", "<tr>");
		sprintf( buff, "%s%s%s", "<td><font size=\"4\" color=\"red\"><strong>", errorstr, "</strong></font></td>" );
		fprintf(stdout,"\t\t%s\n", buff );
		fprintf(stdout,"\t%s\n", "</tr>");
		fprintf(stdout,"\t%s\n", "<tr>");
		fprintf(stdout,"%s\n", "</table>");

		fprintf(stdout,"%s\n", "<hr/>");
		fprintf(stdout,"%s\n", "<p></p>" );

		fprintf(stdout,"%s\n", "<table border=\"0\">");
		fprintf(stdout,"\t%s\n", "<tr>");

		if (!HebrewInterface)
		{
			fprintf(stdout,"\t\t%s\n", "<td><font size=\"4\" color=\"black\"><strong>Please inform the webmaster using the feedback link below.</strong></font></td>");
		}
		else
		{
			fprintf(stdout,"\t\t%s%s%s\n", "<td><font size=\"4\" color=\"black\"><strong>", &heb8[13][0], "</strong></font></td>");
		}

		fprintf(stdout,"\t%s\n", "</tr>");
		fprintf(stdout,"\t%s\n", "<tr>");

		if (!HebrewInterface)
		{
			fprintf(stdout,"%s%s%s\n", "<td><font size=\"4\" color=\"black\"><a href=\"http://", &DomainName[0],"/chai_feedback_eng.php\">Feedback</a></font></td>");
		}
		else
		{
			fprintf(stdout,"%s%s%s%s%s\n", "<td><font size=\"4\" color=\"black\"><a href=\"http://", &DomainName[0],"/chai_feedback_heb.php\">", &heb8[16][0], "</a></font></td>");
		}

		fprintf(stdout,"%s\n", "</tr>");
		fprintf(stdout,"%s\n", "</table>");

		WriteUserLog(0, 0, 0, errorstr); //record the error in the log file
	}

	//////////html header diagnostics/////////////////////////
	//ier = checkenviron();
	///////////////////////////////////////////////

	WriteHtmlClosing();

}
////////////////////////////////////////////////////////////////

///////////////////////////Caldirectories////////////////////////////
short Caldirectories()
///////////////////////////////////////////////////////////////////
//reads in hebcitynames if required
///////////////////////////////////////////
{
	char *sun = "netz";
	//char batfile[255] = "";
	DIR *pdir;
	struct dirent *pent;

	myfile[0] = '\0'; //clear the buffer

	// search for near places for eros city

	switch (types)
	{
		case -1: //<--use sunrise vantage points to calculate zemanim w/o sunrise/sunset
        case 0:
		case 2:
		case 4:
		case 6:
		case 7:
		case 8:
        case 10:
		case 12:
			sun = "netz";

			// need to calculate visible sunrise
			// find "bat" file name of "netz" directory
			if (!eros && !astronplace )
			{
				sprintf( filroot, "%s%s%s", currentdir, sun, "/" );

				if ( !(pdir = opendir( filroot )) )
				{
					ErrorHandler(5, 1);
					return -1;
				}
				else
				{
					while ((pent=readdir(pdir)))
					{
						pent->d_name;
						//strncpy( batfile, pent->d_name, 8 );
						//batfile[8] = 0;
						if ( strstr( pent->d_name, ".bat" ) || strstr( pent->d_name, ".BAT" ) )
						{
							// found bat file, store its name
							sprintf( myfile, "%s%s", filroot, pent->d_name );
							break;
						}
					}
				}
				closedir(pdir);

				//open bat file and add places to the vantage point buffer
				if ( netzskiy( myfile, sun ) ) return -1; //error returns -1

			}
			else if (eros && !astronplace)
			{
				//bat file's name always consists of first 4 letters of currentdir
				sprintf( myfile, "%s%s%s%s%s", currentdir, sun, "/"
						, Mid(erosareabat, 1, 4, buff1), ".bat" );

				//search for places within search radius and add vantage points to buffer
				if ( calnearsearch( myfile, sun ) ) return -1; //error returns -1
			}
			break;
   	}

	switch (types)
	{
	case 1:
	case 3:
	case 5:
	case 6:
	case 7:
	case 8:
    case 11:
	case 13:
		sun = "skiy";

		// need to calculate visible sunset
		// find "bat" file name of "skiy" directory
		if (!eros && !astronplace)
		{
			sprintf( filroot, "%s%s%s", currentdir, sun, "/" );
			if ( !(pdir = opendir( filroot )) )
			{
				ErrorHandler(5, 2);
				return -1;
			}
			else
			{
				while ((pent=readdir(pdir)))
				{
					pent->d_name;
					if ( strstr( pent->d_name, ".bat" ) || strstr( pent->d_name, ".BAT" ) )
					{
						// found bat file, store its name
						sprintf( myfile, "%s%s%s", filroot, pent->d_name, "\0" );
						break;
					}
				}
			}
			closedir(pdir);

			//open bat file and add places to the vantage point buffer
			if ( netzskiy( myfile, sun ) ) return -1; //error returns -1

		}
		else if (eros && !astronplace)
		{
			//bat file name always consists of first 4 letters of currentdir
			sprintf( myfile, "%s%s%s%s%s", currentdir, sun, "/"
					, Mid(erosareabat, 1, 4, buff1 ), ".bat" );

			//search for places within search radius and add vantage points to buffer
			if (calnearsearch( myfile, sun )) return -1; //error returns -1
		}
		break;
	}

	//Hebrew calendar calculations///////////////////////////////////////

	HebCal(g_yrheb); //deter__min if it is hebrew leap year
			//and deter__min the beginning and ending daynumbers of the Hebrew year
	FirstSecularYr = (g_yrheb - 5758) + 1997;
	SecondSecularYr = (g_yrheb - 5758) + 1998;
	StartYr = FirstSecularYr; //default start year for calculations
	yrstrt[1] = 1; //Jan 1st is always beginning of the second civil year that is included
				//in the Hebrew Year.

	return 0;

}

//////////////////////calnearsearch///////////////////////////
short calnearsearch( char *bat, char *sun )
/////////////////////////////////////////////////////////////
//searches for near eros places within the searchradius from the center point
//given by eroslatitude, eroslongitude
//enters vantage places to memory buffer
///////////////////////////////////////////////////////////////////////////
{

	double ydegkm;
	double xdegkm;
	static char citnam[255] = "";
	static char bat_new[255] = "";
	double lon, lat, hgt, dist;
	short findnetz = 0;

	if (!strcmp(bat, "")) //check for nonexisting bat file
	{
		ErrorHandler(6, 1);
		return -1;
	}

	//zero all increments and initialize the ave coordinates
	if (InStr( sun, "netz" ))
	{
		nplac[0] = 0;
		avekmx[0] = 0.0;
		avekmy[0] = 0.0;
		avehgt[0] = 0.0;
	}
	if (InStr( sun, "skiy" ))
	{
		nplac[1] = 0;
		avekmx[1] = 0.0;
		avekmy[1] = 0.0;
		avehgt[1] = 0.0;
	}

	bool foundvantage = false;


	FILE *stream;

	//open bat file and deter__min which vantage points are within the search radiusAstroNoon(
	//try all case version of the bat file

	//first make copy of original filename for manipulations

	strcpy(bat_new, bat);

	short batlen = strlen(bat);

	if (!( stream = fopen( bat, "r" )) )
	{
		strlwr(bat_new + batlen - 8);
		if (!( stream = fopen( bat_new, "r")) )
		{
			strupr(bat_new + batlen - 8);
			if (!( stream = fopen( bat_new, "r")) )
			{
				if (InStr( sun, "skiy" ) && findnetz == 0 && types != 1 && types != 5 && types != 7)
				//probably eros place outside of USA or Eretz Yisroel and calculating mishor/astro times
				//so try using sunrise bat file instead to deter__min the average coordinates of vantage points
				{
					char * pch;
					pch = strstr (bat,"skiy");
					strncpy (pch,"netz",4);
					findnetz = 1;
				}
				if (!findnetz)
				{
					ErrorHandler(6, 1);
					return -1;
				}
			}
		}
	}

	if (findnetz) //look now for sunrise bat file instead of skiy file (see explan. above)
	{

		bat_new[0] = '\0'; //clear manipulation buffer
		strcpy(bat_new, bat);

		batlen = strlen(bat);

		if (!( stream = fopen( bat, "r" )) )
		{
			strlwr(bat_new + batlen - 8);
			if (!( stream = fopen( bat_new, "r")) )
			{
				strupr(bat_new + batlen - 8);
				if (!( stream = fopen( bat_new, "r")) )
				{
					ErrorHandler(6, 1);
					return -1;
				}
			}
		}
	}

	ydegkm = 180.0/ (pi * 6371.315); //degrees latitude per kilometer
	xdegkm = 180.0/(pi * 6371.315 * cos(eroslatitude * pi/180.0)); //degrees longituder per kilometer

	//first read in Hebrew city name, and geotz
	fgets_CR( doclin, 255, stream); //read in line of documentation

	//parse out the hebrew name and the time zone
	const short NUMENTRIES = 4;
	char strarr[NUMENTRIES][MAXDBLSIZE];

	if ( ParseString( doclin, ",", strarr, 2 ) )
	{
		ErrorHandler( 6, 3);; //can't read line of text in cities5 file
		return -1;
	}
	else //extract text
	{
        sprintf( hebcityname, "%s%s", &strarr[0][0], "\0" );
		Trim(Replace( hebcityname, '\"', ' ' )); //get rid of "\"
		geotz = atof( &strarr[1][0] );
	}

	//search for vantage points within the search radius
	while( ! feof(stream) )
	{
		//the eros city string on Windows is always 27 characters long
		//this will be different for Linux, so even though that the following
		//will work for Windows, it will not work for Linux
		//fscanf( stream, "\"%27c\",%Lg,%Lg,%Lg", doclin, &lonbat, &latbat, &hgtbat );
		//so it is necessary to parse out the name and coordinates in order to make
		//this operation platform independent
		fgets_CR( doclin, 255, stream );
		if ( ParseString( doclin, ",", strarr, NUMENTRIES ) )
		{
			ErrorHandler( 6, 3);; //can't read line of text in cities5 file
			return -1;
		}
		else //extract text
		{
			//bat file name
			strcpy( buff, &strarr[0][0] );
			//remove the Windows directory name
			//remove quotation marks if they exits
			Trim(Replace(buff, '"', ' ' ));
			//remove the "\"'s'
			Trim(Replace(buff, '\\', ' ' ));
			//rsuch extract the last 12 letters which compose the profile name
			Mid( buff, strlen(buff) - 11, 12 );
			buff[12] = 0;
			sprintf( citnam, "%s%s%s%s", currentdir, sun, "/", buff );

			lat = atof( &strarr[1][0] );
			lon = atof( &strarr[2][0] );
			hgt = atof( &strarr[3][0] );
		}

		if (InStr( strlwr(doclin), "version" )) //reached end of "bat" file, record version number
		{
			progvernum = (short)lat;
			switch ((short)lon)
			{
			case 0:
				SRTMflag = 0; //GTOPO30/SRTM30 (1000 m)
				accur[0] = cushion[1]; //cushions
				accur[1] = -cushion[1];
				distlim = distlims[1];
				break;
			case 1:
				SRTMflag = 1; //SRTM level 2 (30 m)
				accur[0] = cushion[3]; //cushions
				accur[1] = -cushion[3];
				distlim = distlims[3];
				break;
			case 2:
				SRTMflag = 2; //SRTM level 1 (90 m)
				accur[0] = cushion[2]; //cushions
				accur[1] = -cushion[2];
		        distlim = distlims[2];
				break;
			case 9:
				SRTMflag = 9; //Eretz Yisroel neighborhoods
				accur[0] = cushion[0]; //cushions
				accur[1] = -cushion[0];
				distlim = distlims[0];
				break;
			}
			break;
		}

		if (InStr( eroscountry, "Israel")) //EY neighborhoods
		{
           		dist = sqrt( pow( lat - eroslongitude, 2.0 ) + pow( lon - eroslatitude, 2.0 ) );
		}
		else //rest of the world
		{
           		dist = sqrt( pow((lon - eroslongitude) / xdegkm , 2.0) + pow((lat - eroslatitude) / ydegkm, 2 ));
		}

		if ( dist <= searchradius )
		{
            if ( !foundvantage ) foundvantage = true; //flag that found a vantage point

			//add this vantage point to the time calculation place buffer
			if ( InStr( sun, "netz" ) && findnetz == 0 )
			{

				nplac[0]++;
				strcpy( &fileo[0][nplac[0] - 1][0], citnam );
				latbat[0][nplac[0] - 1] = lat;
				lonbat[0][nplac[0] - 1] = lon;
				hgtbat[0][nplac[0] - 1] = hgt;

				avekmx[0] = avekmx[0] + lon;
			    avekmy[0] = avekmy[0] + lat;
			    avehgt[0] = avehgt[0] + hgt;

			}
			else if ( InStr( sun, "skiy" ) || findnetz == 1 )
			{

				nplac[1]++;
				strcpy( &fileo[1][nplac[1] - 1][0], citnam );
				latbat[1][nplac[1] - 1] = lat;
				lonbat[1][nplac[1] - 1] = lon;
				hgtbat[1][nplac[1] - 1] = hgt;

				avekmx[1] = avekmx[1] + lon;
			    avekmy[1] = avekmy[1] + lat;
			    avehgt[1] = avehgt[1] + hgt;
			}
		}
	}
	fclose( stream );

	if (!foundvantage )
	{
		ErrorHandler(6, 2); //print out that no vantage points were found
		                    //and may be necessary to increase search radius
        return -1;
	}
	else //finish averaging the coordinates
	{
		if ( InStr( sun, "netz" ) && findnetz == 0 )
		{
			avekmx[0] = avekmx[0] / nplac[0];
			avekmy[0] = avekmy[0] / nplac[0];
			avehgt[0] = avehgt[0] / nplac[0];

			if (exactcoordinates) //use eros place's coordinates
			{
				avekmx[0] = eroslongitude;
				avekmy[0] = eroslatitude;
				avehgt[0] = eroshgt;
			}
		}
		else if ( InStr( sun, "skiy" ) || findnetz == 1 )
		{
			avekmx[1] = avekmx[1] / nplac[1];
			avekmy[1] = avekmy[1] / nplac[1];
			avehgt[1] = avehgt[1] / nplac[1];

			if (exactcoordinates) //use eros place's coordinates
			{
				avekmx[1] = eroslongitude;
				avekmy[1] = eroslatitude;
				avehgt[1] = eroshgt;
			}
		}
	}

	return 0;
}

//////////////////////netzskiy///////////////////////////
short netzskiy( char *bat, char *sun )
/////////////////////////////////////////////////////////////
//enters vantage points contained in Israel cities "bat" files to buffer
//(Note: Israel neighborhood tables vantage points are found in routine "calnearsearch")
///////////////////////////////////////////////////////////////////////////
{

	double kmx, kmy, hgt;
	short pos = 0, pos1 = 0, found, addpos;
	short const MAXARRSIZE = 6; //number of data in line = array size of strarr
	char strarr[MAXARRSIZE][MAXDBLSIZE]; //array to contain the data elements


	if (!strcmp(bat, "")) //check for nonexisting bat file
	{
		ErrorHandler(7, 1);
		return -1;
	}

	//zero all increments and initialize the ave coordinates
	if (InStr( sun, "netz" ))
	{
		nplac[0] = 0;
		avekmx[0] = 0.0;
		avekmy[0] = 0.0;
		avehgt[0] = 0.0;
	}
	if (InStr( sun, "skiy" ))
	{
		nplac[1] = 0;
		avekmx[1] = 0.0;
		avekmy[1] = 0.0;
		avehgt[1] = 0.0;
	}

	FILE *stream, *stream2;

	if (!( stream = fopen( bat, "r" )) )
	{
		ErrorHandler(7, 1);
		return -1;
	}

	geotz = 2;

	//search for vantage points within the search radius

	found = 0;
	while( ! feof(stream) )
	{
		//the eros city string on Windows is always 27 characters long
		//this will be different for Linux, so even though that the following
		//will work for Windows, it will not work for Linux
		//fscanf( stream, "\"%27c\",%Lg,%Lg,%Lg", doclin, &lonbat, &latbat, &hgtbat );
		//so it is necessary to parse out the name and coordinates in order to make
		//this operation platform independent
		fgets_CR( doclin, 255, stream );
		pos = InStr( 1, doclin, "," );
		//check if quotation mark precedes is as is the format for old "bat" files
		addpos = 0;
		if ( InStr(pos-1, doclin, "\"") )
		{
			addpos = 1; //move back by one
		}
		Mid(doclin, pos - 12 - addpos, 12, buff );
		strcat( buff, "\0"); //terminate the string
		//remove quotation marks if they exits
		Replace(buff, '"', ' ' );
		sprintf( myfile, "%s%s%s%s", currentdir, sun, "/", Trim(buff) ); //myfile = path/name of city

		pos1 = InStr( pos + 1, doclin, "," );
		kmx = atof( Mid(doclin, pos + 1, pos1 - pos - 1, buff ));

		if (InStr( strlwr(doclin), "version" )) //reached end of "bat" file, record version number
		{
			progvernum = (short)kmx;
			break;
		}

		pos = InStr( pos1 + 1, doclin, "," );
		kmy = atof( Mid(doclin, pos1 + 1, pos - pos1 - 1, buff ));

		hgt = atof( Mid(doclin, pos + 1, strlen(doclin) - pos, buff ));

		///////////////////////////////////////added 101420//////////////////////////////////////////////
		//N.b., hgt read from the bat file may be typically be in error by 1.8 m (the ypical observer height added in)
		//so for more accurate accounting for this for calculating z'manei hayom, then set FixZmanHgts to true

		if (FixZmanHgts) {

			//read the height, hgt from the profile file which is always reliable

			//try opening file with the name as listed in the "bat" file as well
			//as well as all uppercases and all lowercaes
			short fillen = strlen(myfile);
    		if ( !(stream2 = fopen( myfile, "r" )) )
    		{
			 strlwr(myfile + fillen - 12); //try with all lowercases
    			 if ( !(stream2 = fopen( myfile, "r" )) )
    			 {
    				 strupr(myfile + fillen - 12); //try with all uppercase
    				 if ( !(stream2 = fopen( myfile, "r" )) )
    				 {
    	         		goto ny500; //can't open profile file for some reason, ignore error and skip try to find profile's hgt value
    				 }
    			 }
    		}

			doclin[0] = 0; //clear the buffer
			fgets_CR( doclin, 255, stream2 ); //read first header line of text mostly containly text, but sometimes the mean temperataure
			doclin[0] = 0; //clear the buffer
			fgets_CR( doclin, 255, stream2 ); //read second line containing details of scan as one line of text
			fclose(stream2);

			//parse it to obtain height recorded in profile file which is more reliable than the bat file
			if (!ParseString( doclin, ",", strarr, MAXARRSIZE ) )
			{
				hgt = atof( &strarr[2][0] ); //hgt is the third data ןאקצ in this doc line separated by commas
			}
			else
			{
				; //can't read data for some reason, ignore the error and just use the bat file's height
			}

		}

ny500:
		////////////////////////////////end addition 10/14/20////////////////////////////////////

		//add this vantage point to the time calculation place buffer
		if (InStr( sun, "netz" ))
		{

			nplac[0]++;
			strcpy( &fileo[0][nplac[0] - 1][0], myfile );
			latbat[0][nplac[0] - 1] = kmy;
			lonbat[0][nplac[0] - 1] = kmx;
			hgtbat[0][nplac[0] - 1] = hgt;

			avekmx[0] = avekmx[0] + kmx;
			avekmy[0] = avekmy[0] + kmy;
			avehgt[0] = avehgt[0] + hgt;

			found = 1;

		}
		else if (InStr( sun, "skiy" ))
		{

			nplac[1]++;
			strcpy( &fileo[1][nplac[1] - 1][0], myfile );
			latbat[1][nplac[1] - 1] = kmy;
			lonbat[1][nplac[1] - 1] = kmx;
			hgtbat[1][nplac[1] - 1] = hgt;

			avekmx[1] = avekmx[1] + kmx;
			avekmy[1] = avekmy[1] + kmy;
			avehgt[1] = avehgt[1] + hgt;

			found = 1;
		}

	}
	fclose( stream );
	if (found == 0) //didn't find vantage point within the requested search radius
	{
		ErrorHandler(7, 2);
		return -1;
	}

	//finish averaging the coordinates
	if (InStr( sun, "netz" ))
	{
		avekmx[0] = avekmx[0] / nplac[0];
		avekmy[0] = avekmy[0] / nplac[0];
		avehgt[0] = avehgt[0] / nplac[0];
	}
	else if (InStr( sun, "skiy" ))
	{
		avekmx[1] = avekmx[1] / nplac[1];
		avekmy[1] = avekmy[1] / nplac[1];
		avehgt[1] = avehgt[1] / nplac[1];
	}

	distlim = distlims[0];
	accur[0] = cushion[0];
	accur[1] = -cushion[0];

	return 0;

}

////////////////////////////SunriseSunset/////////////////////
short SunriseSunset( short types, short ExtTemp[] )
//////////////////////////////////////////////////////////////
//sets up headers, strings for tables
//////////////////////////////////////////////////////////////
{

	short ier, nsetflag, i = 0,tabletyp;
	double kmxo, kmyo, lt, lg, hgt, aprn;
	int N, E;
	short MinTemp[12],AvgTemp[12],MaxTemp[12];
	short ns1, ns2, ns3, ns4;
	double WinTemp = 0;
	double maxdecl;
	short maxang;
	double meantemp;
	double p;
	short itk;

    short nweather = 3;
	p = Pressure0; //1013.25f;

/*                                  '=0 for summer refraction */
/*                                  '=1 for winter refraction */
/*                                  '=2 for ASTRONOMICAL ALMANAC's refraction */
/*                                  '=3 ad hoc summer length as function of latitude
/*                                  '   ad hoc winter length as function of latitude


    re = 6371315.; //radius of earth in meters
    rekm = 6371.315f; //radius of earth in km

    aprn = 0.5;

/*       Nsetflag = 0 for astr sunrise/visible sunrise/chazos */
/*       Nsetflag = 1 for astr sunset/visible sunset/chazos */
/*       Nsetflag = 2 for mishor sunrise if hgt = 0/chazos or astro. sunrise/chazos */
/*       Nsetflag = 3 for mishor sunset if hgt = 0/chazos or astro. sunset/chazos */
/*       Nsetflag = 4 for mishor sunrise/vis. sunrise/chazos */
/*       Nsetflag = 5 for mishor sunset/vis. sunset/chazos */
/*       Nsetflag = 0 for visible sunrise calculated using 30 m DTM, SRTMflag = 10 */
/*       Nsetflag = 1 for visible sunset calculated using 30 m DTM, SRTMflag = 11 */

    /*
	if (CreateHeaders(types)) return -1; //create headers for all the different table types
         //return -1 if error detected
    */     

	////////////Printing Single Tables of Sunrise or Sunset////////////////////
	if (SRTMflag < 12) {

	///////////////////sunrises//////////////////////////////////////////
	for (i = 0; i < nplac[0]; i++ ) //sunrises
	{

		if (!astronplace)
		{
			strcpy( myfile, &fileo[0][i][0] );
			kmyo = latbat[0][i];
			kmxo = lonbat[0][i];
			hgt = hgtbat[0][i];
		}
		else //invert coordinates due to the way they were defined in the database of cities
		{
			if (types != 10 )
			{
				kmyo = eroslatitude;
				kmxo = eroslongitude;
			}
			else
			{
				kmxo = eroslongitude;
				kmyo = eroslatitude;
			}
			hgt = eroshgt;
			if (types == 10) aprn = erosaprn; //how much to shave in km
		}

	    if ( strstr( TableType,"Chai" ) != NULL && strstr( country, "Israel") != NULL ) //invert coordinates
		{
			kmxo = latbat[0][i];
			kmyo = lonbat[0][i];
		}

		switch (types)
		{
			case 0:
			case 6:
				nsetflag = 0; //visible sunrise/astr. sunrise/chazos
				break;
			case 2:
			case 8:
				nsetflag = 2;
				if (!ast) hgt = 0; //astronomical sunrise for ast = true, otherwise mishor sunrise
				break;
			case 4:
			case 7:
				nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
				break;
		    case 10:
                nsetflag = 0; //vis sunrise using 30m DTM
			    break;
                break;
		}

		if (!geo) 
		{
			//EY, so need to convert ITM to WGS84 lt,lg
			N = (int)(kmyo * 1000 + 1000000); 
			E = (int)(kmxo * 1000);
			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt, &lg); 
			lg = -lg; //convert this program's coordinate system where East longitude is positive
		} 
		else
		{
			//geo coordinates are identical with the kmx,kmy coordinates
			lt = kmyo;
			lg = kmxo;
		}

		ier = Temperatures(lt, -lg, MinTemp, AvgTemp, MaxTemp);

		if (ier) //error detected
		{
			ErrorHandler(17, ier);
			return -1;
		}

		//----Old Method: ad hoc summer and winter ranges for different latitudes----- 
		//----- VDW_REF = true, read WorldClim temperatures and stuff into min,avg, max temperature arrays
		//as well as calculationg the WinTemp used for finding probable inversion layer season within winter months
		//also calculate the longitude and latitude if not inputed

		InitWeatherRegions( &lt, &lg, &nweather,
     						&ns1, &ns2, &ns3, &ns4, &WinTemp,
							MinTemp, AvgTemp, MaxTemp, ExtTemp);

		//for new calcualtion method, deter__min the ground temperatures, minimum, average, and yearly mean
		if (VDW_REF) {

			//find average temperature to use in removing profile's TR 
			meantemp = 0.f;
			for (itk = 1; itk <= 12; ++itk) {
				meantemp += AvgTemp[itk - 1];
			}
			meantemp = meantemp / 12.f + 273.15f;

		}

		ier = netzski6( myfile, lg, lt, hgt, aprn, geotz, nsetflag, i, 
					    &nweather, &meantemp, &p, &ns1, &ns2, &ns3, &ns4, &WinTemp,
						MinTemp, AvgTemp, MaxTemp, ExtTemp);

		if (ier) //error detected
		{
			ErrorHandler(10, ier); //-1 = insufficient azimuthal range
			                       //-2 = can't open profile file
			                       //-3 = can't allocate memory for elev array
			                       //-4 = can't find DTM tile file
			return -1;
		}

	}
	
	if (CreateHeaders(types)) return -1; //create headers for all the different table types
         //return -1 if error detected	

	//now convert calculated sunrise time in fractional hours to time string and
	//print out sunrise table if only single table of sunrise was requested

	switch (types)
	{
	case 0:
	case 4:
	case 10:

		if ((types > 9) && parshiotflag == 2) //set time cushion similar to 30m SRTM (outside EY)
		{
			accur[0] = cushion[3];
			accur[1] = -cushion[3];
		}
		else if ((types > 9) && parshiotflag == 1) //30 m DTM, similar to 30 m DRTM (EY)
		{
			accur[0] = cushion[3];
			accur[1] = -cushion[3];
		}

		tabletyp = 0; //visible sunrise index of tims
		PrintSingleTable( 0, tabletyp ); //converts "tims" array in fractional hours to
		//time strings and prints out table to output
		break;
	case 6:
	case 7:
		if (PrintSeparate)
		{
			//print the sunrise and sunsets as separate tables
			tabletyp = 0; //visible sunrise index of tims
			PrintSingleTable( 0, tabletyp ); //converts "tims" array in fractional hours to
		}
		break;
	case 2: //astronomical/mishor sunrise

		accur[0] = cushion[4];
		accur[1] = -cushion[4];

		if (ast)
		{
			tabletyp = 2; //astronomical sunrise index of tims
		}
		else
		{
			tabletyp = 4; //mishor sunrise index of tims
		}

		PrintSingleTable( 0, tabletyp ); //converts "tims" array in fractional hours to
		//time strings and prints out table to output
		break;
	case 8:
		if (PrintSeparate)
		{

			accur[0] = cushion[4];
			accur[1] = -cushion[4];

			if (ast)
			{
				tabletyp = 2; //astronomical sunrise index of tims
			}
			else
			{
				tabletyp = 4; //mishor sunrise index of tims
			}

			PrintSingleTable( 0, tabletyp ); //converts "tims" array in fractional hours to
						//time strings and prints out table to output
		}
		break;
	}

	////////////////////////////sunsets////////////////////////////////////////////////////////////////


	for (i = 0; i < nplac[1]; i++ ) //sunsets
	{

		if (!astronplace)
		{
			strcpy( myfile, &fileo[1][i][0] );
			kmyo = latbat[1][i];
			kmxo = lonbat[1][i];
			hgt = hgtbat[1][i];
		}
		else //invert coordinates due to the way they were defined in the database of cities
		{
			if (types != 11)
			{
				kmyo = eroslatitude;
				kmxo = eroslongitude;
			}
			else
			{
				kmxo = eroslongitude;
				kmyo = eroslatitude;
			}
			hgt = eroshgt;
			if (types == 11) aprn = erosaprn;  //how much to shave in km
		}

		if ( strstr( TableType,"Chai" ) != NULL && strstr( country, "Israel") != NULL ) //invert coordinates
		{
			kmxo = latbat[1][i];
			kmyo = lonbat[1][i];
		}

		switch (types)
		{
			case 1:
			case 6:
				nsetflag = 1; //visible sunset/astr. sunset/chazos
				break;
			case 3:
			case 8:
				nsetflag = 3;
				if (!ast) hgt = 0; //astronomical sunset for ast = true, otherwise mishor sunset/chazos
				break;
			case 5:
			case 7:
				nsetflag = 5; //visible sunset/ mishor sunset/ chazos
				break;
			case 11:
				nsetflag = 1; //visible sunset using 30 m DTM
				break;
		}

		if (!geo) 
		{
			//EY, need to convert ITM to WGS84 
			N = (int)(kmyo * 1000 + 1000000); 
			E = (int)(kmxo * 1000);
			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt, &lg); 
			lg = -lg; //convert this program's coordinate system where East longitude is positive
		} 
		else
		{
			//geo coordinates are identical with the kmx,kmy coordinates
			lt = kmyo;
			lg = kmxo;
		}

		ier = Temperatures(lt, -lg, MinTemp, AvgTemp, MaxTemp);

		if (ier) //error detected
		{
			ErrorHandler(17, ier);
			return -1;
		}

		//----Old Method: ad hoc summer and winter ranges for different latitudes----- 
		//----- VDW_REF = true, read WorldClim temperatures and stuff into min,avg, max temperature arrays
		//as well as calculationg the WinTemp used for finding probable inversion layer season within winter months
		//also calculate the longitude and latitude if not inputed

		InitWeatherRegions( &lt, &lg, &nweather,
     						&ns1, &ns2, &ns3, &ns4, &WinTemp,
							MinTemp, AvgTemp, MaxTemp, ExtTemp);

		//for new calcualtion method, deter__min the ground temperatures, minimum, average, and yearly mean
		if (VDW_REF) {
			//read WorldClim bil files to deter__min the minimum and average temperatures of the year

			//find average temperature to use in removing profile's TR 
			meantemp = 0.f;
			for (itk = 1; itk <= 12; ++itk) {
				meantemp += AvgTemp[itk - 1];
			}
			meantemp = meantemp / 12.f + 273.15f;
		}

		ier = netzski6( myfile, lg, lt, hgt, aprn, geotz, nsetflag, i, 
					    &nweather, &meantemp, &p, &ns1, &ns2, &ns3, &ns4, &WinTemp,
						MinTemp, AvgTemp, MaxTemp, ExtTemp);

		if (ier) //error detected
		{
			ErrorHandler(10, ier); //-1 = insufficient azimuthal range
			                       //-2 = can't open profile file
			                       //-3 = can't allocate memory for elev array
			return -1;
		}

	}
	
	if (CreateHeaders(types)) return -1; //create headers for all the different table types
         //return -1 if error detected	

	//now convert calculated sunset time in fractional hours to time string and
	//print out sunset table if only single table of sunset was requested
	switch (types)
	{
		case 1:
		case 5:
		case 11:

			if (types > 10 && parshiotflag == 2) //set time cushions similar to 30 m SRTM
			{
				accur[0] = cushion[3];
				accur[1] = -cushion[3];
			}
			else if (types > 10 && parshiotflag == 1) //30m DTM, similar to 30 m SRTM
			{
				accur[0] = cushion[3];
				accur[1] = -cushion[3];
			}

			tabletyp = 1; //visible sunset index of tims
			PrintSingleTable( 1, tabletyp ); //converts "tims" array in fractional hours to
			//time strings and prints out table to output
		break;
	case 6:
	case 7:
		if (PrintSeparate)
		{
			//print the sunrise and sunsets as separate tables
		  tabletyp = 1; //visible sunrise index of tims
		  PrintSingleTable( 1, tabletyp ); //converts "tims" array in fractional hours to
							  //time strings and prints out table to output
		}
		break;
	case 3: //astronomical/mishor sunset

		accur[0] = cushion[4];
		accur[1] = -cushion[4];

		if (ast)
		{
			tabletyp = 3; //astronomical sunset index of tims
		}
		else
		{
			tabletyp = 5; //mishor sunset index of tims
		}

		PrintSingleTable( 1, tabletyp ); //converts "tims" array in fractional hours to
		//time strings and prints out table to output
		break;
	case 8:
		if (PrintSeparate)
		{

		  accur[0] = cushion[4];
		  accur[1] = -cushion[4];

		  if (ast)
		  {
			  tabletyp = 3; //astronomical sunset index of tims
		  }
		  else
		  {
			  tabletyp = 5; //mishor sunset index of tims
		  }

		  PrintSingleTable( 1, tabletyp ); //converts "tims" array in fractional hours to
						//time strings and prints out table to output
		}
	  break;
	}

	////////////////printing sunrises and sunsets together, printing zemanim///////////

	//////////////////Golden Light tables///////////////////////////////////////////
	} else if (SRTMflag > 11) {

		kmxo = eroslongitude;
		kmyo = eroslatitude;
		hgt = eroshgt;
		aprn = erosaprn;  //how much to shave in km
		GoldenLight = 1;

		if (!geo) 
		{
			//EY, need to convert ITM to WGS84 
			N = (int)(kmyo * 1000 + 1000000); 
			E = (int)(kmxo * 1000);
			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt, &lg); 
			lg = -lg; //convert this program's coordinate system where East longitude is positive
		} 
		else
		{
			//geo coordinates are identical with the kmx,kmy coordinates
			lt = kmyo;
			lg = kmxo;
		}

		ier = Temperatures(lt, -lg, MinTemp, AvgTemp, MaxTemp);

		if (ier) //error detected
		{
			ErrorHandler(17, ier);
			return -1;
		}

		//----Old Method: ad hoc summer and winter ranges for different latitudes----- 
		//----- VDW_REF = true, read WorldClim temperatures and stuff into min,avg, max temperature arrays
		//as well as calculationg the WinTemp used for finding probable inversion layer season within winter months
		//also calculate the longitude and latitude if not inputed

		InitWeatherRegions( &lt, &lg, &nweather,
     						&ns1, &ns2, &ns3, &ns4, &WinTemp,
							MinTemp, AvgTemp, MaxTemp, ExtTemp);

		//for new calcualtion method, deter__min the ground temperatures, minimum, average, and yearly mean
		if (VDW_REF) {
			//read WorldClim bil files to deter__min the minimum and average temperatures of the year

			//find average temperature to use in removing profile's TR 
			meantemp = 0.f;
			for (itk = 1; itk <= 12; ++itk) {
				meantemp += AvgTemp[itk - 1];
			}
			meantemp = meantemp / 12.f + 273.15f;
		}

		//N.b., if this is Golden Hour calculation, then first assume same temperatures as observer, calculate the rpfile
		//find the coordinates of the first golden light, then go back and use this point to define the ground temperature, etc.

		//deter__min maximum angular range from latitude
		maxdecl = (66.5 - fabs(lt)) + 23.5; //approximate formula for maximum declination as function of latitude (kmyo)
		//source: chart of declination range vs latitude from http://en.wikipedia.org/wiki/Declination
		maxang = (int)(fabs(asin(cos(maxdecl * cd)))/cd); //approximate formula for half maximum angle range as function of declination
		//source: http://en.wikipedia.org/wiki/Solar_azimuth_angle
		//add another few degrees on either side for a cushion for latitudes greater than 60
		if (fabs(lt) >= 60 && fabs(lt) < 62 )
		{
			maxang = maxang + 3;
		}
		else if (fabs(lt) >= 62)
		{
			maxang = maxang + 10;
		}

		if (maxang < 45) maxang = 45.; //set minimum maxang

		//maxang = MAX_ANGULAR_RANGE;

		if (SRTMflag == 12) {
			SRTMflag = 11; //calculate profile of western horizon (where golden light will first appear)
			ier = readDTM( &lg, &lt, &hgt, &maxang, &aprn, &p, &meantemp, SRTMflag );
			SRTMflag = 12; //restore the SRTMflag value
		} else if (SRTMflag == 13) {
			SRTMflag = 10; //calculate profile of eastern horizon (where golden light will last appear)
			ier = readDTM( &lg, &lt, &hgt, &maxang, &aprn, &p, &meantemp, SRTMflag );
			SRTMflag = 13; //restore the SRTMflag value
		}


		//if highest point is located at very near obstruction, then return just return error message

		if (ier) //error detected
		{
			ErrorHandler(10, ier); //-10 = Golden Light eastern horizon too close for golden light sunset calculations
								   //-11 = Golden Light western horizon too close for golden light sunrise calculations
			return -1;
		}

		if (GoldenLight == 1) {
			//change the coordinates to the highest horizon coordinates
			GoldenLight = 2;
			lt = Gold.vert[3];
			lg = -Gold.vert[2];
			hgt = Gold.vert[5];
		}

		//recalculate coordinate dependent groond temperatures
		ier = Temperatures(lt, -lg, MinTemp, AvgTemp, MaxTemp);

		if (ier) //error detected
		{
			ErrorHandler(17, ier);
			return -1;
		}

		//----Old Method: ad hoc summer and winter ranges for different latitudes----- 
		//----- VDW_REF = true, read WorldClim temperatures and stuff into min,avg, max temperature arrays
		//as well as calculationg the WinTemp used for finding probable inversion layer season within winter months
		//also calculate the longitude and latitude if not inputed

		InitWeatherRegions( &lt, &lg, &nweather,
     						&ns1, &ns2, &ns3, &ns4, &WinTemp,
							MinTemp, AvgTemp, MaxTemp, ExtTemp);

		//for new calcualtion method, deter__min the ground temperatures, minimum, average, and yearly mean
		if (VDW_REF) {
			//read WorldClim bil files to deter__min the minimum and average temperatures of the year

			//find average temperature to use in removing profile's TR 
			meantemp = 0.f;
			for (itk = 1; itk <= 12; ++itk) {
				meantemp += AvgTemp[itk - 1];
			}
			meantemp = meantemp / 12.f + 273.15f;
		}

		//now calculate the sunrise/sunset table
		if (SRTMflag == 12) {
			nsetflag = 0; //sunrise
			tabletyp = 0; //visible sunset index of tims
			SRTMflag = 10; //calculate visible sunrise profile/sunrise table for the golden light coordinates
		} else if (SRTMflag == 13) {
			nsetflag = 1; //sunset
			tabletyp = 1; //visible sunset index of tims
			SRTMflag = 11; //calculate visible sunset profile/sunset table for the  golden light coordinates
		}

		ier = netzski6( myfile, lg, lt, hgt, aprn, geotz, nsetflag, i, 
					    &nweather, &meantemp, &p, &ns1, &ns2, &ns3, &ns4, &WinTemp,
						MinTemp, AvgTemp, MaxTemp, ExtTemp);

		if (ier) //error detected
		{
			ErrorHandler(10, ier); //-1 = insufficient azimuthal range
			                       //-2 = can't open profile file
			                       //-3 = can't allocate memory for elev array
			return -1;
		}

		//restore SRTMflags for Golden Light to flag the correct titles and headers

		if (SRTMflag == 10)
		{
			SRTMflag = 12;
		}
		else if (SRTMflag == 11)
		{
			SRTMflag = 13;
		}

		if (CreateHeaders(types)) return -1; //create headers for all the different table types
			 //return -1 if error detected	

		//now convert calculated sunset time in fractional hours to time string and
		//print out sunset table if only single table of sunset was requested
		if (parshiotflag == 2) //set time cushions similar to 30 m SRTM
			{
				accur[0] = cushion[3];
				accur[1] = -cushion[3];
			}
		else if (parshiotflag == 1) //30m DTM, similar to 30 m SRTM
			{
				accur[0] = cushion[3];
				accur[1] = -cushion[3];
			}

		PrintSingleTable( nsetflag, tabletyp ); //converts "tims" array in fractional hours to
				//time strings and prints out table to output

		////////////////printing sunrises and sunsets together, printing zemanim///////////

		}


	return 0;

}

////////////////////create headers for the different tables////////////////
short CreateHeaders( short types )
///////////////////////////////////////////////////////////////////////////
{

	char l_buff4[1255] = "";
	char l_buff5[1255] = "";
	char l_buffcopy[1255] = "";
	char l_buffwarning[1255] = "";
	char l_buffazimuth[1255] = "";
	char i_buffergolden[1255] = "";
	short ii, iii, iv = 28, iiv = 32; //astronomical sunrise/sunset captions
	char st0[1255] = " of the Astronomical Sunrise of ";
	char st4[1255] = " of the Astronomical Sunset of ";
	char st2[1255] = "Based on the earliest astronomical sunrise for this place";
	char st3[1255] = "Based on the latest astronomical sunset for this place";

	if (ast) //astronomical times
	{
		ii = 25;
		iii = 31;
	}
	else if (!ast)
	{
		ii = 34;
		iii = 37;
	}

	if (optionheb) // 'hebrew captions
	{
		/*
		if (eros || astronplace )
		{
			sprintf( title, "%s%s%s%s", &heb1[1][0], "\"", TitleLine, "\"" );
		}
		else
		{
			sprintf( title, "%s%s%s%s", &heb1[1][0], "\"", TitleLine, "\"" );
		}
*/
		sprintf( title, "%s %s%s%s", &heb1[1][0], "\"", TitleLine, "\"" );

	  	//copyright and warning lines
		if (SRTMflag > 9)
		{
			//sprintf( l_buffcopy, "%s%s%s%s", &heb2[17][0], ". ", &heb2[16][0], " <a href=\"http://www.chaitables.com\">www.chaitables.com</a>"  );
			sprintf( l_buffcopy, "%s%s%s%s%4.1f%s%s%s", &heb2[17][0], ". ", &heb2[23][0], " ", vernum, ", ",&heb2[16][0], " <a href=\"http://www.chaitables.com\">www.chaitables.com</a>"  );
	        sprintf( l_buffwarning, "%s%3.1f%s", &heb2[15][0], distlim, &heb1[72][0] );
	        sprintf( l_buffazimuth, "%s", &heb2[24][0] );
			if (GoldenLight && sqrt(pow(Gold.vert[4],2.0) + pow((Gold.vert[5] - eroshgt) * 0.001,2.0)) > 1.5) 
			{
				sprintf( i_buffergolden, "%s", &heb1[73][0] ); //print warning since golden light is further than techum Shabbos from observer
			}
		}
		else
		{
		  	sprintf( l_buffcopy, "%s%s%s%s%s%s%s%d%s%d", &heb1[24][0], " ", &heb2[2][0], "\""
			                    , &heb2[3][0], " ", &heb2[4][0], datavernum, ".", progvernum);
	        sprintf( l_buffwarning, "%s%s%3.1f%s%s%s", &heb2[5][0], &heb2[6][0], distlim
			                                       , &heb2[7][0], "\"", &heb2[8][0] );
	        sprintf( l_buffazimuth, "%s", &heb2[24][0] );
		}
	}
        else //english captions
        {
		if ( eros || astronplace )
			{
			sprintf( title, "%s%s", TitleLine, " Tables" ); //Chai Tables
			}
		else
			{
			sprintf( title, "%s%s", TitleLine, " Tables" ); //"Bikurei Yosef Tables"
			}

		//copyright and warning lines
		//sprintf( l_buffcopy, "%s%s%s%d%s%d", "All times are according to Standard Time. Shabbosim are underlined. ", "&#169;", " Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: ", datavernum, ".", progvernum );

		if (SRTMflag > 9)
		{
			//sprintf( l_buffcopy, "%s%4.1f", "All <em>zemanim</em> are in Standard Time. <u>Shabbosim are underlined</u>. &#169;Luchos Chai&nbsp;<a href=\"http://www.chaitables.com\">www.chaitables.com</a>.  Version: ", vernum  );
			sprintf( l_buffcopy, "%s%4.1f", "Add a hour for DST. <u>Shabbosim are underlined</u>. &#169;Luchos Chai&nbsp;<a href=\"http://www.chaitables.com\">www.chaitables.com</a>.  Version: ", vernum  );
			sprintf( l_buffwarning, "%s%3.1f%s", "\"Green\" denotes diminished accuracy due to obstructions closer than ", distlim, " km." );
	        sprintf( l_buffazimuth, "%s", "''?'' denotes insufficient azimuth range for calculation (pick a place with an unobstructed view, or shave obstructions)" );
			if (GoldenLight && sqrt(pow(Gold.vert[4],2.0) + pow((Gold.vert[5] - eroshgt) * 0.001,2.0)) > 1.5) 
			{
				sprintf( i_buffergolden, "%s", "Warning: the place where Golden Light will appear may be beyound the techum! Consult a Rav before using this table." );
			}
		}
		else
		{
			//sprintf( l_buffcopy, "%s%s%s%d%s%d", "All times are according to Standard Time. Shabbosim are underlined. ", "&#169;", " Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: ", datavernum, ".", progvernum );
			sprintf( l_buffcopy, "%s%s%s%d%s%d", "Add a hour for DST. Shabbosim are underlined. ", "&#169;", " Luchos Chai, info@chaitables.com; Tel/Fax: +972-2-5713765. Version: ", datavernum, ".", progvernum );
			sprintf( l_buffwarning, "%s%3.1f%s", "The lighter colors denote times that may be inaccurate due to near obstructions (closer than ", distlim, " km).  It is possible that such near obstructions should be ignored." );
	        sprintf( l_buffazimuth, "%s", "''?'' denotes insufficient azimuth range for calculation (pick a place with an unobstructed view, or shave obstructions)" );
		}
	}

	/* //version 2.0 and beyond has extended the range of validity to any year
	if (g_yrheb >= ValidUntil + 3761 || g_yrheb < ValidUntil - 3861) //astronomical constants valid up to ValidUntil, and a 100 years before that
	{
		ErrorHandler(8, 1);  //Warning: contact webmanager to update astronomical constants
		return -1;
	}
	*/

 	if ( astronplace ) //user inputed name as city's hebrew name
	{
		if ( optionheb && InStr(hebcityname, &heb1[5][0] ) ) //add " to USA's hebrew name
		{
			Mid(astname, 1, strlen(astname) - 4, buff1); //remove USA in Hebrew and rewrite it
			hebcityname[0] = '\0'; //clear buffer
			sprintf( hebcityname, "%s%s%s%s", buff1, &heb1[6][0], "\"", heb1[7][0] );
		}
	}


	if (!eros && !astronplace) // Eretz Yisroel cities (not neighborhoods)
	{
		if (optionheb) //Hebrew Eretz Yisroel cities
		{
			//first check for cities_port.sav file containing headers for the visible times.
			//If exists, read its contents, and use the stored strings as the headers
			if (ReadHebHeaders()) return -1;
		}

		if ( !ast )
		{
			st0[0] = '\0';
			strcpy( st0, " of the Mishor Sunrise of " );
			st4[0] = '\0';
			strcpy( st4, " of the Mishor Sunset of " );
			st2[0] = '\0';
			strcpy( st2, "Based on the earliest mishor sunrise of this place" );
			st3[0] = '\0';
			strcpy( st3, "Based on the latest mishor sunset of this place" );
		}

		if (optionheb) //Hebrew Eretz Yisroel cities
		{
			switch (types)
			{
				case 0:
					//nsetflag = 0; //visible sunrise/astr. sunrise/chazos
					if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defaults
					{
						sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
						sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
						sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					}
					else //use headers saved in the _port.sav file
					{
						sprintf( buff1, "%s%s%s%s%s", &storheader[0][0][0], " ", &heb2[1][0], " ", g_yrcal );
						strcpy( &storheader[0][0][0], buff1 );
					}
					sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
					break;
				case 1:
					//nsetflag = 1; //visible sunset/astr. sunset/chazos
					if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 )// use defulats
					{
						sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
						sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
						sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					}
					else //use headers stored in the _port.sav file
					{
						sprintf( buff1, "%s%s%s%s%s", &storheader[1][0][0], " ", &heb2[1][0], " ", g_yrcal );
						strcpy( &storheader[1][0][0], buff1 );
					}
					sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
					break;
				case 2:
					//nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
					sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[ii][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					sprintf( &storheader[0][1][0], "%s", &heb1[iv][0] );
					sprintf( &storheader[0][2][0], "%s", &heb1[27][0] );
					sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					sprintf( &storheader[0][5][0], "%s", "\0");
					sprintf( &storheader[0][6][0], "%s", "\0");
					break;
				case 3:
					//nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
					sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[iii][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					sprintf( &storheader[1][1][0], "%s", &heb1[iiv][0] );
					sprintf( &storheader[1][2][0], "%s", &heb1[27][0] );
					sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					sprintf( &storheader[1][5][0], "%s", "\0" );
					sprintf( &storheader[1][6][0], "%s", "\0" );
					break;
				case 4:
				  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 )//use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
					  }
				  break;
				case 5:
				  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
					  }
				  break;
				case 6:
				  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
                      sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
					  }
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
					  }
					break;
				case 7:
				  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
					  }
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
					  }
				  break;
				case 8:
				  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
				  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s%s", title, &heb1[ii][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
				  sprintf( &storheader[0][1][0], "%s", &heb1[iv][0] );
				  sprintf( &storheader[0][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[0][5][0], "%s", "\0" );
				  sprintf( &storheader[0][6][0], "%s", "\0" );
				  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s%s", title, &heb1[iii][0], hebcityname, " ", &heb2[1][0], " ", g_yrcal );
				  sprintf( &storheader[1][1][0], "%s", &heb1[iiv][0] );
				  sprintf( &storheader[1][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[1][5][0], "%s", "\0" );
				  sprintf( &storheader[1][6][0], "%s", "\0" );
				  break;
			}
		}
		else //English Eretz Yisroel cities
		{
			//strupr( engcityname ); //capitalize first letter of English name
			//strlwr( engcityname + 1); //lower case for rest of string

			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, st0, engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", st2 );
			  sprintf( &storheader[0][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, st4, engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", st3 );
			  sprintf( &storheader[1][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[1][6][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, st0, engcityname, " for the Year ", g_yrheb  );
			  sprintf( &storheader[0][1][0], "%s", st2 );
			  sprintf( &storheader[0][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, st4, engcityname, " for the Year ", g_yrheb  );
			  sprintf( &storheader[1][1][0], "%s", st3 );
			  sprintf( &storheader[1][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  break;
			}
		}
	}
	else //eros or astronplace or both
	{
		ii = 41, iii = 42, iv = 43, iiv = 44; //astronomical sunrise/sunset captions
		st0[0] = '\0';
		strcpy( st0, " of the Astronomical Sunrise of " );
		st4[0] = '\0';
		strcpy( st4, " of the Astronomical Sunset of " );
		st2[0] = '\0';
		strcpy( st2, "Based on the earliest astronomical sunrise of this place" );
		st3[0] = '\0';
		strcpy( st3, "Based on the latest astronomical sunset of this place" );
		if ( !ast )
		{
			ii = 43;
			iii = 44;
			iv = 36;
			iiv = 40;
			st0[0] = '\0';
			strcpy( st0, " of the Mishor Sunrise of " );
			st4[0] = '\0';
			strcpy( st4, " of the Mishor Sunset of " );
			st2[0] = '\0';
			strcpy( st2, "Based on the earliest mishor sunrise of this place" );
			st3[0] = '\0';
			strcpy( st3, "Based on the latest mishor sunset of this place" );
		}

		if (optionheb)
		{
		    if ((!astronplace && types !=2 && types != 3 && types !=8) || (types > 9) )
		    {
				switch( SRTMflag )
				{
				case 0:
					sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[49][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

					sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[50][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

					sprintf(buff, "%s", &heb1[55][0] );
					break;

				case 1:
				  sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[51][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

				  sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[52][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

				  sprintf(buff, "%s", &heb1[56][0] );
				  break;

				case 2:
				  sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[53][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

				  sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[54][0], " ", searchradius, " ", &heb1[61][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

				  sprintf(buff, "%s", &heb1[55][0] );
				  break;

				case 9: //EY neighborhoods
				  sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[58][0], " ", searchradius, " ", &heb1[61][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

				  sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[59][0], " ", searchradius, " ", &heb1[61][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

				  sprintf(buff, "%s", &heb1[21][0] );
				  break;

				case 10: //30 m DTM sunrise  ***************TO DO -- add hebrew for this option --
				  //sprintf( buff2, "%s%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[69][0], eroscity, " ", &heb2[1][0], " ", g_yrcal, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunrise
				  sprintf( buff2, "%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[69][0], eroscity, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunrise

				  sprintf(buff, "%s", &heb1[55][0] );
                  break;
                case 11: //30 m DTM sunset ****************TO DO -- add hebrew for this option --

				  //sprintf( buff3, "%s%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[70][0], eroscity, " ", &heb2[1][0], " ", g_yrcal, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunset
				  sprintf( buff3, "%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[70][0], eroscity, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunset

				  sprintf(buff, "%s", &heb1[56][0] );
				  break;
				case 12: //30 m DTM sunrise  ***************TO DO -- add hebrew for this option --
				  //sprintf( buff2, "%s%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[69][0], eroscity, " ", &heb2[1][0], " ", g_yrcal, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunrise
				  sprintf( buff2, "%s%s%s%s %10.5lf %s %10.5lf %s %6.1lf %s%s%s%s %3.1lf %s%s", &heb1[74][0], eroscity, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunrise
				  sprintf( buff3, "%s %5.1lf %s%s %10.5lf %s %10.5lf %s %6.1lf %s, %5.1lf %s", &heb1[76][0]  
							   , Gold.vert[0], &heb9[14][0], &heb1[18][0], -Gold.vert[2], &heb1[17][0], Gold.vert[3], &heb1[67][0]
							   , Gold.vert[5], &heb1[68][0], Gold.vert[4], &heb9[16][0] );
				  sprintf(buff, "%s", &heb1[55][0] );
                  break;
                case 13: //30 m DTM sunset ****************TO DO -- add hebrew for this option --

				  //sprintf( buff3, "%s%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%s%s%s%3.1lf%s%s", &heb1[70][0], eroscity, " ", &heb2[1][0], " ", g_yrcal, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunset
				  sprintf( buff2, "%s%s%s%s %10.5lf %s %10.5lf %s %6.1lf %s%s%s%s %3.1lf %s%s", &heb1[75][0], eroscity, " (", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, &heb1[67][0], eroshgt, " ", &heb1[68][0], &heb1[71], " ", erosaprn, &heb1[72][0], " ) "  ); //sunset
				  sprintf( buff3, "%s %5.1lf %s%s %10.5lf %s %10.5lf %s %6.1lf %s, %5.1lf %s", &heb1[77][0]  
							   , Gold.vert[0], &heb9[14][0], &heb1[18][0], -Gold.vert[2], &heb1[17][0], Gold.vert[3], &heb1[67][0]
							   , Gold.vert[5], &heb1[68][0], Gold.vert[4], &heb9[16][0] );

				  sprintf(buff, "%s", &heb1[56][0] );
				  break;
				}
			}
			else if ((astronplace || types == 2 || types == 3 || types == 8) && (types < 10) )
			{
				if ( !ast ) // mishor
			   {

					sprintf( l_buff4, "%s%s%s%s%s%s", title, &heb1[34][0], hebcityname, &heb2[1][0], " ", g_yrcal );
					sprintf( l_buff5, "%s%s%s%s%s%s", title, &heb1[37][0], hebcityname, &heb2[1][0], " ", g_yrcal );

        		   if (astronplace) //mishor calculation for a certain coordinate
					{
						sprintf( buff2, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[47][0], hebcityname, &heb1[18][0]
								, eroslatitude, " ,", &heb1[17][0], eroslongitude  ); //sunrise
						sprintf( buff3, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[48][0], hebcityname, &heb1[18][0]
								, eroslatitude, " ,", &heb1[17][0], eroslongitude ); //sunset
					}
				   else //mishor calculation for a searchradius around a certain place
				   {
					   if (SRTMflag != 9)
					   {
							//sunrise
							sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[35][0], " ", searchradius, " ", &heb1[66][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  );

							//sunset
							sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[39][0], " ", searchradius, " ", &heb1[66][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  );
					   }
					   else if (SRTMflag == 9) //EY neighborhoods
					   {

						  sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[35][0], " ", searchradius, " ", &heb1[66][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

						  sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[39][0], " ", searchradius, " ", &heb1[66][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

					   }
				   }

				}
				else if (ast) // astronomical
		        {
					sprintf( l_buff4, "%s%s%s%s%s%s", title, &heb1[25][0], hebcityname, &heb2[1][0], " ", g_yrcal ); //sunrise
					sprintf( l_buff5, "%s%s%s%s%s%s", title, &heb1[31][0], hebcityname, &heb2[1][0], " ", g_yrcal ); //sunset

					if (astronplace) //astronomical calculation for a certain coordinate
				   {
					   sprintf( buff2, "%s%s%s%10.5lf%s%s%10.5lf%s%6.1f%s", &heb1[45][0], hebcityname, &heb1[18][0]
								, eroslongitude, " ,", &heb1[17][0], eroslatitude, &heb1[67][0], eroshgt, &heb1[68][0] ); //sunrise
					   sprintf( buff3, "%s%s%s%10.5lf%s%s%10.5lf%s%6.1f%s", &heb1[46][0], hebcityname, &heb1[18][0]
								, eroslongitude, " ,", &heb1[17][0], eroslatitude, &heb1[67][0], eroshgt, &heb1[68][0] ); //sunset
				   }
				   else //astronomical calculation for a searchradius around a certain place
				   {
					   if (SRTMflag !=9 )
					   {

							sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[28][0], " ", searchradius, " ", &heb1[66][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude
								, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

							sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[32][0], " ", searchradius, " ", &heb1[66][0], " ", eroscity, " ", "(", &heb1[17][0], eroslatitude
								, &heb1[18][0], eroslongitude, " ) "  ); //sunset

					   }
					   else if (SRTMflag == 9) //EY neighborhoods
					   {

						  sprintf( buff2, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[28][0], " ", searchradius, " ", &heb1[66][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunrise

						  sprintf( buff3, "%s%s%3.1f%s%s%s%s%s%s%s%10.5lf%s%10.5lf%s", &heb1[32][0], " ", searchradius, " ", &heb1[66][0], " ", EYhebNeigh, " ", "(", &heb1[17][0], eroslatitude, &heb1[18][0], eroslongitude, " ) "  ); //sunset

					   }
				   }

			   }

				sprintf(buff, "%s", &heb1[26][0] );
			}

			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			   //visible times header
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s", l_buff4 ); //, "%s", &heb1[ii][0] );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s", l_buff5 ); //, "%s", &heb1[iii][0] );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s", l_buff4 ); //, "%s", &heb1[ii][0] );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s", l_buff5 ); //, "%s", &heb1[iii][0] );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[1][6][0], "%s", "\0" );
			  break;
			case 10:
			  //nsetflag = 11; //visible sunrise for 30 m DTM
			  sprintf( &storheader[0][0][0], "%s %s %s %s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
            case 11:
			  //nsetflag = 11; //visible sunset for 30 m DTM
			  sprintf( &storheader[1][0][0], "%s %s %s %s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 12:
			  //nsetflag = 11; //visible sunrise for 30 m DTM
			  sprintf( &storheader[0][0][0], "%s %s %s %s%s%s", title, &heb1[74][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][8][0], "%s", buff3 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[0][7][0], "%s", i_buffergolden ); //add Golden Light warning
			  break;
            case 13:
			  //nsetflag = 11; //visible sunset for 30 m DTM
			  sprintf( &storheader[1][0][0], "%s %s %s %s%s%s", title, &heb1[75][0], hebcityname, &heb2[1][0], " ", g_yrcal );
			  sprintf( &storheader[1][1][0], "%s", buff2 );
			  sprintf( &storheader[1][8][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][7][0], "%s", i_buffergolden ); //add Golden Light warning
			  break;
            }
		}
		else
		{
			if ((!astronplace  && types !=2 && types != 3 && types !=8) || (astronplace && SRTMflag > 9 ))
			{
				switch( SRTMflag )
				{
				   case 0:
					  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "GTOPO30 DTM based calculations of the earliest sunrise within "
							   , searchradius, " km. around ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "GTOPO30 DTM based calculations of the latest sunset within "
							   , searchradius, " km. around ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( buff, "%s", "Terrain model: USGS GTOPO30" );
					  break;

				   case 1:
					  
					   if (InStr(eroscountry, "USA") )
					   {
						  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "30m NED DTM based calculations of the earliest sunrise within "
								   , searchradius, " km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
						  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "30m NED DTM based calculations of the latest sunset within "
								   , searchradius, " km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
						  sprintf( buff, "%s", "Terrain model: 30m NED" );
					   }
					   else
					   {
						  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-2 DTM based calculations of the earliest sunrise within "
								   , searchradius, " km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
						  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-2 DTM based calculations of the latest sunset within "
								   , searchradius, " km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
						  sprintf( buff, "%s", "Terrain model: SRTM-2" );
					   }

					  break;

				   case 2:
					  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-1 DTM based calculations of the earliest sunrise within "
							   , searchradius, " km. around ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-1 DTM based calculations of the latest sunset within "
							   , searchradius, " km. around ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( buff, "%s", "Terrain model: SRTM-1" );
					  break;

				   case 9:
					  sprintf( buff2, "%s%3.1f%s%s%s%7.3lf%s%7.3lf", "Israel DTM based calculations of the earliest sunrise within "
							   , searchradius, " km. around ", eroscity, "; ITMy: "
							   , eroslatitude, ", ITMx: ", eroslongitude );
					  sprintf( buff3, "%s%3.1f%s%s%s%7.3lf%s%7.3lf", "Israel DTM based calculations of the latest sunset within "
							   , searchradius, " km. around ", eroscity, "; ITMy: "
							   , eroslatitude, ", ITMx: ", eroslongitude );
					  sprintf( buff, "%s", "Terrain model: 25m DTM of Eretz Yisroel" );
					  break;
				   case 10:
					   /*
					  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f%s%6.1lf", "30m DTM based earliest sunrise, shaving "
							   , erosaprn, " km. from ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude, ", observer hgt: ", eroshgt + 1.8 );
					  */
					  /*
					  sprintf( buff2, "%s%s%s%d%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30m DTM based visible sunrise of "
							   , eroscity, " for the year: ", g_yrheb, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  */
					  sprintf( buff2, "%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30m DTM based visible sunrise of "
							   , eroscity, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  sprintf( buff, "%s", "Terrain model: 30m DTM" );
                      break;
                   case 11:
					   /*
 					  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f%s%6.1lf", "30m DTM based latest sunset, shaving "
							   , erosaprn, " km. from ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude, ", observer hgt: ", eroshgt + 1.8 );
					  */
					  /*
					  sprintf( buff3, "%s%s%s%d%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30 DTM based visible sunset of "
							   , eroscity, " for the year: ", g_yrheb, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  */
					  sprintf( buff3, "%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30 DTM based visible sunset of "
							   , eroscity, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  sprintf( buff, "%s", "Terrain model: 30m DTM" );
                     break;
				   case 12:
					   /*
					  sprintf( buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f%s%6.1lf", "30m DTM based earliest sunrise, shaving "
							   , erosaprn, " km. from ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude, ", observer hgt: ", eroshgt + 1.8 );
					  */
					  /*
					  sprintf( buff2, "%s%s%s%d%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30m DTM based visible sunrise of "
							   , eroscity, " for the year: ", g_yrheb, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  */
					  sprintf( buff2, "%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "Earliest Western Horizon Golden Light Sunrise of "
							   , eroscity, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  sprintf( buff3, "%s%5.1lf%s%10.5lf%s%10.5lf%s%6.1lf%s%5.1lf%s", "First Golden light will be seen at azimuth:"  
							   , Gold.vert[0], " degrees, longitude: ", -Gold.vert[2], ", latitude: ", Gold.vert[3], ", height: "
							   , Gold.vert[5], " m, ", Gold.vert[4], " km from the observer." );
					  sprintf( buff, "%s", "Terrain model: 30m DTM" );
                      break;
                   case 13:
					   /*
 					  sprintf( buff3, "%s%3.1f%s%s%s%10.5lf%s%10.5f%s%6.1lf", "30m DTM based latest sunset, shaving "
							   , erosaprn, " km. from ", eroscity, "; longitude: "
							   , eroslongitude, ", latitude: ", eroslatitude, ", observer hgt: ", eroshgt + 1.8 );
					  */
					  /*
					  sprintf( buff3, "%s%s%s%d%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "30 DTM based visible sunset of "
							   , eroscity, " for the year: ", g_yrheb, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  */
					  sprintf( buff2, "%s%s%s%10.5lf%s%10.5lf%s%6.1lf%s%3.1lf%s", "Latest Eastern Horizon Golden Light Sunset of "
							   , eroscity, " (longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: "
							   , eroshgt, " m, shaving: ", erosaprn, " km)" );
					  sprintf( buff3, "%s%5.1lf%s%10.5lf%s%10.5lf%s%6.1lf%s%5.1lf%s", "Last Golden light will be seen at azimuth:"  
							   , Gold.vert[0], " degrees, longitude: ", -Gold.vert[2], ", latitude: ", Gold.vert[3], ", height: "
							   , Gold.vert[5], " m, ", Gold.vert[4], " km from the observer." );
					  sprintf( buff, "%s", "Terrain model: 30m DTM" );
                     break;
				}
			}
			else if ((astronplace || types == 2 || types == 3 || types == 8) && (types != 10 && types != 11) )
			{
			   if ( !ast ) //mishor
			   {
					if (SRTMflag != 9 ) //not EY
					{
						sprintf( l_buff4, "%s%s%s%s%4d", title, " of the Mishor Sunrise of ", engcityname, " for the Year ", g_yrheb );
						sprintf( l_buff5, "%s%s%s%s%4d", title, " of the Mishor Sunset of ", engcityname, " for the Year ", g_yrheb );

						sprintf( buff2, "%s%s%s%10.5lf%s%10.5f", "Based on the mishor sunrise of ", eroscity
										, "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
						sprintf( buff3, "%s%s%s%10.5lf%s%10.5f", "Based on the mishor sunset of ", eroscity
										, "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
				  }
				  else  //EY neighborhoods
				  {
						sprintf( buff2, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Earliest mishor sunrise within "
									, searchradius, "km. around ", eroscity, "; ITMy: "
									, eroslatitude, ", ITMx: ", eroslongitude );

						sprintf( buff3, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Earliest mishor sunset within "
									, searchradius, "km. around ", eroscity, "; ITMy: "
									, eroslatitude, ", ITMx: ", eroslongitude );
				  }
			   }
			   else //astronomical
			   {
					if (SRTMflag != 9)
					{
					  sprintf( l_buff4, "%s%s%s%s%4d", title, " of the Astronomical Sunrise of ", engcityname, " for the Year ", g_yrheb );
					  sprintf( l_buff5, "%s%s%s%s%4d", title, " of the Astronomical Sunset of ", engcityname, " for the Year ", g_yrheb );

					  sprintf( buff2, "%s%s%s%10.5lf%s%10.5f%s%6.1f%s", "Based on the astronomical sunrise of ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: ", eroshgt, " m" );
					  sprintf( buff3, "%s%s%s%10.5lf%s%10.5f%s%6.1f%s", "Based on the astronomical sunset of ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude, ", height: ", eroshgt, " m" );
					}
					else //EY neighborhoods
					{
					  sprintf( buff2, "%s%3.1f%s%s%s%7.1lf%s%7.1lf%s%6.1f%s", "Latest astronomical sunrise within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude, ", height: ", eroshgt, " m" );

					  sprintf( buff3, "%s%3.1f%s%s%s%7.1lf%s%7.1lf%s%6.1f%s", "Latest astronomical sunset within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude, ", height: ", eroshgt, " m" );
					}
				}

				sprintf( buff, "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			}

			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s", l_buff4 );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s", l_buff5 );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[1][6][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s", l_buff4 );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[0][6][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s", l_buff5 );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  sprintf( &storheader[1][6][0], "%s", "\0" );
			  break;
			case 10:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
              break;
			case 11:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
              break;
			case 12:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Golden Light Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][8][0], "%s", buff3 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Golden Light Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff2 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[0][7][0], "%s", i_buffergolden );

              break;
			case 13:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Golden Light Sunrise for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[0][1][0], "%s", buff2 );
			  sprintf( &storheader[0][2][0], "%s", buff );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[0][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Golden Light Sunset for the Region of ", engcityname, " for the Year ", g_yrheb );
			  sprintf( &storheader[1][1][0], "%s", buff2 );
			  sprintf( &storheader[1][8][0], "%s", buff3 );
			  sprintf( &storheader[1][2][0], "%s", buff );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][6][0], "%s", l_buffazimuth );
			  sprintf( &storheader[1][7][0], "%s", i_buffergolden );
              break;
			}
		}
	}


	return 0;

}


////////////////////////////ReadHebHeaders/////////////////////
short ReadHebHeaders()
/////////////////////////////////////////////////////////////
//finds "cities_port.sav" file for Israeli cities, reads the
//file, and stores the values in the headers
/////////////////////////////////////////////////////////////////
{
	DIR *pdir;
	struct dirent *pent;
	FILE *stream;
	short nline = 0, tmpsetflag = 0;

	if ( !(pdir = opendir( currentdir )) )
	{
		ErrorHandler(9, 1);
		return -1;
	}
	else
	{
	   while ((pent=readdir(pdir)))
	   {
			pent->d_name;
			if ( strstr( strlwr(pent->d_name), "_port.sav" ) )
			{
				// found header file, open it up and read contents
				sprintf( filroot, "%s%s%s", currentdir, pent->d_name, "\0" );

				if ( !(stream = fopen( filroot, "r" ) ) )
				{
					ErrorHandler(9, 2); //add warning to log file

					storheader[tmpsetflag][0][0] = '\0';
					storheader[tmpsetflag][1][0] = '\0';
					storheader[tmpsetflag][2][0] = '\0';
					storheader[tmpsetflag][3][0] = '\0';
					storheader[tmpsetflag][4][0] = '\0';
					storheader[tmpsetflag][5][0] = '\0';

					return 0;
				}
				else // read contents
				{
					while ( !feof(stream) )
					{
						fgets_CR( doclin, 255, stream );

						if (nline <= 4 && strlen(doclin) > 2 ) //sunrise and sunset info
						{
							storheader[tmpsetflag][nline][0] = '\0'; //clear header array

							if (strlen(doclin) > 2) //sunrise and sunset info
							{
								//strncpy( &storheader[tmpsetflag][nline][0], doclin, strlen(doclin) - 1 );
								//storheader[tmpsetflag][nline][strlen(doclin) - 1] = 0;
								strcpy( &storheader[tmpsetflag][nline][0], doclin );

								if (tmpsetflag == 1 && nline == 1 && strstr( &storheader[tmpsetflag][nline][0], &heb1[12][0] ) )
									strcpy( &storheader[tmpsetflag][nline][0], &heb1[13][0] ); //typo fix for old sav files
							}
						}
						else if (nline == 5) //steps
						{
							if (RoundSeconds == -1 ) //no special RoundSeconds inputed by the cgi -- use defaults
							{
								if ( !strstr( doclin, "NA") ) //use recorded steps
								{
									steps[tmpsetflag] = atoi( doclin ); //round to steps seconds
								}
								else
								{
									steps[tmpsetflag] = 5; //use default steps of 5 seconds
								}
							}
							else //use special RoundSeconds entered by the cgi
							{
								steps[tmpsetflag] = RoundSeconds; //special steps requested
							}
						}
						else if (nline == 6) //accur
						{
							if (!strstr( doclin, "NA") )
							{
								accur[tmpsetflag] = atoi(doclin); // add/subtract cushion
							}
							else //use defaults
							{
								if (tmpsetflag == 0) accur[tmpsetflag] = cushion[0];
								if (tmpsetflag == 1) accur[tmpsetflag] = -cushion[0];
							}
						 }

						nline++;
						doclin[0] = '\0'; //clear the buffer

						if ( tmpsetflag == 0 && nline == 7) //finished sunrise info, start sunsets
						{
							// start over for sunset info
							tmpsetflag = 1;
							nline = 0;
						}
						else if (tmpsetflag == 1 && nline == 7) //finished sunset info
						{
							// start over for next batch of sunrise info, and overwrite last info
							tmpsetflag = 0;
							nline = 0;
						}

				 	}

					fclose( stream );
					break;
				}
			}
		}
	}
	closedir(pdir);

	return 0;
}

/////////////////////////check inputs/////////////////////////////////////////
short CheckInputs(char *strlat, char *strlon, char *strhgt)
/////////////////////////////////////////////////////////////////////////////
//checks if inputed coordinates are numerical and are within the permissable range
//
//eroslatitude must be > -66 and < 66 (below the polar circles)
//eroslongitude must be >= -180 and <= 180 (range of longitudes on the globe)
//eroshgt must be > -700 (larger than the lowest elevation on the earth)
//geotz if 0 < longitude < 180 then geotz must be negative
//      if -180 < longitude < 0 then geotz must be positive
//      if longitude = 0 then geotz must equal zero
//
/////////////////////////////////////////////////////////////////////
{

	short latlen = strlen(strlat);
	short lonlen = strlen(strlon);
	short hgtlen = strlen(strhgt);
	char ch[2] = "";
	double lat;
	double lon;
	double hgt;
	short i;

	//check if lat, lon, hgt are numbers
	for (i = 0; i < latlen; i++)
	{
		if ( !strstr( "-0123456789.", Mid(strlat, i, 1, ch) ) ) return -1; //found non numerical character
	}

	for (i = 0; i < lonlen; i++)
	{
		if ( !strstr( "-0123456789.", Mid(strlon, i, 1, ch) ) ) return -1; //found non numerical character
	}

	for (i = 0; i < hgtlen; i++)
	{
		if ( !strstr( "-0123456789.", Mid(strhgt, i, 1, ch) ) ) return -1; //found non numerical character
	}

	//now convert
	lat = atof( strlat );
	lon = atof( strlon );
	hgt = atof( strhgt );

    /*
    if ( strstr( country, "Israel" ) != NULL )
	{
		if (lat  < 80 || lat > 250 ) return -2;
		if (lon  < 60 || lon > 260 ) return -2;

		//using geographic coordinates in this version so just use check below
		geotz = 2;
	}
	*/
	//else
	//{

	if (lat + 66 < 0 || lat - 66 > 0 ) return -1;
	if (lon + 180 < 0 || lat - 180 > 0 ) return -1;

	if ( strcmp( TableType, "Chai") == 0 ) //Chai Tables
	{
        if ( strcmp(strlat, "0.0") == 0 && strcmp(strlon, "0.0") == 0 && strlen( eroscity ) < 2 ) return -1; //didn't enter coordinates
	}
	else if (strcmp(TableType, "Astr" ) == 0) //also check geotz
	{
		if ((lon > 0 && geotz > 0) || (lon < 0 && geotz < 0)) return -1; //coordinates have wrong sign
		if (strcmp(strlat, "0.0") == 0 && strcmp(strlon, "0.0") == 0  && strlen( astname ) < 2 ) return -1; //didn't enter coordinates
	}

	//if got here, then everything is OK!

	return 0;
}



/////////////////////////Hebrew Calendar/////////////////////////////
void HebCal(short yrheb)
/////////////////////////////////////////////////////////////////////
//calculates Hebrew year months for selected year and other Hebrew Year values
//Calculation follows procedure outlined in the Explanatory Supplement to the
//Astronomical Almanac, pp. 584-588
//////////////////////////////////////////////////////////////////////////
{

/****************declare  functions*************************/

        void newdate(short &nyear, short nnew, short flag, short kyr, short yrstep);

	void engdate();

	void dmh();



//declare local variables

	short n1day = 0;
	double cal1 = 0;
	double cal2 = 0;
	double cal3 = 0;
	short k = 0;
	short iheb = 0;
	short constdif = 0;
	short nnew = 0;
	short kyr = 0;
	short flag = 0;
	short nyear = 0;
	short yrstep = 0;
	short sthebyr = 0;
	//char buff[10];
	char monthhh[8] = "";


//	lpyr:

	//difdec = 0;
	//stdyy = 299; //English date paramters of Rosh Hashanoh year 1
	g_stdyy = 275;   //'English date paramters of Rosh Hashanoh 5758
	//styr = -3760;
	//sthebyr = 1;
	sthebyr = 5758; //starting reference year acc. to Hebrew Calendar
	g_styr = 1997; //starting Reference Year acc. to civil Calendar
	//yl = 366;
	g_yl = 365;
	//rh2 = 2; //Rosh Hashanoh year 1 is on day 2 (Monday)
	g_rh2 = 5;      //'Rosh Hashanoh 5758 is on day 5 (Thursday)
	g_yrr = 2; //5758 is a regular kesidrah year of 354 days
	g_leapyear = 0;
	//ncal1 = 2; //molad of Tishri 1 year 1-- day;hour;chelakim
	//ncal2 = 5;
	//ncal3 = 204;
	g_ncal1 = 5; //molad of Tishri 1 5758 day;hour;chelakim
	g_ncal2 = 4;
	g_ncal3 = 129;
	g_nt1 = g_ncal1;
	g_nt2 = g_ncal2;
	g_nt3 = g_ncal3;
	g_n1rhoo = g_nt1;
	g_leapyr2 = g_leapyear;
	g_n1yreg = 4; //change in molad after 12 lunations of reg. year
	g_n2yreg = 8;
	g_n3yreg = 876;
	g_n1ylp = 5; //change in molad after 13 lunations of leap year
	g_n2ylp = 21;
	g_n3ylp = 589;
	g_n1mon = 1; //monthly change in molad after 1 lunation
	g_n2mon = 12;
	g_n3mon = 793;
    g_n11 = g_ncal1; //initialize molad variables
	g_n22 = g_ncal2;
	g_n33 = g_ncal3;

	//Cls
	//chosen year to calculate monthly moladim
	//yrstep = yrheb - 1;
	yrstep = yrheb - sthebyr;

	//add synodic periods from the reference year in order to
	//deter__min moladim of desired Hebrew Year
	nyear = 0;
	flag = 0;
	for (kyr = 1; kyr <= yrstep; kyr++)
	{
		nyear = nyear + 1;
		nnew = 1;
  		newdate( nyear, nnew, flag, kyr, yrstep);
		g_n1rhooo = g_n1rho;
	}
	//now calculate molad of Tishri 1 of year after the desired Hebrew Year in order to
	//deter__min if the desired year is choser,kesidrah, or sholem
	g_leapyr2 = g_leapyear;
	nyear = nyear + 1;
	flag = 1;
	nnew = 0;
	newdate( nyear, nnew, flag, kyr, yrstep);

	g_n1rh = g_n1rhoo; //reinitialize variables after the "next year" calculation
	g_n2rh = g_nt2;
	g_n3rh = g_nt3;

	constdif = g_n1rhooo - g_rh2;
	rhday = g_rh2;
	if (rhday == 0)
	{
		rhday = 7;
	}

	//record year datum to be used for parshiot and holidays
	g_dayRoshHashono = rhday;
	g_yeartype = g_yrr;
	g_hebleapyear = false;
	if (g_leapyear == 1)
	{
		g_hebleapyear = true; //the desired Heberw Year is a leap year
	}

	dmh();

	g_dyy = g_stdyy; //- difdec%
    //convert stdyy to string and concatenate it to hdryr
	itoa(g_styr, buff, 10);
	strncpy( g_hdryr + 1, buff, 4);

	engdate();
	iheb = 1;
	strcpy( &g_mdates[0][0][0], &monthh[1][0][0]);
	strncpy( &g_mdates[1][0][0], g_dates, 11);
	g_mdates[1][0][12] = 0;
	g_mmdate[0][0] = g_dyy;

	//start and end day numbers of the first year
	yrstrt[0] = g_dyy;
	yrend[0] = 365;
	if (g_leapyr == 1)
	{
		yrend[0] = 366;
	}

	//calculate day number for first shabbos
	fshabos0 = 7 - rhday + g_dyy;

	//now calculate other molados and their english date
	endyr = 12;
	if (g_leapyear == 1)
	{
		endyr = 13;
	}
	for (k = 1; k <= endyr - 1; k++)
	{

		if (k < 5)
		{
			strcpy( monthhh, &monthh[1][k][0] );
		}
		else if (k == 5)
		{
			if (g_leapyear == 0)
			{
				strcpy( monthhh, &monthh[1][13][0] );
			}
			else
			{
				strcpy( monthhh, &monthh[1][5][0] );
			}
		}
		else if (k >= 6)
		{
			if (g_leapyear == 0)
			{
				strcpy( monthhh, &monthh[1][k + 1][0] );
			}
			else
			{
				strcpy( monthhh, &monthh[1][k][0] );
			}
		}

		g_n33 = g_n3rh + g_n3mon;  //calculations of mean molad: hours, minutes, chalokim
		cal3 = (double)g_n33 / 1080;
		g_ncal3 = Cint((cal3 - Fix(cal3)) * 1080.0);
		g_n22 = g_n2rh + g_n2mon;
		cal2 = ((double)g_n22 + Fix(cal3)) / 24;
		g_ncal2 = Cint((cal2 - Fix(cal2)) * 24.0);
		g_n11 = g_n1rh + g_n1mon;
		cal1 = ((double)g_n11 + Fix(cal2)) / 7;
		g_ncal1 = Cint((cal1 - Fix(cal1)) * 7.0);
		g_n1rh = g_ncal1;
		g_n2rh = g_ncal2;
		g_n3rh = g_ncal3;

		dmh();
		n1day = g_n1rh;

		//length of the Heberw months
		if (n1day == 0)
		{
			n1day = 7;
		}
		if (k == 1)
		{
			g_dyy = g_dyy + 30;
		}
		else if (k == 2)
		{
			if (g_yrr != 3)
			{
				g_dyy = g_dyy + 29;
			}
			if (g_yrr == 3)
			{
				g_dyy = g_dyy + 30;
			}
		}
		else if (k == 3)
		{
			if (g_yrr == 1)
			{
				g_dyy = g_dyy + 29;
			}
			if (g_yrr != 1)
			{
				g_dyy = g_dyy + 30;
			}
		}
		else if (k == 4)
		{
			g_dyy = g_dyy + 29;
		}
		else if (k == 5)
		{
			g_dyy = g_dyy + 30;
		}
		else if (k >= 6 && g_leapyear == 0)
		{
			if (k == 6)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 7)
			{
				g_dyy = g_dyy + 30;
			}
			if (k == 8)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 9)
			{
				g_dyy = g_dyy + 30;
			}
			if (k == 10)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 11)
			{
				g_dyy = g_dyy + 30;
			}
		}
		else if (k >= 6 && g_leapyear == 1)
		{
			if (k == 6)
			{
				g_dyy = g_dyy + 30;
			}
			if (k == 7)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 8)
			{
				g_dyy = g_dyy + 30;
			}
			if (k == 9)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 10)
			{
				g_dyy = g_dyy + 30;
			}
			if (k == 11)
			{
				g_dyy = g_dyy + 29;
			}
			if (k == 12)
			{
				g_dyy = g_dyy + 30;
			}
		}
        //strncpy( hdryr, "-", 1 );
        //convert styr to string and concatenate it to hdryr
		itoa(g_styr, buff, 10);
	    strncpy( g_hdryr + 1, buff, 4);

		//store the names, and begninning and ending dates and day
		//numbers of the rosh chadoshim
		g_dyy = g_dyy - 1;
		engdate();
		strncpy( &g_mdates[2][k - 1][0], g_dates, 11 );
		g_mmdate[1][k - 1] = g_dyy;
		g_dyy = g_dyy + 1;
		engdate();
		strcpy( &g_mdates[0][k][0], monthhh );
		strncpy( &g_mdates[1][k][0], g_dates, 11 );
		g_mdates[1][k][12] = 0;
		g_mmdate[0][k] = g_dyy;
	}

	250;
	if (g_styr == 1997)
	{
        g_styr = 1998;
	}
	g_dyy = g_dyy + 28;
    //strncpy( hdryr, "-", 1 );
    //convert styr to string and concatenate it to hdryr
	itoa(g_styr, buff, 10);
    strncpy( g_hdryr + 1, buff, 4);

	engdate(); //: LOCATE 13, 5: Print "end of year: "; dates$
	strncpy( &g_mdates[1][endyr - 1][0], g_dates, 11 );
	g_mdates[1][endyr - 1][12] = 0;
	g_mmdate[1][endyr - 1] = g_dyy;
	yrend[1] = g_dyy;
}

/////////////////////////newdate/////////////////////////
void newdate(short &nyear, short nnew, short flag, short kyr, short yrstep)
/////////////////////////////////////////////////////////
//calculates moladim of Hebrew Years from the reference year
//to the desired Hebrew year.  It uses the laws pertaining to
//the dechiyos to calculate Rosh Hashana for each of these
//years (based on the Explanatory Supplement to the Astronomical
//Almanac article on the Hebrew Calendar (ibid.).
///////////////////////////////////////////////////////////
{
	short difdyy = 0;
	short difrh = 0;
	double cal1 = 0;
	double cal2 = 0;
	double cal3 = 0;
	short n111 = 0;
	short n222 = 0;
	short n333 = 0;

	//initialize synodic period constants

	if (nyear == 20)
	{
		nyear = 1;
	}
	switch (nyear)
	{
		case 3:
		case 6:
		case 8:
		case 11:
		case 14:
		case 17:
		case 19:
			g_leapyear = 1;
			n111 = g_n1ylp;
			n222 = g_n2ylp;
			n333 = g_n3ylp;
			break;
		default:
			g_leapyear = 0;
			n111 = g_n1yreg;
			n222 = g_n2yreg;
			n333 = g_n3yreg;
			break;
	}

	g_n33 = g_n33 + n333;
	cal3 = g_n33 / 1080.0;
	g_ncal3 = Cint((cal3 - Fix(cal3)) * 1080.0);
	g_n22 = g_n22 + n222;
	cal2 = (g_n22 + Fix(cal3)) / 24.0;
	g_ncal2 = Cint((cal2 - Fix(cal2)) * 24.0);
	g_n11 = g_n11 + n111;
	cal1 = (g_n11 + Fix(cal2)) / 7.0;
	g_ncal1 = Cint((cal1 - Fix(cal1)) * 7.0);
	g_n11 = g_ncal1;
	g_n22 = g_ncal2;
	g_n33 = g_ncal3; //molad of Tishri 1 of this iteration

	g_n1rh = g_n11; //intialize variables for dechiyos calculation
	g_n2rh = g_n22;
	g_n3rh = g_n33;
	g_n1rho = g_n1rh;

	//now use dechiyos to deter__min which day of week Rosh Hashanoh falls on
	//////////////////////dechiyos//////////////////////////////////////////
	switch (nyear)
	{
		case 3:
		case 6:
		case 8:
		case 11:
		case 14:
		case 17:
		case 19: //first molad following a leap year
			if (g_n11 == 2 && (g_n22 + (g_n33 / 1080.0) > 15.545) )
			{	/////////////////dechiya # 4////////////////////////////////////////
				//Tishrey molad falls on day 2 at or after 15 hours, 589 chalokim
				//Postpone Rosh Hashona one day to day 3
				//(this prevents a leap year from falling short of 383 days).
				///////////////////////////////////////////////////////////////////
				g_n1rh = g_n1rh + 1;
				//difd% = 1
				goto HC500;
			}
			break;
	}
	if (g_n2rh >= 18) //Tishrey molad occurs at or after 18 hours (i.e., noon)
	{
		///////////////dechiya # 2////////////////////////////////
		//Rosh Hashona is postponed one day
		//this is based on gemorah that if lunar crescent hasn't been
		//sighted by noon, it will not be sighted until after 6 PM the next day
		/////////////////////////////////////////////////////////////////////
		g_n1rh = g_n1rh + 1;
		//difd% = 1
		if (g_n1rh == 8)
		{
			g_n1rh = 1;
		}
		if (g_n1rh == 1 || g_n1rh == 4 || g_n1rh == 6) //molad of Tishrey falls on day 1, or 4, or 6
		{
			////////////////dechiya #1/////////////////////
			//postpone Rosh Hashona by one day in order to prevent Hoshana Rabba from
			//occurring on Shabbos, and prevents Yom Kippur from occurring on the day
			//before or after the Shabbos
			/////////////////////////////////////////////////////////////////////
			g_n1rh = g_n1rh + 1;
		}
		goto HC500;
	}
	if (g_n1rh == 1 || g_n1rh == 4 || g_n1rh == 6) //molad of Tishrey falls on day 1, or 4, or 6
	{
		////////////////dechiya #1/////////////////////
		//postpone Rosh Hashona by one day in order to prevent Hoshana Rabba from
		//occurring on Shabbos, and prevents Yom Kippur from occurring on the day
		//before or after the Shabbos
		/////////////////////////////////////////////////////////////////////
		g_n1rh = g_n1rh + 1;
		//difd% = 1
	}
	if ((flag == 0 || flag == 1 || kyr == yrstep) && g_n1rh == 3 && (g_n2rh + (g_n3rh / 1080.0) > 9.188) )
		//The Tishrey molad falls on day 3 at or after 9 hours and 204 chalokim
	{
		switch (nyear + 1)
		{
			case 20: //ordinary years (not leap years)
			case 2:
			case 4:
			case 5:
			case 7:
			case 9:
			case 12:
			case 13:
			case 15:
			case 16:
			case 18:
				///////////////////dechiya #3////////////////////////////
				//postpone Rosh Hashona two days to day 5 which also satisfies
				//dechiya @1
				//this prevents an ordinary year from exceeding 355 days.  If the
				//////////////////////////////////////////////////////////////
				g_n1rh = 5;
				break;
		}
	}
	if (g_n1rh == 0)
	{
		g_n1rh = 7;
	}

HC500:  //calculate length of Hebrew year
	if (g_rh2 >= g_n1rh)
	{
		difrh = 7 - g_rh2 + g_n1rh;
	}
	if (g_rh2 < g_n1rh)
	{
		difrh = g_n1rh - g_rh2;
	}
	if (nnew == 1)
	{
		g_n1rhoo = g_n1rh;
	}
	if ((g_leapyear == 0 && difrh == 3) || (g_leapyear == 1 && difrh == 5))
	{
		g_yrr = 1;
	}
	else if ((g_leapyear == 0 && difrh == 4) || (g_leapyear == 1 && difrh == 6))
	{
		g_yrr = 2;
	}
	else if ((g_leapyear == 0 && difrh == 5) || (g_leapyear == 1 && difrh == 7))
	{
		g_yrr = 3;
	}
	if (g_leapyear == 0)
	{
		if (g_yrr == 1)
		{
			difdyy = 353;
		}
		if (g_yrr == 2)
		{
			difdyy = 354;
		}
		if (g_yrr == 3)
		{
			difdyy = 355;
		}
	}
	else if (g_leapyear == 1)
	{
		if (g_yrr == 1)
		{
			difdyy = 383;
		}
		if (g_yrr == 2)
		{
			difdyy = 384;
		}
		if (g_yrr == 3)
		{
			difdyy = 385;
		}
	}
	if (flag != 1)
	{
		g_dyy = g_stdyy + difdyy - g_yl; //- difdec%
		g_stdyy = g_dyy;
		g_styr += 1;

		g_yl = YearLength( g_styr ); //length of new civil year

		g_rh2 = g_n1rh; //rh2% is the day of the week (1-7) of the Rosh Hashonoh
		g_leapyr2 = g_leapyear; //store leap year parameter

		g_nt1 = g_n11;
		g_nt2 = g_n22;
		g_nt3 = g_n33;
	}
}

////////////////English Date calculator for HebCal routine////////////////
void engdate()
//////////////////////////////////////////////////////////////////////////
{
	//char buff[10];

	g_yl =  YearLength( g_styr); //length of this civil year

	if (g_dyy > g_yl ) //it is a new civil year
	{
		g_styr += 1;

        //convert styr to string and concatenate it to hdryr
        itoa(g_styr, buff, 10);
        strncpy( g_hdryr + 1, buff, 4);

		g_dyy = g_dyy - g_yl; //current daynumber of new civil year

		g_yl =  YearLength( g_styr); //length of new civil year
	}

	g_leapyr = 0;
	if (g_yl == 366) //leap years
	{
		g_leapyr = 1;
	}

	//actual English date
	EngCalDate( g_dyy, g_styr, g_yl, g_dates );

}

///////////////////////////dmh/////////////////////
void dmh()
///////////////////////////////////////////////////
//calculates printable version of Hours, Minutes, Chalokim for HebCal
////////////////////////////////////////////////////
{
	double minc = 0;

	g_Hourr = g_n2rh + 6;
	if (g_Hourr < 12)
	{
		g_gm = " PM night";
	}
	else if (g_Hourr >= 12)
	{
		g_Hourr = g_Hourr - 12;
		if (g_Hourr < 12)
		{
			if (g_Hourr == 0)
			{
				g_Hourr = 12;
			}
			g_gm = " AM";
		}
		else if (g_Hourr >= 12)
		{
			g_Hourr = g_Hourr - 12;
			if (g_Hourr == 0)
			{
				g_Hourr = 12;
			}
			g_gm = " PM afternoon";
		}
	}
	minc = ((double)g_n3rh * 60.0 ) / 1080.0;
	g_Min = (short)Fix(minc);
	g_hel = Cint((minc - g_Min) * 18);

}

/////////////////////Prints single table/////////////////////
short PrintSingleTable(short setflag, short tbltypes)
/////////////////////////////////////////////////////////
{
	short sp = 0;

	short j = 0;
	short i = 0;
	short k = 0;

	short yl = 0;
	short addmon = 0;
	short dayweek = 0;
	short yrn = 0;
	short ntotshabos = 0;
	short nshabos = 0;
	short myear = 0;
	short fshabos = 0;
	short iheb = 0;

	NearWarning[0] = false;
	NearWarning[1] = false;
	InsufficientAzimuth[0] = false;
	InsufficientAzimuth[1] = false;

	yrn = 1;
	nshabos = 0;
	dayweek = rhday - 1;
	addmon = 0;
	ntotshabos = 0;
	fshabos = fshabos0;
	if (!(optionheb)) iheb = 1;

   	myear = g_yrheb - 3761;

	yl = YearLength( myear );

	sp = 1; //round up
	if (setflag == 1) sp = 2; //round down

	for (i = 1; i <= endyr; i++)
	{

		if (g_mmdate[1][i - 1] > g_mmdate[0][i - 1] )
		{

			if (i > 1 && g_mmdate[0][i - 1] == 1)
			{
				//Rosh Chodesh is on January 1, so increment the year
				//(this happens in the year 2006, 2036)
				yrn++;
			}

			strcpy( &stortim[setflag][i - 1][29][0], "\0" ); //zero time

			k = 0;
			for (j = g_mmdate[0][i - 1]; j <= g_mmdate[1][i - 1]; j++)
			{

				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}
		}
		else if (g_mmdate[1][i - 1] < g_mmdate[0][i - 1])
		{

			strcpy( &stortim[setflag][i - 1][29][0], "\0" ); //zero time


			k = 0;

			for (j = g_mmdate[0][i - 1]; j <= yrend[0]; j++)
			{

				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}


			yrn++;

			for (j = 1; j <= g_mmdate[1][i - 1]; j++)
			{

				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}


		}
	}

	//print out single tables to stdout or to html
	if (endyr == 12)
	{
		heb12monthHTML(setflag);
	}
	else if (endyr == 13)
	{
		heb13monthHTML(setflag);
	}


	return 0;


}


////////////////////EngCalDate/////////////////////
char *EngCalDate(short ndy, short EngYear, short yl, char caldate[] )
//////////////////////////////////////////////
//generates English calendar date for any daynumber, ndy, English year: EngYear
//equivalent to routine that used to be called by netzski6: T1750
//usage:
/*
	short ndy = 255;
	short EngYear = 2007;
	short yl = 365;
	char caldate[12] = "";
	EngCalDate(ndy, EngYear, yl, caldate );
*/
//////////////////////////////////////////////////////////////////////
{
	short nleap = 0;

		//convert dy at particular yr to calendar date
      nleap = 0;

      if (yl == 366) nleap = 1;

      if (ndy <= 31)
		{
			sprintf( caldate, "%s%02d-%4d%s", "Jan-", ndy,  EngYear, "\0" );
		}
        else if ((yl == 365) && (ndy > 31) && (ndy <= 59))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Feb-", ndy - 31,  EngYear, "\0" );
            nleap = 0;
		}
        else if ((yl == 366) && (ndy > 31) && (ndy <= 60))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Feb-", ndy - 31,  EngYear, "\0" );
            nleap = 1;
		}
        else if ((ndy >  59+nleap) && (ndy <=  90+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Mar-", ndy - 59 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  90+nleap) && (ndy <=  120+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Apr-", ndy - 90 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  120+nleap) && (ndy <=  151+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "May-", ndy - 120 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  151+nleap) && (ndy <=  181+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Jun-", ndy - 151 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  181 + nleap) && (ndy <=  212+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Jul-", ndy - 181 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  212+nleap) && (ndy  <=  243+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Aug-", ndy - 212 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  243+nleap) && (ndy  <=  273+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Sep-", ndy - 243 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  273+nleap) && (ndy  <=  304+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Oct-", ndy - 273 - nleap,  EngYear, "\0" );
		}
        else if ((ndy >  304+nleap) && (ndy  <=  334+nleap))
		{
            sprintf( caldate, "%s%02d-%4d%s", "Nov-", ndy - 304 - nleap,  EngYear, "\0" );
		}
        else if (ndy  >  334 + nleap)
		{
            sprintf( caldate, "%s%02d-%4d%s", "Dec-", ndy - 334 - nleap,  EngYear, "\0" );
        }

        return caldate;

}

////////////////////PST_iterate////////////////////////////////////////////
void PST_iterate( short *k, short *j, short *i, short *ntotshabos, short *dayweek
				 , short *yrn, short *fshabos, short *nshabos, short *addmon
				 , short setflag, short sp, short tbltypes, short yl, short iheb )
//////////////////////////////////////////////////////////////////////////
//performs iteration operations for PrinttoSingleFile
{
	char t3subb[12] = "12:12:12 PM";

	*k += 1;

	round( tims[*yrn - 1][tbltypes][*j - 1], steps[setflag], accur[setflag], sp, t3subb );

	if (tbltypes == 0 || tbltypes == 1) //check for near obstructions
	{
		if ( dobs[*yrn - 1][setflag][*j - 1] < distlim ) //add obstruction flag to string
		{
			strcat( t3subb, "*");
			NearWarning[setflag] = true;
		}
		
		//check for insufficient azimuth range
		if (InStr(t3subb, "?"))
		{
            InsufficientAzimuth[setflag] = true;                
        }
	}

	if (*fshabos + *nshabos * 7 == *j) //this is shabbos
	{
		*nshabos += 1; //<<<2--->>>
		*ntotshabos += 1;

		//add marker that it is Shabbos
		strcat( t3subb, "_");

		//check for end of year
		if (*fshabos + *nshabos * 7 > yrend[0])
		{
			*fshabos = 7 - (yrend[0] - (*fshabos + (*nshabos - 1) * 7));
			*nshabos = 0;
		}
	}

	//now actually store the time string
	strcpy( &stortim[setflag][*i - 1][*k - 1][0], t3subb );

}



///////////////////////function round///////////////////////////
char *round( double tim, short stepps, short accurr, short sp, char t3subb[] )
////////////////////////////////////////////////////////////////
//converts fractional hour into time string
//rounds to requested step size and adds cushions
////////////////////////////////////////////////////////////////
//stepps = round to the nearest stepps second
//accurr = add/subtract this cushion (seconds)
/////////////////////////////////////////////////////////////
//sp = 1 forces rounding up
//sp = 2 or -1 forces rounding down
////////////////////////////////////////////////////////////////
//usage:
/*
	double tim;
	char t3subb[12] = "";
	short setflag = 1, sp = 1;
	tim = 17.323;

	round( tim, setflag, sp, t3subb );
*/
/////////////////////////////////////////////////////////////////
{
	short thr = 0, tmin = 0, tsec = 0, secmin = 0, hrhr = 0;
	short minad = 0, secmins = 0, minmin = 0, hrad = 0;
	double t1, ssec;
	char ch[3] = "";

	if ( fabs(tim) == 99999 )
	{
        sprintf( t3subb, "%s", "?" ); //can't calculate visible time due to insufficient azimuth range in profile 
	    return t3subb;
	}
	else if ( tim == -9999 )
	{
        sprintf( t3subb, "%s", "none" ); //no applicable time string found
	    return t3subb;
    }
	else if ( tim == 0.0 )
	{
		sprintf( t3subb, "%s", "-----" ); //candle lighting not applicable
		return t3subb;
	}


	thr = (short)Fix(tim);
	tmin = (short)Fix((tim - thr) * 60.0 );
	t1 = (double)tmin/60.0;
	tsec = (short)((tim - thr - t1) * 3600.0 + 0.5) + accurr; //accur[setflag];

	if (tsec >= 60)
	{
		tsec = tsec - 60;
		tmin = tmin + 1;
	}
	else if (tsec < 0 )
	{
		tsec = 60 + tsec;
		tmin = tmin - 1;
	}

	if (tmin >= 60 )
	{
		tmin = 60 - tmin;
		thr = thr + 1;
	}
	else if (tmin < 0 )
	{
		tmin = 60 + tmin;
		thr = thr - 1;
	}

	if (thr >= 24 ) thr = thr - 24;
	if (thr < 0 ) thr = 24 + thr;
	if (!twentyfourhourclock && thr > 12) thr = thr - 12;

	/////////////seconds///////////////////////////////

	//round up or down
	secmin = tsec;

	minad = 0;
	//if (secmin % steps[setflag] == 0) //no need for rounding
	if (secmin % (int)stepps == 0) //no need for rounding
	{
		goto rnd50;
	}

	if (sp == 1) //round up
	{
		ssec = (double)secmin / stepps; //(double)steps[setflag];
		secmins = Cint(Fix(ssec * 10) * 0.1 + 0.499999);
		if (ssec - (short)Fix(ssec) + 0.000001 < 0.1)
		{
			secmins++;
		}
		secmin = stepps * secmins; //steps[setflag] * secmins;
		if (secmin == 60)
		{
			secmin = 0;
			minad = 1;
		}
	}
	else if (sp == 2 || sp == -1) //round down
	{
		//secmin = steps[setflag] * (short)(Fix((secmin / steps[setflag]) * 10) / 10);
		secmin = stepps * (short)(Fix((secmin / stepps) * 10) / 10);
	}

rnd50:

	///////////////////minutes/////////////////////////
	minmin = tmin + minad;
	if (minmin == 60)
	{
		minmin = 0;
		hrad = 1;
	}

	////////////////////hours/////////////////////////////
	hrhr = thr + hrad;
	if (hrhr >= 24) hrhr = hrhr - 24;
    if (!twentyfourhourclock && hrhr > 12 ) hrhr = hrhr - 12; //don't use railroad time

	////////////////write time string//////////////////////

	if (stepps != 60.0)
	{
		sprintf( t3subb, "%d:%002d:%002d%s", hrhr, minmin, secmin, "\0" );
	}
	else if (stepps == 60.0) //only print the "00" seconds
	{
		sprintf( t3subb, "%d:%002d%s", hrhr, minmin, "\0" );
	}


	return t3subb;

}
//////////emulation of MS VB 6.0  and MS VC++ 6.0 functions//////////////////////

 short iswhite(char c)
 {
   //search for non extended ascii symbols, i.e., codes: 0-20, or >= 255
   if (c == ' '  || c == '\t' ||
          c == '\v' || c == '\f')
   {
	   return -1;
   }
   else
   {
	   return 0;
   }
 }

 short find_first_nonwhite_character(char *str, short start_pos)
 {
   short lpos, lenstr;
   lenstr = strlen(str);
   for (lpos=start_pos; lpos < lenstr; lpos++)
   {

     if (iswhite(str[start_pos + lpos]) != -1)
       break;
   }
   return lpos;
 }

 short find_last_nonwhite_character(char *str, short start_pos)
 {
   short rpos;
   if (start_pos == 0 )
   {
	   start_pos = strlen(str) - 1;
   }

   for (rpos=start_pos; rpos != 0; rpos--)
   {
     if (iswhite(str[rpos]) != -1)
       break;
   }
   return (rpos+1);
 }

 ///////////////emulation of VB6 Trim functions///////////////////////
 //LTrim
 char *LTrim(char * str)
 {
	 short itmp;
	 itmp = find_first_nonwhite_character(str, 0);
	 if (itmp != 0) strcpy( str, &str[itmp] );
	 return str;
 }
 //RTrim
char *RTrim(char * str)
 {
	 short itmp, lenchar;
	 itmp = find_last_nonwhite_character(str, 0);
	 lenchar = strlen( str );
	 if (itmp != lenchar) str[itmp] = 0;
	 return str;
 }
//Trim
char *Trim(char *str)
{
	return LTrim(RTrim(str));
}


//////////////////emulation of Mid function (with external buffer///////////////////////
char *Mid(char *str, short nstart, short nsize, char buff[])
//used for extracting substrings without changing original string
/////////////////////////////////////////////////////////////////////////////////
{
	//works like VB, i.e., substring begins at position 1 and NOT at 0

	//do checks of nstart, nsize
	short lenstr = strlen(str);
	if (nstart > lenstr )  //nstart larger than string size, so don't do anything
	{
		strcpy( buff, str);
		return buff;
	}
    if ( nstart + nsize > lenstr ) nsize = lenstr - nstart + 1; //nstart too big, so reduce it

	strncpy( buff, str + nstart - 1, nsize );
	sprintf( buff + nsize, "%s", "\0" );
	//buff1[nsize] = 0;
	//strcpy( buff, buff1);
	return buff;
}

//////////////////emulation of Mid function with 3 parameters///////////////////////
char *Mid(char *str, short nstart, short nsize)
//used for changing the original string -- i.e., used for truncating string
////////////////////////////////////////////////////////////////////////////////////
{
	//works like VB, i.e., substring begins at nstart = 1 and NOT at 0

	//first do checks of nstart, nsize
	short lenstr = strlen(str);
	if (nstart > lenstr ) return str;  //nstart larger than string size, so don't do anything
   if ( nstart + nsize - 1 > lenstr ) nsize = lenstr - nstart; //nstart too big, so reduce it

	for (short i = 0; i < nsize; i++ )
	{
		str[i] = str[nstart + i - 1];
	}
	str[nsize] = 0;
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
/*
	//static char l_buff[255] = "";
	short lenstr1 = 0;
	short lenstr2 = 0;
	char *pos;
	lenstr1 = strlen(str);
	lenstr2 = strlen(str2);
	if (lenstr2 > lenstr1) return 0; //subdtring bigger than string
	strncpy( buff1, str + nstart - 1, lenstr1 - nstart + 1);
	buff1[lenstr1 - nstart + 1] = 0;
	pos = strstr( buff1, str2) ;
    if (pos == NULL )
	{
		return 0; //substring not found
	}
	else //substring found
	{
		return pos - buff1 + nstart; //like VB position
	}
*/
}

//////////////////emulation of InStr(2 parameters) /////////////////////
int InStr( char * str, char * str2 )
{
	//works like VB, i.e., substrings begin at position 1 and NOT at 0
	return InStr( 1, str, str2 );
/*
	short lenstr1 = 0;
	short lenstr2 = 0;
	char *pos;
	lenstr1 = strlen(str);
	lenstr2 = strlen(str2);
	if (lenstr2 > lenstr1) return 0; //subdtring bigger than string
	strncpy( buff1, str, lenstr1 );
	buff1[lenstr1] = 0;
	pos = strstr( buff1, str2) ;
    if (pos == NULL )
	{
		return 0; //substring not found
	}
	else //substring found
	{
		return pos - buff1 + 1; //like VB position
	}
*/
}

/////////////////emulation of Replace function////////////////////////////
char *Replace( char *str, char Old, char New)
///////////////////////////////////////////////////////////////////////////
//replaces any character "Old" with charcater "New" in the string "str"
//returns the new "str"
//Example: remove "_" from string "random_place"
//
//strcpy( filroot, "random_place" );
//Replace( filroot, '_', ' '); <--need to use ' ', and not " "
//now use Trim(filroot)
//returns "random place"
///////////////////////////////////////////////////////////////////////////
{
	int strLen = strlen(str);

	for(int i=0;i<strLen;i++)
	{
		if(str[i] == Old)
		{
			str[i] = New;
		}
	}

	return str;
}


////////////////emulation of LCase function/////////////////
char *LCase(char *str )//, char buff[])
//example: LCase(citnam, buff)

{
	return strlwr( str) ;
/*
   strcpy( buff, str );

   for (short i=0; i<strlen(buff); i++)
   {
      if (buff[i] >= 0x41 && buff[i] <= 0x5A)
      buff[i] = buff[i] + 0x20;
  }
  return buff;
*/
}

////////////UCase function/////////////////////
char *UCase( char *str )
{
	return strupr( str );
}

/*************  function FIX *********************
This emulates the Visual Basic function Fix() which returns
the integer part of a decimal number, x, as following:
        If x < 0 then FIX(x) = ceil(x)
        If x >= 0 then FIXx) = floor(x)
***************************************************/
double Fix( double x )
{
	if ( x < 0 ) return ceil(x);
	else return floor(x);
}

/*************** function Cint ***************
This emulates the Visual Basic function Cint() which
converts a decimal number into the nearest integer
****************************************************/
short Cint( double x )
{
	double top, bottom;
	top = ceil(x);
	bottom = floor(x);
	if (fabs(top - x) < fabs(bottom - x) )
	{
		return (short)top;
	}
	else
	{
		return (short)bottom;
	}

}


/*****************function Sgn**********************
emulates VB's "Sgn" function (the sign of a real number)
*****************************************************/
double Sgn(double a)
{
	if ( a>0 )
	{
	   return 1.;
	}
	else if (a<0)
	{
	   return -1.;
	}
	else
	{
	   return 0.;
	}
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
/////////////////itoa///////////////////////////
char *itoa ( int value, char *str, int base )
////////////////////////////////////////////////

//replaces MS nonstandard C++ function itoa
{
	if (base == 10 )
	{
	    sprintf(str,"%d",value); // converts to decimal base.
		return str;
    }
	else if (base == 16 )
	{
		sprintf(str,"%x",value); //converts to hexadecimal base.
		return str;
	}
	else if (base == 8 )
	{
		sprintf(str,"%o",value); // converts to octal base.
		return str;
	}

	return " ";

}

///////////////////strlwr//////////////////////
char *strlwr(char *str)
///////////////////////////////////////////////
//emulates MS VC++ strlwr
//////////////////////////////////////////////
{
 char *p;

 for(p=str; *p!='\0'; p++)
  *p=tolower(*p);
 return(str);
}

///////////////strupr//////////////////////
char *strupr(char *str)
//////////////////////////////////////////
//emulates MS VC++ strupr
//////////////////////////////////////////
{
 char *p;

 for(p=str; *p!='\0'; p++)
  *p=toupper(*p);
 return(str);
}

//////////////////LoadTitlesSponsors///////////////////
short LoadTitlesSponsors()
////////////////////////////////////////////////////////
//loads titles of tables and sponsor information
//////////////////////////////////////////////////////
{

	FILE *stream;

	sprintf( myfile, "%s%s%s", drivweb, "SponsorLogo.txt", "\0" );

    	if ( ( stream = fopen( myfile, "r")) )
	{
		while( !feof( stream ) )
		{
			doclin[0] = '\0'; //clear the string buffer

			fgets_CR(doclin, 255, stream); //read in line of text

			if ((strstr( doclin,"[English]")) &&  !optionheb)
			{
				//english sponsor info
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[1][0], "%s%s", doclin, "\0" ); //English memorial
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[3][0], "%s%s", doclin, "\0" ); //English gif
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[6][0], "%s%s", doclin, "\0" ); //Sponsor's internet address
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[7][0], "%s%s", doclin, "\0" ); //Sponsor's name
			}

			if ((strstr( doclin,"[titles]")) && !optionheb)
			{
				//english tables' title info
				if (!eros && !astronplace && !ast) //Eretz Yisroel tables
				{
					fgets_CR(doclin, 255, stream); //read in line of text
					strcpy( TitleLine, doclin );
				}
				else //world tables or astronomical calculations
				{
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
					strcpy( TitleLine, doclin );
				}
			}

			if ((strstr( doclin,"[Hebrew]")) && optionheb)
			{
				//hebrew sponsor info
				fgets_CR(doclin, 255, stream); //read in line of text
				strcpy( &SponsorLine[0][0], doclin ); //Hebrew memorial
				fgets_CR(doclin, 255, stream); //read in line of text
				strcpy( &SponsorLine[2][0], doclin ); //Hebrew gif
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[4][0], "%s%s", doclin, "\0" ); //Sponsor's internet address
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &SponsorLine[5][0], "%s%s", doclin, "\0" ); //Sponsor's name
			}

			if ((strstr( doclin,"[titles]")) && optionheb)
			{
				//hebrew tables' title info
				if (!eros && !astronplace && !ast) //Eretz Yisroel tables
				{
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
 				}
				else //tables for the world or astronomical tables
				{
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
					fgets_CR(doclin, 255, stream); //read in line of text
				}
					strcpy( TitleLine, doclin );

				strncpy(buff, &heb3[1][0], 1);
				buff[1] = 0;

				sprintf( &heb1[3][0],"%s%s%s%s%s%s",&heb1[1][0],"\"",TitleLine,"\""," ",buff);

				strncpy( &heb1[4][0], TitleLine, strlen(TitleLine) );

				strncpy(buff, &heb3[12][9], 16);
				buff[16] = 0;

				sprintf( &heb3[12][0],"%s%s%s%s%s%s",&heb1[1][0],"\"",TitleLine,"\""," ",buff);

			}

			if ( strstr( doclin,"[domain]" ) )
			{
               //domain name
				fgets_CR(doclin, 255, stream); //read in line of text
				sprintf( &DomainName[0], "%s%s", doclin, "\0" ); //domain name
            }

		}
		fclose(stream);
	}
	else //sponsor file not found, use defaults but write message to log file
	{
		if (abortprog) return -1; //abortprog, else write to log file

		WriteUserLog(0, 0, 0, "Can't find or open SponsorLogo.txt file in routine LoadTitleSponsors" );

		if (optionheb)
		{
			if (eros)
			{
				strcpy( TitleLine, &heb1[2][0] );  //Chai in Hebrew
				strcat( TitleLine, "\0");
			}
			else
			{
				strcpy( TitleLine, &heb1[4][0] );  //Bikurei Ysoef in Hebrew
				strcat( TitleLine, "\0");
			}
			//memory of Seidee in Hebrew Unicode
			strcpy( &SponsorLine[1][0], "&#1500;&#1494;&#1499;&#1512;&#1493;&#1503;&#32;&#1506;&#1493;&#1500;&#1501;&#32;&#1493;&#1500;&#1506;&#34;&#1504;&#32;&#1488;&#1489;&#1512;&#1492;&#1501;&#32;&#1497;&#1510;&#1495;&#1511;&#32;&#1489;&#1503;&#32;&#1510;&#1489;&#1497;&#32;&#1494;&#34;&#1500;" );
			strcpy( &SponsorLine[3][0], "afi001.gif" );
			strcpy( &SponsorLine[4][0], "www.shemayisrael.com" );
			strcpy( &SponsorLine[5][0], "&#1512;&#1513;&#1514;;&nbsp;&quot;&#1513;&#1502;&#1506;&nbsp;&#1497;&#1513;&#1512;&#1488;&#1500;&quot;" );
		}
		else if (! optionheb)
		{
			if (eros)
			{
				strncpy( TitleLine, "Chai", 4);
			}
			else
			{
				strncpy( TitleLine, "Bikurei Yosef", 4);
			}
			strcpy( &SponsorLine[1][0], "In loving memory and l'iluy nishmas of Avrohom Yitzhok ben Zvi z\"l" );
			strcpy( &SponsorLine[3][0], "afi001.gif" );
			strcpy( &SponsorLine[6][0], "www.chaitables.com" );
			strcpy( &SponsorLine[7][0], "The Chai Tables" );
		}

	}

	return 0;

}

//////////////////////////LoadConstants//////////////////////////
short LoadConstants( short zmanyes, short typezman, char filzman[] )
///////////////////loads up constanst//////////////////////
//the weather constants must be changed as Global warming
//1. changes the atmosphere.  In partiuclar, the temperature
//zones will increase in latitude
//2. menat's program constants should be pretty constant barring
//new modeling of the atmosphere.
//3. astronomical constants should be updated before 2050
{

	//static char doclin[255] = "";
	//static char buff[255] = "";
	short ipath = 0, numzmantype = 0;
	FILE *stream;

	//load math constants

	pi = acos( -1 );
	cd = pi / 180.;
	dcmrad = 1 / (cd * 1e3); //converts mrads to degrees
	pi2 = pi * 2.0;
	ch = 12. / pi;
	hr = 60.;

    /////////////////////////////////////////////////////////////
	RefYear = 1996; //reference year for ephemerel calculations
	////////////////////////////////////////////////////////////


//load up paths, refraction, atmoshperic, and astronomical constants

//file must reside in directory that this program runs in, i.e., the cgi-bin directory

	if (stream = fopen( "PathsZones_dtm.txt", "r" ) )
	{
		while (!feof(stream))
		{
			fgets_CR( doclin, 255, stream ); //look for headers, etc

			if (InStr( doclin, "[platform]" ))
			{
				ipath = 1;
			}
			if (InStr( doclin, "[paths]" ))
			{
				ipath = 2;
			}
			else if (InStr( doclin, "[zemanim_eng]" ))
			{
				ipath = 3;
				numzmantype = -1;
			}
			else if (InStr( doclin, "[zemanim_heb]" ))
			{
				ipath = 4;
				numzmantype = -1;
			}
			else if (InStr( doclin, "[world zones]" ))
			{
				ipath = 5;
			}
			else if (InStr( doclin, "[Eretz Yisroel zones]" ))
			{
				ipath = 6;
			}
			else if (InStr( doclin, "[adhoc shift]" ))
			{
				ipath = 7;
			}
			else if (InStr( doclin, "[menatsum/win.ren]" ))
			{
				ipath = 8;
			}
			else if (InStr( doclin, "[astron. constants]" ))
			{
				ipath = 9;
			}
			else if (InStr( doclin, "[time cushions]" ))
			{
				ipath = 10;
			}
			else if (InStr( doclin, "[obstruction limit]" ))
			{
				ipath = 11;
			}
			else if (InStr( doclin, "[vdw ref const]" ))
			{
				ipath = 12;
			}
			else
			{
				if (ipath != 1 && ipath != 2 && ipath !=3 && ipath !=4) ipath = 0; //blank detected
			}

			//read constants
			switch (ipath)
			{
			case 1:

				//////////////platform////////////////////
				fgets_CR( doclin, 255, stream);
				if (InStr(doclin, "Windows")) {
					Windows = true;
				}
				ipath = 0;
				break;

			case 2:

				/////////read paths///////////////////////////
				fgets_CR( doclin, 255, stream);
				strcpy( drivjk, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivweb, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivfordtm, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivcities, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivdtm, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivtas, doclin );

				fgets_CR( doclin, 255, stream);
				strcpy( drivTemp, doclin );

				ipath = 0;
				break;

			case 3:
				//////////////if english zmanmim, read zemanim file name//////////
				if (!optionheb && zmanyes && numzmantype == typezman)
				{
					sprintf( filzman, "%s%s", drivcities, doclin );
					ipath = 0;
					break;
				}
				numzmantype++;
				break;

			case 4:
				//////////////if hebrew zmanmim, read zemanim file name//////////
				if (optionheb && zmanyes && numzmantype == typezman)
				{
					sprintf( filzman, "%s%s", drivcities, doclin );
					ipath = 0;
					break;
				}
				numzmantype++;
				break;

			case 5:

				///////////read temperature zones for the rest of the world///////////////////
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[0], &sumwinzone[0][0], &sumwinzone[0][1], &sumwinzone[0][2], &sumwinzone[0][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[1], &sumwinzone[1][0], &sumwinzone[1][1], &sumwinzone[1][2], &sumwinzone[1][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[2], &sumwinzone[2][0], &sumwinzone[2][1], &sumwinzone[2][2], &sumwinzone[2][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[3], &sumwinzone[3][0], &sumwinzone[3][1], &sumwinzone[3][2], &sumwinzone[3][3] );
				ipath = 0;
				break;

			case 6:

				/////////read temperature zone constants for Eretz Yisroel//////////////////
				fscanf( stream, "%d,%d,%d,%d\n", &sumwinzone[4][0], &sumwinzone[4][1], &sumwinzone[4][2], &sumwinzone[4][3] );
				ipath = 0;
				break;

			case 7:

				//read adhoc fix to times (secs)
				fscanf( stream, "%d\n", &nacurrFix );
				ipath = 0;
				break;

			case 8:

				//read refraction constansts = height vs. refraction fit constants to
				//menatsum.ren and meantwin.rem
				fscanf( stream, "%lg\n", &sumrefo );
				fscanf( stream, "%lg\n", &sumref[0] );
				fscanf( stream, "%lg\n", &sumref[1] );
				fscanf( stream, "%lg\n", &sumref[2] );
				fscanf( stream, "%lg\n", &sumref[3] );
				fscanf( stream, "%lg\n", &sumref[4] );
				fscanf( stream, "%lg\n", &sumref[5] );
				fscanf( stream, "%lg\n", &sumref[6] );
				fscanf( stream, "%lg\n", &sumref[7] );
				fscanf( stream, "%lg\n", &winrefo );
				fscanf( stream, "%lg\n", &winref[0] );
				fscanf( stream, "%lg\n", &winref[1] );
				fscanf( stream, "%lg\n", &winref[2] );
				fscanf( stream, "%lg\n", &winref[3] );
				fscanf( stream, "%lg\n", &winref[4] );
				fscanf( stream, "%lg\n", &winref[5] );
				fscanf( stream, "%lg\n", &winref[6] );
				fscanf( stream, "%lg\n", &winref[7] );
				ipath = 0;
				break;

			case 9:

				//read astronomical constants
				fscanf( stream, "%lg\n", &ac0 );//change in the mean anamolay per day
				ac0 = cd * ac0;

				fscanf( stream, "%lg\n", &mc0 );//change in the mean longitude per day
				mc0 = cd * mc0;

				fscanf( stream, "%lg\n", &ec0 );//1st constant of the Ecliptic longitude
				ec0 = cd * ec0;

				fscanf( stream, "%lg\n", &e2c0 );//2nd constant of the Ecliptic longitude
				e2c0 = cd * e2c0;

				fscanf( stream, "%lg\n", &ob10 );//change in the obliquity of ecliptic per day
				ob10 = cd * ob10;

				fscanf( stream, "%lg\n", &ap0 );//mean anomaly for Jan 0, 00:00 at Greenwich meridian
				ap0 = cd * ap0;

				fscanf( stream, "%lg\n", &mp0 );//mean longitude of sun Jan 0, 00:00 at Greenwich meridian
				mp0 = cd * mp0;

				fscanf( stream, "%lg\n", &ob0 );//ecliptic angle for Jan 0, 00:00 at Greenwich meridian
				ob0 = cd * ob0;

				fscanf( stream, "%d\n", &ValidUntil);//Last year that these constants are accurate to 6 seconds
				ipath = 0;
				break;

			case 10:

				//read cushions for Eretz Yisroel (25 m), GTOPO30 (1 km), SRTM-1 (90 m), SRTM-2 (30 m), astronomical/mishor
				fscanf( stream, "%d,%d,%d,%d,%d\n", &cushion[0], &cushion[1], &cushion[2], &cushion[3], &cushion[4] );
				ipath = 0;
				break;

			case 11:

				//read minimum distance to obstruction not to be considered a near obstruction
				fscanf( stream, "%g,%g,%g,%g\n", &distlims[0], &distlims[1], &distlims[2], &distlims[3] );
				ipath = 0;
				break;

			case 12:

				//read fit paramaters for van der Werf dip and atmospheric refraction as function of height
				fscanf( stream, "%lg\n", &vdweps[0] );
				fscanf( stream, "%lg\n", &vdweps[1] );
				fscanf( stream, "%lg\n", &vdweps[2] );
				fscanf( stream, "%lg\n", &vdweps[3] );
				fscanf( stream, "%lg\n", &vdweps[4] );
				fscanf( stream, "%lg\n", &vdwref[0] );
				fscanf( stream, "%lg\n", &vdwref[1] );
				fscanf( stream, "%lg\n", &vdwref[2] );
				fscanf( stream, "%lg\n", &vdwref[3] );
				fscanf( stream, "%lg\n", &vdwref[4] );
				fscanf( stream, "%lg\n", &vdwref[5] );
				//now refraction constant and temperature scaling constants
				fscanf( stream, "%lg\n", &vdwconst[0] );
				fscanf( stream, "%lg\n", &vdwconst[1] );
				fscanf( stream, "%lg\n", &vdwconst[2] );
				fscanf( stream, "%lg\n", &vdwconst[3] );
				//temperature model based on minimum/average/maximum ground temperaters
				fscanf( stream, "%d\n", &TempModel );
				//maximum length of day to look for winter inversion layers
				fscanf( stream, "%lg\n",&MaxWinLen );
				//fraction of average yearly minimum temperature that it is considered likely that
				//an inversion will form
				fscanf( stream, "%lg\n", &WinterPercent );
				//maximum view angle  for adding adhoc inversion corrections
				fscanf( stream, "%lg\n", &HorizonInv );
				//ground pressure 
				fscanf( stream, "%lg\n", &Pressure0 );
				//global warming flag
				fscanf( stream, "%d\n", &GlobalWarmingCorrection );
				ipath = 0;
				break;
			}
		}
		fclose( stream ); //nothing else to read

	}
	else //use defaults
	{
		if (abortprog) return -1; // abortprog program since "PathsZones.txt" not found

		//else if abortprog == false then use defaults

		strcpy( drivjk, "/home/chkeller/public_html/new.chaitables.com/cgi-bin/" ); //path to hebrew strings
		strcpy( drivweb, "/home/chkeller/public_html/new.chaitables.com/cgi-bin/" ); //path to web files
		strcpy( drivfordtm, "/home/chkeller/public_html/new.chaitables.com/cgi-bin/" ); //path to fordtm (tmp dir)
		strcpy( drivcities, "/home/chkeller/public_html/new.chaitables.com/cities/" ); //path to cities directory
		strcpy( drivdtm, "/home/chkeller/public_html/new.chaitables.com/USA/" ); //path to 30m DTM files and GTOPO30
		strcpy( drivtas, "/home/chkeller/public_html/new.chaitalbes.com/3AS/" ); //path to 90 m DTM files
		strcpy( drivTemp, "/home/chaitables/public_html/WorldClim_bil/" ); //path to WorldClim 2.1 bil format temperature files


		//now write error message to Log File
		WriteUserLog( 0, 0, 0, "Missing PathsZones_dtm.txt in routine LoadConstants" );

		if (zmanyes)
		{
			if (!optionheb)
			{
				switch (typezman)
				{
				case 0:
					strcpy( buff, "MA72min.zma" );
					break;
				case 1:
					strcpy( buff, "MA90min.zma" );
					break;
				case 2:
					strcpy( buff, "Gaonmishor.zma" );
					break;
				case 3:
					strcpy( buff, "Gaonastron.zma" );
					break;
				case 4:
					strcpy( buff, "benishchai.zma" );
					break;
				case 5:
					strcpy( buff, "Chabad.zma" );
					break;
				case 6:
					strcpy( buff, "Harav_Melamid.zma" );
					break;
				case 7:
					strcpy( buff, "Yedidia.zma" );
					break;
				}
			}
			else
			{
				switch (typezman)
				{
				case 0:
					strcpy( buff, "MA72min_heb.zma" );
					break;
				case 1:
					strcpy( buff, "MA90min_heb.zma" );
					break;
				case 2:
					strcpy( buff, "Gaonmishor_heb.zma" );
					break;
				case 3:
					strcpy( buff, "Gaonastron_heb.zma" );
					break;
				case 4:
					strcpy( buff, "benishchai_heb.zma" );
					break;
				case 5:
					strcpy( buff, "Chabad_heb.zma" );
					break;
				case 6:
					strcpy( buff, "Harav_Melamid_heb.zma" );
					break;
				case 7:
					strcpy( buff, "Yedidia_heb.zma" );
					break;
				}
			}

			sprintf( filzman, "%s%s", drivcities, buff );

		}

		latzone[0] = 20;
		sumwinzone[0][0] = 55; //winter refraction ends
		sumwinzone[0][1] = 320; //winter refraction begins
		sumwinzone[0][2] = 0;   //adhoc winter fix ends
		sumwinzone[0][3] = 366; //adhoc winter fix begins
		latzone[1] = 30;
		sumwinzone[1][0] = 85; //ditto.
		sumwinzone[1][1] = 290;
		sumwinzone[1][2] = 30;
		sumwinzone[1][3] = 330;
		latzone[2] = 40;
		sumwinzone[2][0] = 115;
		sumwinzone[2][1] = 260;
		sumwinzone[2][2] = 60;
		sumwinzone[2][3] = 300;
		latzone[3] = 50;
		sumwinzone[3][0] = 145;
		sumwinzone[3][1] = 230;
		sumwinzone[3][2] = 90;
		sumwinzone[3][3] = 270;
		//////zones for Eretz Yisroel/////////////////////
		sumwinzone[4][0] = 85;
		sumwinzone[4][1] = 290;
		sumwinzone[4][2] = 30;
		sumwinzone[4][3] = 330;
		////////////////////////ad hoc time fix (sec)/////////////////////////
		nacurrFix = 15;

	/*  Fits to summer and winter refraction values for heights 1 m to 2000 m */
	/*  (1-999 layers) outputed to files menatsum.ren, menatwin.ren.
	/*	Calculated by c:\jk_c\MENAT.FOR */
		sumrefo = 8.899f;
		sumref[0] = 2.791796282;
		sumref[1] = .5032840405;
		sumref[2] = .001353422287;
		sumref[3] = 7.065245866e-4f;
		sumref[4] = 1.050981251;
		sumref[5] = .4931095603;
		sumref[6] = -.02078600882;
		sumref[7] = -.00315052518;
		winrefo = 9.85f;
		winref[0] = 2.779751597;
		winref[1] = .5040818795;
		winref[2] = .001809029729;
		winref[3] = 7.994475831e-4;
		winref[4] = 1.188723157;
		winref[5] = .4911777019;
		winref[6] = -.0221410531;
		winref[7] = -.003454047139;
		//////////////////////////////////////////////////////////////////

		//////////////van der Werf refraction constants/////////////////
		/* 		vdW dip angle vs height polynomial fit coefficients */
		vdweps[0] = 2.77346593151086f;
		vdweps[1] = .497348466526589f;
		vdweps[2] = .00253874620975453f;
		vdweps[3] = 6.75587054940366e-4f;
		vdweps[4] = 3.94973974451576e-5f;
		/* 		vdW atmospheric refraction vs height polynomial fit coefficients */
		vdwref[0] = 1.16577538442405f;
		vdwref[1] = .468149166683532f;
		vdwref[2] = -.019176833246687f;
		vdwref[3] = -.0048345814464145f;
		vdwref[4] = -4.90660400743218e-4f;
		vdwref[5] = -1.60099622077352e-5f;
		/////////////////////////////////////////////////////////////////////////
		/*		vdw refraction and temperature scaling constants  */
		vdwconst[0] = 9.56267125268496f; //refraction from the horizon to infinity at 288.15 degrees Kelvin
		vdwconst[1] = 1.687; //temperature scaling constant for the terrestrial refraction and for vbwconst[0]
		vdwconst[2] = 2.18; //temperataure scaling constant for the refraction component ref, i.e., from the observer to the horizon
		vdwconst[3] = -0.2; //approx. temperature scanling of the vdw eps
		///////////////////TempModel//////////////////////////////////////////////////
		TempModel = 1; //use 30 arcsec WorldClim 2 minimum ground temperatures model (1=minimum ground temps, 2 = average, 3 = maximum)
		MaxWinLen = 11.1; //maximum length of day to look for winter inversion layers
		WinterPercent = 1.0; //default is set to not use this filtering option, i.e.,accept 100 percent
		HorizonInv = 0.01; //only add adhoc inversion fix if calculated view angle is less than HorizonInv
		Pressure0 = 1013.25; //default ground pressure used for calculations
		GlobalWarmingCorrection = 1; //default - use latitude shift model for global warming

		//////////////////astronomical constants///////////////////
		ac0 = cd * .9856003; //change in the mean anamolay per day
		mc0 = cd * .9856474; //change in the mean longitude per day
		ec0 = cd * 1.915; //constants of the Ecliptic longitude
		e2c0 = cd * .02;
		ob10 = cd * -4e-7; //change in the obliquity of ecliptic per day
	/*  mean anomaly for Jan 0, 00:00 at Greenwich meridian */
		ap0 = cd * 357.528;
	/*  mean longitude of sun Jan 0, 00:00 at Greenwich meridian */
		mp0 = cd * 280.461;
	/*  ecliptic angle for Jan 0, 00:00 at Greenwich meridian */
		ob0 = cd * 23.439;
	    /////////////////////////////////////////////////////////////
		ValidUntil = END_SUN_APPROX; //last civil year constants are accurate to 6 seconds
		////////////////////////////////////////////////////////////

		///////////////cushions////////////////////
		cushion[0] = 15; //EY (25 m)
		cushion[1] = 45; //GTOPO30/SRTM30 (1 km)
		cushion[2] = 35; //SRTM-1 (90 m)
		cushion[3] = 20; //SRTM-2 (30 m)
		cushion[4] = 15; //astronomical calculations

		distlims[0] = 5; //EY (25 m)
		distlims[1] = 30; //GTOPO30/SRTM30 (1 km)
		distlims[2] = 10; //SRTM-1 (90 m)
		distlims[3] = 6;  //SRTM-2 (30 m)

	}

	//English months

	strcpy( &monthh[1][0][0], "Tishrey");
	strcpy( &monthh[1][1][0], "Chesvan");
	strcpy( &monthh[1][2][0], "Kislev");
	strcpy( &monthh[1][3][0], "Teves");
	strcpy( &monthh[1][4][0], "Shvat");
	strcpy( &monthh[1][5][0], "Adar I");
	strcpy( &monthh[1][6][0], "Adar II");
	strcpy( &monthh[1][7][0], "Nisan");
	strcpy( &monthh[1][8][0], "Iyar");
	strcpy( &monthh[1][9][0], "Sivan");
	strcpy( &monthh[1][10][0], "Tamuz");
	strcpy( &monthh[1][11][0], "Av");
	strcpy( &monthh[1][12][0], "Elul");
	strcpy( &monthh[1][13][0], "Adar");

	////////////////////////////////////////////////////////////////////////

	return 0;
}


////////////convert from DOS Hebrew to ISO Hebrew////////////
char *IsoToUnicode( char *str )
/////////////////////////////////////////////////////////////
//converts Windows Hebrew (ISO DOS) strings into
// standard (Unicode) Hebrew that can be
//read by Linux, etc.
/////////////////////////////////////////////////////////////
{
	short strln;
	wchar_t dest[255];

	strln = strlen(str);

	mbstowcs(dest, str, strln ); //convert ISO to Unicode

	strcpy( buff, (const char *)dest );

	return buff;
/*
	unsigned char j = 0;

	for(int i=0;i<strLen;i++)
	{
		if( ( (unsigned char)buff[i] >= (unsigned char)220 ) && ( (unsigned char)buff[i] < (unsigned char)252  ) )
		{
			buff[i] += 96;//'?'-128
		}
		else
                {
			j = (unsigned char)buff[i];
                }
	}
	return buff;
*/
}



/////////////////////netzski6////////////////////////////////////////
short netzski6(char *fileon, double lg, double lt, double hgt, double aprn,
			               float geotd, short nsetflag, short nfil,
						   short *nweather, double *meantemp, double *p,
     		               short *ns1, short *ns2, short *ns3, short *ns4, double *WinTemp,
						   short mint[], short avgt[], short maxt[], short ExtTemp[] )
//////////////////////////////////////////////////////////////////////
/* netzski7.f -- translated by f2c (version 20021022).
   You must link the resulting object file with the libraries:
	-lf2c -lm   (in that order)
*/
//this is the subroutine version of netzski6
//outputs times for inputed place
///////////////////////////////////////////////////////////////////////
{

/************ System generated locals ******************/
    short i__1, i__2, i__3;
    double d__1;


/************* Local variables ******************/

    bool adhocset;
    //short nweather;
    double elevintx, elevinty, d__, t, z__;
	//double p;
    short nplaceflg;
    bool adhocrise;
    double a1, t3, t6;
    double ac, mc, ec, e2c, ap, mp, ob;

    short ne;
    double td, tf, es, lr, dy, ms;
    short nn;
    double pt, al1, dy1; 
    //short ns1, ns2, ns3, ns4;
    double sr1;
    short nac;
    double aas;
    double air, ref;
    double azi, eps; 
    short nyd, nyr, nyf;
    double al1o, alt1, azi1, azi2; 

    double hran;
    double hgto;
    double t3sub;

    double dobss;;
    short nyear;
    double dobsy;
    short niter, nloop;
    bool astro;

    short ndirec, nabove;
    short maxang; //maximum angular range
    short nlpend, nfound;
    short nloops, ndystp;
    double refrac1;
    short nyrstp;
    double stepsec, elevint;
    short nacurr;
	double jdn, hrs, ra;
	short mday, mon;
	int nyl;
	short ier = 0;
	double maxdecl;

    //short itk;
    //double meantemp;

	//other parameters new for Version 19
	//double WinTemp = 0;
	short DayCount = 0;
	double DayLength = 0.0;
	bool FindWinter = true; //version 18 way of determining the winter months
	bool EnableSunriseInv = false; //flag to subtract time for suspected inversion at this day
	bool EnableSunsetInv = false; //flag to subtract time for suspected inversion at this day

	double TRfudge = 1.0; //0.95;//increasing positive lapse rate will lower the TR as well as the total atm refraction
	double TRfudgeVis = 1.0;

	//vdw temperature scaling factors
	double vdwsf = 1.0, vdwrefr = 1.0, vbweps = 1.0;
	double trbasis, d__3;
	double hgtdifx = 0., hgtdif = 0., pathlength = 0., exponent = 0., dist = 0., distx = 0.;

	short PartOfDay = 0; //where is zman in the diurnal cycle flag -- deter__mins what temperature is used to calculate the refraction for nweather = 5
						   //= 1 for using minimum @orldClim temperatures
						   //  2 for using average WorldClim temperatures
						   //  3 for using maximum WorldClim temperatures

/**************************************************************/

/*       Version 19 (of netzski6) */

/*       Subroutine version of netzski6.  It is called by main program to */
/*       calculate a maximum of 7 zemanim for each place. */
/*       I.e., the mishor sunrise/sunset, the astronomical sunrise/sunset, */
/*       the visible sunrise/sunset, and chazos. */
/*       The results of the calculations are expressed as an array of */
/*       7 numbers vs day iteration corresponding to the days of the Hebrew year. */

/* ----------------------------------------------------- */
/*       = 20 * maxang + 1 ;number of azimuth points separated by .1 degrees */
/* -------------------------------------------------------------------------------- */
/*       Version 15-16 parameters */
/*       tims contains output zemanim as follows */
/*       tims(0,..) - first secular year of Hebrew Year */
/*       tims(1,..) - second secular year of Hebrew Year */
/*       tims(.,0,.) - visible sunrise */
/*       tims(.,1,.) - visible sunset */
/*       tims(.,2,.) - astronomical sunrise */
/*       tims(.,3,.) - astronomical sunset */
/*       tims(.,4,.) - mishor sunrise */
/*       tims(.,5,.) - mishor sunset */
/*       tims(.,6,.) - chazos (t6) */
/*       azils, dobss contain azimuths and distances to obstructions */
/*       for the visible sunrise and sunset as follows */
/*       azils(0,..), dobss(0,..) - first secular year of Hebrew Year */
/*       azils(1,..), dobss(1,..) - second secular year of Hebrew Year */
/*       azils(.,0,.), dobss(.,0,.) - azimuths, dist. of vis. sunrise */
/*       azils(.,1,.), dobss(.,1,.) - azimuths, dist. of vis. sunset */
/* -------------------------------parameters--------------------------------------------------- */
/*       FirstSecularYr                    'beginning year to calculate */
/*       SecondSecularYr                     'last year to calculate */
/* --------------------------<<<<USER INPUT>>>---------------------------------- */
/* Version 19: Now uses VDW raytracing if VDW_REF = true
/*			   for that method reads the WorldClim temperatures
/*			   For EY, ITM coordinates were already converted to WGS84 lt, lg
/* ---------------------------------------------------------------------------------- */ 
    nplaceflg = 1;
/*                                  '=0 calculates earliest sunrise/latest sunset for selected places in Jerus
alem */
/*                                  '=1 calculation for a specific place--must input place number */
/*                                  '=2 for calculations for LA */
//
/* nsetflag                          = 0 for astr sunrise/visible sunrise/chazos */
/*                                   = 1 for astr sunset/visible sunset/chazos */
/*                                   = 2 for mishor sunrise if hgt = 0/chazos or astro. sunrise/chazos */
/*                                   = 3 for mishor sunset if hgt = 0/chazos or astro. sunset/chazos */
/*                                   = 4 for mishor sunrise/vis. sunrise/chazos */
/*                                   = 5 for mishor sunset/vis. sunset/chazos */
/*                                   = 6 for visible sunrise calculating profile using 30m DTM */
/*                                   = 7 for mishor sunset calculating profile using 30m DTM */
//
//    nweather = 3;
/*                                  '=0 for summer refraction */
/*                                  '=1 for winter refraction */
/*                                  '=2 for ASTRONOMICAL ALMANAC's refraction */
/*                                  '=3 ad hoc summer length as function of latitude
/*                                  '   ad hoc winter length as function of latitude
    ndtm = 5;
    nmenat = 1;
/*                                  '=0 uses tp */
/*                                  '=1 calculates sunrises over Harei Moav using theo. refrac. */
    niter = 1;
/*                                  '=0 doesn't iterate the sunrise times */
/*                                  '=1 iterates the sunrise times */
//    p = Pressure0; //1013.25f;
/*                                  'mean atmospheric pressure (mb) */
    t = 27.;
/*                                  'mean atmospheric temperature (degrees Celsius) */
    stepsec = 2.5f;
/*                                  'step size in seconds */
/*       'N.B. astronomical constants are accurate to 10 seconds only to the year 2050 */
/*       '(Astronomical Almanac, 1996, p. C-24) */
/*       Addhoc fixes: (15 second fixes sunrise based on netz observations at Neveh Yaakov) */
    adhocrise = true;
    adhocset = false;
/* -------------------------------------------------------------------------------------- */

/*    geo = false;

/*       read in elevations from Profile file */
/*       which is a list of the view angle (corrected for refraction) */
/*       as a function of azimuth for the particular observation */
/*       point at kmxo,kmyo,hgt */

	/*
	//----Old Method: ad hoc summer and winter ranges for different latitudes----- 
	//----- VDW_REF = true, read WorldClim temperatures and stuff into min,avg, max temperature arrays
	//as well as calculationg the WinTemp used for finding probable inversion layer season within winter months
	//also calculate the longitude and latitude if not inputed

	InitWeatherRegions( &lt, &lg, &nweather,
     		            &ns1, &ns2, &ns3, &ns4, &WinTemp,
						mint, avgt, maxt, ExtTemp);

	//for new calcualtion method, deter__min the ground temperatures, minimum, average, and yearly mean
	if (VDW_REF) {
		//read WorldClim bil files to deter__min the minimum and average temperatures of the year

		if (nsetflag == 0) {
			//sunrise 
			//now find average temperature to use in removing profile's TR 
		    meantemp = 0.f;
		    for (itk = 1; itk <= 12; ++itk) {
				meantemp += avgt[itk - 1];
		    }
		    meantemp = meantemp / 12.f + 273.15f;
		} else if (nsetflag == 1) {
			//sunset 
			//now find average temperature to use in removing profile's TR /
		    meantemp = 0.f;
		    for (itk = 1; itk <= 12; ++itk) {
				meantemp += avgt[itk - 1];
		    }
		    meantemp = meantemp / 12.f + 273.15f;
		}
	}
	*/

    astro = false;

	if (nsetflag != 2 && nsetflag != 3 && nsetflag != 6 && nsetflag != 7 && SRTMflag < 10) //read horizon profile file
	{
		if (ReadProfile( fileon, &hgt, p, &maxang, meantemp ) ) return -2; //(-2 = file doesn't exist, or can't be read)

	}
	else if ( SRTMflag > 9) //visible sunrise/sunset based on 30m DTM calculations
	{
		//N.b., if this is Golden Hour calculation, then first assume same temperatures as observer, calculate the rpfile
		//find the coordinates of the first golden light, then go back and use this point to define the ground temperature, etc.

		//deter__min maximum angular range from latitude
		maxdecl = (66.5 - fabs(lt)) + 23.5; //approximate formula for maximum declination as function of latitude (kmyo)
		//source: chart of declination range vs latitude from http://en.wikipedia.org/wiki/Declination
		maxang = (int)(fabs(asin(cos(maxdecl * cd)))/cd); //approximate formula for half maximum angle range as function of declination
		//source: http://en.wikipedia.org/wiki/Solar_azimuth_angle
		//add another few degrees on either side for a cushion for latitudes greater than 60
		if (fabs(lt) >= 60 && fabs(lt) < 62 )
		{
			maxang = maxang + 3;
		}
		else if (fabs(lt) >= 62)
		{
			maxang = maxang + 10;
		}

		if (maxang < 30) maxang = 30.; //set minimum maxang

		//maxang = MAX_ANGULAR_RANGE;
        ier = readDTM( &lg, &lt, &hgt, &maxang, &aprn, p, meantemp, SRTMflag );
		if (ier) //error detected
		{
			return ier;
		}


    }
	else //astronomical/mishor times
	{
		astro = true; //no DTM file to read
	}

	//---------------------------------------
	td = geotd;
	// time difference from Greenwich England
	if (nplaceflg == 2)
	{
		td = -8.; //Los Angeles
	}
	//---------------------------------------


    lr = cd * lt;
    tf = td * 60. + lg * 4.;


// ----------------begin astronomical calculations---------------------------

 ////////////define loop parameters for visible times///////////////////////
	if (nsetflag % 2 == 0)
	{
//   sunrises
		nloop = 72;
//    start search 2 minutes above astronomical horizon
		nlpend = 2000;
		PartOfDay = 1;  //refraction scaled according on minimum temperatures
	}
    else if (nsetflag % 2 != 0)
	{
//         sunsets
		nloop = 370;
		if (nplaceflg == 0) nloop = 75;
		nlpend = -10;
		PartOfDay = 2;  //refraction scaled according to average temperatures
	}
////////////////////////////////////////////////////////////////////////////

	i__1 = SecondSecularYr;
    for (nyear = FirstSecularYr; nyear <= i__1; ++nyear)
	{
	/* 	     define year iteration number */
		nyrstp = nyear - FirstSecularYr + 1;
		nyr = nyear;
	/*       year for calculation */
		nyd = nyr - RefYear;

	if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX )
	{
		/* ----initialize astronomical constants to time zone and year--------------------------- */
		   InitAstConst(&nyr, &nyd, &ac, &mc, &ec, &e2c, &ap, &mp, &ob, &nyf, &td);
		/*--------------------------------------------------------------------------------------- */
	}

    //loop over civil years comprising Hebrew year
		d__1 = (double) yrend[nyear - FirstSecularYr];
		//-----------------------------------------------------
		for ( dy = (double) yrstrt[nyear - FirstSecularYr]; dy <= d__1; dy++)
		{
			//day iteration
			ndystp = (int)dy;
			dy1 = dy;
			///////////////////noon/////////////////////////////////////////////////////
			AstroNoon( &jdn, &td, &dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__,
				       &tf, &nyear, &mday, &mon, &nyl, &t6);
			////////////////////////////////////////////////////////////////////////////

			if (nfil == 0)
			{
				tims[nyrstp - 1][6][ndystp - 1] = t6/hr;
			}
			else //record latest noon
			{
				//tims[nyrstp - 1][6][ndystp - 1] = (tims[nyrstp - 1][6][ndystp - 1] * nfil + t6/hr) / (nfil + 1);
				if (t6/hr > tims[nyrstp - 1][6][ndystp - 1] )
						tims[nyrstp - 1][6][ndystp - 1] = t6/hr;

			}
			////////////////////////////////////////////////

	/*          astronomical refraction calculations */
			ref = 0.;
			eps = 0.;
			hgto = 0.;
	L688:
			if (nsetflag >= 4 && hgto == 0.) {
	/*             first calculate the mishor astro. times */
			hgto = hgt;
			hgt = 0.;
			} else if (nsetflag >= 4 && hgto != 0.) {
	/*             going back to calculate the visible times */
			hgt = hgto;
			hgto = 0.;
			}

			///////////Astronomical total refraction/////////////////////////////////////////////////////////
			AstronRefract( &hgt, &lt, nweather, &PartOfDay, &dy, ns1, ns2, &air,
				&a1, mint, avgt, maxt, &vdwsf, &vdwrefr, &vbweps, &nyr);
            ////////////////////////////////////////////////////////////////////////////////////////////////

			//////////////////////////conserved portion of atmospheric refraction//////////////////////////
			if (VDW_REF) 
			{
				//assuming that the terrestrial refraction scales like the atmospheric refraction
				//i.e., with the VDW scaling factor vdwsf, this is unproven but should be at least a close approx.
				trbasis = vdwsf * *p * 8.15f * 1e3f * .0277f / (288.15 * 288.15 * 3600);
			}
			//////////////////////////////////////////////////////////////////////////////////////////////


			if (nsetflag % 2 != 0)
			{
				//calculate astron./mishor sunsets
				//Nsetflag = 1 or Nsetflag = 3 OR Nsetflag = 5 */
				///////////Astronomical/Mishor sunset///////////////////////////////////////////////////////////
            AstronSunset(&ndirec, &dy, &lr, &lt, &t6, &mp, &mc, &ap, &ac, &ms, &aas, &es,
						&ob, &d__, &air, &sr1, nweather, ns3, ns4, &adhocset, &t3,
						&astro, &azi1, &jdn, &mday, &mon, &nyear, &nyl, &td,
						mint, WinTemp, &FindWinter, &EnableSunsetInv);
				///////////////////////////////////////////////////////////////////////////////////////////////

		//write astronomical sunset times to file and loop in days
		if (astro || nsetflag == 5) //astro == true is same as nsetflag == 3
		{
 			if ( !ast )
			{
				//////////////record mishor sunset /////////////
				nacurr = 0;
				if (EnableSunsetInv) {
					//add adhoc inversion fix to mishor/ast sunset
					nacurr = 15;
					EnableSunsetInv = false; //reset flag for next day
				}

				if (nfil == 0)
				{
					tims[nyrstp - 1][5][ndystp - 1] = t3;
				}
				else
				{
					if (t3 > tims[nyrstp - 1][5][ndystp - 1] )
						tims[nyrstp - 1][5][ndystp - 1] = t3;
				}
				////////////////////////////////////////////////
			}
			else
			{
				//////////////record astro sunset /////////////
				nacurr = 0;
				if (EnableSunsetInv) {
					//add adhoc inversion fix to mishor/ast sunset
					nacurr = 15;
					EnableSunsetInv = false; //reset flag for next day
				}

				if (nfil == 0)
				{
					tims[nyrstp - 1][3][ndystp - 1] = t3 + nacurr/3600.0;
				}
				else
				{
					if (t3 > tims[nyrstp - 1][3][ndystp - 1] )
						tims[nyrstp - 1][3][ndystp - 1] = t3 + nacurr/3600.0;
				}
				////////////////////////////////////////////////
			}

			if (astro)
			{
/*             finished calculations for this day */
			   //loop to next day
			   nfound = 1;
		    }
			else
			{
				//go back to calculate visible sunset */
			   goto L688;
		   }
		}
		else
		{
			//////////////record astro sunset /////////////
			nacurr = 0;
			if (EnableSunsetInv) {
				nacurr = 15;
				//don't reset inversion flag since need it for visible sunset calculation
			}

			if (nfil == 0)
			{
				tims[nyrstp - 1][3][ndystp - 1] = t3 + nacurr/3600.0;
			}
			else
			{
				if (t3 > tims[nyrstp - 1][3][ndystp - 1] )
					tims[nyrstp - 1][3][ndystp - 1] = t3 + nacurr/3600.0;
			}
			////////////////////////////////////////////////
//       continue on to calculate vis. sunset */
			nsetflag = 1;
		}
	}
	else if (nsetflag % 2 == 0)
	{
//Nsetflag = 0 or Nsetflag = 2 OR Nsetflag = 4
//    calculate astron./mishor sunrise
		///////////Astronomical/Mishor sunrise///////////////////////////////////////////////////////////
      AstronSunrise(&ndirec, &dy, &lr, &lt, &t6, &mp, &mc, &ap, &ac, &ms, &aas,
						&es, &ob, &d__, &air, &sr1, nweather, ns3, ns4, &adhocrise,
						&t3, &astro, &azi1, &jdn, &mday, &mon, &nyear, &nyl, &td,
						mint, WinTemp, &FindWinter, &EnableSunriseInv);
		///////////////////////////////////////////////////////////////////////////////////////////////

		//write astronomical sunrise time to file and loop in days
		//record chazos
		//write astronomical sunrise times to file and loop in days
		if (astro || nsetflag == 4) //astro == true is same as nsetflag == 2
		{
			if ( !ast )
			{
				//////////////record mishor sunrise /////////////
				nacurr = 0;
				if (EnableSunriseInv) {
					//add adhoc inversion fix to mishor/ast sunrise
					nacurr = -15;
					EnableSunriseInv = false; //reset flag for next day
				}
				
				if (nfil == 0)
				{
					tims[nyrstp - 1][4][ndystp - 1] = t3 + nacurr/3600.0;
				}
				else
				{
					if (t3 < tims[nyrstp - 1][4][ndystp - 1] )
						tims[nyrstp - 1][4][ndystp - 1] = t3 + nacurr/3600.0;
				}
				////////////////////////////////////////////////
			}
			else
			{
				//////////////record astro sunrise /////////////
				nacurr = 0;
				if (EnableSunriseInv) {
					//add adhoc inversion fix to mishor/ast surnise
					nacurr = -15;
					EnableSunriseInv = false; //reset flag for next day
				}

				if (nfil == 0)
				{
					tims[nyrstp - 1][2][ndystp - 1] = t3 + nacurr/3600.0;
				}
				else
				{
					if (t3 < tims[nyrstp - 1][2][ndystp - 1] )
						tims[nyrstp - 1][2][ndystp - 1] = t3 + nacurr/3600.0;
				}
				////////////////////////////////////////////////
			}

			if (astro)
			{
				//finished calculations for this day
			   //loop to next day
			   nfound = 1;
			}
			else
			{
//				go back to calculate astro sunrise, vis sunrise */
			   goto L688;
			}
		}
		else
		{

			//////////////record astro sunrise /////////////
			nacurr = 0;
			if (EnableSunriseInv) {
				//add adhoc inversion fix to mishor/ast sunrise
				nacurr = -15;
				//don't reset inversion flag since need it for visible sunrise calculation
			}

			if (nfil == 0)
			{
				tims[nyrstp - 1][2][ndystp - 1] = t3 + nacurr/3600.0;
			}
			else
			{
				if (t3 < tims[nyrstp - 1][2][ndystp - 1] )
					tims[nyrstp - 1][2][ndystp - 1] = t3 + nacurr/3600.0;
			}
			////////////////////////////////////////////////

			//continue on to calculate vis. sunrise */
		   nsetflag = 0;
			}
		}
L692:
		if (!astro) //the following loop is only for visible sunrise/sunset calculations
		{

			nn = 0;
			nfound = 0;
			nabove = 0;
			i__2 = nlpend;
			i__3 = ndirec;
			for (nloops = nloop; i__3 < 0 ? nloops >= i__2 : nloops <= i__2; nloops += i__3)
			{
				++nn;
				hran = sr1 - stepsec * nloops / 3600.;
				//step in hour angle by stepsec sec
				if (nyear >= BEG_SUN_APPROX && nyear <= END_SUN_APPROX)
				{
					dy1 = dy - ndirec * hran / 24.;
					decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
				}
				else
				{
					hrs = 12.0 - ndirec * hran;
					d__ = Decl3(&jdn, &hrs, &dy, &td, &mday, &mon, &nyear, &nyl,
						        &ms, &aas, &ob, &ra, &es, 1);
				}
				//recalculate declination at each new hour angle
				z__ = sin(lr) * sin(d__) + cos(lr) * cos(d__) * cos(hran / ch);
				z__ = acos(z__);
				alt1 = pi * .5 - z__;
				//altitude in radians of sun's center
				azi1 = cos(d__) * sin(hran / ch);
				azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(hran / ch);
				azi = atan(azi1 / azi2) / cd;
				if (azi > 0.)
				{
					azi1 = 90. - azi;
				}
				if (azi < 0.)
				{
					azi1 = -(azi + 90.);
				}
				if (fabs(azi1) > maxang)
				{
					if (IgnoreInsuffAzimuthRange)
					{
						if (nsetflag == 0) //sunrise
						{
							tims[nyrstp - 1][nsetflag][ndystp - 1] = 99999;
							azils[nyrstp - 1][nsetflag][ndystp - 1] = -9999;
							dobs[nyrstp - 1][nsetflag][ndystp - 1] = -9999;
						}
						else if (nsetflag == 1) //sunset
						{
							tims[nyrstp - 1][nsetflag][ndystp - 1] = -99999;
							azils[nyrstp - 1][nsetflag][ndystp - 1] = -9999;
							dobs[nyrstp - 1][nsetflag][ndystp - 1] = -9999;
						}
						nfound = 1;
						break; //go on to next day
					}
					else
					{
						//azimuthal range insufficient
						return -1; //goto L9050;
					}
				}

				azi1 = ndirec * azi1;

				if (nsetflag == 1) //<---
				{
					//to first order, the top and bottom of the sun have the
					//following altitudes near the sunset
					al1 = alt1 / cd + a1;
					//first guess for apparent top of sun
					pt = 1.; //fudge factor -- default is 1.0 for no fudge

					//deter__min refraction for this solar altitude
					if (!VDW_REF) {
						refract_(&al1, p, &t, &hgt, nweather, &dy, &refrac1, ns1, ns2);
					} else {
						refvdw_(&hgt, &al1, &refrac1);
					}

					//first iteration to find position of upper limb of sun
					al1 = alt1 / cd + .2666f + pt * refrac1;
					//top of sun with refraction
					if (niter == 1)
					{
						for (nac = 1; nac <= 10; ++nac)
						{
							al1o = al1;
							if (!VDW_REF) {
								refract_(&al1, p, &t, &hgt, nweather, &dy, &refrac1, ns1, ns2);
							} else {
								refvdw_(&hgt, &al1, &refrac1);
							}

							//subsequent iterations
							al1 = alt1 / cd + .2666f + pt * refrac1;
							//top of sun with refraction
							if (fabs(al1 - al1o) < .004f) {
								if (!VDW_REF) {
									break;
								} else {
									//scale refraction for the ground temperature at this date
									//using vdw temperature scaling
									al1 = alt1 / cd + .2666f + pt * vdwsf * refrac1;
									break;
								}
							}
	// L693:
						}
					}
	//L695:
				//now search digital elevation data
				ne = (int) ((azi1 + maxang) * 10);
				elevinty = elev[ne + 1].vert[1] - elev[ne].vert[1]; //difference of adjoining view angles
				elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0]; //difference of adjoining aximuths
				elevint = (elevinty / elevintx) * (azi1 - elev[ne].vert[0]) + elev[ne].vert[1]; //extrapolated view angle

				//now modify apparent elevention, elevint, with terrestrial refraction
				//first calculate interpolated difference of heights
				hgtdifx = elev[ne + 1].vert[3] - elev[ne].vert[3]; //difference of neighboring obstruction heights
				hgtdif = (hgtdifx / elevintx) * (azi1 - elev[ne].vert[0]) + elev[ne].vert[3];
				hgtdif -= hgt;
				//convert to kms
				hgtdif *= .001f;
				//now calculate the interpolated distance to obstruction along earth
				distx = elev[ne + 1].vert[2] - elev[ne].vert[2];
				dist = (distx / elevintx) * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
				//now calculate approximate pathlength using Lehn's/ Bruton's parabolic approximation
				//good for nearly horizontal rays with no addition for refraction
				d__3 = abs(hgtdif) - dist * dist * .5f / rearth;
				pathlength = sqrt(dist * dist + d__3 * d__3);
				//now add terrestrial refraction based on modified (Wikipedia) Bomford expression 
				if (abs(hgtdif) > 1e3f) {
					exponent = .99f;
				} else {
					exponent = .9965f;
				}
				elevint += TRfudgeVis * trbasis * pow(pathlength, exponent);
				//////////////////////////////////////////////////////////////////////////////////////////

				if ((elevint >= al1 || nloops <= 0) && nabove == 1)
				{
					//reached sunset
					nfound = 1;

					//add ad hoc inversion fix if flagged
					nacurr = 0;
					if (EnableSunsetInv && elevint <= HorizonInv) {
						//add adhoc inversion fix to mishor/ast sunset for view angles <= 0
						nacurr = 15; //this is minimum adhoc inversion fix (hours); can be a lot later
						EnableSunsetInv = false; //reset flag for this day
					}

					t3sub = t3 + ndirec * (stepsec * nloops + nacurr) / 3600.;
					if (nloops < 0)
					{
						//visible sunset can't be latter then astronomical sunset!
						t3sub = t3;
					}
					//<<<<<<<<<distance to obstruction>>>>>>
					dobsy = elev[ne + 1].vert[2] - elev[ne].vert[2];
					elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
					dobss = dobsy / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
					//record visual sunset results for time,azimuth,dist. to obstructions

            	//only record times if they are earlier (sunrise) or latter (sunset)
					if (nfil == 0) //first set of times////////////////////////////////////
					{
     					tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
						azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
						dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
					}
					else
					{
						if (t3sub > tims[nyrstp - 1][nsetflag][ndystp - 1] )
						{
							tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
							azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
							dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
						}
					}

				/////////////////////////////////////////////////////////////

				//finished for this day, loop to next day
				//goto L925;
				break;
				}
				else if (elevint < al1)
				{
					nabove = 1;
				}
			}
			else if (nsetflag == 0)
			{
				al1 = alt1 / cd + a1;
				//first guess for apparent top of sun
				pt = 1.;  //fudge factor -- default as 1.0, i.e., no fudge

				//deter__min refraction for this solar altitude
				if (!VDW_REF) {
					refract_(&al1, p, &t, &hgt, nweather, &dy, &refrac1, ns1, ns2);
				} else {
					refvdw_(&hgt, &al1, &refrac1);
				}

				//first iteration to find top of sun
				al1 = alt1 / cd + .2666f + pt * refrac1;
				//top of sun with refraction
				if (niter == 1)
				{
					for (nac = 1; nac <= 10; ++nac)
					{
						al1o = al1;
						if (!VDW_REF) {
							refract_(&al1, p, &t, &hgt, nweather, &dy, &refrac1, ns1, ns2);
						} else {
							refvdw_(&hgt, &al1, &refrac1);
						}
						
						//subsequent iterations
						al1 = alt1 / cd + .2666f + pt * refrac1;
						//top of sun with refraction
						if (fabs(al1 - al1o) < .004f) {
							if (!VDW_REF) {
								break;
							} else {
								//scale refraction for the ground temperature at this date
								//using vdw temperature scaling
								al1 = alt1 / cd + .2666f + pt * vdwsf * refrac1;
								break;
							}
						}
// L750:
					}
				}
//L760:
				alt1 = al1;
				//now search digital elevation data
				ne = (int) ((azi1 + maxang) * 10);
				elevinty = elev[ne + 1].vert[1] - elev[ne].vert[1];
				elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
				elevint = elevinty / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[1];

				//now modify apparent elevention, elevint, with terrestrial refraction
				//first calculate interpolated difference of heights
				hgtdifx = elev[ne + 1].vert[3] - elev[ne].vert[3]; //difference of neighboring obstruction heights
				hgtdif = (hgtdifx / elevintx) * (azi1 - elev[ne].vert[0]) + elev[ne].vert[3];
				hgtdif -= hgt;
				//convert to kms
				hgtdif *= .001f;
				//now calculate the interpolated distance to obstruction along earth
				distx = elev[ne + 1].vert[2] - elev[ne].vert[2];
				dist = (distx / elevintx) * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
				//now calculate approximate pathlength using Lehn's/ Bruton's parabolic approximation
				//good for nearly horizontal rays with no addition for refraction
				d__3 = abs(hgtdif) - dist * dist * .5f / rearth;
				pathlength = sqrt(dist * dist + d__3 * d__3);
				//now add terrestrial refraction based on modified (Wikipedia) Bomford expression 
				if (abs(hgtdif) > 1e3f) {
					exponent = .99f;
				} else {
					exponent = .9965f;
				}
				elevint += TRfudgeVis * trbasis * pow(pathlength, exponent);
				//////////////////////////////////////////////////////////////////////////////////////////

				if (elevint <= alt1)
				{
					if (nn > 1)
					{
						nfound = 1; //found visual sunrise

						//add ad hoc inversion fix if flagged
						nacurr = 0;
						if (EnableSunriseInv && elevint <= HorizonInv) {
							//add adhoc inversion fix to mishor/ast sunset for view angles <= 0
							nacurr = -15; //this is minimum adhoc inversion fix (hours); can be a lot later
							EnableSunriseInv = false; //reset flag for this day
						}

						//<<<<<<<<<distance to obstruction>>>>>>
						dobsy = elev[ne + 1].vert[2] - elev[ne].vert[2];
						elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
						dobss = dobsy / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
						
						//account differences between observations and the DTM ``
						t3sub = t3 + (stepsec * nloops + nacurr) / 3600.;
						if (nloops < 0)
						{
							//sunrise can't be before astronomical sunrise
							//so set it equal to astronomical sunrise time
							//minus the adhoc winter fix if any
	               			t3sub = t3;
						}

						//record visual sunrise results for time,azimuth,dist.to obstructions
						if (nfil == 0) //first set of times////////////////////////////////////
						{
     						tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
							azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
							dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
						}
						else
						{
							if (t3sub < tims[nyrstp - 1][nsetflag][ndystp - 1] )
							{
								tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
								azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
								dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
							}
						}
						break; //(goto L900) finished for this day, loop to next day
					}
					else
					{
						break; //(goto L900) reposition the sun further below the horizon (for sunrises)
						       //or further above the horizon (for sunsets)
						       //and restart the search from the new beginning
					}
				;
				}
			}
	// L850:
		}
// L900:
		//-----------------------------------------------------


		if (nfound == 0)
		//no visible sunrise/sunest was found
		//need to reposition the sun further below the horizon (for sunrises)
      //or further above the horizon (for sunsets)
      //and restart the search from the new beginning
		{
			if (nloop > 0 && nsetflag == 1)
			{
				if (nsetflag == 0)
				{
					nloop += -24;
					if (nloop < 0)
					{
						nloop = 0;
					}
				}
				else if (nsetflag == 1)
				{
					nloop += 24;
				}
//				_printf("goto\n");
				goto L692;
			}
			else if (nloop >= -100 && nsetflag == 0)
			{
			   if (nsetflag == 0)
				{
					nloop += -24;
					//IF (nloop .LT. 0) nloop = 0
				}
			  	else if (nsetflag == 1)
			  	{
				  nloop += 24;
			   }
			   goto L692;
			}
			else if (nloop < -100 && nsetflag == 0)
			{
			   nlpend += 500;
			   nloop = 72;
			   goto L692;
			}
		}

////////////////////////////////////////////////////////////////////////////
		//now tweak the nloop parameters for the next day based
		//on the experience of the present day
		if (nsetflag == 1 )
		{
			if (nloops + 30 < nloop || nloops + 30 > nloop)
			{
				nloop = nloops + 30;
			}
		}
		else if (nsetflag == 0)
		{
			if (nloops + 30 < nloop)
			{
				nloop = nloops + 30;
			}
			else if (nloops + 30 > nloop)
			{
			nloop = nloops - 30;
			}
		}
/////////////////////////////////////////////////////////////////////////
	}

//L925: //next day
	}
// L975:
	//next year
    }

//	if (elev)
//	{
//		free( elev ); //free elev memory
//		elev = 0;
//	}

    return 0;
} /* MAIN__ */

/* Subroutine */ short decl_(double *dy1, double *mp, double *mc,
	double *ap, double *ac, double *ms, double *aas,
	double *es, double *ob, double *d__)
{
    // Builtin functions
    //double acos(double), sin(double), asin(double);

    /* Local variables */


	double ec, e2c, ob2;

/*  calculates new declination by advancing all the orbital elements by a day */
    ec = cd * 1.915;
    e2c = cd * .02;
    *ms = *mp + *mc * *dy1;
    if (*ms > pi2) {
	*ms -= pi2;
    }
    *aas = *ap + *ac * *dy1;
    if (*aas > pi2) {
	*aas -= pi2;
    }
    *es = *ms + ec * sin(*aas) + e2c * sin(*aas * 2.);
    if (*es > pi2) {
	*es -= pi2;
    }

	/*change ecliptic angle for each day */
	ob2 = *ob + ob10 * *dy1;

    *d__ = asin(sin(ob2) * sin(*es));
    return 0;
} /* decl_ */


//////////////////////Menat ray tracing atmospheric refract_/////////////////////////////////////////
//calculates atmospheric refraction, "refrac1" of sun at a function of apparent solar altitude,"al1", for
//summer and winter atmospheric models. Units: degrees
//The refraction constants as a function of height and solar altitude were caluclated from ray tracing using
//the program MENAT.FOR
//Afterwards, fits were done as a function of height using the program
//FITFP (NON-LINEAR LEAST SQUARES FITTING PROGRAM USING THE IMSL ROUTINE ZXSSQ
//(employs polynomical fit as function of apparent solar altitude, al1)
//an archive of these fortran programs can be found on directory c:/for
/////////////////////////////////////////////////////////////////////////////////////
short refract_(double *al1, double *p, double *t,
	double *hgt, short *nweather, double *dy, double *
	refrac1, short *ns1, short *ns2)
//////////////////////////////////////////////////////////////////////////////////////
{
    double frwin, frsum, x1; //x;

    if (*nweather == 0)
	{
///////////////////////refsum1//////////////////////////////////
			if (*hgt >= 1150.)
			{
				x1 = *hgt * -1.2745e-5f + .15274f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01047f + *al1 * 1.7e-4f * *al1);
				frsum /= *al1 * .43732f + 1 + *al1 * .06206f * *al1;
			}
			else if (*hgt >= 1050. && *hgt < 1150.)
			{
				x1 = *hgt * -1.232e-5f + .15225f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01207f + *al1 * 1.1e-4f * *al1);
				frsum /= *al1 * .44836f + 1 + *al1 * .06573f * *al1;
			}
			else if (*hgt >= 950. && *hgt < 1050.)
			{
				x1 = *hgt * -1.2599e-5f + .15256f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01589f - *al1 * 3e-5f * *al1);
				frsum /= *al1 * .47504f + 1 + *al1 * .07487f * *al1;
			}
			else if (*hgt >= 850. && *hgt < 950.)
			{
				x1 = *hgt * -1.2876e-5f + .15279f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01631f - *al1 * 4e-5f * *al1);
				frsum /= *al1 * .47688f + 1 + *al1 * .07551f * *al1;
			}
			else if (*hgt >= 750. && *hgt < 850.)
			{
				x1 = *hgt * -1.2262e-5f + .15229f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02014f - *al1 * 1.7e-4f * *al1);
				frsum /= *al1 * .50277f + 1 + *al1 * .08453f * *al1;
			}
			else if (*hgt >= 650. && *hgt < 750.)
			{
				x1 = *hgt * -1.2343e-5f + .15239f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02058f - *al1 * 1.8e-4f * *al1);
				frsum /= *al1 * .50458f + 1 + *al1 * .0851f * *al1;
			}
			else if (*hgt >= 550. && *hgt < 650.)
			{
				x1 = *hgt * -1.3012e-5f + .15286f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02555f - *al1 * 3.6e-4f * *al1);
				frsum /= *al1 * .53764f + 1 + *al1 * .09663f * *al1;
			}
			else if (*hgt >= 450. && *hgt < 550.)
			{
				x1 = *hgt * -1.2399e-5f + .15248f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02054f - *al1 * 1.8e-4f * *al1);
				frsum /= *al1 * .50165f + 1 + *al1 * .08389f * *al1;
			}
			else if (*hgt >= 350. && *hgt < 450.)
			{
				x1 = *hgt * -1.2714e-5f + .15264f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01256f + *al1 * 9e-5f * *al1);
				frsum /= *al1 * .44703f + 1 + *al1 * .06355f * *al1;
			}
			else if (*hgt >= 250. && *hgt < 350.)
			{
				x1 = *hgt * -1.332e-5f + .15287f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01515f);
				frsum /= *al1 * .50458f + 1 + *al1 * .0851f * *al1;
			}
			else if (*hgt >= 150. && *hgt < 250.)
			{
				x1 = *hgt * -1.2843e-5f + .15277f;
				frsum = *p / (*t + 273) * (x1 - *al1 * .0075f + *al1 * 6.7e-4f * *al1);
				frsum /= *al1 * .31655f + 1 + *al1 * .01201f * *al1;
			}
			else if (*hgt >= 50. && *hgt < 150.)
			{
				x1 = *hgt * -1.2951e-5f + .15281f;
				frsum = *p / (*t + 273) * (x1 - *al1 * .00838f + *al1 * 7e-4f * *al1);
				frsum /= *al1 * .3113f + 1 + *al1 * .00986f * *al1;
			}
			else if (*hgt >= -50. && *hgt < 50.)
			{
				x1 = *hgt * -1.3233e-5f + .15237f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02692f - *al1 * 2.1e-4f * *al1);
				frsum /= *al1 * .52849f + 1 + *al1 * .10086f * *al1;
			}
			else if (*hgt >= -150. && *hgt < -50.)
			{
				x1 = *hgt * -1.3322e-5f + .15236f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02718f - *al1 * 2.1e-4f * *al1);
				frsum /= *al1 * .52854f + 1 + *al1 * .10083f * *al1;
			}
			else if (*hgt < -150.)
			{
				x1 = *hgt * -1.3424e-5f + .15234f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02744f - *al1 * 2.2e-4f * *al1);
				frsum /= *al1 * .52857f + 1 + *al1 * .10079f * *al1;
			}

///////////////////////////////////////////////////////////////////
//		refsum1_(hgt, &x, &frsum, p, t);
		*refrac1 = frsum;
    }
	else if (*nweather == 1)
	{
		//x = *al1;
/////////////////////////////refwin1////////////////////////////////////

/*       refwin1:  'calculated from MENAT.FOR/FITP using x = view angle */
		if (*hgt >= 1150.)
		{
			x1 = *hgt * -1.4893e-5f + .16878f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .0148f + *al1 * 1e-4f * *al1);
			frwin /= *al1 * .47806f + 1 + *al1 * .07447f * *al1;
		}
		else if (*hgt >= 1050. && *hgt < 1150.)
		{
			x1 = *hgt * -1.5118e-5f + .16901f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .01575f + *al1 * 7e-5f * *al1);
			frwin /= *al1 * .4834f + 1 + *al1 * .07641f * *al1;
		}
		else if (*hgt >= 950. && *hgt < 1050.)
		{
			x1 = *hgt * -1.5058e-5f + .16896f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .02169f - *al1 * 1.4e-4f * *al1);
			frwin /= *al1 * .52128f + 1 + *al1 * .09025f * *al1;
		} else if (*hgt >= 850. && *hgt < 950.)
		{
			x1 = *hgt * -1.445e-5f + .16838f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .02111f - *al1 * 1.2e-4f * *al1);
			frwin /= *al1 * .5161f + 1 + *al1 * .08825f * *al1;
		}
		else if (*hgt >= 750. && *hgt < 850.)
		{
			x1 = *hgt * -1.4876e-5f + .16876f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .0232f - *al1 * 1.9e-4f * *al1);
			frwin /= *al1 * .52786f + 1 + *al1 * .09266f * *al1;
		}
		else if (*hgt >= 650. && *hgt < 750.)
		{
			x1 = *hgt * -1.5295e-5f + .16907f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .02383f - *al1 * 2.1e-4f * *al1);
			frwin /= *al1 * .5303f + 1 + *al1 * .09358f * *al1;
		}
		else if (*hgt >= 550. && *hgt < 650.)
		{
			x1 = *hgt * -1.5256e-5f + .1691f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .02976f - *al1 * 4.2e-4f * *al1);
			frwin /= *al1 * .56622f + 1 + *al1 * .10671f * *al1;
		}
		else if (*hgt >= 450. && *hgt < 550.)
		{
			x1 = *hgt * -1.6268e-5f + .1696f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .02905f - *al1 * 3.9e-4f * *al1);
			frwin /= *al1 * .56005f + 1 + *al1 * .10422f * *al1;
		}
		else if (*hgt >= 350. && *hgt < 450.)
		{
			x1 = *hgt * -1.5319e-5f + .1691f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .01416f + *al1 * 1.1e-4f * *al1);
			frwin /= *al1 * .46823f + 1 + *al1 * .06829f * *al1;
		}
		else if (*hgt >= 250. && *hgt < 350.)
		{
			x1 = *hgt * -1.5492e-5f + .16916f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .01168f + *al1 * 2e-4f * *al1);
			frwin /= *al1 * .45241f + 1 + *al1 * .06206f * *al1;
		}
		else if (*hgt >= 150. && *hgt < 250.)
		{
			x1 = *hgt * -1.6191e-5f + .16941f;
			frwin = *p / (*t + 273) * (x1 - *al1 * .01107f + *al1 * 8.3e-4f * *al1);
			frwin /= *al1 * .32072f + 1 + *al1 * .0059f * *al1;
		}
		else if (*hgt >= 50. && *hgt < 150.)
		{
			x1 = *hgt * -1.5257e-5f + .16924f;
			frwin = *p / (*t + 273) * (x1 - *al1 * .00922f + *al1 * 8e-4f * *al1);
			frwin /= *al1 * .33125f + 1 + *al1 * .01087f * *al1;
		}
		else if (*hgt >= -50. && *hgt < 50.)
		{
			x1 = *hgt * -1.5938e-5f + .16872f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .0305f - *al1 * 2.1e-4f * *al1);
			frwin /= *al1 * .55065f + 1 + *al1 * .10908f * *al1;
		}
		else if (*hgt >= -150. && *hgt < -50.)
		{
			x1 = *hgt * -1.6064e-5f + .16872f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .03083f - *al1 * 2.2e-4f * *al1);
			frwin /= *al1 * .55083f + 1 + *al1 * .10908f * *al1;
		}
		else if (*hgt < -150.)
		{
			x1 = *hgt * -1.6195e-5f + .1687f;
			frwin = *p / (*t + 273) * (x1 + *al1 * .03119f - *al1 * 2.2e-4f * *al1);
			frwin /= *al1 * .55118f + 1 + *al1 * .10915f * *al1;
		}

/////////////////////////////////////////////////////////////////////////
//		refwin1_(hgt, &x, &frwin, p, t);
		*refrac1 = frwin;
/*       ELSE IF (nweather .EQ. 2) THEN */
/*          refrac1 = FNref(al1) */
    }
	else if (*nweather == 3)
	{
		if (*dy <= (double) (*ns1) || *dy >= (double) (*ns2))
		{
////////////////////////////refwin1/////////////////////////////////////

		/*       refwin1:  'calculated from MENAT.FOR/FITP using x = view angle */
			if (*hgt >= 1150.)
			{
				x1 = *hgt * -1.4893e-5f + .16878f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .0148f + *al1 * 1e-4f * *al1);
				frwin /= *al1 * .47806f + 1 + *al1 * .07447f * *al1;
			}
			else if (*hgt >= 1050. && *hgt < 1150.)
			{
				x1 = *hgt * -1.5118e-5f + .16901f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .01575f + *al1 * 7e-5f * *al1);
				frwin /= *al1 * .4834f + 1 + *al1 * .07641f * *al1;
			}
			else if (*hgt >= 950. && *hgt < 1050.)
			{
				x1 = *hgt * -1.5058e-5f + .16896f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .02169f - *al1 * 1.4e-4f * *al1);
				frwin /= *al1 * .52128f + 1 + *al1 * .09025f * *al1;
			} else if (*hgt >= 850. && *hgt < 950.)
			{
				x1 = *hgt * -1.445e-5f + .16838f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .02111f - *al1 * 1.2e-4f * *al1);
				frwin /= *al1 * .5161f + 1 + *al1 * .08825f * *al1;
			}
			else if (*hgt >= 750. && *hgt < 850.)
			{
				x1 = *hgt * -1.4876e-5f + .16876f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .0232f - *al1 * 1.9e-4f * *al1);
				frwin /= *al1 * .52786f + 1 + *al1 * .09266f * *al1;
			}
			else if (*hgt >= 650. && *hgt < 750.)
			{
				x1 = *hgt * -1.5295e-5f + .16907f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .02383f - *al1 * 2.1e-4f * *al1);
				frwin /= *al1 * .5303f + 1 + *al1 * .09358f * *al1;
			}
			else if (*hgt >= 550. && *hgt < 650.)
			{
				x1 = *hgt * -1.5256e-5f + .1691f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .02976f - *al1 * 4.2e-4f * *al1);
				frwin /= *al1 * .56622f + 1 + *al1 * .10671f * *al1;
			}
			else if (*hgt >= 450. && *hgt < 550.)
			{
				x1 = *hgt * -1.6268e-5f + .1696f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .02905f - *al1 * 3.9e-4f * *al1);
				frwin /= *al1 * .56005f + 1 + *al1 * .10422f * *al1;
			}
			else if (*hgt >= 350. && *hgt < 450.)
			{
				x1 = *hgt * -1.5319e-5f + .1691f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .01416f + *al1 * 1.1e-4f * *al1);
				frwin /= *al1 * .46823f + 1 + *al1 * .06829f * *al1;
			}
			else if (*hgt >= 250. && *hgt < 350.)
			{
				x1 = *hgt * -1.5492e-5f + .16916f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .01168f + *al1 * 2e-4f * *al1);
				frwin /= *al1 * .45241f + 1 + *al1 * .06206f * *al1;
			}
			else if (*hgt >= 150. && *hgt < 250.)
			{
				x1 = *hgt * -1.6191e-5f + .16941f;
				frwin = *p / (*t + 273) * (x1 - *al1 * .01107f + *al1 * 8.3e-4f * *al1);
				frwin /= *al1 * .32072f + 1 + *al1 * .0059f * *al1;
			}
			else if (*hgt >= 50. && *hgt < 150.)
			{
				x1 = *hgt * -1.5257e-5f + .16924f;
				frwin = *p / (*t + 273) * (x1 - *al1 * .00922f + *al1 * 8e-4f * *al1);
				frwin /= *al1 * .33125f + 1 + *al1 * .01087f * *al1;
			}
			else if (*hgt >= -50. && *hgt < 50.)
			{
				x1 = *hgt * -1.5938e-5f + .16872f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .0305f - *al1 * 2.1e-4f * *al1);
				frwin /= *al1 * .55065f + 1 + *al1 * .10908f * *al1;
			}
			else if (*hgt >= -150. && *hgt < -50.)
			{
				x1 = *hgt * -1.6064e-5f + .16872f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .03083f - *al1 * 2.2e-4f * *al1);
				frwin /= *al1 * .55083f + 1 + *al1 * .10908f * *al1;
			}
			else if (*hgt < -150.)
			{
				x1 = *hgt * -1.6195e-5f + .1687f;
				frwin = *p / (*t + 273) * (x1 + *al1 * .03119f - *al1 * 2.2e-4f * *al1);
				frwin /= *al1 * .55118f + 1 + *al1 * .10915f * *al1;
			}


/////////////////////////////////////////////////////////////////////////
//			refwin1_(hgt, &x, &frwin, p, t);
			*refrac1 = frwin;
		}
		else if (*dy > (double) (*ns1) && *dy < (double) (*ns2))
		{

///////////////////////refsum1//////////////////////////////////

			if (*hgt >= 1150.)
			{
				x1 = *hgt * -1.2745e-5f + .15274f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01047f + *al1 * 1.7e-4f * *al1);
				frsum /= *al1 * .43732f + 1 + *al1 * .06206f * *al1;
			}
			else if (*hgt >= 1050. && *hgt < 1150.)
			{
				x1 = *hgt * -1.232e-5f + .15225f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01207f + *al1 * 1.1e-4f * *al1);
				frsum /= *al1 * .44836f + 1 + *al1 * .06573f * *al1;
			}
			else if (*hgt >= 950. && *hgt < 1050.)
			{
				x1 = *hgt * -1.2599e-5f + .15256f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01589f - *al1 * 3e-5f * *al1);
				frsum /= *al1 * .47504f + 1 + *al1 * .07487f * *al1;
			}
			else if (*hgt >= 850. && *hgt < 950.)
			{
				x1 = *hgt * -1.2876e-5f + .15279f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01631f - *al1 * 4e-5f * *al1);
				frsum /= *al1 * .47688f + 1 + *al1 * .07551f * *al1;
			}
			else if (*hgt >= 750. && *hgt < 850.)
			{
				x1 = *hgt * -1.2262e-5f + .15229f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02014f - *al1 * 1.7e-4f * *al1);
				frsum /= *al1 * .50277f + 1 + *al1 * .08453f * *al1;
			}
			else if (*hgt >= 650. && *hgt < 750.)
			{
				x1 = *hgt * -1.2343e-5f + .15239f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02058f - *al1 * 1.8e-4f * *al1);
				frsum /= *al1 * .50458f + 1 + *al1 * .0851f * *al1;
			}
			else if (*hgt >= 550. && *hgt < 650.)
			{
				x1 = *hgt * -1.3012e-5f + .15286f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02555f - *al1 * 3.6e-4f * *al1);
				frsum /= *al1 * .53764f + 1 + *al1 * .09663f * *al1;
			}
			else if (*hgt >= 450. && *hgt < 550.)
			{
				x1 = *hgt * -1.2399e-5f + .15248f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02054f - *al1 * 1.8e-4f * *al1);
				frsum /= *al1 * .50165f + 1 + *al1 * .08389f * *al1;
			}
			else if (*hgt >= 350. && *hgt < 450.)
			{
				x1 = *hgt * -1.2714e-5f + .15264f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01256f + *al1 * 9e-5f * *al1);
				frsum /= *al1 * .44703f + 1 + *al1 * .06355f * *al1;
			}
			else if (*hgt >= 250. && *hgt < 350.)
			{
				x1 = *hgt * -1.332e-5f + .15287f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .01515f);
				frsum /= *al1 * .50458f + 1 + *al1 * .0851f * *al1;
			}
			else if (*hgt >= 150. && *hgt < 250.)
			{
				x1 = *hgt * -1.2843e-5f + .15277f;
				frsum = *p / (*t + 273) * (x1 - *al1 * .0075f + *al1 * 6.7e-4f * *al1);
				frsum /= *al1 * .31655f + 1 + *al1 * .01201f * *al1;
			}
			else if (*hgt >= 50. && *hgt < 150.)
			{
				x1 = *hgt * -1.2951e-5f + .15281f;
				frsum = *p / (*t + 273) * (x1 - *al1 * .00838f + *al1 * 7e-4f * *al1);
				frsum /= *al1 * .3113f + 1 + *al1 * .00986f * *al1;
			}
			else if (*hgt >= -50. && *hgt < 50.)
			{
				x1 = *hgt * -1.3233e-5f + .15237f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02692f - *al1 * 2.1e-4f * *al1);
				frsum /= *al1 * .52849f + 1 + *al1 * .10086f * *al1;
			}
			else if (*hgt >= -150. && *hgt < -50.)
			{
				x1 = *hgt * -1.3322e-5f + .15236f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02718f - *al1 * 2.1e-4f * *al1);
				frsum /= *al1 * .52854f + 1 + *al1 * .10083f * *al1;
			}
			else if (*hgt < -150.)
			{
				x1 = *hgt * -1.3424e-5f + .15234f;
				frsum = *p / (*t + 273) * (x1 + *al1 * .02744f - *al1 * 2.2e-4f * *al1);
				frsum /= *al1 * .52857f + 1 + *al1 * .10079f * *al1;
			}

///////////////////////////////////////////////////////////////////
//			refsum1_(hgt, &x, &frsum, p, t);
			*refrac1 = frsum;
		}
    }
    return 0;
} /* refract_ */

///////////////////////refvdw_ VDW ray tracing based refraction//////////////////////////
/* Subroutine */ short refvdw_(double *hgt, double *x, double *fref)
//////////////////////////////////////////////////////////////////////////////////
//input parameters: observer hgt, hgt (meters)
//					view angle at observer, x (degrees)
//
//output			refraction, fref in degrees
//////////////////////////////////////////////////////////////////////////////////
{
    double ca[11], coef[55]	/* was [5][11] */;

/* 		use the polynomial fit coefficients of the vdw refraction */
/* 		as a function of observer's height, hgt in meters, and the apparent */
/* `		view angle, x in degrees, to calculate the total atms. refraction in mrad */
/* 		polynomical fit coefficients used to calculate */
/* 		vdw total atmospheric refraction vs height */
    coef[0] = 9.56267125268496f;
    coef[1] = -8.6718429211079e-4f;
    coef[2] = 3.1664677349513e-8f;
    coef[3] = -2.04067678864827e-13f;
    coef[4] = -6.21413591282229e-17f;
    coef[5] = -3.54681762174248f;
    coef[6] = 3.05885370538294e-4f;
    coef[7] = -3.48413989765623e-9f;
    coef[8] = -3.27424677578751e-12f;
    coef[9] = 4.85180396156723e-16f;
    coef[10] = 1.00487516923555f;
    coef[11] = -7.12411305623716e-5f;
    coef[12] = -1.30264294792463e-8f;
    coef[13] = 6.08386198681256e-12f;
    coef[14] = -8.26564806865056e-16f;
    coef[15] = -.234676117102f;
    coef[16] = 4.23105602906229e-6f;
    coef[17] = 1.25823603467313e-8f;
    coef[18] = -4.77064146898649e-12f;
    coef[19] = 6.38020633504241e-16f;
    coef[20] = .0455474692911979f;
    coef[21] = 4.20546127818185e-6f;
    coef[22] = -5.71051596715397e-9f;
    coef[23] = 2.05052061222564e-12f;
    coef[24] = -2.73486893484326e-16f;
    coef[25] = -.00707867490693562f;
    coef[26] = -1.77071205323987e-6f;
    coef[27] = 1.51606775168871e-9f;
    coef[28] = -5.33770875527936e-13f;
    coef[29] = 7.12762362198471e-17f;
    coef[30] = 8.32295487796478e-4f;
    coef[31] = 3.56012191465623e-7f;
    coef[32] = -2.51083454939817e-10f;
    coef[33] = 8.78112651062853e-14f;
    coef[34] = -1.17556897803785e-17f;
    coef[35] = -6.96285190742393e-5f;
    coef[36] = -4.19488560840601e-8f;
    coef[37] = 2.62823518700555e-11f;
    coef[38] = -9.18795592984757e-15f;
    coef[39] = 1.23368951173372e-18f;
    coef[40] = 3.85246830558751e-6f;
    coef[41] = 2.93954176835877e-9f;
    coef[42] = -1.6910427905383e-12f;
    coef[43] = 5.92963984786967e-16f;
    coef[44] = -7.98575122972107e-20f;
    coef[45] = -1.25306160093963e-7f;
    coef[46] = -1.13531297031204e-10f;
    coef[47] = 6.10692317760114e-14f;
    coef[48] = -2.15222164778055e-17f;
    coef[49] = 2.90679288057578e-21f;
    coef[50] = 1.80519843190424e-9f;
    coef[51] = 1.8626817474919e-12f;
    coef[52] = -9.47783437562306e-16f;
    coef[53] = 3.36109376083127e-19f;
    coef[54] = -4.55153453427451e-23f;
/* 		write out all loops inline for speed */
    ca[0] = coef[0] + coef[1] * *hgt + coef[2] * *hgt * *hgt + coef[3] * *hgt 
	    * *hgt * *hgt + coef[4] * *hgt * *hgt * *hgt * *hgt;
    ca[1] = coef[5] + coef[6] * *hgt + coef[7] * *hgt * *hgt + coef[8] * *hgt 
	    * *hgt * *hgt + coef[9] * *hgt * *hgt * *hgt * *hgt;
    ca[2] = coef[10] + coef[11] * *hgt + coef[12] * *hgt * *hgt + coef[13] * *
	    hgt * *hgt * *hgt + coef[14] * *hgt * *hgt * *hgt * *hgt;
    ca[3] = coef[15] + coef[16] * *hgt + coef[17] * *hgt * *hgt + coef[18] * *
	    hgt * *hgt * *hgt + coef[19] * *hgt * *hgt * *hgt * *hgt;
    ca[4] = coef[20] + coef[21] * *hgt + coef[22] * *hgt * *hgt + coef[23] * *
	    hgt * *hgt * *hgt + coef[24] * *hgt * *hgt * *hgt * *hgt;
    ca[5] = coef[25] + coef[26] * *hgt + coef[27] * *hgt * *hgt + coef[28] * *
	    hgt * *hgt * *hgt + coef[29] * *hgt * *hgt * *hgt * *hgt;
    ca[6] = coef[30] + coef[31] * *hgt + coef[32] * *hgt * *hgt + coef[33] * *
	    hgt * *hgt * *hgt + coef[34] * *hgt * *hgt * *hgt * *hgt;
    ca[7] = coef[35] + coef[36] * *hgt + coef[37] * *hgt * *hgt + coef[38] * *
	    hgt * *hgt * *hgt + coef[39] * *hgt * *hgt * *hgt * *hgt;
    ca[8] = coef[40] + coef[41] * *hgt + coef[42] * *hgt * *hgt + coef[43] * *
	    hgt * *hgt * *hgt + coef[44] * *hgt * *hgt * *hgt * *hgt;
    ca[9] = coef[45] + coef[46] * *hgt + coef[47] * *hgt * *hgt + coef[48] * *
	    hgt * *hgt * *hgt + coef[49] * *hgt * *hgt * *hgt * *hgt;
    ca[10] = coef[50] + coef[51] * *hgt + coef[52] * *hgt * *hgt + coef[53] * 
	    *hgt * *hgt * *hgt + coef[54] * *hgt * *hgt * *hgt * *hgt;
    *fref = ca[0] + ca[1] * *x + ca[2] * *x * *x + ca[3] * *x * *x * *x + ca[
	    4] * *x * *x * *x * *x + ca[5] * *x * *x * *x * *x * *x + ca[6] * 
	    *x * *x * *x * *x * *x * *x + ca[7] * *x * *x * *x * *x * *x * *x 
	    * *x + ca[8] * *x * *x * *x * *x * *x * *x * *x * *x + ca[9] * *x 
	    * *x * *x * *x * *x * *x * *x * *x * *x + ca[10] * *x * *x * *x * 
	    *x * *x * *x * *x * *x * *x * *x;
	//convert mrad to degrees
	*fref *= dcmrad;
    return 0;
} /* refvdw_ */



/////////////////////casgeo///////////////////////////////////////////////////
short casgeo_(double *gn, double *ge, double *l1, double *l2)
//////////////////////////////////////////////////////////////////////////////
//converts ITM coordinates to geo coordinates (Clark geoid)
///////////////////////////////////////////////////////////
{
    /* System generated locals */
    double d__1;

    /* Local variables */
    double r__, a2, b2, c1, c2, f1, g1, g2, e4, c3, d1, a3, o1, o2,
	     a4, b3, s1, s2, b4, c4, c5, x1, y1, y2, x2, c6, d2, r3, r4, r2,
	    a5, d3;
	bool ggpscorrection = true;

    g1 = *gn * 1e3f;
    g2 = *ge * 1e3f;
    r__ = 57.2957795131f;
    b2 = .03246816f;
    f1 = 206264.806247096f;
    s1 = 126763.49f;
    s2 = 114242.75f;
    e4 = .006803480836f;
    c1 = .0325600414007f;
    c2 = 2.55240717534e-9f;
    c3 = .032338519783f;
    x1 = 1170251.56f;
    y1 = 1126867.91f;
    y2 = g1;
/*       GN & GE */
    x2 = g2;
    if (x2 > 7e5f) {
	goto L5;
    }
    x1 += -1e6f;
L5:
    if (y2 > 5.5e5f) {
	goto L10;
    }
    y1 += -1e6f;
L10:
    x1 = x2 - x1;
    y1 = y2 - y1;
    d1 = y1 * b2 / 2.f;
    o1 = s2 + d1;
    o2 = o1 + d1;
    a3 = o1 / f1;
    a4 = o2 / f1;
/* Computing 2nd power */
    d__1 = sin(a3);
    b3 = 1.f - e4 * (d__1 * d__1);
    b4 = b3 * sqrt(b3) * c1;
/* Computing 2nd power */
    d__1 = sin(a4);
    c4 = 1.f - e4 * (d__1 * d__1);
/* Computing 2nd power */
    d__1 = c4;
    c5 = tan(a4) * c2 * (d__1 * d__1);
/* Computing 2nd power */
    d__1 = x1;
    c6 = c5 * (d__1 * d__1);
    d2 = y1 * b4 - c6;
    c6 /= 3.f;
/*       LAT */
    *l1 = (s2 + d2) / f1;
    r3 = o2 - c6;
    r4 = r3 - c6;
    r2 = r4 / f1;
/* Computing 2nd power */
    d__1 = sin(*l1);
    a2 = 1.f - e4 * (d__1 * d__1);
	//latitude
    *l1 = r__ * *l1;
    a5 = sqrt(a2) * c3;
    d3 = x1 * a5 / cos(r2);
/*       LON */
    *l2 = r__ * ((s1 + d3) / f1);
/*       THIS IS THE EASTERN HEMISPHERE! */
    *l2 = -(*l2);
    if (ggpscorrection) {
        //Use the approximate correction factor
        //in order to agree with GPS.
        *l2 -= 0.0013;
        *l1 += 0.0013;
	}

    return 0;
} /* casgeo_ */

/////////////////geocasc/////////////////////////////////////////////////////
short geocasc_(double *l1, double *l2, double *g1, double *g2)
//////////////////////////////////////////////////////////////////////////////
//converts geo coordinates to ITM (Clark Geoid)
//used for Eretz Yisroel neighborhoods to convert from user
//supplied WGS84 geographic coordinates to Clark geoid ITM coordinates
////////////////////////////////////////////////////////////
{
    /* System generated locals */
    double d__1, d__2;

    /* Local variables */
    double d1, d2, d5, e3, g3, g4, l3, l4;
    int m1;
    double s1, s2, al;

    m1 = 0;
    *g1 = 1e5f;
    *g2 = 1e5f;
    s1 = 31.4896370431f;
    s2 = 34.4727144086f;
    e3 = 1e-4f;
    l3 = *l1;
    l4 = *l2;
L5:
    al = (s1 + l3) / 114.5915590262f;
    g3 = fnm_(&al);
    g4 = fnp_(&al);
    d1 = (l3 - s1) * g3;
    d2 = (l4 - s2) * g4;
    *g1 += d1;
    *g2 += d2;
    ++m1;
/* Computing 2nd power */
    d__1 = d1;
/* Computing 2nd power */
    d__2 = d2;
    d5 = sqrt(d__1 * d__1 + d__2 * d__2);
    if (d5 < e3) {
	goto L10;
    }
    if (m1 > 10) {
	goto L15;
    }
	if (fabs(*g1) > 1e+2) *g1 *= 1e-3; //Eretz Yisroel
	if (fabs(*g2) > 1e+2) *g2 *= 1e-3;
    casgeo_(g1, g2, l1, l2);
	if (fabs(*g1) < 1e+3) *g1 *= 1e+3; //inverse transformation from above
	if (fabs(*g2) < 1e+3) *g2 *= 1e+3;
    s1 = *l1;
    s2 = -*l2; //Eastern Hemisphere
    goto L5;
L10:
    *l1 = l3;
    *l2 = l4;
    return 0;
L15:
    return -1; //didn't converge
} /* geocasc_ */


/////////////////////fnp/////////////////////
	double fnp_(double *p)
////////////////////////////////////////////////
{
    /* System generated locals */
    double ret_val;

/*     Length of a degree parallel: P must be latitude in radians */
    ret_val = cos(*p) * 111415.13f - cos(*p * 3.f) * 94.55f + cos(*p * 5.f) *
	    .012f;
    return ret_val;
} /* fnp_ */

/////////////////////////fnm/////////////////////////
	double fnm_(double *p)
/////////////////////////////////////////////////////////////
{
    /* System generated locals */
    double ret_val;

/*     Length of a degree meridian: P must be latitude in radians */
    ret_val = 111132.09f - cos(*p * 2.f) * 566.05f + cos(*p * 4.f) * 1.2f -
	    cos(*p * 6.f) * .002f;
    return ret_val;
} /* fnm_ */



/////////////////////////HTML HEADER////////////////////////
short WriteHtmlHeader()
////////////////writes HTML header of output///////////////////
{

	char bakColor[8] = "#E1E2D5"; //"#666699"; //"#333366"; //"#CDCEB9"; //"#E1E2D5"; //"#008080"; //Header's background color
	char forColor[8] = "#563300"; //"#333333"; //"#996633"; //"#000000"; //"#FFFFFF"; //Header's foreground color

	if (wroteheader) return 0; //already wrote the header

	fprintf(stdout, "%s\n\n", "Content-type: text/html");
	//--html4--fprintf(stdout, "%s\n", "<!doctype html public \"-//w3c//dtd html 4.0 Transitional//en\">");

	if (! optionheb)
	{
		if (SRTMflag < 10) //html4
		{
			fprintf(stdout, "%s\n", "<!doctype html public \"-//w3c//dtd html 4.0 Transitional//en\">");
		}
		else if (SRTMflag > 9) //html5
		{
			fprintf( stdout, "%s\n", "<!DOCTYPE html>"); //html 5 (needed for progress bar)
			fprintf( stdout, "%s\n", "<html lang=\"en\">"); //<html>"); (html4)
		}
	}
	else
	{
		if (SRTMflag < 10 ) //html4
		{
			fprintf(stdout, "%s\n", "<!doctype html public \"-//w3c//dtd html 4.0 Transitional//en\">");
			sprintf( buff1, "%s%s%s%s%s", "<html dir = ", "\"", "rtl", "\"", ">");
		}
		else if ((SRTMflag > 9) && !wroteheader)  //html5
		{
			fprintf( stdout, "%s\n", "<!DOCTYPE html>"); //html 5 (needed for progress bar)
			sprintf( buff1, "%s", "<html dir=\"rtl\" lang=\"he\">"); //html5
		}
		fprintf( stdout, "%s\n", &buff1[0] );
	}
	fprintf(stdout, "%s\n", "<head>");
	fprintf(stdout, "%s\n", "    <title>Your Chai Table</title>");
	fprintf(stdout, "%s\n", "    <meta http-equiv=\"content-type\" content=\"text/html; charset=UTF-8\"/>");
	fprintf(stdout, "%s%s%s\n", "    <meta name=\"author\" content=\"Chaim Keller, Chai Tables, ", &DomainName[0], "\"/>");
	fprintf(stdout, "%s\n", "    <script language=\"JavaScript\"/>" );
	fprintf(stdout, "%s\n", "    <!-- Begin" );
	fprintf(stdout, "%s\n", "    function printWindow() {" );
	fprintf(stdout, "%s\n", "    bV = parseInt(navigator.appVersion);" );
	fprintf(stdout, "%s\n", "    if (bV >= 4) window.print();" );
	fprintf(stdout, "%s\n", "    }" );
	fprintf(stdout, "%s\n", "    //  End -->" );
	fprintf(stdout, "%s\n", "    </script>" );
	fprintf(stdout, "%s\n", "</head>");
	fprintf(stdout, "%s\n", "<body>");

	//now write picture headers

	if (!optionheb) //English header
	{
        /*
		fprintf(stdout, "%s\n",	"<table style=\"margin-left: auto; margin-right: auto; text-align: left; height: 89px; width: 905px;\" border=\"0\" cellpadding=\"2\" cellspacing=\"2\">" );
		fprintf(stdout, "\t%s\n", "<tbody>" );
		fprintf(stdout, "\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 144px; height: 78px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"100%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td rowspan=\"2\"><a href=\"http://www.zemanim.org\"><img style=\"border: 0px solid ; height: 60px; width: 136px;\" alt=\"Luchos Chai\" title=\"Luchos Chai\" src=\"luchoschai.jpg\" /></a></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 336px; height: 78px\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" width=\"99%\" cellspacing=\"0\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td rowspan=\"2\"><i><b><span lang=\"en-us\"><font size=\"5\" color=\"", forColor, "\">The Chai Tables</font></span></b></i></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s%s%s%s%s%s%s%s%s\n","<td><EM><font color=", forColor, ">(sponsored by </FONT></EM><A href=\"http://", &SponsorLine[6][0],"\"><font color=", forColor, "><EM>", &SponsorLine[7][0], "</EM></FONT></A><font color=", forColor, "><EM>)</EM></FONT></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 405px; height: 78px\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">&nbsp;" );
		fprintf(stdout, "\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"117%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td rowspan=\"3\" width=\"15%\" align=\"center\"><img border=\"0\" src=\"", &SponsorLine[3][0], "\" width=\"48\" height=\"64\"></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"28%\"><p align=\"center\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"28%\" colspan=\"4\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td width=\"23%\"><font color=\"", forColor, "\"><b><span lang=\"he\">&#1489;&#1505;&#34;&#1491;</span></b></font></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s%s%s\n","<td colspan=\"4\"><b><i><font color=\"", forColor, "\">", &SponsorLine[1][0], "</font></i></b></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"15%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"46%\" colspan=\"3\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"38%\" colspan=\"3\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t%s\n","</tbody>" );
		fprintf(stdout, "%s\n","</table>" );
		fprintf(stdout, "%s\n","<hr>" );
		fprintf(stdout, "%s\n","<p>&nbsp;</p>" );
		*/
		/*
		fprintf(stdout, "%s\n","<div align=\"center\">" );
		fprintf(stdout, "%s\n","<table border = \"2\" cellpadding=\"0\" cellspacing=\"0\" width=\"700\" height=\"51\">" );
		fprintf(stdout, "\t%s\n","<tbody>" );
		fprintf(stdout, "\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark = \"", bakColor, "\" >" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"95%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\" align = \"center\" rowspan = \"2\" >&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\" align=\"center\" >" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\" align=\"center\" >" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<a href=\"http://www.zemanim.org\"><img border=\"0\" alt=\"Luchos Chai\" title=\"Luchos Chai\" src=\"luchoschai.jpg\" width=\"63\" height=\"36\"></a></td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\" align=\"center\" >&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"75%\" align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n","<p align=\"left\"><font size=\"5\" color=\"", forColor, "\"><i>The Chai Tables</i></font>" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s\n", "<font size=\"2\" color=\"", forColor, "\">(</font><i><font size=\"2\" color=\"", forColor, "\" >sponsored by" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n", "<a href=\"http://", &SponsorLine[6][0], "\" style=\"text-decoration: none\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s%s%s\n", "<font color=\"", forColor, "\" >", &SponsorLine[7][0], "</font></a></font></i><font size=\"2\" color=\"", forColor, "\">)</font></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 332px; height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"102%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<p align=\"center\">" );
		fprintf(stdout, "\t\t\t\t%s%s%s\n","<img border=\"0\" src=\"", &SponsorLine[3][0], "\" width=\"42\" height=\"52\" ></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"2%\" >&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s%s%s%s%s\n","<td width=\"80%\" ><font color = \"", forColor, "\" ><i>", &SponsorLine[1][0], "</i></font></td>" );
		fprintf(stdout, "\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t%s\n","</tbody>" );
		fprintf(stdout, "%s\n","</table>" );
		fprintf(stdout, "%s\n","</div>" );
        fprintf(stdout, "%s\n","<hr/>" );
        */

		fprintf(stdout, "%s\n", "<div align=\"center\">" );
		fprintf(stdout, "%s\n", "<table border = \"2\" cellpadding=\"0\" cellspacing=\"0\" width=\"700\" height=\"51\">" );
		fprintf(stdout, "\t%s\n", "<tbody>" );
		fprintf(stdout, "\t\t%s\n",	"<tr>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n", "<td style=\"height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n",	"<table border=\"0\" style=\"border-collapse: collapse\" width=\"95%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"2%\" align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"21%\" align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n",	"<a href=\"http://", &DomainName[0], "\"><img border=\"0\" src=\"/luchoschai.jpg\" width=\"63\" height=\"36\" alt=\"Luchos Chai\" title=\"Luchos Chai\"></a></td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"2%\" align=\"center\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"75%\" align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n",	"<p align=\"left\"><font size=\"5\" color=\"", forColor, "\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<i>The Chai Tables </i></font>" );

		if (strlen(&SponsorLine[6][0]) > 1)
		{
			fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s\n",	"<font size=\"2\" color=\"", forColor, "\">(</font><i><font size=\"2\" color=\"", forColor, "\">sponsored by" );
			fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n",	"<a href=\"http://", &SponsorLine[6][0], "\" style=\"text-decoration: none\">" );
			fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s%s%s\n", "<font color=\"", forColor, "\">", &SponsorLine[7][0], "</font></a></font></i><font size=\"2\" color=\"", forColor, "\">)</font></td>" );
		}

		fprintf(stdout, "\t\t\t\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n",	"<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "</table>" );
		fprintf(stdout, "\t\t\t\t%s\n", "</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s%s%s\n", "<td style=\"width: 332px; height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "<table border=\"0\" style=\"border-collapse: collapse\" width=\"102%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td width=\"18%\">" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<p align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s%s%s\n", "<img border=\"0\" src=\"/", &SponsorLine[3][0], "\" width=\"42\" height=\"52\"></td>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s%s%s%s%s\n", "<td width=\"80%\"><font color=\"", forColor, "\"><i>", &SponsorLine[1][0], "</i></font></td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t\t%s\n", "<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n", "</table>" );
		fprintf(stdout, "\t\t\t%s\n", "</td>" );
		fprintf(stdout, "\t\t%s\n", "</tr>" );
		fprintf(stdout, "\t%s\n", "</tbody>" );
		fprintf(stdout, "%s\n", "</table>" );
		fprintf(stdout, "%s\n", "</div>" );
		fprintf(stdout, "%s\n", "<hr>" );


	}
	else //Hebrew header
	{
        /*
		fprintf(stdout, "%s\n",	"<table style=\"margin-left: auto; margin-right: auto; text-align: left; height: 89px; width: 861px;\" border=\"0\" cellpadding=\"2\" cellspacing=\"2\">" );
		fprintf(stdout, "\t%s\n", "<tbody>" );
		fprintf(stdout, "\t\t%s\n", "<tr>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 144px; height: 78px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"100%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td rowspan=\"2\"><a href=\"http://www.zemanim.org\"><img style=\"border: 0px solid ; height: 60px; width: 136px;\" alt=\"Luchos Chai\" title=\"Luchos Chai\" src=\"luchoschai.jpg\" /></a></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td>&nbsp;</td>");
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 267px; height: 78px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" width=\"99%\" cellspacing=\"0\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td rowspan=\"2\"><font size=\"5\" color=\"", forColor, "\"><b><i>&#1500;&#1493;&#1495;&#1493;&#1514;&#32;&#34;&#1495;&#1497;&#34;</i></b></font></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t\t%s\n","<tr>");
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"11\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s%s%s%s%s%s%s%s%s%s%s\n","<td><font color=", forColor, "></font><font color=\"", forColor, "\" size=\"4\"><i>&#40;&#1489;&#1495;&#1505;&#1493;&#1514;-&nbsp;</font><a href=\"http://", &SponsorLine[4][0], "\"><font color=\"", forColor, "\" size=\"4\">", &SponsorLine[5][0], "&#41;</font></a><font color=\"", forColor, "\"></i></font></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>");
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 430px; height: 78px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">&nbsp;" );
		fprintf(stdout, "\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"100%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td rowspan=\"3\" width=\"22%\" align=\"center\"><img border=\"0\" src=\"", &SponsorLine[2][0], "\" width=\"48\" height=\"64\"></td>" );
		fprintf(stdout, "\t\t\t\t\t%s%s%s\n","<td width=\"42%\">&nbsp;<font color=\"", forColor, "\"><b><span lang=\"he\">&#1489;&#1505;&#34;&#1491;</span></b></font></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td dir=\"rtl\" width=\"34%\">" );
        fprintf(stdout, "\t\t\t\t%s\n","<p align=\"right\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s%s%s%s%s\n","<td colspan=\"2\"><font color=\"", forColor, "\"><i><b><font size=\"4\">", &SponsorLine[0][0], "</font></b></i></font></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"22%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"42%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<td width=\"34%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t%s\n","</tbody>" );
		fprintf(stdout, "%s\n","</table>" );
		fprintf(stdout, "%s\n","<hr>" );
		fprintf(stdout, "%s\n","<p>&nbsp;</p>" );
		*/

		fprintf(stdout, "%s\n","<div align=\"center\">" );
		fprintf(stdout, "%s\n","<table border = \"2\" cellpadding=\"0\" cellspacing=\"0\" width=\"700\" height=\"51\">" );
		fprintf(stdout, "\t%s\n","<tbody>" );
		fprintf(stdout, "\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark = \"", bakColor, "\" >" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"95%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\" align=\"center\" rowspan=\"2\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\" align=\"center\" rowspan=\"2\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s\n","<a href=\"http://", &DomainName[0], "\"><img border=\"0\" alt=\"Luchos Chai\" title=\"Luchos Chai\" src=\"/luchoschai.jpg\" width=\"63\" height=\"36\"></a></td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\" align=\"center\" rowspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"75%\" align=\"center\">" );
		fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s\n","<p align=\"right\"><b><i><font size=\"6\" color=\"", forColor, "\">&nbsp;</font><font size=\"5\" color=\"", forColor, "\">&#1500;&#1493;&#1495;&#1493;&#1514;&#32;&#34;&#1495;&#1497;&#34;</font></i></b></td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );

		if (strlen(&SponsorLine[4][0]) > 1)
		{
			fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
			fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"75%\" align=\"center\">" );
			fprintf(stdout, "\t\t\t\t\t\t%s%s%s%s%s%s%s%s%s\n","<p align=\"right\"><b><font size=\"2\" color=\"", forColor, "\">&#40;&#1489;&#1495;&#1505;&#1493;&#1514;-&nbsp;<a href=\"http://", &SponsorLine[4][0], "\" style=\"text-decoration: none\"><font color=\"", forColor, "\">", &SponsorLine[5][0], "</font></a></font><font size=\"2\">)</font></b></td>" );
			fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		}

		fprintf(stdout, "\t\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"2%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td width=\"21%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t\t%s%s%s%s%s\n","<td style=\"width: 332px; height: 51px; text-align:center\" bgcolor=\"", bakColor, "\" bordercolordark=\"", bakColor, "\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<table border=\"0\" style=\"border-collapse: collapse\" width=\"102%\" cellpadding=\"0\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"18%\">" );
		fprintf(stdout, "\t\t\t\t%s\n","<p>" );
		fprintf(stdout, "\t\t\t\t%s%s%s\n","<img border=\"0\" src=\"/", &SponsorLine[2][0], "\" width=\"42\" height=\"52\" align=\"left\"></td>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"3%\" align=\"right\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t\t%s%s%s\n","<td width=\"79%\" align=\"right\">", &SponsorLine[0][0], "</td>" );
		fprintf(stdout, "\t\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","<tr>" );
		fprintf(stdout, "\t\t\t\t%s\n","<td width=\"18%\">&nbsp;</td>" );
		fprintf(stdout, "\t\t\t%s\n","<td colspan=\"2\">&nbsp;</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t\t\t%s\n","</table>" );
		fprintf(stdout, "\t\t\t%s\n","</td>" );
		fprintf(stdout, "\t\t%s\n","</tr>" );
		fprintf(stdout, "\t%s\n","</tbody>" );
		fprintf(stdout, "%s\n","</table>" );
		fprintf(stdout, "%s\n","</div>" );
        fprintf(stdout, "%s\n","<hr/>" );

	}


	wroteheader = true; //flag that header has been written

	return 0;

}

/////////////////////////HTML closing/////////////////////////
short WriteHtmlClosing()
//////////////////////wrties HTML closing lines//////////////
{

	if (wrotecloser) return 0; //already wrote closing

	fprintf(stdout,"%s\n", "<hr>");

	if (!HebrewInterface)
	{
		fprintf(stdout,"%s\n", "<p align = left><input type=\"button\" value=\"Print This Page\" onClick=\"window.print()\"/></p>" );
	}
	else
	{
		fprintf(stdout,"%s%s%s\n", "<p align = right><input type=\"button\" value=\"", &heb8[15][0], "\" onClick=\"window.print()\"/></p>" );
	}

	fprintf(stdout,"%s\n", "<p>");

	if ((SRTMflag > 9) && optionheb) fprintf( stdout, "%s\n", "</bdo>");
	fprintf(stdout, "%s\n", "</body>");
	fprintf(stdout, "%s\n", "</html>");

	/*
	try
	{
		fclose( stdout );
	}
	catch (short ier)
	{
		ier = 0;
	}
	*/

	/*
	if (!fclose(stdout))
	{
		ErrorHandler(15, 1); //can't close file
		return -1;
	}
	*/

	wrotecloser = true;

	return 0;

}

/////////////////////////////////InitAstConst///////////////////////
void InitAstConst(short *nyr, short *nyd, double *ac, double *mc,
                  double *ec, double *e2c, double *ap, double *mp,
				  double *ob, short *nyf, double *td)
////////////////////////////////////////////////////////////////////
//initializes astronomical constant values for each year in loop
{
	double ob1 = 0;
	short nyri, nyltst, nyrtst, nyl, i__2;


	/* -------------------initialize astronomical constants--------------------------- */
		//constansts taken from p. C24 of 1996 Astronomical Almananc, accurate to 6 seconds
		//from 1950 to 2050
		//the initial values are loaded in the function "LoadConstants"
		*ac = ac0; //change in the mean anamolay per day
		*mc = mc0; //change in the mean longitude per day
		*ec = ec0; //constants of the Ecliptic longitude
		*e2c = e2c0;
		ob1 = ob10; //change in the obliquity of ecliptic per day
		*ap = ap0; //mean anomaly for Jan 0, 00:00 at Greenwich meridian
		*mp = mp0; //  mean longitude of sun Jan 0, 00:00 at Greenwich meridian
		*ob = ob0; //  ecliptic angle for Jan 0, 1996 00:00 at Greenwich meridian
	/*-----------------------------------------------------------------------------------*/

	//  shift the ephemerials for the time zone
		*ap = *ap - *td / 24. * *ac;
		*mp = *mp - *td / 24. * *mc;
		*ob = *ob - *td / 24. * ob1;

//	//  use Meeus chapter 7 formula for nyf to speed up things a bit
//
//        if ( *nyr >= 1900 && *nyr < 2100)
//        {
//           jd = 367.0 * *nyr - floor(7.0 *(*nyr + floor((m+9)/12))/4)+floor(275.*m/9)+d+1721013.5+f;
//        }
		//---------------------------------------------
		*nyf = 0;
		if (nyd < 0) {
			i__2 = *nyr;
			for (nyri = RefYear - 1; nyri >= i__2; --nyri) {
			nyrtst = nyri;
			nyltst = 365;
			if (nyrtst % 4 == 0) {
				nyltst = 366;
			}
			if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
				nyltst = 365;
			}
			*nyf = *nyf - nyltst;
			}
		}
		else if (nyd >= 0)
		{
			i__2 = *nyr - 1;
			for (nyri = RefYear; nyri <= i__2; ++nyri) {
			nyrtst = nyri;
			nyltst = 365;
			if (nyrtst % 4 == 0) {
				nyltst = 366;
			}
			if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
				nyltst = 365;
			}
			*nyf = *nyf + nyltst;
			}
		}
		//---------------------------------------------
		nyl = 365;
		if (*nyr % 4 == 0) {
			nyl = 366;
		}
		if (*nyr % 100 == 0 && *nyr % 400 != 0) {
			nyl = 365;
		}
		//---------------------------------------------
		*nyf = *nyf - 1462;
		*ob = *ob + ob10 * *nyf;
		*mp = *mp + *nyf * *mc;
		//---------------------------------------------
		while (*mp < 0.)
		{
			*mp = *mp + pi2;
		}
		//---------------------------------------------
		while (*mp > pi2)
		{
			*mp = *mp - pi2;
		}
		//---------------------------------------------
		*ap += *nyf * *ac;

		while (*ap < 0.)
		{
			*ap = *ap + pi2;
		}
		//---------------------------------------------
		while (*ap > pi2)
		{
			*ap = *ap - pi2;
		}
}

////////////////////////////AstronRefract////////////////////////
void AstronRefract( double *hgt, double *lt, short *nweather, short *PartOfDay,
				    double *dy, short *ns1, short *ns2,	double *air, double *a1,
					short mint[], short avgt[], short maxt[], 
					double *vdwsf, double *vdwrefr, double *vbweps, short *nyr)
//////////////////////////////////////////////////////////////////////////////////
//calculates astronomical refraction for a user's height, hgt
////explanation of refraction terms//////////
//(1) winrefo, sumrefo  = total angular deviation due to refraction for a light ray
//                        traveling from the horizon (hgt = 0 ) to the end of the atmosphere
//                        for the winter, summer mean North American atmospheres, resp.
//(2) eps = apparent depression angle (angle between horizontal and the horizon as
//          measured by the observer at height, hgt.
//(3) ref = atmospheric refraction of a light ray traveling from the observer at height,hgt,
//          to where the light ray grazes the earth (= the new "depressed" horizon point).
//          (this term adds onto (2) to give the total bending of the ray from
//          the new horizon point to the observer at height, hgt).
//
//the total atmoshperic refraction along the path of a light ray from an observer at height, hegt, to the
//depressed horizon point and from there to the end of the atmoshpere (and from there to the sun) is equal to:
//eps + ref + winrefo for a winter atompshere, eps + ref + sumrefo for a summer atmoshpere.
//
//eps (ELTOT) and ref are calculated approximately by using a third order polynomial
//fit vs observer's height to the actual values as a function of height outputted by the program
//MENAT.FOR (MENAT.FOR cam be found in the archive section of this source code).
//eps  = ELTOT, ref = REFT3.  The actual values of eps and ref can be found in
//the menatsum.ren and menatwin.ren files (not included)
//
//fix on 071620 -- added nsetflag as a parameter since the refraction is different
//				   for sunrise and sunset
/////////////////////////////////////////////////////////////////////////////////////
{

double lnhgt, ref = 0.0, eps = 0.0;
double m;
double tk, d__2, refrac1;
	
	if (VDW_REF) 
	{
		//using van der Werf raytracing and standard atmosphere
		//first deter__min the temperature scaling factor for this day
		short nyl = DaysInYear(nyr);
		m = MonthFromDay(&nyl, dy);

		//N.b., if using user set ground temperatures, then they were already
		//stuffed in the mint,avgt,maxt arrays by IntWeatherRegions

		switch (TempModel) 
		{
		case 1:
			//use minimum WorldClim temperature
			if (*PartOfDay == 1) //using minimum temperatures
			{
				tk = mint[(int) m - 1] + 273.15f;
			}
			else //use average temperature for anything else
			{
				tk = avgt[(int) m - 1] + 273.15f;
			}
			break;
		case 2:
			//use average WorldClim temperature for both sunrise and sunset
			tk = avgt[(int) m - 1] + 273.15f;
			break;
		case 3:
			//use maximum WorldClim temperature for both sunrie and sunset
			tk = maxt[(int) m - 1] + 273.15f;
			break;
		}
		
		//calculate van der Werf temperature scaling factors for refraction/
		d__2 = 288.15f / tk;
		*vdwsf = pow(d__2, vdwconst[1]); //scaling for zero to infinity refraction, as for terrestrial refraction
		*vdwrefr = pow(d__2, vdwconst[2]);//scaling for hgt to zero refraction, i.e., ref
		*vbweps = pow(d__2, vdwconst[3]); //scaling for geometric angle contribution to refraction, i.e., eps

		if (*hgt > 0.) {
		    lnhgt = log(*hgt * .001);
		}
		//calculate total atmospheric refraction from the observer's height
		//to the horizon and then to the end of the atmosphere
		// All refraction terms have units of mrad 
		if (*hgt <= 0.) {
		    goto L590;
		}
		//refraction contribution from observer's height to astronomical horizon (height = 0)
		ref = exp(vdwref[0] + vdwref[1] * lnhgt + vdwref[2] * lnhgt * 
			lnhgt + vdwref[3] * lnhgt * lnhgt * lnhgt + vdwref[4] 
			* lnhgt * lnhgt * lnhgt * lnhgt + vdwref[5] * lnhgt * 
			lnhgt * lnhgt * lnhgt * lnhgt);
		//contribution to refraction from the change of the azimuthal coordinate, phi
		eps = exp(vdweps[0] + vdweps[1] * lnhgt + vdweps[2] * lnhgt * 
			lnhgt + vdweps[3] * lnhgt * lnhgt * lnhgt + vdweps[4] 
			* lnhgt * lnhgt * lnhgt * lnhgt);
		//now add the all the contributions together due to the observer's height 
		//along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0 
		//for this calculation, leave the refraction in units of mrad
L590:
		*a1 = 0.;
		refrac1 = vdwconst[0]; //vdW 288.15 degress K total atmospheric refraction from zero to infinity (mrad)
		*air = cd * 90. + (*vbweps * eps + *vdwrefr * ref + *vdwsf * refrac1) / 1e3;
		*a1 = (ref + refrac1) * 1e-3;  //convert to radians from mrad

	}
	else
	{
		//using Menat atomospheres
		if (*hgt > 0.)
		{
			lnhgt = log(*hgt * .001);
		}
		if (*lt >= 0. && *nweather == 3 && (*dy <= (double) *ns1 || *dy >=
			(double) *ns2) || *nweather == 1 || *lt < 0. && *nweather
			== 3 && (*dy > (double) *ns1 && *dy < (double) *ns2))
			{
	/*              winter refraction */
			if (*hgt <= 0.)
			{
				goto L690;
			}
				ref = exp(winref[4] + winref[5] * lnhgt + winref[6] * lnhgt *
					  lnhgt + winref[7] * lnhgt * lnhgt * lnhgt);
				eps = exp(winref[0] + winref[1] * lnhgt + winref[2] * lnhgt *
				lnhgt + winref[3] * lnhgt * lnhgt * lnhgt);
		L690:	*air = cd * 90. + (eps + ref + winrefo) / 1e3; //depression angle of sun at sunrise/sunset as seen from height = hgt
				*a1 = (ref + winrefo) / (cd * 1e3); //total atmospheric refraction for observer at height = hgt (degrees)
		}
		else if (*lt >= 0. && *nweather == 3 && (*dy > (double) *ns1 &&
			*dy < (double) *ns2) || *nweather == 0 || *lt < 0. &&
			*nweather == 3 && (*dy <= (double) *ns1 || *dy >= (
			double) *ns2))
		{
	/*              summer refraction */
			if (*hgt <= 0.)
			{
				goto L691;
			}
			ref = exp(sumref[4] + sumref[5] * lnhgt + sumref[6] * lnhgt *
				lnhgt + sumref[7] * lnhgt * lnhgt * lnhgt);
			eps = exp(sumref[0] + sumref[1] * lnhgt + sumref[2] * lnhgt *
				lnhgt + sumref[3] * lnhgt * lnhgt * lnhgt);
	L691:	*air = cd * 90. + (eps + ref + sumrefo) / 1e3; //depression angle of sun at sunrise/sunset as seen from height = hgt
			*a1 = (ref + sumrefo) / (cd * 1e3); //total atmospheric refraction for observer at height = hgt (degrees)
		}

	}

}

////////////////////InitWeatherRegions/////////////////////////////////////
void InitWeatherRegions( double *lg, double *lt,
						 short *nweather, short *ns1, short *ns2,
						 short *ns3, short *ns4, double *WinTemp,
						 short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[])
//////////////////////////////////////////////////////////////////////
//initializes weather regions and calculates longitude and latitude
//N.b., geo is a global variable, and is no longer passed to this routine as of 091120
///////////////////////////////////////////////////////////////////////
{

    double lt_w = 0.0; //define year dependent latitude variable
	short iTemp;
	short MaxMinTemp = -999;
	short ier = 0;
	short itk = 0;


    if (geo) //rest of the world uses geographic coordinates of longitude and latitude
	{
		//effect of global warming on weather zones from carbon soot
		//assume that global warming will take over the effect within the next 50 years
        //and then be optimistic and assume that it stabilizes

		if (GlobalWarmingCorrection == 1)
		{
			//use simple model for effect of Global Warming
			if (FirstSecularYr < END_SUN_APPROX)
			{
			   if (*lt >= (FirstSecularYr - 2000) * 0.07)
			   {
				   lt_w = *lt - (FirstSecularYr - 2000) * 0.07; //include effect of global warming from Robert J. Allen, UCR,
															  //"Carbon soot may be driving the expansion of the tropics-not CO2
			   }
			   else if (*lt < -(FirstSecularYr - 2000) * 0.07)
			   {
				   lt_w = *lt + (FirstSecularYr - 2000) * 0.07; //include effect of global warming from Robert J. Allen, UCR,
			   }
			   else
			   {
				   lt_w = 0.0;
			   }
			}
			else
			{
			   if (*lt >= 3.5)
			   {
				   lt_w = *lt - 3.5; //include effect of global warming from Robert J. Allen, UCR
			   }
			   else if (*lt < - 3.5)
			   {
				   lt_w = *lt + 3.5; //include effect of global warming from Robert J. Allen, UCR
			   }
			   else
			   {
				   lt_w = 0.0;
			   }
			}
		}
		else
		{
			lt_w = *lt;
		}

		if (!VDW_REF) {
		//---------------------------------------
			if (fabs(lt_w) < latzone[0])
			{
		/*		tropical region, so use tropical atmosphere which */
		/*		is approximated well by the summer mid-latitude */
		/*             atmosphere for the entire year */
				*nweather = 0;
		/*          define days of the year when winter and summer begin */
			}
			else if (fabs(lt_w) >= latzone[0] && fabs(lt_w) < latzone[1])
			{
				*ns1 = sumwinzone[0][0];
				*ns2 = sumwinzone[0][1];
				/*   no adhoc fixes  i.e., = 0,360 */
				*ns3 = sumwinzone[0][2];
				*ns4 = sumwinzone[0][3];
			}
			else if (fabs(lt_w) >= latzone[1] && fabs(lt_w) < latzone[2])
			{
		/*             weather similar to Eretz Israel */
				*ns1 = sumwinzone[1][0];
				*ns2 = sumwinzone[1][1];
		/*             Times for ad-hoc fixes to the visible and astr. sunrise */
		/*             (to fit observations of the winter netz in Neve Yaakov). */
		/*             This should affect  sunrise and sunset equally. */
		/*             However, sunset hasn't been observed, and since it would */
		/*             make the sunset times later, it's best not to add it to */
		/*             the sunset times as a chumrah. */
				*ns3 = sumwinzone[1][2];
				*ns4 = sumwinzone[1][3];
			} else if (fabs(lt_w) >= latzone[2] && fabs(lt_w) < latzone[3]) {
		/*             the winter lasts two months longer than in Eretz Israel */
				*ns1 = sumwinzone[2][0];
				*ns2 = sumwinzone[2][1];
				*ns3 = sumwinzone[2][2];
				*ns4 = sumwinzone[2][3];
			} else if (fabs(lt_w) >= latzone[4]) {
		/*             the winter lasts four months longer than in Eretz Israel */
				*ns1 = sumwinzone[3][0];
				*ns2 = sumwinzone[3][1];
				*ns3 = sumwinzone[3][2];
				*ns4 = sumwinzone[3][3];
			}
		}else if (VDW_REF) {

			if (ExtTemp[0] != -9999) { //using user inputed ground temperatures for all the calculations

				for (itk = 0; itk < 12; itk++) {
					MinTemp[itk] = ExtTemp[itk];
					AvgTemp[itk] = ExtTemp[itk];
					MaxTemp[itk] = ExtTemp[itk];
				}

			}

			for (iTemp = 0; iTemp < 12; iTemp++) {
				if (MaxMinTemp < MinTemp[iTemp] ) MaxMinTemp = MinTemp[iTemp];
			}

			//WinTemp deter__mins range of witner daynumbers that are likely to have inversion layers
			*WinTemp = WinterPercent * MaxMinTemp; //winter is defined as months whose mean min temperature less than WinterPercent of maximium min temp of the year

		}
	}
	else
	{  //Eretz Yisroel

		if (!VDW_REF) {

			*ns1 = sumwinzone[4][0];
			*ns2 = sumwinzone[4][1];
			*ns3 = sumwinzone[4][2];
			*ns4 = sumwinzone[4][3];

		}else if (VDW_REF) {

			if (ExtTemp[0] != -9999) { //using user inputed ground temperatures for all the calculations

				for (itk = 0; itk < 12; itk++) {
					MinTemp[itk] = ExtTemp[itk];
					AvgTemp[itk] = ExtTemp[itk];
					MaxTemp[itk] = ExtTemp[itk];
				}

			} 

			for (iTemp = 0; iTemp < 12; iTemp++) {
				if (MaxMinTemp < MinTemp[iTemp] ) MaxMinTemp = MinTemp[iTemp];
			}

			//WinTemp deter__mins range of winter daynumbers that are likely to have inversion layers
			*WinTemp = WinterPercent * MaxMinTemp; //winter is defined as months whose mean min temperature less than WinterPercent of maximium min temp of the year

		}
	}
}

///////////////////////////VisRiseSet////////////////////////////////////
short VisRiseSet(short *nloop, short *nlpend, short *ndirec, double *dy, double *t3,
				   double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
                   double *es, double *ob, double *lr, short *nsetflag, short *maxang,
				   double *stepsec, short *nyrstp, short *niter, short *ndystp,
				   short *nfil, double *sr1, double *p, double *t, double *hgt,
				   short *nweather, short *ns1, short *ns2, double *a1,
				   double *jdn, short *nyear, short *mday, short *mon, int * nyl, double *td)
/////////////////////////////////////////////////////////////////////////////////
//This routine calculates the visible times by moving the sun as follows:
//sunrise: The sun is moved in altitude and azimuth in stepsec seconds until
//         the upper limb of the sun just appears above the horizon.
//         The function starts at nloop * stepsec seconds above the astronomical
//         horizon and searches for nlpend steps.  If these parameters need
//         to be adjusted, then they are tweaked during the process of finding
//         the first sunrise.
//sunset:  The sun is moved backward in time by stepsec seconds until the upper
//         limb of the sun just appears on the horizon.  The functions starts at
//         nloop*stepsec seconds below the astronomical horizon and continues until
//         a maximum of nlpend steps.  If this needs to be adjusted, then it is
//         tweaked during the proces of finding the first sunset.
//////////////////////////////////////////////////////////////////////////////////
{
	short nn = 0;
	short nloops;
	short nfound = 0;
	short pt, nac, ne, nacurr;
	short nabove = 0;
	short i__2 = *nlpend;
	short i__3 = *ndirec;
	double dy1, hran, d__, z__, alt1, azi1, azi2, azi, al1, al1o;
    double t3sub, dobsy, elevintx, elevinty, dobss, elevint, refrac1;
	double hrs, ra;

L_692:

	for (nloops = *nloop; i__3 < 0 ? nloops >= i__2 : nloops <= i__2; nloops += i__3)
	{
		++nn;
		hran = *sr1 - *stepsec * nloops / 3600.;
/*      step in hour angle by stepsec sec */
		dy1 = *dy - *ndirec * hran / 24.;
		decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
		if (*nyear >= BEG_SUN_APPROX && *nyear <= END_SUN_APPROX)
		{
			dy1 = *dy - *ndirec * hran / 24.;
			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
		}
		else
		{
			hrs = 12.0 - *ndirec * hran;
			d__ = Decl3(jdn, &hrs, dy, td, mday, mon, nyear, nyl,
						ms, aas, ob, &ra, es, 1);
		}

/*      recalculate declination at each new hour angle */
		z__ = sin(*lr) * sin(d__) + cos(*lr) * cos(d__) * cos(hran / ch);
		z__ = acos(z__);
		alt1 = pi * .5 - z__;
/*      altitude in radians of sun's center */
		azi1 = cos(d__) * sin(hran / ch);
		azi2 = -cos(*lr) * sin(d__) + sin(*lr) * cos(d__) * cos(hran / ch);
		azi = atan(azi1 / azi2) / cd;
		if (azi > 0.)
		{
			azi1 = 90. - azi;
		}
		if (azi < 0.)
		{
			azi1 = -(azi + 90.);
		}
		if (fabs(azi1) > *maxang)
		{
			//azimuthal range insufficient
			return -1; //goto L9050;
		}

		azi1 = *ndirec * azi1;

		if (*nsetflag == 1) //<---
		{
/*          to first order, the top and bottom of the sun have the */
/*          following altitudes near the sunset */
			al1 = alt1 / cd + *a1;
/*          first guess for apparent top of sun */
			pt = 1;
			refract_(&al1, p, t, hgt, nweather, dy, &refrac1, ns1, ns2);
/*          first iteration to find position of upper limb of sun */
			al1 = alt1 / cd + .2666f + pt * refrac1;
/*          top of sun with refraction */
			if (*niter == 1)
			{
				for (nac = 1; nac <= 10; ++nac)
				{
					al1o = al1;
					refract_(&al1, p, t, hgt, nweather, dy, &refrac1, ns1, ns2);
/*                  subsequent iterations */
					al1 = alt1 / cd + .2666f + pt * refrac1;
/*                  top of sun with refraction */
					if (fabs(al1 - al1o) < .004f) break;
/* L693: */
				}
			}
//L695:
/*          now search digital elevation data */
			ne = (int) ((azi1 + *maxang) * 10);
			elevinty = elev[ne + 1].vert[1] - elev[ne].vert[1];
			elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
			elevint = elevinty / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[1];
			if ((elevint >= al1 || nloops <= 0) && nabove == 1)
			{
/*              reached sunset */
				nfound = 1;
/*              margin of error */
				nacurr = 0;
				t3sub = *t3 + *ndirec * (*stepsec * nloops + nacurr) / 3600.;
				if (nloops < 0)
				{
	/*              visible sunset can't be latter then astronomical sunset! */
					t3sub = *t3;
				}
/*                 <<<<<<<<<distance to obstruction>>>>>> */
				dobsy = elev[ne + 1].vert[2] - elev[ne].vert[2];
				elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
				dobss = dobsy / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
/*              record visual sunset results for time,azimuth,dist. to obstructions */

            	/////only record times if they are earlier (sunrise) or latter (sunset)
				if (*nfil == 0) //first set of times////////////////////////////////////
				{
     				tims[*nyrstp - 1][*nsetflag][*ndystp - 1] = t3sub;
					azils[*nyrstp - 1][*nsetflag][*ndystp - 1] = azi1;
					dobs[*nyrstp - 1][*nsetflag][*ndystp - 1] = dobss;
				}
				else
				{
					if (t3sub > tims[*nyrstp - 1][*nsetflag][*ndystp - 1] )
					{
						tims[*nyrstp - 1][*nsetflag][*ndystp - 1] = t3sub;
						azils[*nyrstp - 1][*nsetflag][*ndystp - 1] = azi1;
						dobs[*nyrstp - 1][*nsetflag][*ndystp - 1] = dobss;
					}
				}

			/////////////////////////////////////////////////////////////

/*              finished for this day, loop to next day */
			//goto L925;
			break;
			}
			else if (elevint < al1)
			{
				nabove = 1;
			}
		}
		else if (*nsetflag == 0)
		{
			al1 = alt1 / cd + *a1;
/*              first guess for apparent top of sun */
			pt = 1;
			refract_(&al1, p, t, hgt, nweather, dy, &refrac1, ns1, ns2);
/*              first iteration to find top of sun */
			al1 = alt1 / cd + .2666f + pt * refrac1;
/*              top of sun with refraction */
			if (*niter == 1)
			{
				for (nac = 1; nac <= 10; ++nac)
				{
					al1o = al1;
					refract_(&al1, p, t, hgt, nweather, dy, &refrac1, ns1, ns2);
/*                      subsequent iterations */
					al1 = alt1 / cd + .2666f + pt * refrac1;
/*                      top of sun with refraction */
					if (fabs(al1 - al1o) < .004f) break;
/* L750: */
				}
			}
//L760:
			alt1 = al1;
/*          now search digital elevation data */
			ne = (int) ((azi1 + *maxang) * 10);
			elevinty = elev[ne + 1].vert[1] - elev[ne].vert[1];
			elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
			elevint = elevinty / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[1];

			if (elevint <= alt1)
			{
				if (nn > 1)
				{
					nfound = 1; //found visual sunrise

/*                      <<<<<<<<<distance to obstruction>>>>>> */
					dobsy = elev[ne + 1].vert[2] - elev[ne].vert[2];
					elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
					dobss = dobsy / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
/*                      sunrise */
					nacurr = 0;
/*                      margin of error (seconds) to take into */
/*                      account differences between observations and the DTM `` */
					t3sub = *t3 + (*stepsec * nloops + nacurr) / 3600.;
					if (nloops < 0)
					{
/*                         sunrise can't be before astronomical sunrise */
/*                         so set it equal to astronomical sunrise time */
/*                         minus the adhoc winter fix if any */
	                   t3sub = *t3;
					}

/*                      record visual sunrise results for time,azimuth,dist. to obstructions */
					if (*nfil == 0) //first set of times////////////////////////////////////
					{
     					tims[*nyrstp - 1][*nsetflag][*ndystp - 1] = t3sub;
						azils[*nyrstp - 1][*nsetflag][*ndystp - 1] = azi1;
						dobs[*nyrstp - 1][*nsetflag][*ndystp - 1] = dobss;
					}
					else
					{
						if (t3sub < tims[*nyrstp - 1][*nsetflag][*ndystp - 1] )
						{
							tims[*nyrstp - 1][*nsetflag][*ndystp - 1] = t3sub;
							azils[*nyrstp - 1][*nsetflag][*ndystp - 1] = azi1;
							dobs[*nyrstp - 1][*nsetflag][*ndystp - 1] = dobss;
						}
					}
					break; //(goto L900) finished for this day, loop to next day
				}
				else
				{
					break; //(goto L900) reposition the sun further below the horizon (for sunrises)
						   //or further above the horizon (for sunsets)
						   //and restart the search from the new beginning
				}
			;
			}
		}
/* L850: */
	}

	if (nfound == 0)
	//no visible sunrise/sunest was found
	//need to reposition the sun further below the horizon (for sunrises)
    //or further above the horizon (for sunsets)
    //and restart the search from the new beginning
	{
		if (*nloop > 0 && *nsetflag == 1)
		{
			if (*nsetflag == 0)
			{
				*nloop = *nloop - 24;
				if (*nloop < 0)
				{
				*nloop = 0;
				}
			}
			else if (*nsetflag == 1)
			{
				*nloop = *nloop + 24;
			}
//			_printf("goto\n");
			goto L_692;
		}
		else if (*nloop >= -100 && *nsetflag == 0)
		{
		   if (*nsetflag == 0)
		   {
			  *nloop = *nloop - 24;
/*            IF (nloop .LT. 0) nloop = 0 */
		   }
		   else if (*nsetflag == 1)
		   {
			  *nloop = *nloop + 24;
		   }
		   goto L_692;
		}
		else if (*nloop < -100 && *nsetflag == 0)
		{
		   *nlpend = *nlpend + 500;
		   *nloop = 72;
		   goto L_692;
		}
	}

////////////////////////////////////////////////////////////////////////////
	//now tweak the nloop parameters for the next day based
	//on the experience of the present day
	if (*nsetflag == 1 )
	{
		if (nloops + 30 < *nloop || nloops + 30 > *nloop)
		{
			*nloop = nloops + 30;
		}
	}
	else if (*nsetflag == 0)
	{
		if (nloops + 30 < *nloop)
		{
			*nloop = nloops + 30;
		}
		else if (nloops + 30 > *nloop)
		{
		*nloop = nloops - 30;
		}
	}
/////////////////////////////////////////////////////////////////////////

	return 0;
}


/////////////////////ReadProfile//////////////////////////
short ReadProfile( char *fileon, double *hgt, double *p, short *maxang, double *meantemp )

////////////////////////////////////////////////////////////
//reads the contents of a profile file and stores the values
//in the elev arrays
////////////////////////////////////////////////////////////
{
	double tmp1,tmp2,tmp3,tmp4;
	short nazn, nazimuth,ntrcalc;
	double deltd, distd, defm, defb, avref;
	double trrbasis, d__1, d__2, d__3;
	double exponent,pathlength;
	char Tgroundch[7] = "";
	double Tground = 288.15;

	short const MAXARRSIZE = 6; //number of data in line = array size of strarr
	char strarr[MAXARRSIZE][MAXDBLSIZE]; //array to contain the data elements

	FILE *stream;

//	try opening file with the name as listed in the "bat" file as well
//	as well as all uppercases and all lowercaes
	short fillen = strlen(fileon);
    	if ( !(stream = fopen( fileon, "r" )) )
    	{
	     strlwr(fileon + fillen - 12); //try with all lowercases
    	     if ( !(stream = fopen( fileon, "r" )) )
    	     {
    	         strupr(fileon + fillen - 12); //try with all uppercase
    	         if ( !(stream = fopen( fileon, "r" )) )
    	         {
    	         	return -2; //can't open profile file'
    	         }
    	     }
    	}

	doclin[0] = 0; //clear the buffer
	fgets_CR( doclin, 255, stream ); //read first header line of text mostly containly text, but sometimes the mean temperataure
									 //used for calculating the terrestrial refraction

	if (VDW_REF) {
		//deter__min version of profile file in order to remove the terrestrial refraction
		if (InStr(doclin, "APPRNR")) {
			if (InStr(doclin, ",0")) {
				//new format w/o TR calculated, no need to renormalize the view angles
				ntrcalc = 0;
			} else if (InStr(doclin, ",1")) {
				//average temp TR contrib added to view angles, so remove it from view angles
				//this assumes that the change in the horizon profile will be small when changing the temperature
				ntrcalc = 1;
			}else{
				//old profile file with old TR calculation, so remove it from the view angles
				//Note: assumption that the horizon profile will not change is more iffy in this case */
				ntrcalc = 2;
			}
		} else {
			//new TR added, but Tground is declared, so extract it from the header string and use it as the meantemp
			ntrcalc = 3;
			Mid( doclin,41,6, Tgroundch);
			Tground = atof(Tgroundch);
		}
	}

	doclin[0] = 0; //clear the buffer
	fgets_CR( doclin, 255, stream ); //read second line containing details of scan as one line of text
	//parse it to obtain height recorded in profile file which is more reliable than the bat file
	if (!ParseString( doclin, ",", strarr, MAXARRSIZE ) )
	{
		*hgt = atof( &strarr[2][0] ); //hgt is the third data ןאקצ in this doc line separated by commas
	}
	else
	{
		return -1; //can't read data
	}


	fgets_CR( doclin, 255, stream ); //read first line of data to deter__min angle range
	short doclen = strlen(doclin);
	//doclin[doclen - 1] = 0; //remove carriage return

	if (!ParseString( doclin, ",", strarr, MAXARRSIZE ) )
	{
		tmp1 = atof( &strarr[0][0] ); //azimuth
		tmp2 = atof( &strarr[1][0] ); //view angle
		tmp3 = atof( &strarr[4][0] ); //distance (km) to obstruction
		tmp4 = atof( &strarr[5][0] ); //height of obstruction (m)
	}
	else
	{
		return -1; //can't read data
	}

	//*maxang = (short)fabs(atof( &strarr[0][0]));
	*maxang = (short)fabs(tmp1);
	if (*maxang > MAXANGRANGE) return -3; //array size too big


	//now remove the terrestrial refraction value of this first line of data
	//terrestrial refraction needs to be removed since it was only a temporary
	//value calculated when calculating the profile
	///////////////////////////////////////////////////////////////////////////////

	if (VDW_REF) {
		if (ntrcalc == 2) {
/* 		   remove old terrestrial refraction values from the view angle */
			deltd = *hgt - tmp4;
			distd = tmp3;
			if (deltd <= 0.f) {
				defm = 7.82e-4 - deltd * 3.11e-7;
				defb = deltd * 3.4e-5 - .0141;
			} else if (deltd > 0.f) {
				defm = deltd * 3.09e-7 + 7.64e-4;
				defb = -.00915 - deltd * 2.69e-5;
			}
			avref = defm * distd + defb;
			if (avref < 0.f) {
				avref = 0.f;
			}
			tmp2 -= avref;
		} else if (ntrcalc == 1 || ntrcalc == 3) {

			deltd = *hgt - tmp4;
			distd = tmp3;
			//calculate conserved portion of Wikipedia's terrestrial refraction expression 
			//Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere
			//so have p * 8.15 * 1000 * (0.0343 - 0.0065)/(T * T * 3600) degrees

			if (ntrcalc == 1) {
				Tground = *meantemp;
			} //else for ntrcalc == 3, Tground was extracted above from the header string

			trrbasis = *p * 8.15f * 1e3f * .0277f / (Tground * Tground * 3600);
			//in the winter, dt/dh will be positive for inversion layer and refraction is increased
			//now calculate pathlength of curved ray using Lehn's parabolic approximation 
			d__1 = distd;
			d__3 = distd;
			d__2 = abs(deltd) * .001f - d__3 * d__3 * .5f / rearth;
			pathlength = sqrt(d__1 * d__1 + d__2 * d__2);
			//now add terrestrial refraction based on modified Wikipedia expression */
			if (abs(deltd) > 1e3f) {
				exponent = .99f;
			} else {
				exponent = .9965f;
			}
			tmp2 -= trrbasis * pow(pathlength, exponent);
		}
	}
	///////////////////////////////////////////////////////////////////

	//stuff the array according to the azimuth value
    nazimuth = (short)floor(( atof( &strarr[0][0] ) + *maxang) * 10.0);
	elev[nazimuth].vert[0] = tmp1; //azimuth
	elev[nazimuth].vert[1] = tmp2; //view angle
	elev[nazimuth].vert[2] = tmp3; //distance (km) to obstruction
	elev[nazimuth].vert[3] = tmp4; //height (m) of obstruction

	//now process the rest of the data lines in the profile file

	while ( !feof(stream) )
	{

		for (nazn = 1; nazn < 20 * *maxang + 1; nazn++)
		{
			fgets_CR( doclin, 255, stream ); //read first line of data
			if (!ParseString( doclin, ",", strarr, MAXARRSIZE ) )
			{
				tmp1 = atof( &strarr[0][0] ); //azimuth
				tmp2 = atof( &strarr[1][0] ); //view angle
				tmp3 = atof( &strarr[4][0] ); //distance (km) to obstruction
				tmp4 = atof( &strarr[5][0] ); //height of obstruction (m)

                /* //old code -- for reference only
				//stuff the array according to the azimuth value
                nazimuth = (short)floor(( atof( &strarr[0][0] ) + *maxang) * 10.0);
				elev[nazimuth].vert[0] = atof( &strarr[0][0] ); //azimuth
				elev[nazimuth].vert[1] = atof( &strarr[1][0] ); //view angle
				elev[nazimuth].vert[2] = atof( &strarr[4][0] ); //distance (km) to obstruction
				elev[nazimuth].vert[3] = atof( &strarr[5][0] ); //height (m) of obstruction
				*/

				//now deal with terrestrial refraction
				///////////////////////////////////////////////////////////////////////////////

				if (VDW_REF) {
					if (ntrcalc == 2) {
			/* 		   remove old terrestrial refraction values from the view angle */
						deltd = *hgt - tmp4;
						distd = tmp3;
						if (deltd <= 0.f) {
							defm = 7.82e-4 - deltd * 3.11e-7;
							defb = deltd * 3.4e-5 - .0141;
						} else if (deltd > 0.f) {
							defm = deltd * 3.09e-7 + 7.64e-4;
							defb = -.00915 - deltd * 2.69e-5;
						}
						avref = defm * distd + defb;
						if (avref < 0.f) {
							avref = 0.f;
						}
						tmp2 -= avref;
					} else if (ntrcalc == 1 || ntrcalc == 3) {

						deltd = *hgt - tmp4;
						distd = tmp3;
						//calculate conserved portion of Wikipedia's terrestrial refraction expression 
						//Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere
						//so have p * 8.15 * 1000 * (0.0343 - 0.0065)/(T * T * 3600) degrees

						if (ntrcalc == 1) {
							Tground = *meantemp;
						} //else for ntrcalc == 3, Tground was extracted above from the header string

						trrbasis = *p * 8.15f * 1e3f * .0277f / (Tground * Tground * 3600);

						//in the winter, dt/dh will be positive for inversion layer and refraction is increased
						//now calculate pathlength of curved ray using Lehn's parabolic approximation 
						d__1 = distd;
						d__3 = distd;
						d__2 = abs(deltd) * .001f - d__3 * d__3 * .5f / rearth;
						pathlength = sqrt(d__1 * d__1 + d__2 * d__2);
						//now add terrestrial refraction based on modified Wikipedia expression */
						if (abs(deltd) > 1e3f) {
							exponent = .99f;
						} else {
							exponent = .9965f;
						}
						tmp2 -= trrbasis * pow(pathlength, exponent);
					}
				}
				///////////////////////////////////////////////////////////////////

				//stuff the array according to the azimuth value
				nazimuth = (short)floor(( atof( &strarr[0][0] ) + *maxang) * 10.0);
				elev[nazimuth].vert[0] = tmp1; //azimuth
				elev[nazimuth].vert[1] = tmp2; //view angle
				elev[nazimuth].vert[2] = tmp3; //distance (km) to obstruction
				elev[nazimuth].vert[3] = tmp4; //height (m) of obstruction

			}
			else
			{
				return -1; //can't read data
			}
		}
		break;
	}

	fclose(stream);
	return 0;

}


///////////////////noon/////////////////////////////////////////////////////
void AstroNoon( double *jdn, double *td, double *dy1, double *mp, double *mc, double *ap,
		   double *ac, double *ms, double *aas, double *es, double *ob,
		   double *d__, double *tf, short *nyear, short *mday, short *mon, int *nyl, double *t6)
////////////////////////////////////////////////////////////////////////////
{
	double ra, df, et, hrs;

	//change other astronomical constants for the current day
	if (*nyear >= BEG_SUN_APPROX && *nyear <= END_SUN_APPROX)
	{
		//use formula valid to 6 sec for these 100 years only
		decl_(dy1, mp, mc, ap, ac, ms, aas, es, ob, d__);
		ra = atan(cos(*ob) * tan(*es));
	}
	else
	{
		//use general formula for declination valid to 6 sec for any year
		//Since this is first call to Decl3 for this day, so set mode = 0
		//in order to calculate jdn, mday, mon, nyl for all the
		//calls to Decl3 on this day
		hrs = 12.0;
		*d__ = Decl3(jdn, &hrs, dy1, td, mday, mon, nyear,
			         nyl, ms, aas, ob, &ra, es, 0);
	}

	if (ra < 0.)
	{
		ra += pi;
	}
	df = *ms - ra;
	//-----------------------------------------------------
	//
	while ( fabs(df) > pi * 0.5 )
	{
		if (df > 0.)
		{
			df -= df / fabs(df) * pi;
		}
		else if (df < 0.)
		{
			df += df / fabs(df) * pi;
		}
	}
	//-----------------------------------------------------
	et = df / cd * 4.;
	/*          equation of time (minutes) */
	*t6 = *tf + 720. - et;
	/*          astronomical noon in minutes */
}


///////////Astronomical/Mishor sunset///////////////////////////////////////////////////////////
void AstronSunset(short *ndirec, double *dy, double *lr, double *lt, double *t6,
			 double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
			 double *es, double *ob, double *d__, double *air, double *sr1,
			 short *nweather, short *ns3, short *ns4, bool *adhocset, double *t3,
			 bool *astro, double *azi1, double *jdn, short *mday, short *mon, 
			 short *nyear, int *nyl, double *td, short mint[], double *WinTemp,
			 bool *FindWinter, bool *EnableSunsetInv)
///////////////////////////////////////////////////////////////////////////////////////////////
{
	/*              calculate astron./mishor sunsets */
	/*              Nsetflag = 1 or Nsetflag = 3 OR Nsetflag = 5 */

	double fdy, dy1, sr, azi2, azio, hrs, ra;
	short nacurr = 0, kTemp, mTemp;
	double DayLength = 12.0;

	*ndirec = -1;

	fdy = acos(-tan(*d__) * tan(*lr)) * ch / 24.;
	dy1 = *dy - *ndirec * fdy;
	/*              calculate sunset/sunrise */
	if (*nyear >= BEG_SUN_APPROX && *nyear <= END_SUN_APPROX)
	{
		decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, d__);
	}
	else
	{
		hrs = 12.0 - *ndirec * fdy * 24.0;
		*d__ = Decl3(jdn, &hrs, dy, td, mday, mon, nyear, nyl, ms, aas, ob, &ra, es, 1);
	}
	//add solar semidiamter to depression angle, *air, and compensate for variable
	//semidiameter as a function of the anomaly to first approx (based on AA-1996 C24)
	*air = *air + (1. + cos(*aas) * .0167) * .2666 * cd;
	*sr1 = acos(-tan(*lr) * tan(*d__) + cos(*air) / (cos(*lr) * cos(*d__))) * ch;
	sr = *sr1 * hr;
	/*              hour angle in minutes */
	*t3 = (*t6 - *ndirec * sr) / hr;

	/* *************************BEGIN SUNSET WINTER ADHOC FIX************************ */
	if (!VDW_REF) {
		if (*lt >= 0. && *nweather == 3 && (*dy <= (double) *ns3 ||
			*dy >= (double) *ns4) && *adhocset) {
		/*                 add winter addhoc fix */
			nacurr = nacurrFix; //15;
		}
	} else {
		//new of version 18 -- deter__min portion of year prone to inversion layers using daylength
		nacurr = 0; //add the adhoc fix in the calling module if not calculating the visible sunset
		if (*adhocset) {
			if (*FindWinter) {
				DayLength = 2 * *sr1;
				//make sure that it is cold by checking miniimum temperature for the relevant month
				//use Meeus's forumula p. 66 to convert daynumber to month,
				//no need to interpolate between temepratures -- that is overkill
				kTemp = 2;
				if (*nyl == 366) kTemp = 1;
				mTemp = floor(9 * (kTemp + *dy) / 275 + 0.98);

				if (WinterPercent < 1) {
					if (DayLength <= MaxWinLen && mint[mTemp-1] <= *WinTemp) {
						*EnableSunsetInv = true;
					}
				} else { //WinTemp filtering not being used
					if (DayLength <= MaxWinLen) {
						*EnableSunsetInv = true;
					}
				}
			}else{
				//add winter adhoc fix without day length qualification
				if (*lt >= 0. && *nweather == 0 && (*dy <= (double) *ns3 
					|| *dy >= (double) *ns4) ) {
					*EnableSunsetInv = true;
				}
				if (*lt < 0. && *nweather == 0 && (*dy >= (double) *ns3 
					|| *dy <= (double) *ns4) ) {
					*EnableSunsetInv = true;
				}
			}
		}
	}
	/* *************************END SUNSET WINTER ADHOC FIX************************ */


	*t3 = *t3 + nacurr / 3600.;
	/*               t3sub = t3 */
	/*              astronomical sunrise/sunset in hours.frac hours */
	if (astro)
	{
		*azi1 = cos(*d__) * sin(*sr1 / ch);
		azi2 = -cos(*lr) * sin(*d__) + sin(*lr) * cos(*d__) * cos(*sr1/ ch);
		azio = atan(*azi1 / azi2);
		if (azio > 0.) {
		*azi1 = 90. - azio / cd;
		}
		if (azio < 0.) {
		*azi1 = -(azio / cd + 90.);
		}
	}
	//-----------------------------------------------------
}


///////////Astronomical/Mishor sunrise///////////////////////////////////////////////////////////
void AstronSunrise(short *ndirec, double *dy, double *lr, double *lt, double *t6,
			 double *mp, double *mc, double *ap, double *ac, double *ms, double *aas,
			 double *es, double *ob, double *d__, double *air, double *sr1,
			 short *nweather, short *ns3, short *ns4, bool *adhocrise, double *t3,
			 bool *astro, double *azi1, double *jdn, short *mday, short *mon,
			 short *nyear, int *nyl, double *td, short mint[], double *WinTemp,
			 bool *FindWinter, bool *EnableSunriseInv)
///////////////////////////////////////////////////////////////////////////////////////////////
{
	double fdy, dy1, sr2, sr, azi2, azio, tss, hrs, ra;
	double DayLength = 12.0;
	short nacurr = 0, kTemp, mTemp;

	*ndirec = 1;

	fdy = acos(-tan(*d__) * tan(*lr)) * ch / 24.;
	//first calculate sunset
	dy1 = *dy + fdy;
	if (*nyear >= BEG_SUN_APPROX && *nyear <= END_SUN_APPROX)
	{
		decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, d__);
	}
	else
	{
		hrs = 12.0 + *ndirec * fdy * 24.0;
		*d__ = Decl3(jdn, &hrs, dy, td, mday, mon, nyear, nyl, ms, aas, ob, &ra, es, 1);
	}
	//add solar semidiamter to depression angle, *air, and compensate for variable
	//semidiameter as a function of the anomaly to first approx (based on AA-1996 C24)
	*air = *air + (1. + cos(*aas) * .0167) * .2666 * cd;
	/*              calculate sunset */
	sr2 = acos(-tan(*lr) * tan(*d__) + cos(*air) / (cos(*lr) * cos(*d__))) * ch;
	sr2 *= hr;
	tss = (*t6 + sr2) / hr;
	/*              sunset in hours.fraction of hours */
	dy1 = *dy - fdy;
	/*              calculate sunrise */
	if (*nyear >= BEG_SUN_APPROX && *nyear <= END_SUN_APPROX)
	{
		decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, d__);
	}
	else
	{
		hrs = 12.0 - *ndirec * fdy * 24.0;
		*d__ = Decl3(jdn, &hrs, dy, td, mday, mon, nyear, nyl, ms, aas, ob, &ra, es, 1);
	}
	*sr1 = acos(-tan(*lr) * tan(*d__) + cos(*air) / (cos(*lr) * cos(*d__))) * ch;
	sr = *sr1 * hr;
	/*              hour angle in minutes */
	*t3 = (*t6 - sr) / hr;

	/* *************************BEGIN SUNRISE WINTER ADHOC FIX************************ */
	if (!VDW_REF) {
		if (*lt >= 0. && *nweather == 3 && (*dy <= (double) *ns3 ||
			*dy >= (double) *ns4) && *adhocrise) {
			//add winter adhoc fix for sunset
			nacurr = nacurrFix; //15;
		}
	} else {
		//new of version 18 -- deter__min portion of year prone to inversion layers using daylength
		nacurr = 0;  //only add fix in calling module if not calculating the visible sunset
		if (*adhocrise) {
			if (*FindWinter) {
				DayLength = 2 * *sr1;
				//make sure that it is cold by checking miniimum temperature for the relevant month
				//use Meeus's forumula p. 66 to convert daynumber to month,
				//no need to interpolate between temepratures -- that is overkill
				kTemp = 2;
				if (*nyl == 366) kTemp = 1;
				mTemp = floor(9 * (kTemp + *dy) / 275 + 0.98);

				if (WinterPercent < 1) {
					if (DayLength <= MaxWinLen && mint[mTemp-1] <= *WinTemp) {
						*EnableSunriseInv = true;
					}
				} else { //WinTemp filtering not being used
					if (DayLength <= MaxWinLen) {
						*EnableSunriseInv = true;
					}
				}

			}else{
				//add winter adhoc fix
				if (*lt >= 0. && *nweather == 0 && (*dy <= (double) *ns3 
					|| *dy >= (double) *ns4) ) {
					*EnableSunriseInv = true;
				}
				if (*lt < 0. && *nweather == 0 && (*dy >= (double) *ns3 
					|| *dy <= (double) *ns4) ) {
					*EnableSunriseInv = true;
				}
			}
		}
	}
	/* *************************END SUNSET WINTER ADHOC FIX************************ */

	*t3 = *t3 + nacurr / 3600.;
	/*               t3sub = t3 */
	/*              astronomical sunrise in hours.frac hours */
	/*              sr1= hour angle in hours at astronomical sunrise */
	if (*astro) {
		*azi1 = cos(*d__) * sin(*sr1 / ch);
		azi2 = -cos(*lr) * sin(*d__) + sin(*lr) * cos(*d__) * cos(*sr1/ ch);
		azio = atan(*azi1 / azi2);
		if (azio > 0.) {
		*azi1 = 90. - azio / cd;
		}
		if (azio < 0.) {
		*azi1 = -(azio / cd + 90.);
		}
	}
}

////////////////////PrintMultiColTable//////////////////////////////////
short PrintMultiColTable(char filzman[], short ExtTemp[] )
//////////////////////////////////////////////////////////////
//loads zemanim info into zemanim buffer then calculates
//and prints zemanim to output stream
//
//filzman is the name of the zemanim file (resides in directory dircities)
///////////////////////////////////////////////////////////////
{

	//static char doclin[255] = "";
	static char zmannamestmp[50][100];
	static char zmanstring[255] = "";
	static char TitleZman[255] = "";
//	char zmannamestmp[50][100];
	short typeread = 0;
	short pos1 = 0, pos2 = 0, numzman = 0, numsort = 0;
	double avekmxzman = 0, avekmyzman = 0, avehgtzman = 0;
	short MinTemp[12], AvgTemp[12], MaxTemp[12];
	double lt,lg;
	int N,E;

//	bool adhocset;

	FILE *stream;

	//initiate array that will contains sort numbers
	for (short i = 0; i < maxzemanim; i++)
	{
        sortzman[i] = -1;
		//-1 is flag that this array element is not to be printed
	}

	//open the zma file, and read the contents looking for "sedra/holidays" or "sort","order"



	if ( !( stream = fopen( filzman, "r")) )

	{
		ErrorHandler(12, 1);
		return -1; //zma file not found
	}
	else
	{
		while (!feof(stream) )
		{
			if (typeread == 0)
			{
				fgets_CR(doclin, 255, stream); //read in line of text

				if (InStr(doclin, "Title/Description"))
				{
					typeread = 1; //Title of tble
				}
				else if (InStr(doclin, "sedra/holidays" ))
				{
					//sedra info flag detected, so next line must contain sedra type
					typeread = 2;
				}
				else if (InStr(doclin, "sort"))
				{
					//sort order flag detected, so next lines until EOF contain sort order
					typeread = 3;
				}
				else if (strlen(Trim(doclin)) > 1 && typeread !=3 )
				{
					strncpy( zmanstring, doclin, strlen(doclin) - 1);
					zmanstring[strlen(doclin) - 1] = 0;

					//parse out the zemanim elements
					pos1 = 1;
					for (short j = 0; j < 8; j++)
					{
						//now parse zman components and enter them into
						//unsorted zman description array: zmannames

						pos2 = InStr(pos1 + 1, zmanstring, "," );
						if ( pos2 != 0 )
						{
							strncpy (&zmannames[numzman][j][0], &zmanstring[pos1], pos2 - pos1 - 1);
							Trim(Replace( &zmannames[numzman][j][0], '\"', ' ' )); //get rid of VB style quotation marks
							pos1 = pos2 + 1;
						}
						else //last parameter
						{
							strncpy (&zmannames[numzman][j][0], &zmanstring[pos1], strlen(doclin) - pos1 - 1);
							Trim(Replace( &zmannames[numzman][j][0], '\"', ' ' )); //get rid of VB style quotation marks
							}
					}

					numzman++;

				}
			}

			if (typeread == 1)
			{
				fgets_CR( doclin, 255, stream );
				strcpy( TitleZman, doclin );
				typeread = 0; //reset flag to find other parameters
			}
			else if (typeread == 2)
			{
				fgets_CR( doclin, 255, stream );
				parshiotflag = atoi(doclin);
				//  parshiotflag = 0 for no parshiot
				//               = 1 for EY parshiot
				//               = 2 for diaspora parshiot

				typeread = 0; //reset flag to find sort info if any

			}
			else if (typeread == 3)
			{
				doclin[0] = '\0'; //clear buffer
                fgets_CR( doclin, 255, stream);
				if (strlen(Trim(doclin)) > 0 )
				{
					sortzman[numsort] = atoi(doclin);
					numsort++;
                    //typeread = 2; //keep on reading sort numbers until EOF
				}
				else
				{
					break; //end of information reached, close template ".zma" file
				}

			}

		}

		fclose(stream);

		if (numsort == 0) //no sort info, so use the order that they appear in zma file
		{
			numsort = numzman;
			for (short j = 0; j < numzman; j++)
			{
				sortzman[j] = j;
			}
		}
	}


    //edit zemanin titles array with the zemanim titiles
   for (short j = 0; j < numzman; j++ )
   {
	   strcpy( &zmannamestmp[j][0], &zmannames[j][0][0] ); //copy to temporary array

	   //remove the "Zmanim", etc from the name
	   pos1 = InStr(&zmannamestmp[j][0], ":");
	   if (pos1 != 0)
	   {
		   Trim(Mid( &zmannamestmp[j][0], pos1 + 1, strlen(&zmannamestmp[j][0]) - pos1, doclin) );
		   sprintf( &zmannamestmp[j][0], "%s", doclin );
	   }
	   
	   /*
	   if (optionheb) // Then 'convert zemanim labels into hebrew
	   {
		  if (InStr(strlwr(&zmannamestmp[j][0]), "mishor sunrise") )
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[18][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "mishor sunset") )
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[21][0] );
		  }
		  else if ( InStr(strlwr(&zmannamestmp[j][0]), "astronomical sunrise") )
		  {
             sprintf( &zmannamestmp[j][0], "%s", &heb3[17][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "astronomical sunset") )
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[20][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "chazos") )
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[28][0] );
		  }
	   }
	   */

	   if (optionheb) // Then 'convert zemanim labels into hebrew
	   {
		  if (InStr(strlwr(&zmannamestmp[j][0]), "mishor") && InStr(strlwr(&zmannamestmp[j][0]), "sunrise"))
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[18][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "mishor") && InStr(strlwr(&zmannamestmp[j][0]), "sunset"))
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[21][0] );
		  }
		  else if ( InStr(strlwr(&zmannamestmp[j][0]), "astronomical") && InStr(strlwr(&zmannamestmp[j][0]), "sunrise"))
		  {
             sprintf( &zmannamestmp[j][0], "%s", &heb3[17][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "astronomical") && InStr(strlwr(&zmannamestmp[j][0]), "sunset"))
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[20][0] );
		  }
		  else if (InStr(strlwr(&zmannamestmp[j][0]), "chazos") )
		  {
			 sprintf( &zmannamestmp[j][0], "%s", &heb3[28][0] );
		  }
	   }	   
   }

   //now sort the zmanim titles to appear in the order of increasing shaos zemanios
   for (short k = 0; k < numsort; k++ )
   {
   	   sprintf( &zmantitles[k][0], "%s", &zmannamestmp[sortzman[k]][0] );
   }


    //calculate coordinates for zemanim calculation////////////////////////////////

	if (avekmx[0] != 0 && avekmx[1] == 0)
	{
		avekmxzman = avekmx[0];
		avekmyzman = avekmy[0];
		avehgtzman = avehgt[0];
	}
    else if (avekmx[0] == 0 && avekmx[1] != 0)
	{
		avekmxzman = avekmx[1];
		avekmyzman = avekmy[1];
		avehgtzman = avehgt[1];
	}
    else if (avekmx[0] != 0 && avekmx[1] != 0)
	{
		avekmxzman = 0.5 * (avekmx[0] + avekmx[1]);
		avekmyzman = 0.5 * (avekmy[0] + avekmy[1]);
		avehgtzman = 0.5 * (avehgt[0] + avehgt[1]);
	}

	//now calculate and print simultaneous sunrise and sunset tables
	//and/or the requested zemanim

	if (!geo) 
	{
		//EY, need to convert ITM to WGS84 
		N = (int)(avekmyzman * 1000 + 1000000); 
		E = (int)(avekmxzman * 1000);
		//use J. Gray's conversion from ITM to WGS84
		//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
		ics2wgs84(N, E, &lt, &lg); 
		lg = -lg; //convert this program's coordinate system where East longitude is positive
	} 
	else 
	{

		lt = avekmyzman;
		lg = avekmxzman;
	}

	short ier = Temperatures(lt, -lg, MinTemp, AvgTemp, MaxTemp);

	if (ier) //error detected
	{
		ErrorHandler(17, ier);
		return -1;
	}

	if ( WriteTables(TitleZman, numsort, numzman,
		             &lg, &lt, &avehgtzman,
					 MinTemp, AvgTemp, MaxTemp, ExtTemp ) )
	{
		//error in parameters detected
		ErrorHandler(13, 1);
		return -1;
	}

	return 0;
}


///////////////////////cal/////////////////////////////////////////////////////////
short cal(double *ZA, double *air, double *lr, double *tf, double *dyo, double *hgto,
		  short *yro, double *td, short *yr, double *dy, double *mp, double *mc,
		  double *ap, double *ac, double *ms,double *aas, double *es, double *ob,
		  double *fdy, double *ec, double *e2c, short *ns1,short *ns2, short *ns3,short *ns4,
		  short *nweather, short *PartOfDay, 
		  double *avekmxzman, double *avekmyzman, double *avehgtzman,
		  double *t6, double *t3, double *t3sub, double *t3sub99,
		  double *jdn, int *nyl, short *mday, short *mon, 
		  short mint[], short avgt[], short maxt[], short ExtTemp[])
////////////////////////////////////////////////////////////////////////////////////
{

	double dy1, ra, df, d__, et, air2, sr, sr1, sr2, ZA1;
	double a1;
	short nyr, nyd, nyf;
	double hrs;
	double WinTemp = 0.0;
	double vdwsf = 1.0, vdwrefr = 1.0, vbweps = 1.0;

	if (*yr != *yro) //initialize
	{
		*yro = *yr;

		nyr = *yr;

		//initialize weather constants and coordinates
		if (!Israel_neigh) //invert coordinate due to nonstandard convention
		{
			InitWeatherRegions( avekmxzman, avekmyzman, nweather,
								ns1, ns2, ns3, ns4, &WinTemp, mint, avgt, maxt, ExtTemp ); //<<<<<<<<<<check this for coordinates
		}
		else
		{
			InitWeatherRegions( avekmyzman, avekmxzman, nweather,
								ns1, ns2, ns3, ns4, &WinTemp, mint, avgt, maxt, ExtTemp ); //<<<<<<<<<<check this for coordinates
		}

		//---------------------------------------
		*td = (double)geotz;
		/*       time difference from Greenwich England */
		//---------------------------------------

		*lr = cd * *avekmyzman;;
		*tf = *td * 60. + *avekmxzman * 4.;

		if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
		{
			nyd = nyr - RefYear;

    		// ----initialize astronomical constants to time zone and year---------------------------
			InitAstConst(&nyr, &nyd, ac, mc, ec, e2c, ap, mp, ob, &nyf, td);
			//---------------------------------------------------------------------------------------
		}

	}
	else
	{
		*lr = cd * *avekmyzman;
		*tf = *td * 60. + *avekmxzman * 4.;
		nyr = *yr;
	}


	if (*dy != *dyo  || *avehgtzman != *hgto ) //check once per day
	{

		*dyo = *dy;
		*hgto = *avehgtzman;

		//initialize astronomical constants once a day

		///////////Astronomical total refraction//////////////////////////////////
		AstronRefract( avehgtzman, avekmyzman, nweather, PartOfDay, dy, ns1, ns2,
			air, &a1, mint, avgt, maxt, &vdwsf, &vdwrefr, &vbweps, &nyr);
        ///////////////////////////////////////////////////////////////////

		//change from noon of previous day to noon of today
		if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
		{
			dy1 = *dy;
			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			ra = atan(cos(*ob) * tan(*es));
		}
		else
		{
			//use general formula for declination valid to 6 sec for any year
			//Since this is first call to Decl3 for this day, so set mode = 0
			//in order to calculate jdn, mday, mon, nyl for all the
			//calls to Decl3 on this day
			hrs = 12.0;
			d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
						 nyl, ms, aas, ob, &ra, es, 0);
		}

        //**************check for polar circles**************
		if ( (pi - Sgn(*lr) * (d__ + *lr) <= *air + 0.002) || (fabs(*lr - d__) >= pi * 0.5 - 0.02) )
		{
           //sun doesn't set that day
           *t3sub = -9999;
           *t3sub99 = -1;
           return -2;
		}
        //***************************************************

		if (ra < 0) ra += pi;
		df = *ms - ra;
		while (fabs(df) > pi * 0.5 )
		{
			df -= Sgn(df) * pi;
		}
		et = df / cd * 4;
		*t6 = 720 + *tf - et;

		if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
		{
			dy1 = *dy;
			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
		}
		//not sure why ^ is recalculated, so don't repeat for other years

		if ( fabs( -tan(d__) * tan(*lr) ) > 1.0 )
		{
           //sun doesn't set that day
           *t3sub = -9999;
           *t3sub99 = -1;
           return -2;
		}

		//calculate sunrise
        *fdy = acos(-tan(d__) * tan(*lr)) * ch / 24;
		if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
		{
			dy1 = *dy - *fdy;
			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
		}
		else
		{
			//second call to Decl3, so mode = 1
			hrs = 12.0 - *fdy * 24.0;
			d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
						 nyl, ms, aas, ob, &ra, es, 1);
		}

	    *air = *air + (1. + cos(*aas) * .0167) * .2666 * cd;
		if ( fabs(((-tan(*lr) * tan(d__)) + (cos(*air) / cos(*lr) / cos(d__)))) > 1.0 )
		{
			//sun doesn't rise that day
			*t3sub = -9999;
			*t3sub99 = -1;
			return -2;
		}

		//higher accuracy sunrise
		sr1 = acos((-tan(*lr) * tan(d__)) + (cos(*air) / cos(*lr) / cos(d__))) * ch;
		sr = sr1 * hr;
		*t3 = *t6 - sr;

		//convert times to fractions of hours
        *t3 = *t3 / hr;
		*t6 = *t6 / hr;

	}
   else
	{
      if (*t3sub99 == -1)
		{
			*t3sub = -9999; //no zemanim for this day since sun doesn't rise/set
			return -1;
		}
	}


   if ( *ZA < -90 ) //Then 'dawns
	{
		if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
		{
			dy1 = *dy - *fdy; //hour angle at sunrise <- first approx
			//approximate declination
			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
		}
		else
		{
			hrs = 12.0 - *fdy * 24.0; //hour angle at sunrise <--first approx.
			d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
						 nyl, ms, aas, ob, &ra, es, 1);
		}

		ZA1 = fabs(*ZA);
		air2 = ZA1 * cd; // 'zenith angle of dawn in radians
		//********************check for no dawn*************
		if ( air2 > pi - Sgn(*lr) * (d__ + *lr) || air2 < fabs(*lr - d__ ) )
		   {
				//no dawn
				*t3sub = -9999;
				return -1;
		   }
           //'**************************************************
         if (fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
		   {
				//no dawn
				*t3sub = -9999;
				return -1;
		   }

			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;

			//now calculate declination to higher precision
			if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
			{
				dy1 = *dy - sr2 / 24.0;
    			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			}
			else
			{
				hrs = 12.0 - sr2; //hour angle at sunrise <--first approx.
				d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
							 nyl, ms, aas, ob, &ra, es, 1);
			}

			if ( fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
			{
				//no dawn
				*t3sub = -9999;
				return -1;
			}

			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;
			*t3sub = *t6 - sr2; //in fractions of hours
		}
		else if ( *ZA == -90 ) // Then astronomical or mishor sunrise
		{
			*t3sub = *t3;
		}
		else if ( *ZA == 0 ) // Then 'true noon
		{
			*t3sub = *t6;
		}
       else if ( *ZA == 90 ) // Then astronomical or mishor sunset
		{
			if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX) 
			{
				dy1 = *dy + *fdy;
				//approximate declination
    			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			}
			else
			{
				hrs = 12.0 + *fdy * 24.0; //hour angle at sunset <--first approx.
				d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
							 nyl, ms, aas, ob, &ra, es, 1);
			}

			air2 = *air; //'zenith angle
			//********************check for no sunset (polar circle)*************
			if ( air2 > pi - Sgn(*lr) * (d__ + *lr) || air2 < fabs(*lr - d__) )
		   {
				//no sunset
				*t3sub = -9999;
				return -1;
		   }
			//'**************************************************
			if ( fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
		   {
				//no sunset
				*t3sub = -9999;
				return -1;
			}
			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;

			//calculate declination to higher precision
			if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
			{
				dy1 = *dy + sr2 / 24.0;
    			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			}
			else
			{
				hrs = 12.0 + sr2; //hour angle at sunset <--2nd iteration
				d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
							 nyl, ms, aas, ob, &ra, es, 1);
			}

			if ( fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
		   {
				//no dawn
				*t3sub = -9999;
				return -1;
		   }
			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;
			*t3sub = *t6 + sr2; //in fractions of hours
		}
		else if ( *ZA > 90 ) // Then 'twilights
		{
			//first calculate approximate declination
			if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
			{
				dy1 = *dy + *fdy;
				decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			}
			else
			{
				hrs = 12.0 + *fdy * 24.0; //hour angle at sunset <--1st iteration
				d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
							 nyl, ms, aas, ob, &ra, es, 1);
			}

			air2 = *ZA * cd; //zenith angle
			//'********************check for no twilight*************
			if ( air2 > pi - Sgn(*lr) * (d__ + *lr) || air2 < fabs(*lr - d__) )
			{
				//no twilight
				*t3sub = -9999;
				return -1;
			}
			//'**************************************************
			if ( fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
			{
				//no twilight
				*t3sub = -9999;
				return -1;
			}
			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;
			//'calculate declination to higher accuracy
			if (nyr >= BEG_SUN_APPROX && nyr <= END_SUN_APPROX)
			{
				dy1 = *dy + sr2 / 24.0;
    			decl_(&dy1, mp, mc, ap, ac, ms, aas, es, ob, &d__);
			}
			else
			{
				hrs = 12.0 + sr2; //hour angle at sunset <--2nd iteration
				d__ = Decl3(jdn, &hrs, dy, td, mday, mon, &nyr,
							 nyl, ms, aas, ob, &ra, es, 1);
			}

			if ( fabs(((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__)))) > 1.0 )
			{
				//no twilight
				*t3sub = -9999;
				return -1;
			}
			sr2 = acos((-tan(*lr) * tan(d__)) + (cos(air2) / cos(*lr) / cos(d__))) * ch;
			*t3sub = *t6 + sr2; //in fractions of hours
		}

	return 0;
}

////////////////////newzemanim////////////////////////////
short newzemanim(short *yr, short *jday, double *air, double *lr, double *tf, double *dyo, double *hgto,
						short *yro, double *td, double *dy, char *caldayHeb, short *dayweek,
						double *mp, double *mc, double *ap, double *ac,  double *ms, double *aas,
						double *es, double *ob, double *fdy, double *ec, double *e2c,
						double *avekmxzman, double *avekmyzman, double *avehgtzman,
						short *ns1, short *ns2, short *ns3, short *ns4, short *nweather,
						short *numzman, short *numsort, double *t6, double *t3, double *t3sub99,
						double *jdn, int *nyl, short *mday, short *mon,
						short MinTemp[], short AvgTemp[], short MaxTemp[], short ExtTemp[])
///////////////////////////////////////////////////////////////
//parses the .zma files to calculate the requested zemanim
///////////////////////////////////////////////////////////////////
//ZA is zenith angle (degrees)
// PartOfDay = 2; //1 -- sunrise uses averaged minimum temperatures to calculate the refraction
				  //     everything else uses average temperature
				  //2 -- uses average temperatures for the zemanim calculations
				  //3 -- uses maximum temperatures for the zemanim calculations

/////////////////////////////////////////////////////////////////
{

   //static char buff[255] = "";
   char t3subb[12] = "";
   short ier = 0, pos = 0, num = 0;
   double ZA, t3sub, sunrise, sunsets, dawn, twilight, hourszemanios;
   short mishornetznum, mishorskiynum, astnetznum, astskiynum, chazosnum;
   double mishorhgt = 0.0;
   short PartOfDay = 2; //always use average WorldClim temepratures for zemanim

   for (short i = 0; i < *numzman; i++ )
   {
      //'calculate the zeman and record it according to its Combo2,Combo3 number
      if ( InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 4, buff), "dawn" ) )
		{
      	//If Mid$(az$, 1, 4) = "Dawn" Then

			switch (atoi(&zmannames[i][2][0])) //cz$
			{
				case 1: //'dawn defined by solar depression
					ZA = -90 - atof(&zmannames[i][3][0]);
					*dy = (double)*jday;
					ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
								aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
								avekmxzman, avekmyzman,avehgtzman, t6, t3, &t3sub, t3sub99,
								jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

					zmantimes[i] = t3sub;
					zmannumber[0][atoi(&zmannames[i][1][0])] = i;
					break;
				case 2: //dawn defined by minutes zemanios (as defined by mishor sunrise/sunset)
					t3sub = zmantimes[mishornetznum]; //this is mishor sunrise
					if (t3sub != -9999) //Then
					{
						sunrise = t3sub;
						t3sub = zmantimes[mishorskiynum]; // 'this is mishor sunset
						if (t3sub != -9999) //Then
						{
							sunsets = t3sub;
							hourszemanios = (sunsets - sunrise) / 12.0;
							//subtract minutes zemnanios
							if ( (sunrise - (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios) < 0 ) //Then
							{
									t3sub = -9999;
							}
							else
							{
								t3sub = sunrise - (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios;
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[0][atoi(&zmannames[i][1][0])] = i;
					break;
				case -2: //dawn defined by minutes zemanios (as defined by astron. sunrise/sunset)
					t3sub = zmantimes[astnetznum]; //this is astronomical sunrise
					if ( t3sub == -9999 ) //Then
					{
						sunrise = t3sub;
						t3sub = zmantimes[astskiynum]; //this is astronomical sunset
						if (t3sub != -9999 ) // Then
						{
							sunsets = t3sub;
							hourszemanios = (sunsets - sunrise) / 12.0;
							//'subtract minutes zemnanios
							if ( sunrise - (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios < 0  ) //Then
							{
								t3sub = -9999;
							}
							else
							{
								t3sub = sunrise - (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios;
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[0][atoi(&zmannames[i][1][0])] = i;
					break;
				case 3:
					t3sub = zmantimes[mishornetznum]; //this is mishor sunrise
					if (t3sub != -9999) //Then
					{
						if (t3sub != -9999) //Then
						{
							sunrise = t3sub;
							if ( sunrise - (atof(&zmannames[i][5][0]) / 60.0) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else
							{
								t3sub = sunrise - (atof(&zmannames[i][5][0]) / 60.0); //subtract fixed minutes
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[0][atoi(&zmannames[i][1][0])] = i;
					break;
				case -3:
					t3sub = zmantimes[astnetznum]; //this is astron. sunrise
					if (t3sub != -9999) //Then
					{
						if (t3sub != -9999) //Then
						{
							sunrise = t3sub;
							if ( sunrise - (atof(&zmannames[i][5][0]) / 60.0) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else
							{
								t3sub = sunrise - (atof(&zmannames[i][5][0]) / 60.0); //subtract fixed minutes
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[0][atoi(&zmannames[i][1][0])] = i;
					break;
				default:
					return -1; //error in zemanim perameter
					break;
			}

		}
		else if ( InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 8, buff), "twilight" ) )
		{

			switch (atoi(&zmannames[i][2][0])) //cz$
			{
				case 1: // 'twilight defined by solar depression
					ZA = 90 + atof(&zmannames[i][3][0]);
					*dy = (double)*jday;
					ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
							aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
							avekmxzman, avekmyzman,avehgtzman, t6, t3, &t3sub, t3sub99, 
							jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

					zmantimes[i] = t3sub;
					zmannumber[1][atoi(&zmannames[i][1][0])] = i;
					break;
				case 2: //twilight defined by minutes zemanios (as defined by mishor sunrise/sunset)
					t3sub = zmantimes[mishornetznum]; //this is mishor sunrise
					if (t3sub != -9999) //Then
					{
						sunrise = t3sub;
						t3sub = zmantimes[mishorskiynum]; // 'this is mishor sunset
						if (t3sub != -9999) //Then
						{
							sunsets = t3sub;
							hourszemanios = (sunsets - sunrise) / 12.0;
							//subtract minutes zemnanios
							if ( (sunsets + (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else //add minutes zemnanios to sunset
							{
								t3sub = sunsets + (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios;
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[1][atoi(&zmannames[i][1][0])] = i;
					break;
				case -2: //'twilight defined by minutes zemanios (as defined by astron. sunrise/sunset)
					t3sub = zmantimes[astnetznum]; //this is astronomical sunrise
					if ( t3sub == -9999 ) //Then
					{
						sunrise = t3sub;
						t3sub = zmantimes[astskiynum]; //this is astronomical sunset
						if (t3sub != -9999 ) // Then
						{
							sunsets = t3sub;
							hourszemanios = (sunsets - sunrise) / 12.0;
							//subtract minutes zemnanios
							if ( (sunsets + (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else //add minutes zemnanios to sunset
							{
								t3sub = sunsets + (atof(&zmannames[i][4][0]) / 60.0) * hourszemanios;
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[1][atoi(&zmannames[i][1][0])] = i;
					break;
				case 3:
					t3sub = zmantimes[mishorskiynum]; //this is mishor sunset
					if (t3sub != -9999) //Then
					{
						if (t3sub != -9999) //Then
						{
							sunsets = t3sub;
							if ( sunsets + (atof(&zmannames[i][5][0]) / 60.0) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else
							{
								t3sub = sunsets + (atof(&zmannames[i][5][0]) / 60.0); //subtract fixed minutes
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[1][atoi(&zmannames[i][1][0])] = i;
					break;
				case -3: //
					t3sub = zmantimes[astskiynum]; //this is astron. sunset
					if (t3sub != -9999) //Then
					{
						if (t3sub != -9999) //Then
						{
							sunsets = t3sub;
							if ( sunsets + (atof(&zmannames[i][5][0]) / 60.0) < 0 ) //Then
							{
								t3sub = -9999;
							}
							else
							{
								t3sub = sunsets + (atof(&zmannames[i][5][0]) / 60.0); //subtract fixed minutes
							}
						}
						else
						{
							t3sub = -9999;
						}
					}
					zmantimes[i] = t3sub;
					zmannumber[1][atoi(&zmannames[i][1][0])] = i;
					break;
				default:
					return -1; //error in zemanim perameter
					break;
			}
		}
		else if ( InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 7, buff), "sunrise" ) )
		{

			if ( ( InStr( strlwr( &zmannames[i][0][0] ), "astronomical" )) ||
				( optionheb && (InStr( &zmannames[i][0][0], &heb3[24][0] ) ||
				InStr( &zmannames[i][0][0], &heb3[25][0] ))) )
			 {
				 //yr = Mid(&stortim[2][*imonth - 1][*kday - 1][0], 8, 4, buff ); //<<<<<<<<<better to pass the yr
             ZA = -90;
             *dy = (double)*jday;
             ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
							aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
							avekmxzman, avekmyzman,avehgtzman, t6, t3, &t3sub, t3sub99,
							jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

				 zmantimes[i] = t3sub;
             astnetznum = i;
			}
			else if ( ( InStr( strlwr(&zmannames[i][0][0]), "mishor" )) ||
			   ( optionheb && (InStr( &zmannames[i][0][0], &heb3[26][0] ) ||
			   InStr( &zmannames[i][0][0], &heb3[27][0] ))) )
			{
				ZA = -90;
				*dy = (double)*jday;
			// set avehgtzman = 0.0 (mishor)
				ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
						aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
						avekmxzman, avekmyzman, &mishorhgt, t6, t3, &t3sub, t3sub99,
						jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

				 zmantimes[i] = t3sub;
             mishornetznum = i;
		  }

		  zmannumber[0][atoi(&zmannames[i][2][0])] = i;

		}
		else if ( InStr( strlwr(Mid(&zmannames[i][0][0], 1, 6, buff)), "sunset" ) )
		{

			if ( ( InStr( strlwr(&zmannames[i][0][0]), "astronomical" )) ||
				( optionheb && (InStr( &zmannames[i][0][0], &heb3[24][0] ) ||
			   InStr( &zmannames[i][0][0], &heb3[25][0] ))) )
			{
				 //yr = Mid(&stortim[2][*imonth - 1][*kday - 1][0], 8, 4, buff ); //<<<<<<<<<better to pass the yr
             ZA = 90;
             *dy = (double)*jday;
             ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
						aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
					   avekmxzman, avekmyzman,avehgtzman, t6, t3, &t3sub, t3sub99,
					   jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

				 zmantimes[i] = t3sub;
             astskiynum = i;
		  	}
         else if ( ( InStr( strlwr(&zmannames[i][0][0]), "mishor" )) ||
						( optionheb && (InStr( &zmannames[i][0][0], &heb3[26][0] ) ||
					  	InStr( &zmannames[i][0][0], &heb3[27][0] ))) )
			{
				ZA = 90;
				*dy = (double)*jday;
			// set avehgtzman = 0.0 (mishor)
				ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
							aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
						   avekmxzman, avekmyzman, &mishorhgt, t6, t3, &t3sub, t3sub99,
						   jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

				 zmantimes[i] = t3sub;
             mishorskiynum = i;
		  }

		  zmannumber[1][atoi(&zmannames[i][2][0])] = i;

		}
      else if ( InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 7, buff), "candles" ) )
		{

			switch (atoi(&zmannames[i][1][0])) //bz$
			{

				case 1: //minutes before mishor sunset
					t3sub = zmantimes[mishorskiynum]; //this is mishor sunset
					t3sub = t3sub - atof(&zmannames[i][2][0]) / 60.0;
					//sprintf( &zmannames[i][0][0], "%s", &zmannames[i][0][0] ); //<<<<<----not clear that this is necessary
					zmantimes[i] = t3sub;
					break;

				default:
					return -1; //error in zemanim perameter
					break;
			}

		}
		else if ( InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 6, buff), "chazos" ) ||
			      InStr( Mid(strlwr(&zmannames[i][0][0]), 1, 6, buff), "noon" ))
		{

			ZA = 0;
			*dy = (double)*jday;
			ier = cal(&ZA, air, lr, tf, dyo, hgto, yro, td, yr, dy, mp, mc, ap, ac, ms,
					aas, es, ob, fdy, ec, e2c, ns1, ns2, ns3, ns4, nweather, &PartOfDay,
					avekmxzman, avekmyzman,avehgtzman, t6, t3, &t3sub, t3sub99,
					jdn, nyl, mday, mon, MinTemp, AvgTemp, MaxTemp, ExtTemp);

			zmantimes[i] = t3sub;
			sprintf( &zmannames[i][0][0], "%s","Noon (exact): " );
			chazosnum = i;
			zmannumber[0][atoi(&zmannames[i][2][0])] = i;

		}
   }

///////////////////////////////////////////////////////////////////////////////////
//   'now go back and do the zemanim, finish candle lighting, convert to rounded
//   time string, and sort the times for printing
///////////////////////////////////////////////////////////////////////////////////

   for (short n = 0; n < *numzman; n++ )
   {

	  if ( InStr( Mid(strlwr(&zmannames[n][0][0]), 1, 6, buff), "zmanim" ) )
	  {

		// Select Case Val(bz$) 'optionz%
		switch (atoi(&zmannames[n][1][0])) //bz$
		 {
			 case 1: //use hours zemanios
			   dawn = zmantimes[zmannumber[0][atoi(&zmannames[n][2][0])]];
			   twilight = zmantimes[zmannumber[1][atoi(&zmannames[n][3][0])]];
			   if (twilight == -9999 || dawn == -9999 ) // Then
			   {
				  t3sub = -9999;
			   }
			   else
			   {
				  hourszemanios = (twilight - dawn) / 12.0;
				  t3sub = dawn + atof(&zmannames[n][4][0]) * hourszemanios;
				  if ( t3sub < 0 || t3sub >= 24 ) // Then
				  {
					 t3sub = -9999;
				  }
			   }
			   break;
			 case 2: //use clock hours after dawn
			   dawn = zmantimes[zmannumber[0][atoi(&zmannames[n][2][0])]];
			   t3sub = dawn + atof(&zmannames[n][5][0]);
			   if ( t3sub < 0 || t3sub >= 24 ) //Then
			   {
				  t3sub = -9999;
			   }
			   break;
			 default:
				 return -1; //errror in zman paramter
				 break;
		 }

		 zmantimes[n] = t3sub;
		 round( t3sub, atoi(&zmannames[n][7][0]), 0, atoi(&zmannames[n][6][0]), t3subb );
		 sprintf( &c_zmantimes[n][0], "%s", t3subb );
		 //sprintf( &zmantitles[num][0], "%s", &zmannames[n][0][0] );

	  }
	  else if ( InStr( Mid(strlwr(&zmannames[n][0][0]), 1, 7, buff), "candles" ) )
	  {

		   //first deal with Erev Shabbos candle lighting times
		   //if (InStr( &stortim[4][*imonth - 1][*kday - 1][0],&heb4[6][0]) ) //Then
		   if ( *dayweek == 6 ) //Then Erev Shabbos
		   {
			  t3sub = zmantimes[n];
			  round( t3sub, atoi(&zmannames[n][7][0]), 0, atoi(&zmannames[n][6][0]), t3subb );
			  sprintf( &c_zmantimes[n][0], "%s", t3subb );

		   }
		   else if ( *dayweek != 6 ) // Then not Erev Shabbos
		   {
			  //also list the candle lighting for erev yom tov

			  if ( !strcmp( caldayHeb, &heb6[1][0] ) ||
	            !strcmp( caldayHeb, &heb6[2][0] ) ||
				   !strcmp( caldayHeb, &heb6[3][0] ) ||
				   !strcmp( caldayHeb, &heb6[4][0] ) ||
				   !strcmp( caldayHeb, &heb6[5][0] ) ||
				   !strcmp( caldayHeb, &heb6[6][0] ) ||
				   !strcmp( caldayHeb, &heb6[7][0] ) ||
				   !strcmp( caldayHeb, "29-Elul" ) ||
				   !strcmp( caldayHeb, "9-Tishrey" ) ||
				   !strcmp( caldayHeb, "14-Tishrey" ) ||
				   !strcmp( caldayHeb, "21-Tishrey" ) ||
				   !strcmp( caldayHeb, "14-Nisan" ) ||
				   !strcmp( caldayHeb, "20-Nisan" ) ||
				   !strcmp( caldayHeb, "5-Sivan" ) )
			  {

					if ( *dayweek == 7 || *dayweek == 0 ) // Then Shabbos
					{
						//this is erev yom-tov that falls on shabbos (NO CANDLE LIGHTING!)
						round( 0.0, 0, 0, 0, t3subb ); //outputs "-----"
						sprintf( &c_zmantimes[n][0], "%s", t3subb );
						sprintf( &zmantitles[n][0], "%s", &zmannames[n][0][0] );
					}
					else
					{
					t3sub = zmantimes[n];
					round( t3sub, atoi(&zmannames[n][7][0]), 0, atoi(&zmannames[n][6][0]), t3subb );
					sprintf( &c_zmantimes[n][0], t3subb );
					}

				}
				else
				{
				round( 0.0, 0, 0, 0, t3subb ); //outputs "-----"
				sprintf( &c_zmantimes[n][0], "%s", t3subb );
				}
		   }
	   }
	   else
	   {
		  t3sub = zmantimes[n];
		  round( t3sub, atoi(&zmannames[n][7][0]), 0, atoi(&zmannames[n][6][0]), t3subb );
		  sprintf( &c_zmantimes[n][0], "%s", t3subb );
	   }


//		}

	}

    return 0;

}


////////////////////YearLength////////////////////////////
short YearLength( short myear)
//////////////////////////////////////////////////////////
//calculates the length of the civil year, myear, in days
//////////////////////////////////////////////////////////
{
	short yd, yl;

	yd = myear - 1988;
    yl = 365;
    if (yd % 4 == 0) yl = 366;
    if ((yd % 4 == 0) && (myear % 100 == 0) && (myear % 400 != 0)) yl = 365;

	return yl;
}

//////////////////////////ParseData///////////////////////////
short ParseData( char *doclin, char *sep, double arr[], short NUMENTRIES, short MAXARRSIZE )
///////////////////////////////////////////////////////////////////////////////////////////
//Parses a line of double precision data separated by the string "sep" (e.g., ",", " ", etc)
//stuffs the extracted double precision data into the array, arr
//number of data entries to extract = NUMENTRIES
//maximum number of data entries that can be extracted is defined by the array size = MAXARRSIZE
//////////////////////////////////////////////////////////////////
{
	if ( NUMENTRIES == 0 || NUMENTRIES > MAXARRSIZE ) return -1;

	short pos1 = 1;
	short pos2 = 0;
	//static char buff[255] = "";
	short strnglen = 0;

	strnglen = strlen( doclin );

	//begin parsing by looking for sep string
	for (short i = 0; i < NUMENTRIES; i++ )
	{
		pos2 = InStr( pos1, doclin, sep ); //sep found at position "pos2"
		if (pos2)
		{
			buff[0] = '\0'; //clear buffer
			Mid( doclin, pos1, pos2 - pos1, buff );
			arr[i] = atof( buff );
			pos1 = pos2 + 1;
		}
		else
		{
			//no more sep strings found, so parse last bit of doclin and return
			if ( pos1 < strnglen )
			{
				buff[0] = '\0'; //clear buffer
				Mid( doclin, pos1, strnglen - pos1, buff );
				arr[i] = atof( buff );
			}
			else //sep at end of string, so vlaue of data is = 0
			{
				arr[i] = 0.;
			}

			return 0;

		}
	}

	return 0;
}

//////////////////////////ParseString///////////////////////////
short ParseString( char *doclin, char *sep, char arr[][MAXDBLSIZE], short NUMENTRIES )
///////////////////////////////////////////////////////////////////////////////////////////
//Parses a line of text which is composed of tokens separated by "sep" (e.g., ",", " ", etc)
//stuffs the extracted strings into the array, arr
//number of tokens entries to extract = NUMENTRIES
//each token can be a maximum MAXDBLSIZE characters long
//maximum number of data entries that can be extracted is defined by the array size = MAXARRSIZE
//////////////////////////////////////////////////////////////////
{
	if ( NUMENTRIES == 0 ) return -1;

	char *token;
	short numtoken = 0;

	//get first token
	token = strtok( doclin, sep );
	strcpy( &arr[0][0], token );

	//get other tokens
    while( token != NULL )
    {
       if (numtoken + 1 >= NUMENTRIES) break;
	   token = strtok( NULL, sep );
	   numtoken++;
	   strcpy( &arr[numtoken][0], token );
    }

	return 0;
}


/*********************************************************************
cgiinput - get the input data from the cgi to global array called vars
*********************************************************************/
short cgiinput (short *NumInputs) //<--returns actual number of cgi inputs
   {

   char *input[MAXCGIVARS]; //<--MAXCGIVARS = maximum possible number of cgi inputs
   int j = 0;
   int k;
   char * pch;

   int x;
   int y;

   if (INPUT_TYPE == 0) //post
   {
      const char *len1 = getenv("CONTENT_LENGTH");
      if (len1 == NULL) //the header string containing the cgi input can't be read
      {
         return -1;
      }

      int contentlength;
      contentlength=atoi(len1);
      fread(buff, contentlength, 1, stdin);

   }
   else if (INPUT_TYPE == 1) //get
   {
      const char *get_data = getenv("QUERY_STRING");
      strcpy(buff,get_data);

      ////////////diagnostics-prints out query string///////////////
      //short ier;
      //ier = WriteUserLog( 0, 0, 0, buff );

   }

   if (strstr(buff, "Geofinder"))
   {
	   Geofinder = true; //using cgi rather than php to do the redirect to the Geofinder
   }

   //make all vars variable = ""

   for (x=0; x < 2; x++)
   {
		for (y=0; y < MAXCGIVARS; y++)
		{
			vars[x][y]="";
		}
   }

	for (x = 0, y = 0; x < strlen(buff); x++, y++)
	{
		switch (buff[x])
         {
         /* Convert all + chars to space chars */
				case '+':
					buff2[y] = ' ';
					break;

				/* Convert all %xy hex codes into ASCII chars */
				case '%':

					/* Copy the two bytes following the % */
					strncpy(buff3, &buff[x + 1], 2);

					/* Skip over the hex */
					x = x + 2;

					/* Convert the hex to ASCII */
					/* Prevent user from altering URL delimiter sequence */
					if( ((strcmp(buff3,"26")==0)) || ((strcmp(buff3,"3D")==0)) )
					{
						buff2[y]='%';
						y++;
						strcpy(buff2,buff3);
						y=y+2;
						break;
					}

					unsigned char tmp;
					sscanf(buff3,"%x",&tmp);
					buff2[y]=tmp;
					break;

				/* Make an exact copy of anything else */
				default:
					buff2[y] = buff[x];
					break;
			}
      }


   /* splitting buff2 to tokens with ampersand as separator*/

   pch = strtok (buff2,"&");

   while (pch != NULL)
     {
       if (!strstr(pch, "submit")) //ignore submit button value
       {
           input[j] = pch;
           j++;
           //fprintf ( stdout, "%s\n",pch);
           pch = strtok (NULL, "&");
       }
       else
       {
           break;
       }
     }

   *NumInputs = j;

   /* splitting name and value */

   for (k = 0; k < j; k++)
   {
		vars[0][k] = strtok (input[k],"=");
		vars[1][k] = strtok(NULL,"=");
   }

	return 0;

}

/************************************************************************
getpair - function for getting variable from the pseudo assocaitive array
************************************************************************/
char * getpair(short *NumInputs, char * name)

{
	int i;

	if (*NumInputs == 0) return "Error - Not found";

	for (i=0; i < *NumInputs; i++)
	{
        if (strcmp(vars[0][i],name)==0)
		{
            if (vars[1][i] == NULL ) return "Error - Not found"; //catch nulls
			return vars[1][i];
		}
  	}

	return "Error - Not found";
}

////////////////////RedirectToGeofinder////////////////////////
short RedirectToGeofinder()
///////////////////////////////////////////////////////////////
//redirect to chai_geofinder.php using Javascript redirect
////////////////////////////////////////////////////////////////
{
	fprintf(stdout, "%s\n\n", "Content-type: text/html");
	fprintf(stdout, "%s\n", "<!doctype html public \"-//w3c//dtd html 4.0 Transitional//en\">");
	fprintf(stdout, "%s\n", "<html>");
	fprintf(stdout, "%s\n", "<head>");
	fprintf(stdout, "%s\n", "    <title>Your Chai Table</title>");
	fprintf(stdout, "%s\n", "    <meta http-equiv=\"content-type\" content=\"text/html; charset=UTF-8\"/>");
	fprintf(stdout, "%s%s%s\n", "    <meta name=\"author\" content=\"Chaim Keller, Chai Tables, ", &DomainName[0], "\"/>");
	fprintf(stdout, "%s\n", "</head>");
	fprintf(stdout, "%s\n", "<body onload=\"submitform()\">");
	fprintf(stdout, "%s\n", "<form name=\"Geofinder\" action=\"/chai_geofinder.php\">" );
	fprintf(stdout, "%s%s%s\n", "<input type=\"hidden\" id=\"cgi_country\" name=\"cgi_country\" value=\"", eroscountry, "\" >" );
	fprintf(stdout, "%s%s%s\n", "<input type=\"hidden\" id=\"metroname\" name=\"metroname\" value=\"", erosareabat, "\" >" );
	if (HebrewInterface) //Hebrew interface
	{
		fprintf(stdout, "%s\n", "<input type=\"hidden\" id=\"interface\" name=\"cgi_Language\" value=\"Hebrew\" >" );
	}
	else //English interface
	{
		fprintf(stdout, "%s\n", "<input type=\"hidden\" id=\"interface\" name=\"cgi_Language\" value=\"English\" >" );
	}
	fprintf(stdout, "%s\n", "</form>" );
	fprintf(stdout, "%s\n", "    <script language=\"JavaScript\"/>" );
	fprintf(stdout, "%s\n", "    <!--" );
	fprintf(stdout, "%s\n", "    document.Geofinder.submit();" );
	fprintf(stdout, "%s\n", "    //-->" );
	fprintf(stdout, "%s\n", "    </script>" );
	fprintf(stdout, "%s\n", "</body>");
	fprintf(stdout, "%s\n", "</html>");
	exit(0); //finished with the cgi, success
}

//////////////////checkenviron//////////////////////////
short checkenviron()
/////////////////////////////////////////////////////
//check html header info -- used for diagnosis
/////////////////////////////////////////////////////
{

	char *pEnvPtr;

	// SERVER_SOFTWARE
	pEnvPtr= getenv ("SERVER_SOFTWARE");
	fprintf( stdout, "%s\n", "SERVER_SOFTWARE= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// SERVER_NAME
	pEnvPtr= getenv ("SERVER_NAME");
	fprintf( stdout, "%s\n", "SERVER_NAME= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// SERVER_PROTOCOL
	pEnvPtr= getenv ("SERVER_PROTOCOL");
	fprintf( stdout, "%s\n", "SERVER_PROTOCOL= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// SERVER_PORT
	pEnvPtr= getenv ("SERVER_PORT");
	fprintf( stdout, "%s\n", "SERVER_PORT= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REQUEST_URI
	pEnvPtr= getenv ("REQUEST_URI");
	fprintf( stdout, "%s\n", "REQUEST_URI= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REQUEST_METHOD
	pEnvPtr= getenv ("REQUEST_METHOD");
	fprintf( stdout, "%s\n", "REQUEST_METHOD= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// GATEWAY_INTERFACE
	pEnvPtr= getenv ("GATEWAY_INTERFACE");
	fprintf( stdout, "%s\n", "GATEWAY_INTERFACE= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// HTTP_CONNECTION
	pEnvPtr= getenv ("HTTP_CONNECTION");
	fprintf( stdout, "%s\n", "HTTP_CONNECTION= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// PATH_INFO
	pEnvPtr= getenv ("PATH_INFO");
	fprintf( stdout, "%s\n", "PATH_INFO= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// PATH_TRANSLATED
	pEnvPtr= getenv ("PATH_TRANSLATED");
	fprintf( stdout, "%s\n", "PATH_TRANSLATED= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REMOTE_HOST
	pEnvPtr= getenv ("REMOTE_HOST");
	fprintf( stdout, "%s\n", "REMOTE_HOST= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REMOTE_ADDR
	pEnvPtr= getenv ("REMOTE_ADDR");
	fprintf( stdout, "%s\n", "REMOTE_ADDR= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REMOTE_PORT
	pEnvPtr= getenv ("REMOTE_PORT");
	fprintf( stdout, "%s\n", "REMOTE_PORT= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// REMOTE_IDENT
	pEnvPtr= getenv ("REMOTE_IDENT");
	fprintf( stdout, "%s\n", "REMOTE_IDENT= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// SCRIPT_FILENAME
	pEnvPtr= getenv ("SCRIPT_FILENAME");
	fprintf( stdout, "%s\n", "SCRIPT_FILENAME= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// SCRIPT_NAME
	pEnvPtr= getenv ("SCRIPT_NAME");
	fprintf( stdout, "%s\n", "SCRIPT_NAME= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// QUERY_STRING
	pEnvPtr= getenv ("QUERY_STRING");
	fprintf( stdout, "%s\n", "QUERY_STRING= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// CONTENT_TYPE
	pEnvPtr= getenv ("CONTENT_TYPE");
	fprintf( stdout, "%s\n", "CONTENT_TYPE= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");

	// CONTENT_LENGTH
	pEnvPtr= getenv ("CONTENT_LENGTH");
	fprintf( stdout, "%s\n", "CONTENT_LENGTH= ");
	if (!pEnvPtr)
	fprintf( stdout, "%s\n", "<NULL-POINTER>");
	else
	fprintf( stdout, "%s\n", pEnvPtr);
	fprintf( stdout, "%s\n", "<br>");
	return (EXIT_SUCCESS);

}

/********************Conversion Functions from ITM to WGS84***********************/
//=================================================
// Israel New Grid (ITM) to WGS84 conversion
//=================================================
void itm2wgs84(int N,int E,double& lat,double &lon)
{
	// 1. Local Grid (ITM) -> GRS80
	double lat80,lon80;
	Grid2LatLon(N,E,lat80,lon80,gITM,eGRS80);

	// 2. Molodensky GRS80->WGS84
	double lat84,lon84;
	Molodensky(lat80,lon80,lat84,lon84,eGRS80,eWGS84);

	// final results
	lat = lat84*180/pi;
	lon = lon84*180/pi;
}

//=================================================
// WGS84 to Israel New Grid (ITM) conversion
//=================================================
void wgs842itm(double lat,double lon,int& N,int& E)
{
	double latr = lat*pi/180;
	double lonr = lon*pi/180;

	// 1. Molodensky WGS84 -> GRS80
	double lat80,lon80;
	Molodensky(latr,lonr,lat80,lon80,eWGS84,eGRS80);

	// 2. Lat/Lon (GRS80) -> Local Grid (ITM)
	LatLon2Grid(lat80,lon80,N,E,eGRS80,gITM);
}

//=================================================
// Israel Old Grid (ICS) to WGS84 conversion
//=================================================
void ics2wgs84(int N,int E,double *lat,double *lon)
{
	// 1. Local Grid (ICS) -> Clark_1880_modified
	double lat80,lon80;
	//printf("inside N, E = %d, %d", N, E);
	//pause();
	Grid2LatLon(N,E,lat80,lon80,gICS,eCLARK80M);
	//printf("grid lat, lon = %f, %f", lat80*180/pi, lon80*180/pi);
	//pause();

	// 2. Molodensky Clark_1880_modified -> WGS84
	double lat84,lon84;
	Molodensky(lat80,lon80,lat84,lon84,eCLARK80M,eWGS84);

	// final results
	*lat = lat84*180/pi;
	*lon = lon84*180/pi;
	//printf("final lat, lon = %f, %f", lat, lon);
	//pause();

}

//=================================================
// WGS84 to Israel Old Grid (ICS) conversion
//=================================================
void wgs842ics(double lat,double lon,int& N,int& E)
{
	double latr = lat*pi/180;
	double lonr = lon*pi/180;

	// 1. Molodensky WGS84 -> Clark_1880_modified
	double lat80,lon80;
	Molodensky(latr,lonr,lat80,lon80,eWGS84,eCLARK80M);

	// 2. Lat/Lon (Clark_1880_modified) -> Local Grid (ICS)
	LatLon2Grid(lat80,lon80,N,E,eCLARK80M,gICS);
}

//====================================
// Local Grid to Lat/Lon conversion
//====================================
void Grid2LatLon(int N,int E,double& lat,double& lon,gGrid from,eDatum to)
{
	//================
	// GRID -> Lat/Lon
	//================
	//printf("gridfrom.falsen = %f", Grid[from].false_n);
	//printf("datumto.a = %f", Datum[to].a);
	//pause();

	double y = N + Grid[from].false_n;
	double x = E - Grid[from].false_e;
	double M = y / Grid[from].k0;

	double a = Datum[to].a;
	double b = Datum[to].b;
	double e = Datum[to].e;
	double esq = Datum[to].esq;

	double mu = M / (a*(1 - e*e/4 - 3*pow(e,4)/64 - 5*pow(e,6)/256));

	double ee = sqrt(1-esq);
	double e1 = (1-ee)/(1+ee);
	double j1 = 3*e1/2 - 27*e1*e1*e1/32;
	double j2 = 21*e1*e1/16 - 55*e1*e1*e1*e1/32;
	double j3 = 151*e1*e1*e1/96;
	double j4 = 1097*e1*e1*e1*e1/512;

	// Footprint Latitude
	double fp =  mu + j1*sin(2*mu) + j2*sin(4*mu) + j3*sin(6*mu) + j4*sin(8*mu);

	double sinfp = sin(fp);
	double cosfp = cos(fp);
	double tanfp = sinfp/cosfp;
	double eg = (e*a/b);
	double eg2 = eg*eg;
	double C1 = eg2*cosfp*cosfp;
	double T1 = tanfp*tanfp;
	double R1 = a*(1-e*e) / pow(1-(e*sinfp)*(e*sinfp),1.5);
	double N1 = a / sqrt(1-(e*sinfp)*(e*sinfp));
	double D = x / (N1*Grid[from].k0);

	double Q1 = N1*tanfp/R1;
	double Q2 = D*D/2;
	double Q3 = (5 + 3*T1 + 10*C1 - 4*C1*C1 - 9*eg2*eg2)*(D*D*D*D)/24;
	double Q4 = (61 + 90*T1 + 298*C1 + 45*T1*T1 - 3*C1*C1 - 252*eg2*eg2)*(D*D*D*D*D*D)/720;
	// result lat
	lat = fp - Q1*(Q2-Q3+Q4);

	double Q5 = D;
	double Q6 = (1 + 2*T1 + C1)*(D*D*D)/6;
	double Q7 = (5 - 2*C1 + 28*T1 - 3*C1*C1 + 8*eg2*eg2 + 24*T1*T1)*(D*D*D*D*D)/120;
	// result lon
	lon = Grid[from].lon0 + (Q5 - Q6 + Q7)/cosfp;
}

//====================================
// Lat/Lon to Local Grid conversion
//====================================
void LatLon2Grid(double lat,double lon,int& N,int& E,eDatum from,gGrid to)
{
	// Datum data for Lat/Lon to TM conversion
	double a = Datum[from].a;
	double e = Datum[from].e; 	// sqrt(esq);
	double b = Datum[from].b;

	//===============
	// Lat/Lon -> TM
	//===============
	double slat1 = sin(lat);
	double clat1 = cos(lat);
	double clat1sq = clat1*clat1;
	double tanlat1sq = slat1*slat1 / clat1sq;
	double e2 = e*e;
	double e4 = e2*e2;
	double e6 = e4*e2;
	double eg = (e*a/b);
	double eg2 = eg*eg;

	double l1 = 1 - e2/4 - 3*e4/64 - 5*e6/256;
	double l2 = 3*e2/8 + 3*e4/32 + 45*e6/1024;
	double l3 = 15*e4/256 + 45*e6/1024;
	double l4 = 35*e6/3072;
	double M = a*(l1*lat - l2*sin(2*lat) + l3*sin(4*lat) - l4*sin(6*lat));
	//double rho = a*(1-e2) / pow((1-(e*slat1)*(e*slat1)),1.5);
	double nu = a / sqrt(1-(e*slat1)*(e*slat1));
	double p = lon - Grid[to].lon0;
	double k0 = Grid[to].k0;
	// y = northing = K1 + K2p2 + K3p4, where
	double K1 = M*k0;
	double K2 = k0*nu*slat1*clat1/2;
	double K3 = (k0*nu*slat1*clat1*clat1sq/24)*(5 - tanlat1sq + 9*eg2*clat1sq + 4*eg2*eg2*clat1sq*clat1sq);
	// ING north
	double Y = K1 + K2*p*p + K3*p*p*p*p - Grid[to].false_n;

	// x = easting = K4p + K5p3, where
	double K4 = k0*nu*clat1;
	double K5 = (k0*nu*clat1*clat1sq/6)*(1 - tanlat1sq + eg2*clat1*clat1);
	// ING east
	double X = K4*p + K5*p*p*p + Grid[to].false_e;

	// final rounded results
	E = (int)(X+0.5);
	N = (int)(Y+0.5);
}

//======================================================
// Abridged Molodensky transformation between 2 datums
//======================================================
void Molodensky(double ilat,double ilon,double& olat,double& olon,eDatum from,eDatum to)
{
	// from->WGS84 - to->WGS84 = from->WGS84 + WGS84->to = from->to
	double dX = Datum[from].dX - Datum[to].dX;
	double dY = Datum[from].dY - Datum[to].dY;
	double dZ = Datum[from].dZ - Datum[to].dZ;

	double slat = sin(ilat);
	double clat = cos(ilat);
	double slon = sin(ilon);
	double clon = cos(ilon);
	double ssqlat = slat*slat;

	//dlat = ((-dx * slat * clon - dy * slat * slon + dz * clat)
	//        + (da * rn * from_esq * slat * clat / from_a)
	//        + (df * (rm * adb + rn / adb )* slat * clat))
	//       / (rm + from.h);

	double from_f = Datum[from].f;
	double df = Datum[to].f - from_f;
	double from_a = Datum[from].a;
	double da = Datum[to].a - from_a;
	double from_esq = Datum[from].esq;
	double adb = 1.0 / (1.0 - from_f);
	double rn = from_a / sqrt(1 - from_esq * ssqlat);
	double rm = from_a * (1 - from_esq) / pow((1 - from_esq * ssqlat),1.5);
	double from_h = 0.0; // we're flat!

	double dlat = (-dX*slat*clon - dY*slat*slon + dZ*clat
				   + da*rn*from_esq*slat*clat/from_a +
				   + df*(rm*adb + rn/adb)*slat*clat) / (rm+from_h);

	// result lat (radians)
	olat = ilat+dlat;

	// dlon = (-dx * slon + dy * clon) / ((rn + from.h) * clat);
	double dlon = (-dX*slon + dY*clon) / ((rn+from_h)*clat);
	// result lon (radians)
	olon = ilon+dlon;
}

/////////////////////////readDTM/////////////////////////////
short readDTM( double *kkmxo, double *kkmyo, double *khgt, short *kmaxang, 
			   double *kaprn, double *p, double *meantemp, short nsetflag )
/////////////////////////////////////////////////////////////
//calculates horizon profiles from 30m DTM
/////////////////////////////////////////////////////////////
{
    double beglog;
    double endlog;
    double beglat;
    double endlat;
    double lat0,lon0,hgt0,hgt;
    //short int landflag;
    short int mode = 0;
    float modeval = 0;
    int ang; //maximum angular range
    bool directx = false;
    short m_progresspercent = 0;
	short m_progresspercent2 = 0;
    short oldpercent = 0;
	short oldpercent2 = 0;

	short ier = 0;

	///////////used for printing/////////////
	//short alreadyprinted = 0;
	double vang_max = -9999;
	double vang_min = 9999;
	bool NearObstructions = false;
    //////////////////////////////////////////////////////////////

	/////////////GoldenLight variables//////////////
	int nk_max = 0;
	double dist_max = 0.;
	double hgt_max = -9999.;
	short ier_max = 0;
	//////////////////////////////////////////////

    //bool manytiles, record,
    bool poles;
    //short numtiles;
    short whichCDs[4];
    double ntmp, tmp, nx;
    int ny, lg1, lg11, lt1, lt11;

    int totnumtiles;
    int numtotrec;

    double xdim,ydim;
    int nrows, ncols;
    int gototag, reply;//, rep;
    long j,ii,ikmy,ikmx,itnum; //i,jj,kk,
    long iikmy,iikmx;
    long maxrange = 0;
	long maxranget = 0;
    double kmx,kmy;
    long pos, numrc, numrec, numCD, tmprc, numrctot;
    //long numrc2; //used for removing radar shadow
    long numstepkmx, numstepkmy;//,numx,numy;
    double kmxo,kmyo;
    //long nverts;
    //char st2[16];
    char quest[255] = "";
    //char buffer[5];
    //char CDbuffer[1];
    //char filt[255] = "";
    //char filtmp[8];
    char filn[255] = "";
    char filnn[255] = "";
    char DTMfile[255] = "";
    char lg1ch[15];
    char lt1ch[15];
    char lch[3];
    char EW[2];
    char NS[2];

    //char worlddtm[]="";
    //bool worlddtmcd;
    //char tmpfil[81];
    int DTMflag = 1; //default DTM = 30 m DTM masquerading as 30 m SRTM hgt files
    bool IgnoreMissingTiles;
	IgnoreMissingTiles = g_IgnoreMissingTiles;
    short calcProfile = 1; //=1 for calculating profile, = 0 for not calculating profile
    double AzimuthStep = 0.1; //step size in azimuth for horizon profiles

    int nkmx1;
    short minview = 10;
	double dstpointx;
	double dstpointy;
	double maxang = 0;
	double aprn = 0;
	double mang;
    double maxang0 = 0;
	short nomentr = 0;
	double apprnr, nnetz;
	bool netz = false;
    bool found = false;
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
	double begkmx, endkmx, skipkmx, skipkmy;
	int nbegkmx, numkmx, i__, i__1, i__2, i__3, i__4, nk;
	double re, rekm, range, dkmy, testazi;
	float hgt2, treehgt = 0, fudge = 0, xentry;
	double lg2, lt2, x1, x2, y1, y2, z1, z2, d__1, d__2, d__3;
	double distd, re1, re2, deltd, dist1, dist2, angle, viewang;
	double d__, x1d, y1d, z1d, azicos, x1s, y1s, z1s, x1p, y1p, z1p, azisin, azi;
	double defm, defb, avref, kmxp, kmyp, bne1, bne2;
	int istart1, ndela, ndelazi, ne, nstart2;
	double az2,vang2,kmy2,az1,vang1,kmy1,bne, vang, proftmp, delazi, dist;
	int invAzimuthStep = 10;

	//char pszFileName[256];
	//char lpszText[256];
    short worldCD[29];
    unsigned int bytesread;

	//delcare and allocate a minimum memory for other arrays
	/*
	int lg1chk[MAX_NUM_TILES];
	int lt1chk[MAX_NUM_TILES];
	double kmxb1[MAX_NUM_TILES];
	double kmxb2[MAX_NUM_TILES];
	double kmyb1[MAX_NUM_TILES];
	double kmyb2[MAX_NUM_TILES];
	long numstepkmyt[MAX_NUM_TILES];
	long numstepkmxt[MAX_NUM_TILES];
	char filt1[MAX_NUM_TILES][255];
	*/

	FILE *stream;
	FILE* f;

//////////////////////////////////
   //1 km dtm of the world
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

   //pass over addressses and convert to local values
   lat0 = *kkmyo * 1.f;
   lon0 = *kkmxo * -1.f;
   hgt0 = *khgt * 1.f;
   ang = *kmaxang * 1;
   aprn = __max(*kaprn * 1.f, 0.03f); //closest approach is 30 meters

   modeval = 0; //ignore this parmeter

   //calculate beglog, endlog, beglat, endlat -- as done in Maps & More
   switch (nsetflag)
   {
       case 10: //sunrise (eastern) horizon profile
	   case 12:
            mode = nsetflag;
            beglat = lat0 - erosdiflat; // * 0.5f; 'not twice width as in Maps&More
            endlat = beglat + 2.0 * erosdiflat;
            beglog = lon0 - 0.1;
            endlog = beglog + erosdiflog; // * 0.5f;
            break;
       case 11: //sunset (western) horizon profile
	   case 13:
            mode = -nsetflag;
            beglat = lat0 - erosdiflat; // * 0.5f;
            endlat = beglat + 2.0 * erosdiflat;
            endlog = lon0 + 0.1;
            beglog = endlog - erosdiflog; // * 0.5f;
            break;
   }

	xdim = 0.008333333333333; //GTOPO30
	ydim = 0.008333333333333;

	xdim = xdim/30.0; // 30m DTM
	ydim = ydim/30.0;

	//deter__min number of columns of rows to be extracted
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
	  if (totnumtiles > MAX_NUM_TILES) return -6;

	}
	//deallocate memory for tile boundaries and names
	//and then allocate the required amount
	if (GoldenLight < 2) {

		try
		{
			delete [] lg1chk;
			delete [] lt1chk;
			delete [] kmxb1;
			delete [] kmxb2;
			delete [] kmyb1;
			delete [] kmyb2;
			delete [] numstepkmyt;
			delete [] numstepkmxt;
			delete [] filt1;
		}
		catch (short ier)
		{
			ier = 0;
		}

	}

	//allocate memory as required for totnumtiles of tiles
	int *lg1chk = new int[totnumtiles];
	int *lt1chk = new int[totnumtiles];
	double *kmxb1 = new double[totnumtiles];
	double *kmxb2 = new double[totnumtiles];
	double *kmyb1 = new double[totnumtiles];
    double *kmyb2 = new double[totnumtiles];
	long *numstepkmyt = new long[totnumtiles];
	long *numstepkmxt = new long[totnumtiles];
	char (*filt1)[255] = new char[totnumtiles][255];
	
    //************************progress notification********************************//
    //appraise user of calculation progress
    WriteHtmlHeader();

	fprintf( stdout, "%s\n", "<bdo dir=\"ltr\">" );

	/*
	if ( !optionheb )
	{
		sprintf( buff1, "%s", "<div style=\"position: absolute; top:150px; width:300px; left:50px; height:25px\" background-color: white>Initializing.....please wait</div>" );
		//					   <DIV style="position: absolute; top:80px; left:400px; width:200px; height:25px">80 from the top and 400 from the left</DIV>
	    fprintf( stdout, "%s\n", buff1);
	}
	else
    {
         //to do
         //sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb1[1][0], " \"", TitleLine, "\" ", &heb7[12][0], hebcityname, &heb3[2][0], g_yrcal, "</h2>");
    }
	*/

    //*********************************************************************************//

		// deter__min if need to read multiple tiles
		gototag = 1;
		goto findtile;

ret1:   ;
		//now open the .DEM file
		gototag = 1;
		goto dtmfiles;

ret2:	;


		//begin writing to the BIN file


	//allocate memory for main data array (add a little more just to be safe)
	//union unich *uni5;
	if (GoldenLight < 2) free( uni5 ); //free the memory allocated before for uni1, and reallocate
    uni5 = (struct unich *) malloc(maxrange * sizeof(struct unich));

	if (uni5 == NULL) return -5; //memmory wasn't allocated, return error code

	/*
	if ( !optionheb )
	{
		sprintf( buff1, "%s", "<div style=\"position: absolute; top:150px; width:300px; left:50px; height:25px\" background-color: white><font color=white>Initializing.....Please wait</font></div>" );
		//					   <DIV style="position: absolute; top:80px; left:400px; width:200px; height:25px">80 from the top and 400 from the left</DIV>
		fprintf( stdout, "%s\n", buff1);
	}
	else
    {
         //to do
         //sprintf( buff1, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", &heb1[1][0], " \"", TitleLine, "\" ", &heb7[12][0], hebcityname, &heb3[2][0], g_yrcal, "</h2>");
    }
	*/


	short iter;
	//long nvert;
	numrctot = 0;
	oldpercent = -999;
	oldpercent2 = -999;

	/*
	sprintf( buff1, "%s", "<div style=\"position: absolute; top:170px; width:400px; left:50px; height:25px\" background-color: white>Extracting DTM...please wait</div>" );
    fprintf( stdout, "%s\n", buff1);
	*/

    for ( iter = 0; iter < totnumtiles ; ++iter )
		{

        ////////////////////progress bar for slower server////////////////////////////////////////
	    m_progresspercent = (int)((iter /  (float)(totnumtiles - 1)) * 100);

		if ( m_progresspercent != oldpercent )
			{

			oldpercent = m_progresspercent;	
			/////////////////draw progress bar using canvas//////////////////////////////
			////////////////////////////////////////////////////////////////////////////////////////////////////////////
			fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 250px; left: 45%;\" >" );
			fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
			fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
			fprintf(stdout, "%s\n", "</canvas>" );

			fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
			fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
			fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
			fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );

			fprintf(stdout, "%s\n", "ctx1.strokeStyle = \"#000000\";" );
			fprintf(stdout, "%s\n", "ctx1.lineWidth = 2;" );
			fprintf(stdout, "%s\n", "ctx1.strokeRect(0, 0, 101, 26);" );

			fprintf(stdout, "%s\n", "ctx1.font = 'italic 10px sans-serif';" );
			fprintf(stdout, "%s\n", "ctx1.textBaseline = 'top';" );
			/////////////////////////////////////////////////////////////////////////////////////////////////////////////
			fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#99CC00\";" );
			//fprintf(stdout, "%s\n", "ctx1.clearRect(1, 1, 101, 26);" );
			fprintf(stdout, "%s%d%s\n", "ctx1.fillRect(1,1,", m_progresspercent,",24);" );
			fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#000000\";" );
			fprintf(stdout, "%s%d%s\n", "ctx1.fillText('", m_progresspercent, "%', 45, 5);" );
			/////////////////////////////////////////////////////////////////////////////
			fprintf(stdout, "%s\n", "</script>" );
			fprintf(stdout, "%s\n", "</div>" );

			}


////////////////////////////////////////////////////////////////////////////////////////////////////////////

		if (iter != 0) // open new tile
			{
			// check that this is first time to process this tile
			// need to open new tile, so close old tile
			fclose ( stream );
			if (DTMflag == 0) //GTOPO30 data
			{
			//strncpy( &filt[0], &filt1[iter][0], 18 );
			sprintf( &filroot[0], "%s", drivdtm);
			//strncpy( filn, (const char *)worlddtm, 1);
			//strncpy( filn + 1, ":\\", 2 );

			//strncpy( filn + 3, &filt1[iter][0], 7 );
			//strncpy( filn + 10, "/", 1 );
			//strncpy( filn + 11, &filt1[iter][0], 7 );
			sprintf( &filn[0], "%s%s%s%s%s", drivdtm, &filt1[iter][0], "/", &filt1[iter][0], "\0" );
			numCD = whichCDs[iter];
			}
			else //SRTM data
			{
			//strncpy( &filt[0], (const char *)filt1[iter], 7 );
			//strncpy( filn + 7, (const char *)filt1[iter], 7);
			sprintf(&filn[0], "%s%s", drivdtm, &filt1[iter][0] );
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

    			kmy = kmyb1[iter] - ydim * (double)(ii - 1);
    			//i = Nint((endlat - kmy)/ydim) + 1;
    			double kmyff;
    			kmyff = kmy;
    			poles = false;
    			if ( kmy < -90 )
    				{
    				kmyff = -180 - kmy;
    				poles = true;
    				}
    			else if ( kmy > 90 )
    				{
    				kmyff = 180 - kmy;
    				poles = true;
    				}

    			ikmy = Nint((lt1chk[iter] - kmyff) / ydim) + 1;

    			{
    				kmx = kmxb1[iter]; //beglog;  // for polar regions manytiles = TRUE
    				//ikmx = (int)(floor((kmx - lg1) / xdim) + 1);
    				ikmx = Nint((kmx - lg1chk[iter]) / xdim) + 1;
    				numrec = (ikmy - 1) * ncols + ikmx - 1;
    				//position DEM file at starting coordinate for reading
    				pos = fseek( stream, numrec * 2L, SEEK_SET );
    				//deter__min beginning record number of BIN file to write to
    				iikmx = Nint((kmx - beglog)/xdim) + 1;
    				iikmy = Nint((endlat - kmy)/ydim) + 1;
    				numrc = (iikmy - 1) * numstepkmx + iikmx - 1;
    				bytesread = fread( &uni5[numrc + 1].i5, numstepkmxt[iter] * sizeof(struct unich), 1, stream );
    				numrctot += numstepkmxt[iter];
    			}

//   			//update progress in reading tiles
	   			m_progresspercent2 = (int)(ii / (float)(numstepkmyt[iter])) * (100/(totnumtiles-1));
	   			if ( m_progresspercent2 != oldpercent2 )
   				{
					oldpercent2 = m_progresspercent2;
					fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 250px; left: 45%;\" >" );
					fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
					fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
					fprintf(stdout, "%s\n", "</canvas>" );

					fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
					fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
					fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
					fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );

					fprintf(stdout, "%s\n", "ctx1.strokeStyle = \"#000000\";" );
					fprintf(stdout, "%s\n", "ctx1.lineWidth = 2;" );
					fprintf(stdout, "%s\n", "ctx1.strokeRect(0, 0, 101, 26);" );

					fprintf(stdout, "%s\n", "ctx1.font = 'italic 10px sans-serif';" );
					fprintf(stdout, "%s\n", "ctx1.textBaseline = 'top';" );
					/////////////////////////////////////////////////////////////////////////////////////////////////////////////
					fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#0040FF\";" );
					//fprintf(stdout, "%s\n", "ctx1.clearRect(1, 1, 101, 26);" );
					fprintf(stdout, "%s%d%s\n", "ctx1.fillRect(1,1,", m_progresspercent + m_progresspercent2,",24);" );
					fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#000000\";" );
					fprintf(stdout, "%s%d%s\n", "ctx1.fillText('", m_progresspercent + m_progresspercent2, "%', 45, 5);" );
					/////////////////////////////////////////////////////////////////////////////
					fprintf(stdout, "%s\n", "</script>" );
					fprintf(stdout, "%s\n", "</div>" );
				}
			}
		}
		fclose( stream );

		//erase progress bar
		fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 190px; left: 50%;\" >" );
		fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
		fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
		fprintf(stdout, "%s\n", "</canvas>" );

		fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
		fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
		fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
		fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );
		fprintf(stdout, "%s\n", "</script>" );
		fprintf(stdout, "%s\n", "</div>" );


		if (!(directx)) //calculate profile and write the big-endian heights to the BIN file if flagged
			{
			//int ier = Profile(lat0,lon0,hgt0,ang,aprn,mode,modeval,
	        //                  beglog,endlog,beglat,endlat );
			//allocate memory for prof array
			free( prof ); //free the memory allocated before for prof, and reallocate
			prof = (union profile *) malloc((2 * ang * invAzimuthStep + 10) * sizeof(union profile));

			//allocate memory for dtms array
			free( dtms ); //free the memory allocated before for dtms, and reallocate
			dtms = (union dtm *) malloc((14000) * sizeof(union dtm));

			//check that memory was properly allocated to all the arrays
			if ( prof == NULL || dtms == NULL )
			{
				//CString lpszText = "Can't allocate enough memory!";
				//reply = AfxMessageBox( lpszText, MB_OK | MB_ICONSTOP | MB_APPLMODAL );
                //<<<<<<<<<<<<<<<<<<<<<<error message>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
				//deallocate memory
				try
				{
					delete [] lg1chk;
					delete [] lt1chk;
					delete [] kmxb1;
					delete [] kmxb2;
					delete [] kmyb1;
					delete [] kmyb2;
					delete [] numstepkmyt;
					delete [] numstepkmxt;
					delete [] filt1;
				}
				catch ( short ier )
				{
					ier = 0;
				}
				
				//OnCancel();
			}

			if (calcProfile == 1) goto Profiles;
endProfiles:
			//don't write the binary file anymore
			////address of the beginning of the data array is uni5[1]
            //byteswritten = fwrite(&uni5[1].i5, numrctot * 2L, 1, stream2 );
		    //fclose( stream2 );
			;
			}


		//deallocate memory
		free ( uni5 );
		//free ( heights );
		try
		{
			delete [] lg1chk;
			delete [] lt1chk;
			delete [] kmxb1;
			delete [] kmxb2;
			delete [] kmyb1;
			delete [] kmyb2;
			delete [] numstepkmyt;
			delete [] numstepkmxt;
			//delete [] filt1;
		}
		catch ( short ier )
		{
			ier = 0;
		}

return ier_max; //end of readDTM

/**************findtile*****************/
findtile:
	double longitude, latitude;
	int itnum1,itnum2,uniquetilesfound;
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
				//_itoa( lg11, lg1ch, 10);
				//sprintf(lg1ch, "%d", lg11);
				itoa( lg11, lg1ch, 10);
			}
			else
			{
				//strncpy(lg1ch, "0", 1);
				//strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);

				//sprintf(lg1ch, "%s%s", "0", (const char *)_itoa( lg11, lch, 10));
				sprintf(lg1ch, "%s%s", "0", itoa( lg11, lch, 10));
			}
			if (lg1 < 0)
			{
				//strncpy( EW, "W", 1 );
				sprintf( EW, "%s", "W");
			}
			else
			{
				//strncpy( EW, "E", 1 );
				sprintf( EW, "%s", "E");
			}
			ny = (int)((90 - latitude) * 0.02);
			lt1 = 90 - 50 * ny;
			lt11 = abs(lt1);
			//_itoa( lt11, lt1ch, 10);
			itoa( lt11, lt1ch, 10);
			//sprintf(lg1ch, "%d", lt11);
			if (lt1 > 0)
			{
				//strncpy( NS, "N", 1 );
				sprintf( NS, "%s", "N");
			}
			else
			{
				//strncpy( NS , "S", 1 );
				sprintf( NS, "%s", "S");
			}
			//strncpy( filt1[itnum], (const char *)EW, 1 );
			//strncpy( filt1[itnum] + 1, (const char *)lg1ch, 3 );
			//strncpy( filt1[itnum] + 4, (const char *)NS, 1 );
			//strncpy( filt1[itnum] + 5, (const char *)lt1ch, 2 );
			//strncpy( filt1[itnum] + 7, "\0", 1);
			sprintf( &filt1[itnum][0], "%s%s%s%s%s", (const char *)EW, (const char *)lg1ch, (const char *)NS, (const char *)lt1ch, "\0" );
			////strncpy( filnn, (const char *)worlddtm, 1);
			//strncpy( filnn + 1, ":\\", 2 );
			//strncpy( filnn + 3, (const char *)filt1[itnum], 7 );
			//strncpy( filnn + 10, "/", 1 );
			//strncpy( filnn + 11, (const char *)filt1[itnum], 7 );
			sprintf( filnn, "%s%s%s%s", (const char *)drivdtm, &filt1[itnum][0], "/", &filt1[itnum][0]);
			nrows = 6000;
			ncols = 4800;
			whichCDs[itnum] = worldCD[ny * 9 + (int)ntmp + 1];

			//deter__min kmx,kmy boundaries of analysis within this tile
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
				//_itoa( lg11, lg1ch, 10);
				itoa( lg11, lg1ch, 10);
				//sprintf(lg1ch, "%d", lg11);
				}
			else if (( lg11 < 100 ) && ( lg11 != 0 ))
				{
				//strncpy(lg1ch, "0", 1);
				//strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);

				//sprintf(lg1ch, "%s%s", "0", (const char *)_itoa( lg11, lch, 10) );
				sprintf(lg1ch, "%s%s", "0", itoa( lg11, lch, 10) );
				}
			else if (lg11 == 0 )
				{
				//strncpy(lg1ch, "000", 3);
				sprintf(lg1ch, "%s", "000");
				}
			if (lg1 <= 0)
				{
				//strncpy( EW, "W", 1 );
				sprintf( EW, "%s", "W");
				}
			else
				{
				//strncpy( EW, "E", 1 );
				sprintf( EW, "%s", "E");
				}
			//strncpy( NS, "S", 1 );
			sprintf( NS, "%s", "S");
			lt1 = -60;
			//strncpy(lt1ch, "60", 2);
			sprintf(lt1ch, "%s", "60");
			//strncat( filt1[itnum], EW, 1 );
			//strncat( filt1[itnum], lg1ch, 3 );
			//strncat( filt1[itnum], NS, 1 );
			//strncat( filt1[itnum], lt1ch, 2 );
			//strncpy( filt1[itnum] + 7, "\0", 1);
			sprintf( &filt1[itnum][0], "%s%s%s%s%s", EW, lg1ch, NS, lt1ch, "\0" );
			////strncpy( filnn, (const char *)worlddtm, 1);
			//sprintf( filnn, "%s", drivdtm);
			//strncpy( filnn + 1, ":\\", 2 );
			//strncpy( filnn + 3, (const char *)filt1[itnum], 7 );
			//strncpy( filnn + 10, "/", 1 );
			//strncpy( filnn + 11, (const char *)filt1[itnum], 7 );
			sprintf( filnn, "%s%s%s%s", drivdtm, &filt1[itnum][0], "/", &filt1[itnum][0] );
			nrows = 3600;
			ncols = 7200;
			whichCDs[itnum] = 5;

			//deter__min kmx,kmy boundaries of analysis within this tile
			kmxb1[itnum] = __max(lg1, beglog);
			kmxb2[itnum] = __min(lg1 + 60, endlog);
			kmyb1[itnum] = __min(lt1 ,endlat);
			kmyb2[itnum] = __max(lt1 - 30, beglat);

		}

		lt1chk[itnum] = lt1;
		lg1chk[itnum] = lg1;

        //deter__min number of steps in this tile
		numstepkmyt[itnum] = Nint((fabs(kmyb1[itnum] - kmyb2[itnum])) / ydim) + 1;
		numstepkmxt[itnum] = Nint((fabs(kmxb1[itnum] - kmxb2[itnum])) / xdim) + 1;

		if ( itnum == 0 )
			{
			//strncpy( filn, filnn, 18 );
			sprintf( filn, "%s", filnn);
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
			    //if (_strnicoll( (const char *)filt1[itcheck], (const char *)filt1[itnum], 7) == 0 )
			    if (strncmp(&filt1[itcheck][0], &filt1[itnum][0], 7) == 0 )
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
				strncpy(&filt1[uniquetilesfound][0], &filt1[itcheck][0], 8);
				numstepkmyt[uniquetilesfound] = numstepkmyt[itcheck];
				numstepkmxt[uniquetilesfound] = numstepkmxt[itcheck];
				kmxb1[uniquetilesfound] = kmxb1[itcheck];
				kmxb2[uniquetilesfound] = kmxb2[itcheck];
				kmyb1[uniquetilesfound] = kmyb1[itcheck];
				kmyb2[uniquetilesfound] = kmyb2[itcheck];
				}
			}


}

else if (DTMflag > 0) //SRTM data, deter__min number of tiles

{
	  //now deter__min the tile names
	  itnum = 0;
	  for ( itnum1 = 0; itnum1 <= abs(lg2SRTM - lg1SRTM); itnum1++ )
	  {
		  for ( itnum2 = 0; itnum2 <= abs(lt2SRTM - lt1SRTM); itnum2++ )
		  {
			lg1 = lg1SRTM + itnum1;
			lg11 = abs(lg1);

			if (lg11 >= 100)
			{
				//_itoa( lg11, lg1ch, 10);
				itoa( lg11, lg1ch, 10);
				//sprintf(lg1ch, "%d", lg11);
			}
			else if (lg11 >= 10 && lg11 < 100 )
			{
				//strncpy(lg1ch, "0", 1);
				//strncpy(lg1ch + 1, (const char *)_itoa( lg11, lch, 10), 2);

				//sprintf(lg1ch, "%s%s", "0", (const char *)_itoa( lg11, lch, 10) );
				sprintf(lg1ch, "%s%s", "0", itoa( lg11, lch, 10) );
			}
			else if (lg11 < 10 )
			{
				//strncpy(lg1ch, "00", 2);
				//strncpy(lg1ch + 2, (const char *)_itoa( lg11, lch, 10), 1);

				//sprintf(lg1ch, "%s%s", "00", (const char *)_itoa( lg11, lch, 10) );
				sprintf(lg1ch, "%s%s", "00", itoa( lg11, lch, 10) );
			}


			if (lg1 < 0) //<--check out that name convention for lg1=0
						 //is E000 and not W000 -- otherwise use "<="
			{
				//strncpy( EW, "W", 1 );
				sprintf( EW, "%s", "W");
			}
			else
			{
				//strncpy( EW, "E", 1 );
				sprintf( EW, "%s", "E");
			}

			lt1 = lt2SRTM + itnum2;
			lt11 = abs(lt1);

			if (lt11 >= 10 )
			{
			   //_itoa( lt11, lt1ch, 10);
			   //sprintf(lt1ch, "%d", lt11);
			   itoa( lt11, lt1ch, 10);
			}
			else
			{
        		//strncpy(lt1ch, "0", 1);
				//strncpy(lt1ch + 1, (const char *)_itoa( lt11, lch, 10), 1);

				//sprintf(lt1ch, "%s%s", "0", (const char *)_itoa( lt11, lch, 10) );
				sprintf(lt1ch, "%s%s", "0", itoa( lt11, lch, 10) );
			}

			if (lt1 >= 0)
			{
				//strncpy( NS, "N", 1 );
				sprintf( NS, "%s", "N");
			}
			else
			{
				//strncpy( NS , "S", 1 );
				sprintf( NS, "%s", "S");
			}

			//strncpy( filt1[itnum], (const char *)NS, 1 );
			//strncpy( filt1[itnum] + 1, (const char *)lt1ch, 2 );
			//strncpy( filt1[itnum] + 3, (const char *)EW, 1 );
			//strncpy( filt1[itnum] + 4, (const char *)lg1ch, 3 );
			//strncpy( filt1[itnum] + 7, "\0", 1);
			sprintf( &filt1[itnum][0], "%s%s%s%s%s", (const char *)NS, (const char *)lt1ch, (const char *)EW, (const char *)lg1ch, "\0" );
			////strncpy( filnn, (const char *)worlddtm, 1);
			//sprintf( filnn, "%s", drivdtm);
			//strncpy( filnn + 1, ":\\", 2 );

			if (DTMflag == 1) // 30 meter
			{
				nrows = 3601;
				ncols = 3601;
				////strncpy( filnn + 3, "USA", 3 );
				////strncpy( filnn + 3, "dtm", 3 );
				////strncpy( filnn + 6, "/", 1 );
				////strncpy( filnn + 7, (const char *)filt1[itnum], 7 );
				sprintf( filnn, "%s%s", drivdtm, &filt1[itnum][0] );
			}
			else if (DTMflag == 2) //100 meter
			{
				nrows = 1201;
				ncols = 1201;
				////strncpy( filnn + 3, "3AS", 3 );
				////strncpy( filnn + 6, "/", 1 );
				//strncpy( filnn + 7, (const char *)filt1[itnum], 7 );
				sprintf( filnn, "%s%s", drivtas, &filt1[itnum][0] );
			}

			lt1chk[itnum] = (int)(lt1 + 1.0); //SRTM DTM is named by SW corner
									   //unlike GTOPO30 which is named by NW corner

            lg1chk[itnum] = lg1;

			//deter__min kmx,kmy boundaries of analysis within this tile
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
				//strncpy( filn, filnn, 14 );
				sprintf( filn, "%s", filnn );
				}

			maxranget += numstepkmyt[itnum] * numstepkmxt[itnum];
			itnum++;
		  }
	  }
}

    //current tile name
	//strncpy( filt, filt1[0], 7 );
	//sprintf( &filt[0], "%s", &filt1[0][0] );
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
            sprintf( DTMfile, "%s%s%s", filn, ".DEM", "\0" );
		    //strncpy( DTMfile, filn, 18 );
		    //strncpy( DTMfile + 18, ".DEM", 4 );
		    //strncpy( DTMfile + 22, "\0", 1);
		}
		else //SRTM DTM data
		{
            sprintf( DTMfile, "%s%s%s", filn, ".hgt", "\0");

			//strncpy( DTMfile, filn, 14 );
			//strncpy( DTMfile + 14, ".hgt", 4 );
			//strncpy( DTMfile + 18, "\0", 1);
		}

retry:	if ( (stream = fopen( (const char *)DTMfile, "rb" )) == NULL)
		if ( stream == NULL )
			{
			if (DTMflag == 0)
				{

				//erase progress bar
				fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 190px; left: 50%;\" >" );
				fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
				fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
				fprintf(stdout, "%s\n", "</canvas>" );

				fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
				fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
				fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
				fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );
				fprintf(stdout, "%s\n", "</script>" );
				fprintf(stdout, "%s\n", "</div>" );

				return -9; //DTM tile not found, return error code

				// tell user which CD it is on
				//strncpy( quest, "Please insert USGS EROS CD#", 27 );

				/*
				sprintf( quest, "%s", "Please insert USGS EROS CD#" );
				strncpy( quest + 27, (const char *)_itoa( numCD, CDbuffer, 10 ), 1 );
				strncpy( quest + 28, "\0", 1);
				*/

				//reply = AfxMessageBox( (const char *)quest, MB_OKCANCEL | MB_ICONINFORMATION | MB_APPLMODAL );
				//<<<<<<<<<<<<<<error message>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
				if ( reply == 2 ) // Cancel button selected
					{

						//deallocate memory
						try
						{
							delete [] lg1chk;
							delete [] lt1chk;
							delete [] kmxb1;
							delete [] kmxb2;
							delete [] kmyb1;
							delete [] kmyb2;
							delete [] numstepkmyt;
							delete [] numstepkmxt;
							delete [] filt1;
						}
						catch (short ier)
						{
							ier = 0;
						}

					//OnCancel();
					}
				else
					{
					goto retry;
					}
				}
			else if (DTMflag > 0) //warn user
				{
				if (!IgnoreMissingTiles)
					{

					//erase progress bar
					fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 190px; left: 50%;\" >" );
					fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
					fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
					fprintf(stdout, "%s\n", "</canvas>" );

					fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
					fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
					fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
					fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );
					fprintf(stdout, "%s\n", "</script>" );
					fprintf(stdout, "%s\n", "</div>" );
		
					return -9; //return error code since missing tile found

                    sprintf( quest, "%s%s%s%s%s", "Can't find the SRTM tile: ", DTMfile, "\n", "Do you wan't to use an empty tile instead?", "\0" );
                    //strncpy( quest, "Can't find the SRTM tile: ", 26 );
					//strncpy( quest + 26, (const char *)DTMfile, 18 );
					//strncpy( quest + 44, "\n", 1);
					//strncpy( quest + 45, "Do you wan't to use an empty tile instead?", 42);
					//strncpy( quest + 87, "\0", 1);
					//reply = AfxMessageBox( (const char *)quest, MB_YESNOCANCEL | MB_ICONINFORMATION | MB_APPLMODAL );
					//<<<<<<<<<<<<<<<<<<<<<<<<<<error message>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
					}
				else if (IgnoreMissingTiles) //automatically use default empty tile
					{
					reply = 6;
					}

				if (reply == 6) //Replace the missing srtm file with an empty tile
					{
					//goto retry;
						if (DTMflag == 1)
						{
						//strncpy( DTMfile, (const char *)worlddtm, 1);
						sprintf( DTMfile, "%s%s%s", drivdtm, "Z000000.hgt", "\0");
    					//strncpy( DTMfile + 1, ":\\USA\\Z000000.hgt\0", 18);
						goto retry;
						}
						else if (DTMflag == 2)
						{
						//strncpy( DTMfile, (const char *)worlddtm, 1);
						sprintf( DTMfile, "%s%s%s", drivtas, "Z000000.hgt", "\0");
    					//strncpy( DTMfile + 1, ":\\3AS\\Z000000.hgt\0", 18);
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
						try
						{
							delete [] lg1chk;
							delete [] lt1chk;
							delete [] kmxb1;
							delete [] kmxb2;
							delete [] kmyb1;
							delete [] kmyb2;
							delete [] numstepkmyt;
							delete [] numstepkmxt;
							delete [] filt1;
						}
						catch (short ier)
						{
							ier = 0;
						}

					//OnCancel();
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

/*end of timer(calculation) routine*/

/////////////////Profiles inline routine///////////////////
//int ier = Profile(lat0,lon0,hgt0,ang,aprn,mode,modeval,
//                  beglog,endlog,beglat,endlat );

//int Profile( double kmyo,double kmxo, double hgt, double mang, double aprn,double nnetz,double modval,
//		double lgo,double endlg, double beglt, double endlt)

Profiles:

	
	short iTK = 0;
	double TRpart = 0.0;
	double Tground = 0.0;
	double Exponent = 0.0;
	double PATHLENGTH = 0.0;
	double TRtemp, trrbasis,pathlength,exponent;

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

	if (hgt == 0) //user probably wants to use the NED height,
				 //so query the NED height
	{
		kmx = kmxo;
		ikmx = Nint((kmx - beglog) / dstpointx) + 1;
		kmy = kmyo;
		ikmy = Nint((lto - kmy) / dstpointy) + 1;
		numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1; //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

		if (numrc >= 0 && numrc <= numtotrec )
		{
			hgt = (float)height(numrc);
			if (hgt <= -9999 ) hgt = 0; //null height corresponding to no-data ocean pixel in NED tile
			eroshgt = hgt;
		}
		else //the required tile hasn't been processed, so have to keep zero height
		{
			hgt = 0;
			eroshgt = hgt;
		}
		hgt0 = hgt;
	}
	
	/*
	else //test the height to see if it falls within an existing tile, and if the inputed height is close to the tile's height
	{
		double hgt_test = 0;
		kmx = kmxo;
		ikmx = Nint((kmx - beglog) / dstpointx) + 1;
		kmy = kmyo;
		ikmy = Nint((lto - kmy) / dstpointy) + 1;
		numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1; //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

		if (numrc >= 0 && numrc <= numtotrec )
		{
			hgt_test = (float)height(numrc);
			if (fabs(hgt - hgt_test) > MAX_HEIGHT_DISCREP ) //large discrepency
			{
				return -7;
			}
		}
		else //can't find the tile containing the origin, this means that this country is not supported
		{
			return -8;
		}
	}
	*/

    //numcol = (int) ((endlog - beglog) / dstpointx) + 1;
    //numrow = (int) ((lto - beglat) / dstpointy) + 1;

    if (mang != 0.) {
	maxang = mang;
    }
    if (maxang > 80.) {
	maxang = 80.;
    }
    if (maxang < 24.) {
	maxang = 24.;
    }
    maxang0 = maxang;
	*kmaxang = maxang0; //reset the estimate of maxang

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

	// now set optimization range //<--not using this in these calculations, rather calculated proper maxang for each latitude
    optang = (short)maxang; //(short) (maxang - 12);
    numskip = 1;


    kmyoo = kmyo;
    kmxoo = kmxo;
    lt = kmyoo;
    lg = kmxoo;
    hgt += hgtobs;


	if (VDW_REF) {
		//deter__min the mean ground temperature to be used for calculating the terrestrial refraction

		//calculate conserved portion of terrain refraction based on Wikipedia's expression
		//lapse rate is set at -0.0065 K/m, Ground Pressure = *p = 1013.25 mb
		TRpart = 8.15 * *p * 1000.0 * (0.0342 - 0.0065)/(*meantemp * *meantemp * 3600); //units of deg/km

	}
    
    //now update the height used for calculations
    *khgt = hgt;


	if (hgt >= 0.) {
	    anav = sqrt(hgt) * -.0348f;
	}
	if (hgt < 0.) {
	    anav = -1.f;
	}

	for ( i__ = 0; i__ < nomentr; ++i__) {
		prof[i__].ver[0] = (float)anav - minview;}

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

	numrctot = 0;
	oldpercent = 0;
	oldpercent2 = 0;

    nbegkmx = Nint(begkmx / dstpointx);
    begkmx = nbegkmx * dstpointx;
    skipkmx = dstpointx;
    skipkmy = dstpointy;
    numkmx = Nint(fabs(endkmx - begkmx)/skipkmx) + 1;

    re = 6371315.;
    rekm = 6371.315f;
    i__1 = numkmx;
    i__2 = numskip;
    for (i__ = nkmx1; i__2 < 0 ? i__ >= i__1 : i__ <= i__1; i__ += i__2) {


		kmx = begkmx + (i__ - 1) * skipkmx;
		ikmx = Nint((kmx - beglog) / skipkmx) + 1;

		if (ikmx == 0) {
			goto L550;
		}
/*      now deter__min optimized kmy angular range */
/*      first convert longitude range into kilometers */
		range = fabs(kmx - kmxo) * cd * rekm * cos(kmyo * cd);
		if (nnetz == 1 || SRTMflag == 11) {
/*          if range >= 20 km, then narrow azimuthal range to MAXANG=OptAng */

			//only good for mid latitudes
			/*
			if (range >= 20. && maxang > (double) optang)
			{
				maxang = (double) optang;
			}
			*/

/*          if range >= 40 km and SRTM-1, then reduce step size to ~ 60m */
			if (range > 40. && DTMflag == 1 && SRTMflag != 11)
			{
				numskip = 2;
			}

			if (SRTMflag == 11 && DTMflag == 1) //further optimizations to reduce server cpu usage
			{
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

		} else if (nnetz == 0 || SRTMflag == 10) {
/*          if range >= 20 km, then narrow azimuthal range to MAXANG=OptAng */

			/*
			if (range > 20. && maxang > (double) optang)
			{
				maxang = (double) optang;
			}
			*/

/*          if range >= 40 km and SRTM-1, then increase step size to ~ 60 m */
			if (range > 40. && DTMflag == 1 && SRTMflag != 10)
			{
				numskip = 2;
			}

			if (SRTMflag == 10 && DTMflag == 1) //further optimizations in lieu of unnecessary accuracy to reduce server cpu usage
			{
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
			if (range <= 40. && DTMflag == 1) {
				numskip = 1;}
		}

/*      optimized half angular range in latitude (degrees) = dkmy */
		dkmy = range / (rekm * cd) * tan((maxang + 2) * cd);
		//numkmy = (Nint(dkmy / skipkmy) << 1) + 1;
		d__1 = dkmy / skipkmy;
		numkmy = 2 * Nint(d__1) + 1;
		numhalfkmy = Nint(d__1) + 1;
		numdtm = (int) ((maxang0 - maxang) * 10);
		i__3 = numkmy;
		i__4 = numskip;
		for (j = 1; i__4 < 0 ? j >= i__3 : j <= i__3; j += i__4) {


			kmy = kmyo - (numhalfkmy - j + 1) * skipkmy;
			if (kmy < beglat || kmy > endlat) {
			goto L120;  }


			//further optimization: use crude approximation of azimuth to
			//deter__min if data will lie within the required azimuth range
			if (kmx != lg)
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
			numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1; //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
			if (numrc >= 0 && numrc <= numtotrec )
			{
				//hgt2 = (float)(uni5[numrc].chi5[0] << 8 | uni5[numrc].chi5[1]);
				//hgt2 = ((uni5[numrc].i5 >> 8) & 0xff) + ((uni5[numrc].i5 << 8) & 0xf00);

				hgt2 = (float)height(numrc);
				if (hgt2 <= -9999 ) hgt2 = 0; //null height corresponding to no-data ocean pixel in NED tile
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
			/*
			if (hgt2 == -9999. || hgt2 == -32768. || fabs(hgt2) > 8848.) {
			hgt2 = 0.f;}
			*/
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

			if (!VDW_REF) {
				// add old contribution of atmospheric refraction to view angle/

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

			} else {

				//use terrestrial refraction formula from Wikipedia, suitably modified to fit van der Werf ray tracing
				PATHLENGTH = sqrt(distd*distd + pow(fabs(deltd)*0.001 - 0.5*(distd*distd/rearth),2.0));
				if (fabs(deltd) > 1000.0) {
					Exponent = 0.99;
				}else{
					Exponent = 0.9965;
				}

				avref = TRpart * pow(PATHLENGTH, Exponent); //degrees

			}

			++numdtm;
			dtms[numdtm].ver[0] = azi;
			dtms[numdtm].ver[1] = viewang / cd + avref;
			dtms[numdtm].ver[2] = kmy;

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
							prof[ne].ver[0] = (float)vang;
							prof[ne].ver[1] = (float)(begkmx + (i__ - 1) * skipkmx);
							prof[ne].ver[2] = (float)((kmy2 - kmy1) / (az2 - az1) * (delazi - az1) + kmy1);
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

				/////////////////draw progress bar using canvas//////////////////////////////
				////////////////////////////////////////////////////////////////////////////////////////////////////////////
				fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 250px; left: 45%;\" >" );
				fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
				fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
				fprintf(stdout, "%s\n", "</canvas>" );

				fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
				fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
				fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
				fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );

				fprintf(stdout, "%s\n", "ctx1.strokeStyle = \"#000000\";" );
				fprintf(stdout, "%s\n", "ctx1.lineWidth = 2;" );
				fprintf(stdout, "%s\n", "ctx1.strokeRect(0, 0, 101, 26);" );

				fprintf(stdout, "%s\n", "ctx1.font = 'italic 10px sans-serif';" );
				fprintf(stdout, "%s\n", "ctx1.textBaseline = 'top';" );
				/////////////////////////////////////////////////////////////////////////////////////////////////////////////
				fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#99CC00\";" );
				//fprintf(stdout, "%s\n", "ctx1.clearRect(1, 1, 101, 26);" );
				fprintf(stdout, "%s%d%s\n", "ctx1.fillRect(1,1,", m_progresspercent,",24);" );
				fprintf(stdout, "%s\n", "ctx1.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s%d%s\n", "ctx1.fillText('", m_progresspercent, "%', 45, 5);" );
				/////////////////////////////////////////////////////////////////////////////
				fprintf(stdout, "%s\n", "</script>" );
				fprintf(stdout, "%s\n", "</div>" );

			}

		//fprintf( stdout, "%s,%d,%lg,%lg\n", " completed interpreting: ",i__, kmx, kmy );
L550:
		;
    }

	//erase progress bar
	fprintf(stdout, "%s\n", "<div align = \"center\" style=\"position: absolute; top: 190px; left: 50%;\" >" );
	fprintf(stdout, "%s\n", "<canvas id=\"profiles\" width=\"101\" height=\"26\" >" ); // style="border:1px solid #000000;">" );
	fprintf(stdout, "%s\n", "<p>Progress can't be displayed since your browser doesn't support canvas.</p>" );
	fprintf(stdout, "%s\n", "</canvas>" );

	fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
	fprintf(stdout, "%s\n", "var canvas1 = document.getElementById('profiles');" );
	fprintf(stdout, "%s\n", "var ctx1 = canvas1.getContext('2d');" );
	fprintf(stdout, "%s\n", "ctx1.clearRect(0, 0, 101, 26);" );
	fprintf(stdout, "%s\n", "</script>" );
	fprintf(stdout, "%s\n", "</div>" );


    //no file output for this version of the program
	if (!CGI || Windows)
	{
		if ( f =fopen("C:/JK_C/EROS.TMP", "w" ) )
		{
			if (!VDW_REF) {
				fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR" );
			}else{
				fprintf( f, "%s\n", "Lati,Long,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR,1" );
			}
			fprintf( f, "%lg,%lg,%lg,%lg,%lg,%lg,%lg,%lg\n", lat0, -lon0, hgt0, -startkmx, -sofkmx, dstpointx, dstpointy, apprnr); //longitude sign convention
		}
	}


/* 260     FORMAT(F9.4,2X,F9.4,2X,F8.2,4(3X,F8.3),3X,F5.3) */
    i__2 = nomentr;
    for (nk = 1; nk <= i__2; nk++) {
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
		numrc = (ikmy - 1) * numstepkmx + ikmx; // - 1; //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		hgt2 = 0;
		if (numrc >= 0 && numrc <= numtotrec )
		{
			hgt2 = (float)height(numrc);
			if (hgt2 <= -9999) hgt2 = 0; //null pixel ocean void in NED tile
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

		if (!CGI || Windows) //print out output
		{
			fprintf( f, "%lg,%lg,%lg,%lg,%lg,%lg\n", xentry, prof[nk].ver[0], prof[nk].ver[1], prof[nk].ver[2], dist, hgt2);
		}

		//else //store in elev array and draw graph on canvas
		///////////////////diagnositcs -continue with calculating the zemanim///////////////////////////
		{
			//////////////////////////remove temporary terrestrial refraction//////////////////////////
			if (!VDW_REF) {
				//remove old terrestrial refraction values from the view angle 
				deltd = hgt - hgt2;
				distd = dist;
				if (deltd <= 0.f) {
					defm = 7.82e-4 - deltd * 3.11e-7;
					defb = deltd * 3.4e-5 - .0141;
				} else if (deltd > 0.f) {
					defm = deltd * 3.09e-7 + 7.64e-4;
					defb = -.00915 - deltd * 2.69e-5;
				}
				avref = defm * distd + defb;
				if (avref < 0.f) {
					avref = 0.f;
				}
				TRtemp = avref;

			} else {

				deltd = hgt - hgt2;
				distd = dist;
				//calculate conserved portion of Wikipedia's terrestrial refraction expression 
				//Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere
				//so have p * 8.15 * 1000 * (0.0343 - 0.0065)/(T * T * 3600) degrees
				trrbasis = *p * 8.15f * 1e3f * .0277f / (*meantemp * *meantemp * 3600);
				//in the winter, dt/dh will be positive for inversion layer and refraction is increased
				//now calculate pathlength of curved ray using Lehn's parabolic approximation 
				d__1 = distd;
				d__3 = distd;
				d__2 = abs(deltd) * .001f - d__3 * d__3 * .5f / rearth;
				pathlength = sqrt(d__1 * d__1 + d__2 * d__2);
				//now add terrestrial refraction based on modified Wikipedia expression */
				if (abs(deltd) > 1e3f) {
					exponent = .99f;
				} else {
					exponent = .9965f;
				}
				TRtemp = trrbasis * pow(pathlength, exponent);
			}
			///////////////////////////////////////////////////////

			elev[nk - 1].vert[0] = xentry; //azimuth (deg)
			elev[nk - 1].vert[1] = prof[nk].ver[0] - TRtemp; //view angle (deg) (after removing temporary TR)
			elev[nk - 1].vert[2] = dist; //distance to obstruction (km)
			elev[nk - 1].vert[3] = hgt2; //height of obstruction (m)

			if (dist < distlim) NearObstructions = true;

//			if (fabs(prof[nk].ver[0]) < optang - 2 ) //find min, max for graph within graphed range
//			{
				vang_max = __max(vang_max, prof[nk].ver[0] );
				vang_min = __min(vang_min, prof[nk].ver[0] );

				if (GoldenLight == 1) {
					//record nk position of largest view angle
					//also record the distance to the horizon at this anlge
					hgt_max = __max(hgt_max, elev[nk - 1].vert[3] );
					
					//if (hgt_max == elev[nk - 1].vert[3]) {
					if (vang_max == prof[nk].ver[0]) {
						nk_max = nk; //record where the maximum occurs
						hgt_max = elev[nk - 1].vert[3];
						dist_max = elev[nk - 1].vert[2];
					}
				}
//			}
		}

/* L300: */
    }

	if (GoldenLight == 1) {
		//not clear what portion of the hozion will be lit up first, ifit is neaest highest point
		//or the highest point in the horizon no matter how far away.  Since the latter might be too far
		//away to actually see very soon, the first criteria is used.
		//However, if the near point is extremely close, i.e., < 1 km, then an error message is returned

		//dist_max = elev[nk_max - 1].vert[2];

		if (dist_max < MinGoldenDist || vang_max > MaxGoldenViewAng) {
			if (SRTMflag == 10) {
				//Eastern horizon too close to accuratly deter__min Golden Light horizon for sunsets
				ier_max = -10;
			}
			else if (SRTMflag == 11) {
				//Western horizon too close to accurately deter__min Golden Light horizon for sunrises
				ier_max = -11;
			}
		}

		//record coordinates of highest point in profile to be used for calculating Golden Light table
		Gold.nkk = nk_max;
		Gold.vert[0] = elev[nk_max - 1].vert[0]; //azimuth of above point
		Gold.vert[1] = elev[nk_max - 1].vert[1]; //viewangle of above point
		Gold.vert[2] = prof[nk_max].ver[1]; //longitude of above point
		Gold.vert[3] = prof[nk_max].ver[2]; //latitude of above point
		Gold.vert[4] = elev[nk_max - 1].vert[2]; //distance to above point
		Gold.vert[5] = elev[nk_max - 1].vert[3]; //height of above point
	}

	if (!CGI || Windows)
	{
		fclose( f );
	}
	//else //draw graph on canvas
	{
		PlotGraphHTML5( nomentr, vang_max, vang_min, NearObstructions, optang);
	}

   goto endProfiles;
}

/*********************************************************************************/
/////////////////////////////////////////////
//	  emulation of Fortran 95 Nint(double x)
/////////////////////////////////////////////////////////////
//source: http://www.lahey.com/docs/lfprohelp/F95ARNINTFn.htm
////////////////////////////////////////////////////////////
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

	return (int)x;

}

/**********************************************************************************/
////////////////PlotGraphHTML5/////////////////
void PlotGraphHTML5(short nomentr, double vang_max, double vang_min, bool NearObstructions, short optang)
/////////////////////////////////////////////
//plots html5 canvas graph view angle vs azimuth of the profile calculated by routine readDTM
////////////////////////////////////////////////////////////////////////////////////////////
{
	short nk;
	double azi, vang, dist;
	double x, y, xo, yo;
	short ycoord;

	ycoord = Nint((vang_max/(vang_max - vang_min)) * 341);

	//sprintf( buff1, "%s", "<div style=\"position: absolute; top:210px; width:300px; left:50px; height:25px\" background-color: white>Profile calculation...please wait</div>" );
	//fprintf( stdout, "%s\n", buff1);

	if (!optionheb)
	{
		fprintf( stdout, "%s\n", "<bdo dir=\"ltr\">" ); //graph is in english!
	}
	else
	{
		fprintf( stdout, "%s\n", "<bdo dir=\"rtl\">" ); //headers of graph in Hebrew!
	}

	fprintf(stdout, "%s\n", "<div align = \"center\" >" );

	//Graph header
	fprintf(stdout, "%s\n", "<table align = \"center\" >" );
	/*
	if (GoldenLight == 0) {
		//create space between header
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	}
	*/

	if (SRTMflag == 10)
	{
		if (!optionheb)
		{
			fprintf(stdout, "%s\n", "<tr><td align = \"center\" ><b>Eastern Horizon looking due East</b></td></tr>");
		}
		else
		{
			fprintf(stdout, "%s%s%s\n", "<tr><td align = \"center\" ><b>", &heb9[1][0], "</b></td></tr>");
		}
	}
	else if (SRTMflag == 11)
	{
		if (!optionheb)
		{
			fprintf(stdout, "%s\n", "<tr><td align = \"center\" ><b>Western Horizon looking due West</b></td></tr>");
		}
		else
		{
			fprintf(stdout, "%s%s%s\n", "<tr><td align = \"center\" ><b>", &heb9[2][0], "</b></td></tr>");
		}
	}
	if (!optionheb) 
	{
		fprintf(stdout, "%s\n", "<tr><td align = \"center\" >X Axis: azimuth (deg.); Y Axis: view angle (deg.)</td></tr>" );
	}
	else
	{
		fprintf(stdout, "%s%s%s\n", "<tr><td align = \"center\" >", &heb9[3][0], "</td></tr>" );
	}
	//fprintf(stdout, "%s%5.1lf%s%5.1lf%s%2.1lf%s%2.1lf%s\n", "<tr><td>Azimuth range: -45 to 45 degrees, View Angle range: ", vang_min, " to ", vang_max, ", lat hw: ", erosdiflat, " deg., lon hw: ", erosdiflog, " deg.</td></tr>" );
	
	if (!optionheb)
	{
		fprintf(stdout, "%s%2.1lf%s%2.1lf%s\n", "<tr><td align = \"center\" >latitude scan half width: ", erosdiflat, " deg., longitude scan width: ", erosdiflog, " deg.</td></tr>" );
	}
	else
	{
		fprintf(stdout, "%s%s %2.1lf %s %2.1lf %s%s\n", "<tr><td align = \"center\" >", &heb9[4][0], erosdiflat, &heb9[5][0], erosdiflog, &heb9[14][0], "</td></tr>" );
	}
	
	
	if (NearObstructions)
	{
		if (!optionheb)
		{
			fprintf(stdout, "%s%3.1f%s\n", "<tr><td align = \"center\" ><font color = \"#009900\" >(When the horizon is closer than ", distlim, " km, it is drawn in green.)</font></td></tr>");

		}
		else
		{
			fprintf(stdout, "%s%s %3.1f %s%s\n", "<tr><td align = \"center\" ><font color = \"#009900\" >,(",&heb9[8][0], distlim, &heb9[7][0],")</font></td></tr>");
		}
	}
	if (GoldenLight == 1)
	{
		if (SRTMflag == 10)
		{
			if (optionheb)
			{
				fprintf(stdout, "%s%s%s\n", "<tr><td align = \"center\" ><font color = \"#DF7401\" >", &heb9[10][0], "</font></td></tr>");
			}
			else
			{
				fprintf(stdout, "%s\n", "<tr><td align = \"center\" ><font color = \"#DF7401\" >Marked azimuth in this color is predicted position of last golden light.</font></td></tr>");
			}
		}
		else if (SRTMflag == 11)
		{
			if (optionheb)
			{
				fprintf(stdout, "%s%s%s\n", "<tr><td align = \"center\" ><font color = \"#DF7401\" >", &heb9[9][0], "</font></td></tr>");
			}
			else
			{
				fprintf(stdout, "%s\n", "<tr><td align = \"center\" ><font color = \"#DF7401\" >Marked azimuth in this color is predicted position of first golden light.</font></td></tr>");
			}
		}
	}		

	if (GoldenLight <= 1 ) {

		fprintf(stdout, "%s\n", "<canvas id=\"plot\" width=\"668\" height=\"334\" >" );
		fprintf(stdout, "%s\n", "<p>Graph can't be displayed since your browser doesn't support canvas.</p>" );
		fprintf(stdout, "%s\n", "</canvas>" );

		fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
		fprintf(stdout, "%s\n", "var canvas = document.getElementById('plot');" );

	} else if (GoldenLight == 2) {

		fprintf(stdout, "%s\n", "<canvas id=\"plot2\" width=\"668\" height=\"334\" >" );
		fprintf(stdout, "%s\n", "<p>Graph can't be displayed since your browser doesn't support canvas.</p>" );
		fprintf(stdout, "%s\n", "</canvas>" );

		fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
		fprintf(stdout, "%s\n", "var canvas = document.getElementById('plot2');" );
	}

	fprintf(stdout, "%s\n", "var ctx = canvas.getContext('2d');" );
	fprintf(stdout, "%s\n", "ctx.clearRect(0, 0, 668, 334);" );

	fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" );
	fprintf(stdout, "%s\n", "ctx.lineWidth = 3;" );
	fprintf(stdout, "%s\n", "ctx.strokeRect(0, 0, 668, 334);" );

	//add axis labels
	fprintf(stdout, "%s\n", "ctx.font = 'bold 10px sans-serif';" );
	fprintf(stdout, "%s\n", "ctx.textBaseline = 'top';" );
	fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );

	if (!optionheb) 
	{
		fprintf(stdout, "%s\n", "ctx.fillText('view angle (deg.)', 344, 10);" );
	}
	else
	{
		fprintf(stdout, "%s", "ctx.fillText('" );
		fprintf(stdout, "%s",heb9[15] );
		fprintf(stdout, "%s\n", "', 320, 10);" );
	}

	if (!optionheb)
	{
		fprintf(stdout, "%s%5.1lf%s\n", "ctx.fillText('", vang_max,"', 300, 10);" );
		fprintf(stdout, "%s%5.1lf%s\n", "ctx.fillText('", vang_min,"', 300, 310);" );
	}
	else
	{
		fprintf(stdout, "%s%5.1lf%s\n", "ctx.fillText('", vang_max,"', 370, 10);" );
		fprintf(stdout, "%s%5.1lf%s\n", "ctx.fillText('", vang_min,"', 370, 310);" );
	}

	fprintf(stdout, "%s%d%s\n", "ctx.fillText('0', 319,", ycoord - 15, ");" );

	if (SRTMflag == 10)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due East', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			if (!optionheb)
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s\n", "ctx.fillText('Eastern horizon looking due east', 488, 295);" );
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#009900\";" );
				fprintf(stdout, "%s%3.1f%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 410, 310);" );
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s", "ctx.fillText('");
				fprintf(stdout, "%s", &heb9[1][0] );
				fprintf(stdout, "%s\n", "', 288, 295);" );

				//fprintf(stdout, "%s%s%s\n", "ctx.fillText('", &heb9[1][0], "', 488, 295);" );

				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#009900\";" );

				fprintf(stdout, "%s", "ctx.fillText('(" );
				fprintf(stdout, "%s %3.1f %s", &heb9[11][0], distlim, &heb9[7][0] );
				fprintf(stdout, "%s\n", ")', 310, 310);" );


				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			}
		}
		else
		{
			if (!optionheb)
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s\n", "ctx.fillText('Eastern horizon looking due east', 488, 295);" );
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s", "ctx.fillText('" );
				fprintf(stdout, "%s", &heb9[1][0] );
				fprintf(stdout, "%s\n","', 200, 295);" );

			}
		}
	}
	else if (SRTMflag == 11)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due West', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			if (!optionheb)
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s\n", "ctx.fillText('Western horizon looking due west', 488, 295);" );
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#009900\";" );
				fprintf(stdout, "%s%3.1f%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 410, 310);" );
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );

				fprintf(stdout, "%s", "ctx.fillText('" );
				fprintf(stdout, "%s", &heb9[2][0] );
				fprintf(stdout, "%s\n", "', 288, 295);" );

				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#009900\";" );

				fprintf(stdout, "%s", "ctx.fillText('(" );
				fprintf(stdout, "%s %3.1f %s",&heb9[11][0], distlim, &heb9[7][0] );
				fprintf(stdout, "%s\n", ")', 310, 310);" );

				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			}
		}
		else
		{
			if (!optionheb) 
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s\n", "ctx.fillText('Western horizon looking due west', 488, 295);" );
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
				fprintf(stdout, "%s", "ctx.fillText('");
				fprintf(stdout, "%s",heb9[2] );
				fprintf(stdout, "%s\n", "', 288, 295);" );
			}
		}
	}

	if (ycoord > 0 && ycoord < 334)
	{
		/*
		//draw labels on mishor axis
		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", MAX_ANGULAR_RANGE * -1, "', 30,", ycoord - 15 , ");" );
		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", MAX_ANGULAR_RANGE, "', 625,", ycoord - 15 , ");" );
		fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
		fprintf(stdout, "%s%d%s\n", "ctx.fillText('mishor horizon', 15,", ycoord + 5, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.fillText('azimuth (deg.)', 55,", ycoord - 15, ");" );
		*/

		if (!optionheb)
		{
			fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 15,", 334 - 15 , ");" );
			fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640,", 334 - 15 , ");" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('azimuth (deg.)', 55,", 334 - 15, ");" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('mishor horizon', 15,", ycoord + 5, ");" );
		}
		else
		{
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 35, 319 );" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640, 319 );" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );

			fprintf(stdout, "%s", "ctx.fillText('" );
			fprintf(stdout, "%s",heb9[12] );
			fprintf(stdout, "%s\n", "', 555, 319 );" );


			fprintf(stdout, "%s", "ctx.fillText('" );
			fprintf(stdout, "%s",heb9[13] );
			fprintf(stdout, "%s%d%s\n", "', 555,", ycoord + 5, ");" );

		}
	}
	else
	{
		if (!optionheb)
		{
			fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 15,", 334 - 15 , ");" );
			fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640,", 334 - 15 , ");" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('azimuth (deg.)', 55,", 334 - 15, ");" );
		}
		else
		{
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 35, 319);" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640, 319);" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );

			fprintf(stdout, "%s", "ctx.fillText('" );
			fprintf(stdout, "%s", heb9[12] );
			fprintf(stdout, "%s\n", "', 555, 319 );" );

		}
	}

	if (ycoord > 0 && ycoord < 334)
	{
		// draw x, y axes
		fprintf(stdout, "%s\n", "ctx.beginPath();" );
		fprintf(stdout, "%s\n", "ctx.lineWidth = 1;" );
		fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#00568E\";" );
		fprintf(stdout, "%s\n", "ctx.moveTo(334, 10);" );
		fprintf(stdout, "%s\n", "ctx.lineTo(334, 319);" );
		fprintf(stdout, "%s%d%s\n", "ctx.moveTo(10,", ycoord, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.lineTo(652,", ycoord, ");" );
		fprintf(stdout, "%s\n", "ctx.stroke();" );
	}
	else
	{
		fprintf(stdout, "%s\n", "ctx.beginPath();" );
		fprintf(stdout, "%s\n", "ctx.lineWidth = 1;" );
		fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#00568E\";" );
		fprintf(stdout, "%s\n", "ctx.moveTo(334, 10);" );
		fprintf(stdout, "%s\n", "ctx.lineTo(334, 319);" );
		fprintf(stdout, "%s\n", "ctx.stroke();" );
	}

	// draw points
	fprintf(stdout, "%s\n", "ctx.save();" );
	fprintf(stdout, "%s\n", "ctx.translate(334 ,167);" );
	fprintf(stdout, "%s\n", "ctx.scale(0.9, -0.9);" ); //invert y scale to make larger y's to the top of the screen
	//fprintf(stdout, "%s%g%s%g%s\n", "ctx.scale(", vang_max, ",", vang_min, ");" ); //invert y scale to make larger y's to the top of the screen
	//fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" );
	fprintf(stdout, "%s\n", "ctx.lineWidth = 3;" );

    for (nk = 1; nk < nomentr; nk++) //<<<<<<<<<<<<<<<<<<<< should be nk <= nomentr >>>>>>>>>>>>>>>> ????
	{
		azi = elev[nk - 1].vert[0]; //azimuth
		vang = elev[nk - 1].vert[1]; //view angle
		dist = elev[nk - 1].vert[2]; //distance to obstruction

		x = 334 * azi / ((double)optang - 2.);
		y = 334 * (vang - vang_min)/(vang_max - vang_min) - 167;

		if (nk == 1)
		{
			xo = x;
			yo = y;
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.beginPath();" );
			if (dist >= distlim )
			{
				fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" ); //black for "far" mountains
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#33cc66\";" ); //green for near mountains
			}
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.moveTo(", xo, ",", yo, ");" );
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.lineTo(", x, ",", y, ");" );

			xo = x;
			yo = y;
		}

		fprintf(stdout, "%s\n", "ctx.stroke();" );

		if (GoldenLight == 1) {

			if (nk == Gold.nkk) {
				//reached azimuth of the Golden light defining point in the opposing horizon
			fprintf(stdout, "%s\n", "ctx.beginPath();" );
			fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#DF7401\";" ); //black for "far" mountains

			fprintf(stdout, "%s%g%s\n", "ctx.moveTo(", xo, ",10);" );
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.lineTo(", x, ",", yo, ");" );
			fprintf(stdout, "%s\n", "ctx.stroke();");

			}

		}

	}

	/*} else if (GoldenLight > 1) {

		short yshift = 400;


	fprintf(stdout, "%s\n", "<canvas id=\"plot\" width=\"668\" height=\"334\" >" );
	fprintf(stdout, "%s\n", "<p>Graph can't be displayed since your browser doesn't support canvas.</p>" );
	fprintf(stdout, "%s\n", "</canvas>" );

	fprintf(stdout, "%s\n", "<script language=\"JavaScript\"/>" );
	fprintf(stdout, "%s\n", "var canvas = document.getElementById('plot');" );
	fprintf(stdout, "%s\n", "var ctx = canvas.getContext('2d');" );
	fprintf(stdout, "%s%d%s\n", "ctx.clearRect(0, 0, 668,", 334 + yshift, ");" );

	fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" );
	fprintf(stdout, "%s\n", "ctx.lineWidth = 3;" );
	fprintf(stdout, "%s%d%s\n", "ctx.strokeRect(0, 0, 668,", 334 + yshift, ");" );

	//add axis labels
	fprintf(stdout, "%s\n", "ctx.font = 'bold 10px sans-serif';" );
	fprintf(stdout, "%s\n", "ctx.textBaseline = 'top';" );
	fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
	fprintf(stdout, "%s%d%s\n", "ctx.fillText('view angle (deg.)', 344,", 10 + yshift, ");" );
	fprintf(stdout, "%s%5.1lf%s%d%s\n", "ctx.fillText('", vang_max,"', 300, ", 10 + yshift, ");" );
	fprintf(stdout, "%s%5.1lf%s%d%s\n", "ctx.fillText('", vang_min,"', 300, ", 310 + yshift, ");" );
	fprintf(stdout, "%s%d%s\n", "ctx.fillText('0', 319,", ycoord - 15 + yshift, ");" );

	if (SRTMflag == 10)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due East', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			fprintf(stdout, "%s\n", "ctx.fillText('Eastern horizon looking due east', 488, ", 295 + yshift, ");" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#009900\";" );
			fprintf(stdout, "%s%3.1f%s%d%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 410, ", 310 + yshift, ");" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('Eastern horizon looking due east', 488, ", 295 + yshift, ");" );
		}
	}
	else if (SRTMflag == 11)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due West', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			fprintf(stdout, "%s\n", "ctx.fillText('Western horizon looking due west', 488, ", 295 + yshift, ");" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillStyle = \"#009900\";" );
			fprintf(stdout, "%s%3.1f%s%d%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 410, ", 310 + yshift, ");" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
			fprintf(stdout, "%s%d%s\n", "ctx.fillText('Western horizon looking due west', 488, ", 295 + yshift, ");" );
		}
	}

	if (ycoord > 0 && ycoord < 334)
	{

		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 15,", 334 - 15 + yshift , ");" );
		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640,", 334 - 15 + yshift , ");" );
		fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
		fprintf(stdout, "%s%d%s\n", "ctx.fillText('azimuth (deg.)', 55,", 334 - 15 + yshift, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.fillText('mishor horizon', 15,", ycoord + 5 + yshift, ");" );

	}
	else
	{
		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2) * -1, "', 15,", 334 - 15 + yshift , ");" );
		fprintf(stdout, "%s%d%s%d%s\n", "ctx.fillText('", (optang - 2), "', 640,", 334 - 15 + yshift , ");" );
		fprintf(stdout, "%s\n", "ctx.fillStyle = \"#00568E\";" );
		fprintf(stdout, "%s%d%s\n", "ctx.fillText('azimuth (deg.)', 55,", 334 - 15 + yshift, ");" );
	}

	if (ycoord > 0 && ycoord < 334)
	{
		// draw x, y axes
		fprintf(stdout, "%s\n", "ctx.beginPath();" );
		fprintf(stdout, "%s\n", "ctx.lineWidth = 1;" );
		fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#00568E\";" );
		fprintf(stdout, "%s\n", "ctx.moveTo(334, ", 10 + yshift, ");" );
		fprintf(stdout, "%s\n", "ctx.lineTo(334, ", 319 + yshift, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.moveTo(10,", ycoord + yshift, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.lineTo(652,", ycoord + yshift, ");" );
		fprintf(stdout, "%s\n", "ctx.stroke();" );
	}
	else
	{
		fprintf(stdout, "%s\n", "ctx.beginPath();" );
		fprintf(stdout, "%s\n", "ctx.lineWidth = 1;" );
		fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#00568E\";" );
		fprintf(stdout, "%s%d%s\n", "ctx.moveTo(334, ", 10 + yshift, ");" );
		fprintf(stdout, "%s%d%s\n", "ctx.lineTo(334, ", 319 + yshift, ");" );
		fprintf(stdout, "%s\n", "ctx.stroke();" );
	}

	// draw points
	fprintf(stdout, "%s\n", "ctx.save();" );
	fprintf(stdout, "%s%d%s\n", "ctx.translate(334 ,", 167 + yshift, ");" );
	fprintf(stdout, "%s\n", "ctx.scale(0.9, -0.9);" ); //invert y scale to make larger y's to the top of the screen
	//fprintf(stdout, "%s%g%s%g%s\n", "ctx.scale(", vang_max, ",", vang_min, ");" ); //invert y scale to make larger y's to the top of the screen
	//fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" );
	fprintf(stdout, "%s\n", "ctx.lineWidth = 3;" );

    for (nk = 1; nk < nomentr; nk++) //<<<<<<<<<<<<<<<<<<<< should be nk <= nomentr >>>>>>>>>>>>>>>> ????
	{
		azi = elev[nk - 1].vert[0]; //azimuth
		vang = elev[nk - 1].vert[1]; //view angle
		dist = elev[nk - 1].vert[2]; //distance to obstruction

		x = 334 * azi / ((double)optang - 2.);
		y = 334 * (vang - vang_min)/(vang_max - vang_min) - 167;

		if (nk == 1)
		{
			xo = x;
			yo = y;
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.beginPath();" );
			if (dist >= distlim )
			{
				fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#000000\";" ); //black for "far" mountains
			}
			else
			{
				fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#33cc66\";" ); //green for near mountains
			}
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.moveTo(", xo, ",", yo + yshift, ");" );
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.lineTo(", x, ",", y + yshift, ");" );

			xo = x;
			yo = y;
		}

		fprintf(stdout, "%s\n", "ctx.stroke();" );

		if (GoldenLight == 1) {

			if (nk == Gold.nkk) {
				//reached azimuth of the Golden light defining point in the opposing horizon

			fprintf(stdout, "%s\n", "ctx.beginPath();" );
			fprintf(stdout, "%s\n", "ctx.strokeStyle = \"#DF7401\";" ); //black for "far" mountains

			fprintf(stdout, "%s%g%s%d%s\n", "ctx.moveTo(", xo, ", ", 10 + yshift, ");" );
			fprintf(stdout, "%s%g%s%g%s\n", "ctx.lineTo(", x, ",", yo + yshift, ");" );

			}

		}
	}

	/*
	if (SRTMflag == 10)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due East', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			fprintf(stdout, "%s\n", "ctx.fillText('Eastern horizon looking due east', 408, 264);" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#33cc66\";" );
			fprintf(stdout, "%s%4.1lf%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 357, 284);" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.fillText('Eastern horizon looking due east', 317, 309);" );
		}
	}
	else if (SRTMflag == 11)
	{
		//fprintf(stdout, "%s%d%s\n", "ctx.fillText('Due West', 472,", Nint(vang_max/(vang_max - vang_min) * 540) + 15, ");" );
		if (NearObstructions) //print near obstructions warning
		{
			fprintf(stdout, "%s\n", "ctx.fillText('Western horizon looking due west', 408, 264);" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#33cc66\";" );
			fprintf(stdout, "%s%4.1lf%s\n", "ctx.fillText('(Obstructions closer than ", distlim, " km drawn in green)', 357, 284);" );
			fprintf(stdout, "%s\n", "ctx.fillStyle = \"#000000\";" );
		}
		else
		{
			fprintf(stdout, "%s\n", "ctx.fillText('Western horizon looking due west', 317, 309);" );
		}
	}
	*/
	/*

	//finish graph tags

	}
	*/

	fprintf(stdout, "%s\n", "</script>" );

	if (optionheb)
	{
		fprintf( stdout, "%s\n", "</bdo>");
	}

	if (GoldenLight == 0) {

	//paginate by using a table
	fprintf(stdout, "%s\n", "<table align = \"center\" >" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
		fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	
	fprintf(stdout, "%s\n", "</table>" );

	}
	else if (GoldenLight == 1)
	{
	fprintf(stdout, "%s\n", "<table align = \"center\" >" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "</table>" );
	}
	else if (GoldenLight == 2)
	{
	fprintf(stdout, "%s\n", "<table align = \"center\" >" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "<tr><td>&nbsp;</td></tr>" );
	fprintf(stdout, "%s\n", "</table>" );
	}

	fprintf(stdout, "%s\n", "</div>" );

	if (optionheb)
	{
		fprintf( stdout, "%s\n", "<bdo dir=\"rtl\">" );
	}


}


/////////////rotate bytes for windows//////////////////
short int height(long numrc)
///////////////////////////////////////////////////////
//extract heights from where it was written to memory
///////////////////////////////////////////////////////
{
	
	short int hgt = 0;

	i4sw[0].i4s = uni5[numrc].i5;

	hgt = (i4sw[0].chi6[0] << 8) | i4sw[0].chi6[1];

	if (abs(hgt) == 9999 || abs(hgt) == 32768 || abs(hgt) > 8848)
	{
		hgt = 0; //ocean and radar voids or glitches (since max height on Earth = 8848 meters)
	}

/*
	union i6switch //used to rotating little endian to big endian
	{
		short int i6s;
		unsigned char chi7[2]; //two bytes for a short integer
	} i6sw[1];
*/

/*
	if (!Linux) //need to rotate bits
	{
		i4sw[0].i4s = uni5[numrc].i5;

		hgt = (i4sw[0].chi6[0] << 8) | i4sw[0].chi6[1];
	}
	else
	{
		hgt = uni5[numrc].i5;
	}

	//if ((abs(hgt) == 9999 || abs(hgt) == 32768 || abs(hgt) > 8848)  && !Linux)
	if (abs(hgt) == 9999 || abs(hgt) == 32768 || abs(hgt) > 8848)
	{
		hgt = 0; //ocean and radar voids or glitches (since max height on Earth = 8848 meters)
	} 
	else if (abs(hgt) > 8848  && Linux) //try rotating
	{
		i4sw[0].i4s = uni5[numrc].i5;

    	//i6sw[0].chi7[0] = i4sw[0].chi6[1];
    	//i6sw[0].chi7[1] = i4sw[0].chi6[0];
    	//hgt = i6sw[0].i6s;

		hgt = (i4sw[0].chi6[0] << 8) | i4sw[0].chi6[1];

		if (abs(hgt) > 8848) hgt = 0; //now set to zero
	} 

*/

	return hgt;

}

/////////////////////functions for epehm calc valid for any year/////
//source: homeplanet code and Ofek Mathlab routines
///////////////////////////////////////////////////////////////////////

/*  OBLIQEQ  --  Calculate the obliquity of the ecliptic for a given Julian
				 date.  This uses Laskar's tenth-degree polynomial fit
				 (J. Laskar, Astronomy and Astrophysics, Vol. 157, page 68 [1986])
				 which is accurate to within 0.01 arc second between AD 1000
				 and AD 3000, and within a few seconds of arc for +/-10000
				 years around AD 2000.  If we're outside the range in which
				 this fit is valid (deep time) we simply return the J2000 value
				 of the obliquity, which happens to be almost precisely the mean.  */
//source Astro.c in homeplanet code.
//////////////////////obliqueq////////////////////////////////////
double obliqeq(double *jd)
////////////////////////////////////////////////////////////////////
{
#define Asec(x)	((x) / 3600.0)

	//built in functions
    //double abs(double);

	static double oterms[10] = {
		Asec(-4680.93),
		Asec(   -1.55),
		Asec( 1999.25),
		Asec(  -51.38),
		Asec( -249.67),
		Asec(  -39.05),
		Asec(    7.12),
		Asec(   27.87),
		Asec(    5.79),
		Asec(    2.45)
	};

	double eps = 23 + (26 / 60.0) + (21.448 / 3600.0), u, v;
	int i;

	v = u = (*jd - J2000) / (JulianCentury * 100);

    if (abs(u) < 1.0) {
		for (i = 0; i < 10; i++) {
			eps += oterms[i] * v;
			v *= u;
		}
	}
	return eps;
}

//  KEPLER  --	Solve the equation of Kepler.  //
//////////////////////kepler/////////////////////////////
double kepler(double *m, double *ecc)
/////////////////////////////////////////////////////////
{
	#define EPSILON 1E-6

    double e, delta;
	double sin(double), cos(double);

    e = *m = dtr(*m);
    do {
		delta = e - *ecc * sin(e) - *m;
		e -= delta / (1 - *ecc * cos(e));
    } while (abs(delta) > EPSILON);
    return e;
}

///////////////////////////Decl3/////////////////////////////////////////////////////////////
double Decl3(double *jdn, double *hrs, double *dy, double *td, short *mday, short *mon,
				 short *year, int *yl, double *ms, double *aas, double *ob, double *ra,
				 double *rv, int mode)
//////////////////////////////////////////////////////////////////////////////////////////////////
// calculate the solar declination to low accuracy valid for any year after the reformation of the Gregorian calendar.
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// input variables: hrs = hours from midnight
//					td = time zone (negative for west of greenwhich)
//					year = the Gregorian year
//	output variables: ms  =	Geometric mean longitude of the Sun, referred to the
//						    mean equinox of the date. 
//					  aas = Sun's mean anomaly.
//					  ob  = Obliquity of the ecliptic.	
//					  ra  = Sun's right accension in radians
//					  ra  = Sun's radius in astronomical units
//					  mday = day (1-366)
//					  mon = month (1-12)
//					  jdn = portion of Julain day number at 0:00:00 not referred to any equinox
//  mode			  = 0 to calculate the month and day and julian day and other astron. cosntanst
//					  = 1 to skip everything that was already calculated
//	parameters		  apparent = should be TRUE_ if  the  apparent  position
//	 				  (corrected  for  nutation  and aberration) is desired (defuult).
//  dstflag			  daylight saving time flag = 1 for DST, = 0 for Standard time
//  routine returns the Sun's declination in radians
////////////////////////////////////////////////////////////////////////////////////////////////		
{

	//first calculate the julain day number
	int y, m, a, b;
	double UT, jd, Decl;
    double t, t2, t3, l, ma, e, ea, v, theta, omega, eps;
	bool apparent = true;
	int dstflag = 0;
	int month = 0;
	int day;
	//month (beginning and ending days for normal and leap years
	int monthsN[13] = {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365};
	int monthsL[13] = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366};

	/////////////////month and day///////////////////////

	if (mode == 0) //calculate the month and day
	{
		//deter__min if leap year
		*yl = 365;
		if (*year % 4 == 0) {
		*yl = 366;
		}
		if (*year % 100 == 0 && *year % 400 != 0) {
		*yl = 365;
		}

		if (*yl == 365)
		{
			while (monthsN[month] < *dy) {
				// Figure out the correct month.
				month++;
			}

			// Get the day thanks to the months array
			day = *dy - monthsN[month - 1];
		}
		else
		{
			while (monthsL[month] < *dy) {
				// Figure out the correct month.
				month++;
			}

			// Get the day thanks to the months array
			day = *dy - monthsL[month - 1];
		}

		*mon = month;
		*mday = day;

	//////////////////////////////////////////////////////////////////////////

		//calculate the julain date
		//use Meeus formula 7.1 //modified from home planet code function: ucttoj::Astro.c
		

		m = *mon;          //bookmark
		y = *year;

		if (m <= 2) {
			y--;
			m += 12;
		}

		/* Deter__min whether date is in Julian or Gregorian calendar based on
		   canonical date of calendar reform. */

		if ((*year < 1582) || ((*year == 1582) && ((*mon < 9) || (*mon == 9 && *mday < 5)))) {
			b = 0;
		} else {
			a = ((int) (y / 100));
			b = 2 - a + (a / 4);
		}

		*jdn = (((long) (365.25 * (y + 4716))) + ((int) (30.6001 * (m + 1))) +
					*mday + b - 1524.5);
	}

	//Julian Date
	UT = *hrs - *td - dstflag; //universal time
	jd = *jdn + UT / 24.0;


	///////////////////////SunPos (modified from home planet source code: SunPos: Astro.c)/////////////////////////

    //double t, t2, t3, l, m, e, ea, v, theta, omega, eps;

    //double cos(double), tan(double), atan(double), 
	//	sqrt(double), sin(double), asin(double), atan2(double, double);


    /* Time, in Julian centuries of 36525 ephemeris days,
       measured from the epoch 1900 January 0.5 ET. */
	//jd = 2456540.91666666667; //for diagnostics

    t = (jd - J2000) / JulianCentury;           //bookmark
    t2 = t * t;
    t3 = t2 * t;

    /* Geometric mean longitude of the Sun, referred to the
       mean equinox of the date. */

    l = fixangle(280.46645 + 36000.76983 * t + 0.0003032 * t2);

	*ms = dtr(l);

    /* Sun's mean anomaly. */

    ma = fixangle(357.5291 + 35999.0503 * t - 0.0001559 * t2 - 0.00000048 * t3);

	*aas = dtr(ma);

    /* Eccentricity of the Earth's orbit. */

    e = 0.016708617 - 0.000042037 * t - 0.0000001236 * t2;

    /* Eccentric anomaly. */

    ea = kepler(&ma, &e);

	ma = rtd(ma); //convert back to degrees

    /* True anomaly */

    v = fixangle(2 * rtd(atan(sqrt((1 + e) / (1 - e))  * tan(ea / 2))));

    /* Sun's true longitude. */

    theta = fixangle(l + v - ma);

    /* Obliquity of the ecliptic. */

    eps = obliqeq(&jd);

	*ob = dtr(eps);

    /* Corrections for Sun's apparent longitude, if desired. */

    if (apparent) {
       omega = fixangle(259.18 - 1934.142 * t);
       theta = theta - 0.00569 - 0.00479 * sin(dtr(omega));
       eps += 0.00256 * cos(dtr(omega));
    }

    /* Return Sun's longitude and radius vector */

    //*slong = theta; //this routine doesn't need to return the solar longitude
    *rv = (1.0000002 * (1 - e * e)) / (1 + e * cos(dtr(v)));

    /* Deter__min solar co-ordinates in radians. */

    *ra =
	fixangr(atan2(cos(dtr(eps)) * sin(dtr(theta)), cos(dtr(theta)))); //in radians

    Decl = fixangr(asin(sin(dtr(eps)) * sin(dtr(theta)))); //return declination in radians

	if (Decl > pi && pi2 - Decl < pi * 0.5) Decl = Decl - pi2;  //declination's range is 0 - pi/2

	return Decl;
}

///////////////////////Temperatures//////////////////////////////////////
short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] )
///////////////////////////////////////////////////////////////////////////
//reads temperature at latitude, longitude lt,lg, exports Minimum and Average Temperatures to 
//arrays MinTemp, and AvgTemp
//files reside on server at drive drivTemp
/////////////////////////////////////////////////////////////////////////////////////////////
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

	FILE *stream;

	strcpy(FilePathBil,drivTemp );

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

	if (Tempmode == 0) { // minimum temperatures
		strcat(FilePathBil,"min_");
	}
	else if ( Tempmode == 1) { // average temperatures
		strcpy(FilePathBil,drivTemp );
		strcat(FilePathBil,"avg_");
	}
	else if ( Tempmode == 2) { // maximum temperatures
		strcpy(FilePathBil,drivTemp );
		strcat(FilePathBil,"max_");
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

		//fprintf(stdout, "%s\n", FileNameBil );

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
			if ( Tempmode == 0 ) {
				MinTemp[i - 1] = IO;
			} else if( Tempmode == 1 ) {
				AvgTemp[i - 1] = IO;
			} else if ( Tempmode == 2 ) {
				MaxTemp[i - 1] = IO;
			}
        
			fclose(stream);

		}else{
			return Tempmode * 12 + i; //return index of missing bill file for error report
		}
     
	}

  //now go back and calculate the AvgTemps
	if ( Tempmode == 0) {
        Tempmode = 1;
        goto T50;
	} else if (Tempmode == 1) {
		Tempmode = 2;
		goto T50;
	}

	return 0;
}

//////////////lenMonth(x) deter__mins number of days in month x, 1<=x<=12////////////////////
//source: https://cmcenroe.me/2014/12/05/days-in-month-formula.html
short lenMonth( short x) {
	return 28 + (short)(x + floor(x/8)) % 2 + 2 % x + 2 * floor(1/x);
}

///////////////DaysInYear deter__mins the number of days in the year///////////////
short DaysInYear(short *nyr)
//////////////////////////////////////////////////////////////////////////////////
{
	short nyl = 365;
	if (*nyr % 4 == 0) {
		nyl = 366;
	}
	if (*nyr % 100 == 0 && *nyr % 400 != 0) {
		nyl = 365;
	}
	return nyl;
}

////////////////MonthFromDay////////////////////////////////
//find the month of the year given the day and the length of the year
double MonthFromDay(short *nyl, double *dy)
{
	short k = 2;

	if (*nyl == 366) {
		    k = 1.;
	}

	return (double) ((int) ((k + *dy) * 9 / 275 + .98f));

}


////////////////////////ARCHIVES////////////////////////////
/////////////////////////////////////////////////////////////

////////////////////////MENAT.FOR/////////////////////
/*
C       PROGRAM MENAT--CALCULATES ASTRONOMICAL ATMOSPHERIC REFRACTION
c       FOR ANY OBSERVER HEIGHT by using ray tracing through a
c       simplified layered atmosphered
c       ITS OUTPUT IS REPRESENTED in the files menatsum.ren, menatwin.ren
        PROGRAM MENAT
        IMPLICIT REAL*8(A-H,O-Z)
        COMMON/ZZ/HJ(7),TJ(7),PJ(7),AT(7),CT(7)
        DIMENSION TECE(4),PEMB(4),HS(7),TS(7),PS(7),HW(7),TW(7),PW(7)
        DOUBLE PRECISION EPG
        REAL ENT(4,2000)
        bool NEG,BEG,ANGL
        CHARACTER FINAM*10,CH1*1,CH2*2,CH3*3,CH7*7,CH*1
        ANGL=.TRUE.

1       PRINT *,' HELLO'
        DATA HS/0.D0,13.D0,18.D0,25.D0,47.D0,50.D0,70.D0/
        DATA HW/0.D0,10.D0,19.D0,25.D0,30.D0,50.D0,70.D0/
        DATA TS/299.D0,215.5D0,216.5D0,224.D0,273.D0,276.D0,218.D0/
        DATA TW/284.D0,220.0D0,215.0D0,216.D0,217.D0,266.D0,231.D0/
        DATA PS/1013.D0,179.D0,81.2D0,27.7D0,1.29D0,.951D0,0.067D0/
        DATA PW/1018.D0,256.8D0,62.8D0,24.3D0,11.1D0,.682D0,0.0467D0/
        DATA RA,PI2,CO,FR/6.371D3,1.570796D0,1.7453293D-2,1.0010D0/
        DATA HZ,DH,HSOF,EPG/801.D-3,2.0D-3,65.D0,0.0D0/
        DATA DIS1,DIS2,DIS3,DIS4/3.48D1,3.52D1,4.280D1,4.320D1/
        DATA HE1,HE2,HE3,HE4/8.47D-1,8.53D-1,8.970D-1,9.03D-1/
        DATA HE5,HE6,HE7,HE8/9.47D-1,9.53D-1,9.97D-1,1.003D0/
        DATA HE9,HE10,HE11,HE12/1.047D0,1.053D0,7.970D-1,8.03D-1/
        DATA PIE,RAMG,Q3,Q6/3.1415926D0,5.729578D-2,1.D3,1.D6/
        DATA TC,PMB/11.D0,1.018D3/
        DATA TECE/11.5D0,26.D0,24.D0,16.D0/
        DATA PEMB/1.0155D3,1.006D3,1.010D3,1.012D3/
        AT(7)=0.D0
        CT(7)=0.D0
        DH=2.0D-3
        WRITE(*,'(A,F6.4)')' DH=',DH
2       WRITE(*,'(A)')' INPUT BEGINNING HEIGHT FOR CALCULATION (M)-->'
        READ(*,*,ERR = 2)HZ1
        HZ1 = HZ1/1.0D3
3       WRITE(*,'(A)')' INPUT END HEIGHT FOR CALCULATION (M)-->'
        READ(*,*,ERR = 3)HZ2
        HZ2 = HZ2/1.0D3
4       WRITE (*,'(A)')' INPUT STEP HEIGHT FOR CALCULATION (M)-->'
        READ(*,*,ERR = 4)DHZ
        DHZ = DHZ/1.0D3
        NHGT = NINT((HZ2 - HZ1)/DHZ) + 1

C        WRITE (*,'(A,F6.1)')' PRESENT HEIGHT FOR CALCULATION =',HZ*1000
C       WRITE (*,'(A\)')' WANT NEW HEIGHT ? (Y/N)'
C        READ(*,'(A)')CH
C        IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) THEN
C2               WRITE(*,'(A\)')' INPUT NEW HEIGHT (M)-->'
C                READ(*,*,ERR=2)HZ
C                HZ=HZ/1.0D3
C                END IF
C       DO 550 ISN=1,1
C       WRITE(*,5)ISN
C5       FORMAT('2',5X,'SUMMER=1 @ WINTER=2','   ISN=',I1//)
6       WRITE(*,'(A)')' SUMMER=1 (DEF), WINTER=2 -->'
        READ(*,'(I1)',ERR=6)ISN
        IF (ISN .NE. 2) THEN
           ISN=1
        ELSE IF (CHISN .EQ. 2) THEN
           ISN=2
           END IF

        DO 7 K=1,7
           HJ(K)=HS(K)
           TJ(K)=TS(K)
           PJ(K)=PS(K)
           IF (ISN.EQ.1) GOTO 7
           HJ(K)=HW(K)
           TJ(K)=TW(K)
           PJ(K)=PW(K)
7       CONTINUE
        DO 10 K=1,7
           L=K+1
           IF (K.LE.6)AT(K)=(TJ(L)-TJ(K))/(HJ(L)-HJ(K))
           IF (K.LE.6)CT(K)=DLOG(PJ(L)/PJ(K))/DLOG(TJ(L)/TJ(K))
           WRITE(*,8)K,HJ(K),PJ(K),TJ(K),AT(K),CT(K)
8       FORMAT(1X,'---8',I3,F6.1,F9.3,F7.1,' AT=',F7.4,' CT=',F9.4)
10      CONTINUE
12      WRITE(*,'(A)')' INPUT MINIMUM ANG ALT (DEG)-->'
        READ(*,*,ERR=12)EPG1
13      WRITE(*,'(A)')' INPUT MAXIMUM ANG ALT (DEG)-->'
        READ(*,*,ERR=13)EPG2
14      WRITE(*,'(A)')' INPUT STEP ANG ALT (DEG)-->'
        READ(*,*,ERR=14)ESTEP
15      FINAM=' '
        DO 900 NKHGT = 1, NHGT
           IF (ISN .EQ. 1) THEN
              FINAM = 'SUMREF.'
              print *,'here'
              print *,FINAM
              pause 'continue'
           ELSE IF (ISN .EQ. 2) THEN
              FINAM ='WINREF.'
              ENDIF
           HZ = HZ1 + (NKHGT - 1)*DHZ
           MHGT = NINT(HZ*1000)
           IF (MHGT .LT. 10) THEN
              WRITE(CH1,'(I1)')MHGT
              FINAM(8:8)='0'
              FINAM(9:9)='0'
              FINAM(10:10)=CH1(1:1)
           ELSE IF ((MHGT.GE.10).AND.(MHGT.LT.100)) THEN
              WRITE(CH2,'(I2)')MHGT
              FINAM(8:8) = '0'
              FINAM(9:10)=CH2(1:2)
           ELSE
              WRITE(CH3,'(I3)')MHGT
              FINAM(8:10)=CH3(1:3)
              print *,FINAM(8:10)
              pause 'continue'
              END IF
         print *,FINAM
         pause 'continue'
         PRINT *,FINAM
C        WRITE(*,'(A\)')' INPUT FILE NAME (DEF=REFRAC.DAT)-->'
C        READ(*,'(A)',ERR=15)FINAM
C        IF (FINAM.EQ.' ') FINAM='REFRAC.DAT'
        OPEN(1,FILE=FINAM,STATUS='NEW')
        NANG=NINT((EPG2-EPG1)/ESTEP)+1
18      WRITE(*,'(A)')' WANT DIST/HEIGHT OUTPUT? (Y/N)-->'
        READ(*,'(A)',ERR=18)CH
        IF ((CH.EQ.'N').OR.(CH.EQ.'n')) ANGL=.FALSE.

        DO 500 KGR=1,NANG
           DTG2 = 0.0D0
           DH=2.0D-3
           EPG=EPG1+(KGR-1)*ESTEP
           IF ((EPG.GT.0).AND.(EPG.LT.ESTEP)) THEN
              EPG=0.0D0
              END IF
           IF (ANGL) THEN
              MANG = NINT(DABS(EPG*1000))
              FINAM=' '
              IF (MANG .LT. 10) THEN
                 WRITE(CH1,'(I1)')MANG
                 FINAM(8:8) = '0'
                 FINAM(9:9) = '0'
                 FINAM(10:10)=CH1(1:1)
              ELSE IF ((MANG.GE.10).AND.(MANG.LT.100)) THEN
                 WRITE(CH2,'(I2)')MANG
                 FINAM(8:8) = '0'
                 FINAM(9:10)=CH2(1:2)
              ELSE
                 WRITE(CH3,'(I3)')MANG
                 FINAM(8:10)=CH3(1:3)
                 END IF
              IF (ISN .EQ. 1) THEN
                 FINAM(1:7) = 'MENSUM.'
              ELSE IF (ISN .EQ. 2) THEN
                 FINAM(1:7) = 'MENWIN.'
                 ENDIF
              PRINT *,FINAM
              OPEN(2,FILE=FINAM,STATUS='NEW')
              END IF
           EPZ=EPG*CO
c           write(*,*)' epz=',epz
           EPZM=EPZ*Q3
           RSOF=RA+HSOF
           IF (EPZ.LT.0) THEN
              IF ((HZ-DH/2.0D0).GT.0) THEN
                 HEN1=HZ-DH/2.0D0
                 DH=-2.0D-3
              ELSE IF ((HZ-DH/2.0D0).LE.0) THEN
                 HEN1=HZ
                 END IF
           ELSE IF (EPZ .GE. 0) THEN
                HEN1=HZ+DH/2.0D0
                END IF
           CZ=DCOS(EPZ)*(RA+HZ)*(1.0D0+FUN(HEN1))
           BZ=PI2-EPZ
           BN=BZ
c           write(*,*)' bn=',bn
           EP1=Q3*EPZ
c           WRITE(*,*)HZ,DH,EP1,EPG,FR,TJ(1),PJ(1)
C13         FORMAT(X13,'HZ=',F6.3,'  DH=',F6.3,' EPZMIL=',F8.3,' EPZG=',
C       *   F4.2,'  FR=',F6.4,'   TK=',F5.1,'   PMB=',F6.1)
C           WRITE(*,15)
C15         FORMAT(4X,'N',4X,'DHM',3X,'EPS13',5X,'H',5X,'DEL',4X,'ELTOT',
C       *   4X,'ALF6',3X,'EPS23',3X,'EN17',5X,'REF6',3X,'REFT3   REG3')
           T=0.0D0
           S1=0.0D0
           AT2=0.0D0
           NZ=(HSOF+DABS(DH)-HZ)/DABS(DH)+1.1D0
C           WRITE(*,*)'n=',NZ
C           READ(*,'(A)')CH
           H=HZ-DH
           BEG=.TRUE.
           NEG=.FALSE.
           NENT = 0
20         DO200 N=1,NZ
             B=BN
             E=PI2-B
22           E1=Q3*E
c             write(*,*)' E(DEG)=',E/CO
c>>>          read(*,'(a\)')ch
             IF((EPG .NE. 0.0D0) .AND. (DABS(E/CO).LT.(5.0D-5))) THEN
                E=0
                DH=2.0D-3
                DTG2 = DTG
                END IF
C             read(*,'(a)')ch

C             IF ((E.LT.0).AND.(BEG)) THEN
C                NEG=.TRUE.
C                BEG=.FALSE.
C                DN=1.0D0
C23              DH=-DH/DN
C                IF (H+DH/DN .GT. 0) THEN
C                ELSE
C                   DN=DN*2.0D0
C                   GOTO 23
C                   END IF
C                NZ=NZ+1
C             ELSE IF ((E.GE.0).AND.(NEG)) THEN
C                NEG=.FALSE.
C                DH=-DH*DN
C                END IF
             IF (DH.LT.0) NZ=NZ+1
             H=H+DH
             IF (H.LT.0)H=0
             RT=RA+H
25           RU=RT+DH
             HEN=H+DH/2.0D0
C>>          write(*,*)' hen=',hen
             IF (N.EQ.1) THEN
                HMIN=HEN
                HMAX=HEN
             ELSE
                IF (HEN.LT.HMIN) HMIN=HEN
                IF (HEN.GT.HMAX) HMAX=HEN
                ENDIF
             HEV=HEN+DH
c             WRITE(*,'(A)')'H,RT,RU,HEN,HEV,B,E,E1'
c             WRITE(*,*)H,RT,RU,HEN,HEV,B,E,E1
             EL=SQT(E,DH,RT)
C>>          WRITE(*,*)' N,EL(KM)=',N,EL
C             READ(*,'(A)')CH
             S=EL*DCOS(E)
             if (s/ru .gt. 1) then
                PRINT *,' s/ru=',s/ru
                read(*,'(a)')ch
                s=ru
                end if
             A2=DASIN(S/RU)
30           DEN=A2*RA
             AT2=AT2+A2
             S2=AT2*RA
             A3=Q6*A2
             EN=FUN(HEN)
             EN1=1.0D7*EN
35           G=B-A2
C>>          WRITE(*,*)' G(DEG)=',G/CO
             G1=Q3*G
             E2=Q3*(PI2-G)
             EM=1.0D0+FUN(HEV)
c             WRITE(*,'(A)')'e1,EN,EN1,G,G1,E2,EM='
c             WRITE(*,*)E1,EN,EN1,G,G1,E2,EM
C             WRITE(*,'(A\)')' CONTINUE?'
C             READ(*,'(A)')ANS
C             IF ANS.EQ.'N' GOTO 999
             SBN=(1.0D0+EN)*DSIN(G)/EM
C>>          WRITE(*,*)' SBN=',SBN
             if (sbn .gt. 1) then
                PRINT *,' sbn=',sbn
                sbn=1.0
c                read(*,'(a)')ch
                end if
             BN=DASIN(SBN)
             IF (G.GT.PI2) THEN
                BN=PIE-BN
                DH=-EL*DABS(G-PI2)
C>>             WRITE(*,*)' EL,G,-EL*G=',EL,G,-EL*DABS(G-PI2)
                END IF
C>>          WRITE(*,*)' BN(DEG)=',BN/CO
40           D=BN-G
             D6=Q6*D
             T=T+D
             DT1=Q3*T
             DTG=DT1*RAMG
             FR=1.0D0
             IF(H.GE.1.5D0)FR=1.0001D0
             DH=FR*DH
             DG=DH*Q3
c        IF ((HEN.LE.1.0).AND.(ANGL)) THEN
        IF (HEN .LE. 5.0) THEN
c         IF (S2 .LE. 250) THEN
c          HENR = HEN - RA * (1.0 - DCOS(S2/RA))
           NENT = NENT + 1
           ENT(1,NENT) = E1
           ENT(2,NENT) = HEN
           ENT(3,NENT) = S2
           ENT(4,NENT) = DT1
c        WRITE(2,'(I4,1X,F6.2,1X,F5.3,1X,F7.3,1X,F5.2)')N,E1,HEN,S2,DT1
           END IF
c        WRITE(*,'(I3,1X,F6.2,1X,F5.3,1X,F7.3,1X,F5.2)')N,E1,HEN,S2,DT1
c>>        s22 = 101.1
c>>        henn = .055
c>>        IF ((S2.LE.s22 + 1).AND.(S2.GE.s22 - 1)) THEN
c           IF ((HEN.LE.henn + .002).AND.(HEN.GE.henn - .002)) THEN
c>>           IF (HEN .GE. henn) THEN
c>>              WRITE(*,*)' Dist =',S2
c>>              WRITE(*,*)' Height =',HEN
c>>              WRITE(*,*)' View Angle =',EPG
c>>              WRITE(*,*)' Refraction =', DTG2
c>>              WRITE(*,'(A\)')' CONTINUE?'
c>>              READ(*,'(A)')ANS
c>>              END IF
c>>          END IF
C            WRITE(*,*)' D6,DT1=',D6,DT1
C49           IF(N.LE.10)WRITE(*,*)N,DG,E1,HEN,DEN,S2,A3,E2,EN1,D6,DT1,DTG
C      .OR.(N.EQ.20).OR.(N.EQ.40).OR.(N.EQ.60).OR.
C      *       (N.EQ.80).OR.(N.EQ.100).OR.(N.EQ.200).OR.(N.EQ.300))
C      *        WRITE(*,*)N,DG,E1,HEN,DEN,S2,A3,E2,EN1,D6,DT1,GDTG
C56           FORMAT(1X,I5,2F7.2,2F8.3,F9.3,F7.2,F7.2,F9.2,F8.4,2F7.3)
c             WRITE(*,*)N,RT,RSOF
60           IF(RT.LT.RSOF)GOTO200
             FIEM=EPZM-4.665D0-DT1
             FIEG=FIEM*RAMG
77           PRINT *,RT,HEN,S2,N,DT1,DTG,FIEM,FIEG
c                write(*,*)' here'
c             IF (ANGL) write(*,*) ' here2'
             IF (ANGL) THEN
                WRITE(2,'(F7.3)')DT1
                PRINT *,NENT
                DO 80 NN = 1, NENT
                   EN1 = ENT(1,NN)
                   EN2 = ENT(2,NN)
                   EN3 = ENT(3,NN)
                   EN4 = ENT(4,NN)
                   WRITE(2,78)NN,EN1,EN2,EN3,EN4
78                 FORMAT(I4,1X,F6.2,1X,F5.3,1X,F7.3,1X,F5.2)
80              CONTINUE
c                PRINT *,' finished'
                END IF
C80           FORMAT(//1X,' RT=',F9.3,'  HT=',F6.3,'  TOTL=',F8.3,2X,' NTOT=',
C       *     Z'  FIEZG=',F7.4/)
             PRINT *,' epg,dt1,dtg,fieg=',epg,dt1,dtg,fieg
C             write(*,'(a\)')' cr-->'
C             read(*,'(a)')ch
             IF (EPG .NE. 0.0D0) THEN
                WRITE(1,100)EPG,DT1,DTG,FIEG,HMIN,HMAX
c                WRITE(1,100)EPG - DTG2,DT1,DTG,FIEG,HMIN,HMAX
               PRINT *,epg - dtg2, hmin, hmax
c               write(*,'(a\)')' cr-->'
c               read(*,'(a)')ch
                ENDIF
100          FORMAT(F6.2,3X,F8.4,3X,F8.4,3X,F8.4,3X,F8.3,3X,F8.3)
c             WRITE(1,'(F6.2,3X,F8.4,3X,F8.4,3X,F8.4)')EPG,DT1,DTG,FIEG
             GOTO 300
200          CONTINUE
300          CONTINUE
             CLOSE(2)
500          CONTINUE
C550          CONTINUE
             CLOSE(1)
900          CONTINUE
980          WRITE(*,'(A)')' WANT TO EXIT? (Y/N)-->'
             READ(*,'(A)',ERR=980)CH
             IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) GOTO 999
             GOTO 1
999          END
C
        FUNCTION SQT(E,DH,R)
        IMPLICIT REAL*8(A-H,O-Z)
        Q=R*DSIN(E)
        Y=2.0D0*R+DH
C>>     WRITE(*,*)' DH=',DH
        IF (DH.LT.0) SQT=-Q-DSQRT(DABS(Q*Q+DH*Y))
        IF (DH.GE.0) SQT=-Q+DSQRT(DABS(Q*Q+DH*Y))
c        if (q*q+dh*y .lt. 0) then
c           write(*,*)'-q,quo,sqt=',-q,q*q+dh*y,sqt
c           end if
        RETURN
        END
C
        FUNCTION FUN(H)
        IMPLICIT REAL*8(A-H,O-Z)
        COMMON/ZZ/HJ(7),TJ(7),PJ(7),AT(7),CT(7)
        DO10 N=1,7
           M=N-1
           IF(H-HJ(N))15,15,10
10      CONTINUE
15      T=TJ(M)+AT(M)*(H-HJ(M))
        P=PJ(M)*(T/TJ(M))**CT(M)
        FUN=79.D-6*P/T
25      RETURN
        END

////////////////////MENAT4.FOR/////////////////
        PROGRAM MENAT4
c       MENAT4 CALCULATES DIP ANGLES FOR GIVEN DIST-HGT FILE
c       version 2.00
        IMPLICIT REAL*8(A-H,O-Z)
        PARAMETER(NPNT = 2760)
        COMMON/ZZ/HJ(7),TJ(7),PJ(7),AT(7),CT(7)
        DIMENSION TECE(4),PEMB(4),HS(7),TS(7),PS(7),HW(7),TW(7),PW(7)
        DOUBLE PRECISION EPG
        bool NEG,BEG,EXISTS,FOUND,SKIP,STR1,STR2
        REAL HGTS(3,NPNT),HT(NPNT)
        INTEGER*2 NPLAC(50)
        CHARACTER*1 CH
        CHARACTER CHRD1*23, CHRD2*44, FILN*12, FILNN*20, PLACEN*8
        CHARACTER CH8*8,CH5*5,CH58*58

1       WRITE(*,'(A)')' HELLO'
        DATA HS/0.D0,13.D0,18.D0,25.D0,47.D0,50.D0,70.D0/
        DATA HW/0.D0,10.D0,19.D0,25.D0,30.D0,50.D0,70.D0/
        DATA TS/299.D0,215.5D0,216.5D0,224.D0,273.D0,276.D0,218.D0/
        DATA TW/284.D0,220.0D0,215.0D0,216.D0,217.D0,266.D0,231.D0/
        DATA PS/1013.D0,179.D0,81.2D0,27.7D0,1.29D0,.951D0,0.067D0/
        DATA PW/1018.D0,256.8D0,62.8D0,24.3D0,11.1D0,.682D0,0.0467D0/
        DATA RA,PI2,CO,FR/6.371D3,1.570796D0,1.7453293D-2,1.0010D0/
        DATA HZ,DH,HSOF,EPG/801.D-3,2.0D-3,65.D0,0.0D0/
        DATA DIS1,DIS2,DIS3,DIS4/3.48D1,3.52D1,4.280D1,4.320D1/
        DATA HE1,HE2,HE3,HE4/8.47D-1,8.53D-1,8.970D-1,9.03D-1/
        DATA HE5,HE6,HE7,HE8/9.47D-1,9.53D-1,9.97D-1,1.003D0/
        DATA HE9,HE10,HE11,HE12/1.047D0,1.053D0,7.970D-1,8.03D-1/
        DATA PIE,RAMG,Q3,Q6/3.1415926D0,5.729578D-2,1.D3,1.D6/
        DATA TC,PMB/11.D0,1.018D3/
        DATA TECE/11.5D0,26.D0,24.D0,16.D0/
        DATA PEMB/1.0155D3,1.006D3,1.010D3,1.012D3/

        NPLN = 1
        NZ = 1
C       nplacesn is the number of places to loop through
        NPLAC(1) = 1
c       loop between places with NPLACE num NPLAC(1),...,NPLAC(NPLN)

        INQUIRE(FILE='MENAT44.TMP',EXIST=EXISTS)
        IF (EXISTS) THEN
           OPEN(2,FILE='MENAT44.TMP',STATUS='OLD')
           READ(2,*)NPL1,NPLIT,ISN1,NHT,NTOT
           CLOSE(2)
        ELSE
           NPL1=1
           ISN1=1
           NH1=0
           GOTO 5
           ENDIF
        IF ((NHT.EQ.NTOT).AND.(ISN1.EQ.2)) THEN
           IF (NPL1.LT.NPLIT) THEN
              NPL1=NPL1+1
              ISN1=1
              NH1=1
           ELSE IF (NPL1.EQ.NPLIT) THEN
              GOTO 999
              END IF
        ELSE IF ((NHT.EQ.NTOT).AND.(ISN1.EQ.1)) THEN
           ISN1=2
           NH1=1
        ELSE IF (NHT.LT.NTOT) THEN
           NH1=NHT+1
           END IF

5       DO 990 NPLI=NPL1,NPLN
        NPLACE=NPLAC(NPLI)
        IF (NPLACE .EQ. 1) THEN
           PLACEN = 'Hgtbel'
           END IF

        NFILN=0
        DO 33 I=1,8
           IF (PLACEN(I:I).NE.' ') THEN
              NFILN=NFILN+1
              FILN(I:I)=PLACEN(I:I)
              END IF
33      CONTINUE
        FILN(NFILN+1:NFILN+4)='.DAT'
        FILNN(1:20)='                    '
        FILNN(1:6)='C:\JK\'
        FILNN(7:10+NFILN)=FILN(1:NFILN+4)

        NH=1
        OPEN(3,FILE=FILNN,STATUS='OLD')
        READ(3,'(A58)')CH58
50      READ(3,*,END=55)N,hgt,HGTS(1,NH),HGTS(2,NH),HGTS(3,NH)
        HT(NH) = hgt * .001
        HGTS(3,NH)=HGTS(3,NH)*.001
        NH=NH+1
        GOTO 50
55      CLOSE(3)
        NH=NH-1

        AT(7)=0.D0
        CT(7)=0.D0
        DH=2.0D-3
        NLEAP=0
C        WRITE (*,'(A,F6.4)')' DH=',DH


        DO 950 ISN=ISN1,2
c4       WRITE(*,*)'SUMMER=1 (DEF), WINTER=2 -->'
c        READ(*,'(I1)',ERR=4)ISN
c        IF (ISN .NE. 2) THEN
c           ISN=1
c        ELSE IF (CHISN .EQ. 2) THEN
c           ISN=2
c           END IF

        IF (ISN.EQ.1) THEN
           FILNN(7+NFILN:10+NFILN)='.SUM'
        ELSE IF (ISN.EQ.2) THEN
           FILNN(7+NFILN:10+NFILN)='.WIN'
           END IF
        PRINT *,'PLACE FOR CALCULATION: ',PLACEN
        PRINT *,'FILE NAME OF OUTPUT FILE: ',FILNN
        PRINT *,'ISN,NH1=',ISN1,NH1
        INQUIRE(FILE=FILNN,EXIST=EXISTS)
        IF (EXISTS) THEN
           PRINT *, 'FILE ALREADY STARTED--CONTINUE IT'
           OPEN(1,FILE=FILNN,STATUS='OLD')
           READ(1,'(A8,1X,A5,1X,I3)')CH8,CH5,ISNN
           READ(1,'(A44)')CHRD2
           NH1=1
           DO 2 IINQ=1,NH
              READ(1,*,END=3)XENT,X11,E22,E11,D11,D11R
              NH1=NH1+1
2          CONTINUE
3          CLOSE(1)
           X22=-30.0+(NH1-2)*.1
           NLEAP=NINT((XENT-X22)*10)
           PRINT *,'NLEAP=',NLEAP
           NH1=NH1+NLEAP
           END IF

6       DO 7 K=1,7
           HJ(K)=HS(K)
           TJ(K)=TS(K)
           PJ(K)=PS(K)
           IF (ISN.EQ.1) GOTO 7
           HJ(K)=HW(K)
           TJ(K)=TW(K)
           PJ(K)=PW(K)
7       CONTINUE
        DO 10 K=1,7
           L=K+1
           IF (K.LE.6)AT(K)=(TJ(L)-TJ(K))/(HJ(L)-HJ(K))
           IF (K.LE.6)CT(K)=DLOG(PJ(L)/PJ(K))/DLOG(TJ(L)/TJ(K))
c           WRITE(*,8)K,HJ(K),PJ(K),TJ(K),AT(K),CT(K)
c8       FORMAT(1X,'---8',I3,F6.1,F9.3,F7.1,'   AT=',F7.4,'  CT=',F9.4)
10      CONTINUE

        FOUND=.FALSE.
        DSTSRCH0 = 1.1
        DISTSRCH=DSTSRCH0
c       old value 1
        NUMLIM=2

C        INQUIRE(FILE='C:\DTMI\ARMONHGT.OUT',EXIST=EXISTS)
C        IF (EXISTS) THEN
C           OPEN(1,FILE='C:\DTMI\ARMONHGT.OUT',STATUS='OLD')
C        ELSE
C           OPEN(1,FILE='C:\DTMI\ARMONHGT.OUT',STATUS='NEW')
C           END IF
C        WRITE(1,'(A,1X,I1)')'ARMON HANATZIV N., ISN=',ISN
C        WRITE(1,'(A)')'OBS HGT, VIEWANGLE, DIST TO POINT, POINT HGT'

        STR1=.TRUE.
        STR2=.TRUE.
        HGTOO=HGTS(2,NH1)
        DO 900 NKHGT = NH1,NH
           HZ=HT(NKHGT)
           SKIP=.FALSE.
           IF (HGTS(2,NKHGT).LE.15) THEN
              EPG=HGTS(1,NKHGT)
              WRITE(*,*)'***DIST <= 15 KM***'
              WRITE(*,*)'JUST COPY .HGT FILE VALUE'
              WRITE(*,*)'Dist =',HGTS(2,NKHGT)
              WRITE(*,*)'View Angle =',EPG
              WRITE(*,*)'OBSERVER HEIGHT=',HT(NKHGT)
              DTG3=0.0
              DT1=0.0
              SKIP=.TRUE.
              END IF

11         EPG1=HGTS(1,NKHGT)
           IF ((STR1).OR.(STR2).OR.(HGTS(2,NKHGT).NE.HGTOO)) THEN
              IF (HGTS(2,NKHGT).NE.HGTOO) THEN
                 STR1=.TRUE.
                 STR2=.TRUE.
                 HGTOO=HGTS(2,NKHGT)
                 END IF
              EPG2=HGTS(1,NKHGT)+.1
           ELSE
              IF (DIFEPG2.GT.0.0001) THEN
                 EPGT=HGTS(1,NKHGT)+DIFEPG1 -.005
                 IF (EPGT.GT.EPG1) EPG1=EPGT
                 EPG2=HGTS(1,NKHGT)+DIFEPG1+.02
              ELSE IF (DIFEPG2 .LT. 0.0) THEN
                 EPGT=HGTS(1,NKHGT)+DIFEPG1-.005
                 IF (EPGT.GT.EPG1) EPG1=EPGT
                 EPG2=HGTS(1,NKHGT)+DIFEPG1+.005
              ELSE IF (DIFEPG2 .LT. 0.0001) THEN
                 IF (DIFEPG1 .LT. 0.0001) THEN
                    EPG2=HGTS(1,NKHGT)+.005
                 ELSE
                    EPGT=HGTS(1,NKHGT)+DIFEPG1-.005
                    IF (EPGT.GT.EPG1) EPG1=EPGT
                    EPG2=HGTS(1,NKHGT)+DIFEPG1+.005
                    END IF
                 END IF
              END IF

c          old value + .02
           ESTEP=.001
        NANG=NINT((EPG2-EPG1)/ESTEP)+1

        DO 500 KGR=1,NANG
12         IF (SKIP) GOTO 20
           DTG2 = 0.0D0
           DH=2.0D-3
           EPG=EPG1+(KGR-1)*ESTEP
           IF ((EPG.GT.0).AND.(EPG.LT.ESTEP)) THEN
              EPG=0.0D0
              END IF
           EPZ=EPG*CO
c           write(*,*)' epz=',epz
           EPZM=EPZ*Q3
           RSOF=RA+HSOF
           IF (EPZ.LT.0) THEN
c>>>>         IF ((HZ-DH/2.0D0).GT.0) THEN
                 HEN1=HZ-DH/2.0D0
                 DH=-2.0D-3
c>>>>>        ELSE IF ((HZ-DH/2.0D0).LE.0) THEN
c>>>>>           HEN1=HZ
c>>>>>           END IF
           ELSE IF (EPZ .GT. 0) THEN
                HEN1=HZ+DH/2.0D0
                END IF
           CZ=DCOS(EPZ)*(RA+HZ)*(1.0D0+FUN(HEN1))
           BZ=PI2-EPZ
           BN=BZ
c           write(*,*)' bn=',bn
           EP1=Q3*EPZ
c           WRITE(*,*)HZ,DH,EP1,EPG,FR,TJ(1),PJ(1)
C13         FORMAT(X13,'HZ=',F6.3,'  DH=',F6.3,' EPZMIL=',F8.3,' EPZG=',
C       *   F4.2,'  FR=',F6.4,'   TK=',F5.1,'   PMB=',F6.1)
C           WRITE(*,15)
C15         FORMAT(4X,'N',4X,'DHM',3X,'EPS13',5X,'H',5X,'DEL',4X,'ELTOT',
C       *   4X,'ALF6',3X,'EPS23',3X,'EN17',5X,'REF6',3X,'REFT3   REG3')
           T=0.0D0
           S1=0.0D0
           AT2=0.0D0
           NZ=(HSOF+DABS(DH)-HZ)/DABS(DH)+1.1D0
C           WRITE(*,*)'n=',NZ
C           READ(*,'(A)')CH
           H=HZ-DH
           BEG=.TRUE.
           NEG=.FALSE.
20         DO200 N=1,NZ
             IF (SKIP) GOTO 43
             B=BN
             E=PI2-B
22           E1=Q3*E
C>>>>>>>     write(*,*)' E(DEG)=',E/CO
             IF(DABS(E/CO).LT.(5.0D-5)) THEN
                E=0
                DH=2.0D-3
                DTG2 = DTG
                DTG3 = DTG2/RAMG
                END IF
C             read(*,'(a)')ch

C             IF ((E.LT.0).AND.(BEG)) THEN
C                NEG=.TRUE.
C                BEG=.FALSE.
C                DN=1.0D0
C23              DH=-DH/DN
C                IF (H+DH/DN .GT. 0) THEN
C                ELSE
C                   DN=DN*2.0D0
C                   GOTO 23
C                   END IF
C                NZ=NZ+1
C             ELSE IF ((E.GE.0).AND.(NEG)) THEN
C                NEG=.FALSE.
C                DH=-DH*DN
C                END IF
             IF (DH.LT.0) NZ=NZ+1
             H=H+DH
c>>>>>>      IF (H.LT.0)H=0
             RT=RA+H
25           RU=RT+DH
             HEN=H+DH/2.0D0
c            write(*,*)' hen=',hen
             IF (N.EQ.1) THEN
                HMIN=HEN
                HMAX=HEN
             ELSE
                IF (HEN.LT.HMIN) HMIN=HEN
                IF (HEN.GT.HMAX) HMAX=HEN
                ENDIF
             HEV=HEN+DH
c             WRITE(*,'(A)')'H,RT,RU,HEN,HEV,B,E,E1'
c             WRITE(*,*)H,RT,RU,HEN,HEV,B,E,E1
             EL=SQT(E,DH,RT)
C>>          WRITE(*,*)' N,EL(KM)=',N,EL
C             READ(*,'(A)')CH
             S=EL*DCOS(E)
             if (s/ru .gt. 1) then
                write(*,*)' s/ru=',s/ru
                read(*,'(a)')ch
                s=ru
                end if
             A2=DASIN(S/RU)
30           DEN=A2*RA
             AT2=AT2+A2
             S2=AT2*RA
             A3=Q6*A2
             EN=FUN(HEN)
             EN1=1.0D7*EN
35           G=B-A2
C>>          WRITE(*,*)' G(DEG)=',G/CO
             G1=Q3*G
             E2=Q3*(PI2-G)
             EM=1.0D0+FUN(HEV)
c             WRITE(*,'(A)')'e1,EN,EN1,G,G1,E2,EM='
c             WRITE(*,*)E1,EN,EN1,G,G1,E2,EM
C             WRITE(*,'(A\)')' CONTINUE?'
C             READ(*,'(A)')ANS
C             IF ANS.EQ.'N' GOTO 999
             SBN=(1.0D0+EN)*DSIN(G)/EM
C>>          WRITE(*,*)' SBN=',SBN
             if (sbn .gt. 1) then
                write(*,*)' sbn=',sbn
                sbn=1.0
                print *,'**skip to next trial angle**'
                GOTO 500
c                read(*,'(a)')ch
                end if
             BN=DASIN(SBN)
             IF (G.GT.PI2) THEN
                BN=PIE-BN
                DH=-EL*DABS(G-PI2)
C>>             WRITE(*,*)' EL,G,-EL*G=',EL,G,-EL*DABS(G-PI2)
                END IF
C>>          WRITE(*,*)' BN(DEG)=',BN/CO
40           D=BN-G
             D6=Q6*D
             T=T+D
             DT1=Q3*T
             DTG=DT1*RAMG
             FR=1.0D0
             IF(H.GE.1.5D0)FR=1.0001D0
             DH=FR*DH
             DG=DH*Q3
c        IF (HEN.LE.1.0) THEN
c        WRITE(2,'(I3,1X,F5.2,1X,F5.3,1X,F7.3,1X,F5.2)')N,E1,HEN,S2,DT1
c        END IF
c        WRITE(*,'(I3,1X,F6.2,1X,F5.3,1X,F7.3,1X,F5.2)')N,E1,HEN,S2,DT1
c         print *,'s2,dst,hen,hgt=',s2,hgts(2,nkhgt),hen,hgts(3,nkhgt)
c         pause 'continue'
           DIST1=HGTS(2,NKHGT)-DISTSRCH
           DIST2=HGTS(2,NKHGT)+DISTSRCH
           IF ((S2.GE.DIST1).AND.(S2.LE.DIST2)) THEN
c           PRINT *,'S2,DST,HEN,HGT=',S2,HGTS(2,NKHGT),HEN,HGTS(3,NKHGT)
c              PAUSE 'CONTINUE'
           H11=HGTS(3,NKHGT)-.0015
           H22=HGTS(3,NKHGT)+.0015
              IF ((HEN.GE.H11).AND.(HEN.LE.H22)) THEN
C              IF ((HGTS(3,NKHGT).GE.HEN) THEN
                 WRITE(*,*)'Dist =',S2,HGTS(2,NKHGT)
                 WRITE(*,*)'Height =',HEN,HGTS(3,NKHGT)
                 WRITE(*,*)'View Angle =',EPG, HGTS(1,NKHGT)
                 WRITE(*,*)'Refraction =', DTG2
                 WRITE(*,*)'OBSERVER HEIGHT=',HT(NKHGT)
43      INQUIRE(FILE='MENAT44.TMP',EXIST=EXISTS)
        IF (EXISTS) THEN
           OPEN(2,FILE='MENAT44.TMP',STATUS='OLD')
        ELSE
           OPEN(2,FILE='MENAT44.TMP',STATUS='NEW')
           ENDIF
        WRITE(2,'(4(I4,1X))')NPLI,NPLN,ISN,NKHGT,NH
        CLOSE(2)
        IF (ISN.EQ.1) THEN
           FILNN(7+NFILN:10+NFILN)='.SUM'
        ELSE IF (ISN.EQ.2) THEN
           FILNN(7+NFILN:10+NFILN)='.WIN'
           ENDIF
        INQUIRE(FILE=FILNN,EXIST=EXISTS)
        IF (EXISTS) THEN
           OPEN(1,FILE=FILNN,STATUS='OLD')
           READ(1,'(A8,1X,A5,1X,I3)')CH8,CH5,ISNN
           READ(1,'(A44)')CHRD2
           DO 47 IINQ=1,NKHGT-1-NLEAP
              READ(1,*,END=48)XEN,X11,E22,E11,D11,D11R
47         CONTINUE
        ELSE
           OPEN(1,FILE=FILNN,STATUS='NEW')
           WRITE(1,'(A8,1X,A5,1X,I3)')PLACEN,' ISN=',ISN
           WRITE(1,'(A)')'ENT,HGT, VIEWANGLE, DIST TO POINT, POINT HGT'
           END IF
           XENTRY=-30.0+(NKHGT-1)*.1
48      hgto=HT(NKHGT)*1000
        M=NKHGT
        WRITE(1,49)XENTRY,hgto,HGTS(1,M),EPG,HGTS(2,M),HGTS(3,M)
49      FORMAT(F5.1,3X,F5.0,2(3X,F7.4),2(3X,F8.3))
        CLOSE(1)
        IF ((STR1).AND.(STR2)) THEN
           STR1=.FALSE.
           DIFEPG0=EPG-HGTS(1,M)
        ELSE IF ((.NOT.STR1).AND.(STR2)) THEN
           STR2=.FALSE.
           DIFEPG1=EPG-HGTS(1,M)
           DIFEPG2=DIFEPG1-DIFEPG0
           DIFEPG0=DIFEPG1
        ELSE IF ((.NOT.STR1).AND.(.NOT.STR2)) THEN
           DIFEPG1=EPG-HGTS(1,M)
           DIFEPG2=DIFEPG1-DIFEPG0
           DIFEPG0=DIFEPG1
           END IF
C                WRITE(*,*)' STOP NOW (Y/N)?-->'
C                READ(*,'(A)',ERR=980)CH
C                IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) GOTO 999
                 FOUND=.TRUE.
                 DISTSRCH=DSTSRCH0
                 GOTO 890
                 END IF
              END IF
c56      CONTINUE
c        CLOSE(3)
c>>        s22 = 101.1
c>>        henn = .055
c>>        IF ((S2.LE.s22 + 1).AND.(S2.GE.s22 - 1)) THEN
c           IF ((HEN.LE.henn + .002).AND.(HEN.GE.henn - .002)) THEN
c>>           IF (HEN .GE. henn) THEN
c>>              WRITE(*,*)' Dist =',S2
c>>              WRITE(*,*)' Height =',HEN
c>>              WRITE(*,*)' View Angle =',EPG
c>>              WRITE(*,*)' Refraction =', DTG2
c>>              WRITE(*,'(A\)')' CONTINUE?'
c>>              READ(*,'(A)')ANS
c>>              END IF
c>>          END IF
C            WRITE(*,*)' D6,DT1=',D6,DT1
C49           IF(N.LE.10)WRITE(*,*)N,DG,E1,HEN,DEN,S2,A3,E2,EN1,D6,DT1,DTG
C      .OR.(N.EQ.20).OR.(N.EQ.40).OR.(N.EQ.60).OR.
C      *       (N.EQ.80).OR.(N.EQ.100).OR.(N.EQ.200).OR.(N.EQ.300))
C      *        WRITE(*,*)N,DG,E1,HEN,DEN,S2,A3,E2,EN1,D6,DT1,GDTG
C56           FORMAT(1X,I5,2F7.2,2F8.3,F9.3,F7.2,F7.2,F9.2,F8.4,2F7.3)
c             WRITE(*,*)N,RT,RSOF
60           IF(RT.LT.RSOF)GOTO200
             FIEM=EPZM-4.665D0-DT1
             FIEG=FIEM*RAMG
C77>>           WRITE(*,*)RT,HEN,S2,N,DT1,DTG,FIEM,FIEG
C             WRITE(2,'(F5.3)')DT1
C80           FORMAT(//1X,' RT=',F9.3,'  HT=',F6.3,'  TOTL=',F8.3,2X,' NTOT=',
C       *     Z'  FIEZG=',F7.4/)
C>>             write(*,*)' epg,dt1,dtg,fieg=',epg,dt1,dtg,fieg
C             write(*,'(a\)')' cr-->'
C             read(*,'(a)')ch
             IF (EPG .NE. 0.0D0) THEN
c                WRITE(1,100)EPG - DTG2,DT1,DTG,FIEG,HMIN,HMAX
C>>               write(*,*)epg - dtg2, hmin, hmax
c               write(*,'(a\)')' cr-->'
c               read(*,'(a)')ch
                ENDIF
100          FORMAT(F6.2,3X,F8.4,3X,F8.4,3X,F8.4,3X,F8.3,3X,F8.3)
c             WRITE(1,'(F6.2,3X,F8.4,3X,F8.4,3X,F8.4)')EPG,DT1,DTG,FIEG
             GOTO 300
200          CONTINUE
300          CONTINUE
500          CONTINUE
C550          CONTINUE
890          XENTRY=-30.0+(NKHGT-1)*.1
             IF (FOUND) THEN
                FOUND=.FALSE.
                GOTO 900
             ELSE
                WRITE(*,*)'SEARCH WAS NOT SUCCESSFUL AT AZI',XENTRY
                DISTSRCH=DISTSRCH+DSTSRCH0
                DISTLIMS=NUMLIM*DSTSRCH0
                IF (DISTSRCH.GT.DISTLIMS) THEN
                PRINT *,'**REACHED LIMIT OF SEARCH, SKIP THIS ANGLE***'
                   NLEAP=NLEAP+1
                   DISTSRCH=DSTSRCH0
                   GOTO 900
                   END IF
C               IN ORDER TO AVOID INFINITE LOOP DON'T TRY SEARCHING FURTHER THAN NUMLIM*DSTSRCH0
                WRITE(*,*)'INCREASING DELTA DIST TO:',DISTSRCH
                GOTO 11
C                WRITE (*,*)' WANT TO EXIT? (Y/N)-->'
C                READ(*,'(A)')CH
C                IF ((CH.EQ.'Y').OR.(CH.EQ.'y')) GOTO 999
                END IF
900          CONTINUE
C             CLOSE(1)
C             CLOSE(2)
950          CONTINUE
             NH1=1
             ISN1=1
990          CONTINUE
999          END
C
        FUNCTION SQT(E,DH,R)
        IMPLICIT REAL*8(A-H,O-Z)
        Q=R*DSIN(E)
        Y=2.0D0*R+DH
C>>     WRITE(*,*)' DH=',DH
        IF (DH.LT.0) SQT=-Q-DSQRT(DABS(Q*Q+DH*Y))
        IF (DH.GE.0) SQT=-Q+DSQRT(DABS(Q*Q+DH*Y))
c        if (q*q+dh*y .lt. 0) then
c           write(*,*)'-q,quo,sqt=',-q,q*q+dh*y,sqt
c           end if
        RETURN
        END
C
        FUNCTION FUN(H)
        IMPLICIT REAL*8(A-H,O-Z)
        COMMON/ZZ/HJ(7),TJ(7),PJ(7),AT(7),CT(7)
        IF (H.LT.0) THEN
           M = 1
           GOTO 15
           END IF
        DO10 N=1,7
           M=N-1
           IF(H-HJ(N))15,15,10
10      CONTINUE
15      T=TJ(M)+AT(M)*(H-HJ(M))
        P=PJ(M)*(T/TJ(M))**CT(M)
        FUN=79.D-6*P/T
25      RETURN
        END

*/

/* a few references on function usage


		nyr = nyear;
	//  year for calculation
		nyd = nyr - RefYear;

        ////////////initialize weather constants for refraction//////////////
		InitWeatherRegions( &geo, &avekmxzman, &avekmyzman, &lt, &lg, &nweather,
							&ns1, &ns2, &ns3, &ns4 );
        /////////////////////////////////////////////////////////////////

		//---------------------------------------
		td = geotz;
		//       time difference from Greenwich England
		//---------------------------------------


		lr = cd * lt;
		tf = td * 60. + lg * 4.;


	// ----initialize astronomical constants to time zone and year---------------------------
       InitAstConst(&nyr, &nyd, &ac, &mc, &ec, &e2c, &ap, &mp, &ob, &nyf, &td);
    //---------------------------------------------------------------------------------------

			dy1 = dy; //day

			///////////////////noon/////////////////////////////////////////////////////
			AstroNoon( &dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__, &tf, &t6);
			////////////////////////////////////////////////////////////////////////////


	        //astronomical refraction calculations
			//mishor hgt = 0
			//astronomical hgt = real hgt

			///////////Astronomical total refraction/////////////////////////////////////////////////////////
			AstronRefract( &avehgtzman, &lt, &nweather, &dy, &ns1, &ns2, &air, &a1 );
            ///////////////////////////////////////////////////////////////////

//              calculate astron./mishor sunsets
//              Nsetflag = 1 or Nsetflag = 3 OR Nsetflag = 5
			///////////Astronomical/Mishor sunset///////////////////////////////////////////////////////////
            AstronSunset(&ndirec, &dy, &lr, &lt, &t6, &mp, &mc, &ap, &ac, &ms, &aas, &es,
						&ob, &d__, &air, &sr1, &nweather, &ns3, &ns4, &adhocset, &t3,
						&astro, &azi1, &jdn, &mday, &mon, &nyear, &nyl, &td,
						mint, &WinTemp, &FindWinter, &EnableSunsetInv);
			///////////////////////////////////////////////////////////////////////////////////////////////

//            Nsetflag = 0 or Nsetflag = 2 OR Nsetflag = 4
// 	         calculate astron./mishor sunrise
			///////////Astronomical/Mishor sunrise///////////////////////////////////////////////////////////
			AstronSunrise(&ndirec, &dy, &lr, &lt, &t6, &mp, &mc, &ap, &ac, &ms, &aas,
						&es, &ob, &d__, &air, &sr1, &nweather, &ns3, &ns4, &adhocrise,
						&t3, &astro, &azi1, &jdn, &mday, &mon, &nyear, &nyl, &td,
						mint, &WinTemp, &FindWinter, &EnableSunriseInv);

 */
