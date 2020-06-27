

#include "Main.h"
#include <fcntl.h>
#include <string.h>
#include <stdlib.h>

/* #include <io.h>
#include <math.h>
#include <stdio.h>
#include <sys/types.h>
#include <sys/stat.h>

#include <algorithm>
#include <mbstring.h>
#include <tchar.h>
#include <iostream>
#include "dirent.h" 
#include <errno.h> */

#define bool int
#define true 1
#define false 0

extern float latzone[4]; //four weather latitude zones excluding Eretz Yisroel
                  //used both for Northern Hemisphere, and mirror image for Southern Hemisphere
                  //latzone[0] = maximum latitude for tropical atmosphere with year long summer refraction
                  //latzone[1] = Mediterranian climate between latitudes latzone[0] and latzone[1]
                  //latzone[2] = mid northern climate between latitudes latzone[1] and latzone [2]
                  //latzone[3] = northern climate between latitudes latzone[3] to latzone[4]
				  //latzone[4] = far northern climate for latitudes larger than latzone[4]
extern short sumwinzone[5][4]; //weather zone winter summer ranges
                        //sumwinzones[0-4] correspond to latzones[0-4]
                        //sumwinzones[5] is for Eretz Yisroel
                        //sumwinzone[0] = end of winter refractions
                        //sumwinzone[1] = beginning of winter refraction
                        //sunwinzome[2] = end of ad hoc fix for winter refraction
                        //sumwinzone[3] = beginning of ad hoc fix for winter refraction
extern short nacurrFix;  //ad hoc shift of winter and summer time in seconds

//based on output from program Menat.for for winter and summer Northern Hemisphere mean-latitude atmospheres
//for heights of 1 - 999 meters.
extern double winrefo, sumrefo;
extern double sumref[8];
extern double winref[8];

/////////astronomical constansts/////////////////
extern double ac0, mc0, ec0, e2c0, ob10, ap0, mp0, ob0;
extern short RefYear;

////////////////time cushions////////////////////
extern float cushion[5]; //EY, GTOPO30, SRTM-1, SRTM-2, astronomical calculations

/************************************************************/


extern ELEVAT *elev;
extern double dobs[2][2][366]; 
extern double tims[2][7][366]; 
extern double azils[2][2][366];

extern char heb1[71][100]; //[sunrisesunset] NOTICE: largest hebrew string is 99 characters long
extern char heb2[21][100];//[newhebcalfm]
extern char heb3[33][100];//[zmanlistfm]
extern char heb4[7][6]; //[hebweek] NOTICE: largest hebrew string is 5 characters long
extern char heb5[40][9];//[holidays] NOTICE largest hebrew string is 8 characters long
extern char holidays[2][12][40]; //NOTICE: largest holiday string is 40 characters long
extern char arrStrParshiot[2][63][40]; //NOTICE: largest parshia name is 40 characters long
extern char arrStrSedra[2][55][40]; //NOTICE: Sedra length must equla Parshiot name	
extern char stortim[7][13][31][40]; //actually stors times, parshios, days, etc.
                              //NOTICE: maximum length = length of Parshios
extern char zmantimes[49][385][9];
extern short mmdate[2][12];
extern char mdates[3][13][12];
extern char zmannames[49][100];
extern short yrend[2];
extern short yrstrt[2];
extern short FirstSecularYr;
extern short SecondSecularYr;
extern char storheader[2][6][255];
extern char monthh[2][14][8];

/////////buffer of places to calculate times////////
extern short nplac[2];
extern char fileo[2][200][100]; 
extern double latbat[2][200];
extern double lonbat[2][200];
extern double hgtbat[2][200];
extern double avekmx[2];
extern double avekmy[2];
extern double avehgt[2];

////////strings///////////////
extern char TitleLine[100];
extern char SponsorLine[2][255];
extern char eroscity[50];
extern char citnamp[255];
extern char title[150];
extern char currentdir[100];
extern char eroscountry[50];
extern char erosareabat[50];
extern char astname[255];
extern char engcityname[100];
extern char hebcityname[100];
extern char MetroArea[50];
extern char country[50];
extern char TableType[5];

//////global non string variables///////
extern short datavernum;
extern short progvernum;
extern short yr;
extern float distlim;
extern float distlims[4];
extern bool parshiotEY;
extern bool parshiotdiaspora;
extern double eroslatitude;
extern double eroslongitude;
extern double eroshgt;
extern float searchradius;
extern double pi, cd, pi2, ch, hr; //3.14..., conv deg to rad, 2*pi, hours/rad, min/hour
extern short fshabos0; //keeps tracks of first Shabbos
extern short rhday; //keeps track of day of Rosh Hashono
extern short endyr;

extern float geotz;

/////////platform flag//////////
extern bool Linux;

//////////flags//////////////
extern bool ast; //false for mishor, true for astronomical
extern bool astronplace; //= true for non eros non EY astronomical times
extern bool optionheb; // = true for Hebrew tables
extern bool eros; // = true for GTOPO/SRTM/EY neighborhoods
extern bool geo; // = true for using geo coordinates of latitude and longitude
extern bool eroscityflag;
extern bool NearWarning[2];
extern short types; //see below
extern short SRTMflag; //0 for GTOPO30/SRTM30, 1 for SRTM-2, 2 for SRTM-1, 9 for EY neighborhoods

/////output for hebyr//////////
extern bool hebleapyear;
extern short yeartype;
extern short dayRoshHashono;
extern short yrheb; //Hebrew Year in numbers
extern char yrcal[6]; //Hebrew Year in Hebrew characters
extern short steps[2]; //round to steps seconds
extern short accur[2]; //cushions
extern short RoundSeconds;

////global constants for Hebrew Calendar///////////////
extern short n3mon; // monthly change in molad after 1 lunation
extern short n2mon;
extern short n1mon;
extern short n3ylp; // change in molad after 13 lunations of leap year
extern short n2ylp;
extern short n1ylp;
extern short n3yreg; //change in molad after 12 lunations of reg. year
extern short n2yreg;
extern short n1yreg;
extern short n33;
extern short n22;
extern short n11;
extern short n3rh;
extern short n2rh;
extern short n1rh;
extern short n1rhooo;
extern short n1rho;
extern short ncal3;
extern short ncal2;
extern short ncal1;
extern short n1rhoo;
extern short nt3;
extern short nt2;
extern short nt1;
extern short rh2;
extern short leapyear;
extern short yl;
extern short dyy;
extern short Hourr;
extern short Min;
extern short hel;
extern char dates[12];
extern char *tm;
extern short stdyy;
extern short leapyr2;
extern char hdryr[];
extern short styr;
extern short yrr;

