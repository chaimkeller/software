//Globals.c

#include "Main.h"

#define bool int
#define true 1
#define false 0

///////////////////////////////////////////////////////////////////////

/************file folder paths/weather modeling****************/
//eventually these will be stored in a file with a php interface
//with values for Windows and values for Linux
char drivjk[100] = ""; //path to hebrew strings
char drivweb[100] = ""; //path to web files
char drivfordtm[100] = ""; //path to fordtm (temp directory)
char drivcities[100] = ""; //path to cities directory

//////////weather modeling////////////////
float latzone[4]; //four weather latitude zones excluding Eretz Yisroel
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

/////////astronomical constansts/////////////////
double ac0, mc0, ec0, e2c0, ob10, ap0, mp0, ob0;
short RefYear;

////////////////time cushions////////////////////
float cushion[5]; //EY, GTOPO30, SRTM-1, SRTM-2, astronomical calculations

/************************************************************/

//////////////multi-dimen.arrayrs//////////
//variables pertaining to netzski6
double dobs[2][2][366]; 
double tims[2][7][366]; 
double azils[2][2][366];

char heb1[71][100]; //[sunrisesunset] NOTICE: largest hebrew string is 99 characters long
char heb2[21][100];//[newhebcalfm]
char heb3[33][100];//[zmanlistfm]
char heb4[7][6]; //[hebweek] NOTICE: largest hebrew string is 5 characters long
char heb5[40][9];//[holidays] NOTICE largest hebrew string is 8 characters long
char holidays[2][12][40]; //NOTICE: largest holiday string is 40 characters long
char arrStrParshiot[2][63][40]; //NOTICE: largest parshia name is 40 characters long
char arrStrSedra[2][55][40]; //NOTICE: Sedra length must equla Parshiot name	
char stortim[7][13][31][40]; //actually stors times, parshios, days, etc.
                              //NOTICE: maximum length = length of Parshios
char zmantimes[49][385][9];
short mmdate[2][12];
char mdates[3][13][12];
char zmannames[49][100];
short yrend[2];
short yrstrt[2];
short FirstSecularYr = 0;
short SecondSecularYr = 0;
char storheader[2][6][255];
char monthh[2][14][8];


ELEVAT *elev=0;

/////////buffer of places to calculate times////////
short nplac[2] = {0, 0};
char fileo[2][200][100]; 
double latbat[2][200];
double lonbat[2][200];
double hgtbat[2][200];
double avekmx[2];
double avekmy[2];
double avehgt[2];

////////strings///////////////
char TitleLine[100] = "";
char SponsorLine[2][255];
char eroscity[50] = "";
char citnamp[255] = "";
char title[150] = "";
char currentdir[100] = "";
char eroscountry[50] = "---";
char erosareabat[50] = "---";
char astname[255] = "";
char engcityname[100] = "";
char hebcityname[100] = "";
char MetroArea[50] = "";
char country[50] = "";
char TableType[5] = "";

//////global non string variables///////
short datavernum = 0;
short progvernum = 0;
short yr = 5767;
float distlim = 0;
float distlims[4];
bool parshiotEY = 0;
bool parshiotdiaspora = 0;
double eroslatitude = 0;
double eroslongitude = 0;
double eroshgt = 0;
float searchradius = 3;
double pi, cd, pi2, ch, hr; //3.14..., conv deg to rad, 2*pi, hours/rad, min/hour
short fshabos0 = 0; //keeps tracks of first Shabbos
short rhday = 0; //keeps track of day of Rosh Hashono
short endyr = 12;

float geotz = 2.0;

/////////platform flag//////////
bool Linux = false;

//////////flags//////////////
bool ast = false; //false for mishor, true for astronomical
bool astronplace = false; //= true for non eros non EY astronomical times
bool optionheb = false; // = true for Hebrew tables
bool eros = false; // = true for GTOPO/SRTM/EY neighborhoods
bool geo = false; // = true for using geo coordinates of latitude and longitude
bool eroscityflag = false;
bool NearWarning[2];
short types = 0; //see below
short SRTMflag = 0; //0 for GTOPO30/SRTM30, 1 for SRTM-2, 2 for SRTM-1, 9 for EY neighborhoods

/////output for hebyr//////////
bool hebleapyear = 0;
short yeartype = 0;
short dayRoshHashono = 0;
short yrheb = 0; //Hebrew Year in numbers
char yrcal[6] = ""; //Hebrew Year in Hebrew characters
short steps[2]; //round to steps seconds
short accur[2]; //cushions
short RoundSeconds = 0;

////global constants for Hebrew Calendar///////////////
short n3mon = 793; // monthly change in molad after 1 lunation
short n2mon = 12;
short n1mon = 1;
short n3ylp = 589; // change in molad after 13 lunations of leap year
short n2ylp = 21;
short n1ylp = 5;
short n3yreg = 876; //change in molad after 12 lunations of reg. year
short n2yreg = 8;
short n1yreg = 4;
short n33 = 0;
short n22 = 0;
short n11 = 0;
short n3rh = 0;
short n2rh = 0;
short n1rh = 0;
short n1rhooo = 0;
short n1rho = 0;
short ncal3 = 0;
short ncal2 = 0;
short ncal1 = 0;
short n1rhoo = 0;
short nt3 = 0;
short nt2 = 0;
short nt1 = 0;
short rh2 = 0;
short leapyear = 0;
short yl = 0;
short dyy = 0;
short Hourr = 0;
short Min = 0;
short hel = 0;
char dates[12] = "";
char *tm = " PM afternoon";
short stdyy = 0;
short leapyr2 = 0;
char hdryr[] = "-2007";
short styr = 0;
short yrr = 0;
