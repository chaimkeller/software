// stdafx.cpp : source file that includes just the standard includes
//	StringRoutines.pch will be the pre-compiled header
//	stdafx.obj will contain the pre-compiled type information

#include "Main.h"
#include <fcntl.h>
#include <io.h>
#include <math.h>
#include <stdio.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <string.h>
#include <stdlib.h>
#include <algorithm>
#include <mbstring.h>
#include <tchar.h>
#include <iostream>
#include "dirent.h" 
#include <errno.h>


//**********declare functions*****************/
void hebnum(short k, char hebnum[]);
char *hebweek( short dayweek, char ch[] );
short RdHebStrings();
short LoadParshiotNames();
void ParseSedra(char doclin[255], short PMode);
void InsertHolidays(char *calday, short i, short k);
void HebCal(short yrheb);
//int isspace ( char c );
char *LTrim(char * p);
char *RTrim(char * p);
char *Trim(char * str);
char *Mid(char *str, short nstart, short nsize, char buff[]);
char *Mid(char *str, short nstart, short nsize);
int InStr( short nstart, char * str, char * str2 );
int InStr( char * str, char * str2 );
//char *Replace( char *str, char *Old, char *New, char buff[] );
char *Replace( char *str, char Old, char New);
char *LCase(char *str, char buff[]);
//char *MyStrcpy(char *src,char *des);
double Fix( double x );
short Cint( double x );
char *hebyear(short yr);
void WriteTables(FILE *stream2);
void heb12monthHTML(FILE *stream2, short setflag);
void heb13monthHTML(FILE *stream2, short setflag);
void timestring( short ii, short k, short iii, FILE *stream2 );
void parseit( char outdoc[], FILE *stream2, bool closerow, short linnum);
short CheckInputs(char *strlat, char *strlon, char *strhgt);
short Caldirectories();
short SunriseSunset(short types, FILE *stream);
short CreateHeaders(short types);
short ReadHebHeaders();
short PrintSingleTable(short setflag, short types, FILE *stream2);
short calnearsearch( char *batfile, char *sun );
short netzskiy( char *batfile, char *sun );
short WriteUserLog( int UserNumber, short zmanyes, short typezman );
void ErrorHandler( short mode, short ier );

short iswhite(char c);
short find_first_nonwhite_character(char *str, short start_pos);
short find_last_nonwhite_character(char *str, short start_pos);

//char *DosToIso( char *str, char buff[] );
char *DosToIso( char *str, char *buff );
short LoadConstants();
char *round( double tim, short setflag, short sp, char t3subb[] );
char *EngCalDate(short ndy, short EngYear, short yl, char *caldate );
void PST_iterate( short *k, short *j, short *i, short *ntotshabos, short *dayweek
				 , short *yrn, short *fshabos, short *nshabos, short *addmon
				 , short setflag, short sp, short types, short yl, short iheb );
short WriteHtmlHeader( FILE *stream2 );
short WriteHtmlClosing( FILE *stream2 );

////////////////////////netzski6///////////////////////////////////////
extern "C"
{
short netzski6(char *fileo, double kmyo, double kmxo, double hgt
			              , float geotd, short nsetflag, short nfil);
}
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

/////////////////////////////////////////////////////////////////////

int main(int argc, char* argv[])
{
	///////local variables for main program///////////////////////////////////
	int UserNumber = 0;
	char doclin[255] = "";
	short foundhebcity = 0;
	short typeflag;
	short zmanyes;
	short typezman;
	short zmantype;
	short numUSAcities1 = 0;
	short numUSAcities2 = 0;
	float searchradiusin = 0;
	short yesmetro = 0;
	short cnum = 0;
	short pos = 0;
	char *eroslatitudestr = "";
	char *eroslongitudestr = "";
	char *eroshgtstr = "";
	double eroslatitudein = 0;
	double eroslongitudein = 0;
	double eroshgtin = 0;
    short RoundSecondsin = 0;
	char eroscityin[50] = "";
	double lat, lon, hgt;
    double avekmxnetz = 0; //parameters used for astronomical z'manim tables
    double avekmynetz = 0;
    double avehgtnetz = 0;
    double avekmxskiy = 0;
    double avekmyskiy = 0;
    double avehgtskiy = 0;
    bool aveusa = true; //invert coordinates in SunriseSunset due to
                        //silly nonstandard coordinate format
	bool viseros = false;
	short k = 23;
	char filroot[100] = "";
	char buff[255] = "";
	short ier = 0;
	FILE *stream2;


	/////////////LoadConstants////////////////////////////////

	LoadConstants();

	//find version number of netzski6 calculations//////////////
	sprintf( filroot, "%s%s", drivcities, "version.num" );

	if ( (stream2 = fopen( filroot, "r" ) ) )
	{
		fscanf( stream2, "%d\n", &datavernum );
        fclose( stream2 );
	}
	else //can't find version file, use default
	{
		datavernum = 12;
	}

	////////////////process inputed parameters for php pages///////////////
	strcpy( TableType, "Chai" );//"Astr" );//"Chai" )//"BY" );
	yesmetro = 0;
    strcpy( MetroArea, "achituv" ); //"telz_stone_ravshulman"; //"טלז_סטון";
    strcpy( country, "Israel" );//"USA";
	UserNumber = 302343;
    yrheb = 5767;
	typeflag = 1; //0 = visible, 1 = mishor, 3 = astron.
	zmanyes = 1; //0 = no zemanim, 1 = zemanim
	typezman = 1; // acc. to which opinion
	optionheb = true;
	zmantype = 1;
	RoundSeconds = 5;
	///unique to Chai tables
    numUSAcities1 = 1; //Metro area
    numUSAcities2 = 1; //city area
    searchradius = 4; //search area
    yesmetro = 0;  //yesmetro = 0 if found city, = 1 if didn't find city and providing coordinates
	eroslatitudein = 34.0;
	eroslongitudein = 118.0;
	searchradiusin = 3;
	eroshgtin = 130;
    UserNumber = 302343;
    yrheb = 5767;
	geotz = -8;

	///////////////////astronomical type//////////////////////
	ast = true; //= false for mishor, = true for astronomical
	//////////////////////////////////////////////////////////

    ////////////table type (see below for key)///////
	types = 1; //3; 
	////////////////////////////////////////////////

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
	  nsetflag = 3; //astronomical sunset for ght != 0, otherwise mishor sunset/chazos
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
	default:
	  {
	  ErrorHandler( 1, 2);; //uncrecogniable "types"
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



///////////////////////////////end of input parameters/////////////////////////

	//check the inputs for errors --- inputs will be strings that need to be converted, etc
	//before converting them, check that they are really numbers and that the numbers
	//fall between certain values
	if (yesmetro == 1 || TableType == "Astr") //inputing coordinates so check them
	{
	  eroslatitudestr = "34.432";
	  eroslongitudestr = "118.342";
	  eroshgtstr = "83.2";

      if (CheckInputs( eroslatitudestr, eroslongitudestr, eroshgtstr))
	  {
		  ErrorHandler(1, 0); //something wrong with inputs <<<<<<<<<<<<<<
		  return -1;
	  }

    	//now convert from the input string to the values
	   eroslatitudein = atof( eroslatitudestr);
	   eroslongitudein = atof( eroslongitudestr);
	   eroshgtin = atof(eroshgtstr);
    }

    if (RoundSecondsin != 0)
    {
 	   RoundSeconds = RoundSecondsin;
    }

	
    steps[0] = RoundSeconds;
    steps[1] = RoundSeconds;

///////////////////determine which table will be generated///////////////////////

	if (optionheb) hebyear( yrheb ); //convert Hebrew year in English to Hebrew

	short langMetro = 0;

    if ( InStr( TableType, "BY") ) //Eretz Yisroel tables based on 25 m DTM
	{
		
		//if ( (char)Mid( MetroArea, 1, 1, buff) < (char)(128) ) //English table
		if (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", Mid( MetroArea, 1, 1, buff)) ) //English table
		{
			langMetro = 0;
		}
		else // if ( (char)Mid( MetroArea, 1, 1, buff) >= (char)(128) )  //Hebrew table
		{
			langMetro = 1;
		}

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
			fgets(doclin, 255, stream2); //read in line of text
			strncpy( engcityname, doclin, strlen(doclin) - 1);
			engcityname[strlen(doclin) - 1] = 0;

			fgets(doclin, 255, stream2); //read in line of text
			strncpy( hebcityname, doclin, strlen(doclin) - 1);
			hebcityname[strlen(doclin) - 1] = 0;

			if (Linux) 
			{
				if (langMetro == 0)
				{
				  DosToIso( MetroArea, buff );
				  strcpy( MetroArea, buff );
				}
				DosToIso( hebcityname, buff );
				strcpy( hebcityname, buff );
			}

			if ( ( langMetro == 0 && strstr(MetroArea, engcityname) ) ||
				 ( langMetro == 1 && strstr(MetroArea, hebcityname) )    )
				{
				sprintf( currentdir, "%s%s%s", drivcities, engcityname, "/" );
				foundhebcity = 1;
				break;
				}
			}

			fclose( stream2 );
			if (!foundhebcity ) //Hebrew city equivalent not found
			{
				ErrorHandler(1, 4);
				return -1;
			}

		}

        strcpy( eroscountry, "Israel" );
	    strcpy( eroscity, engcityname );

		eros = false;
		geo = false;
		geotz = 2;

		//remove the underlines in the city names
		Replace( engcityname, '_', ' ');
		Replace( hebcityname, '_', ' ');

	}
	else if ( InStr( TableType, "Chai") ) //World-wide tables based on SRTM
	{
        parshiotdiaspora = true;
	    strcpy( eroscountry, country );
	    eros = true;
	    geo = true;
	    eroscityflag = true;
	    //now read in parameters
	    eroscity[0] = 0;
	    eroslatitude = eroslatitudein;
	    eroslongitude = eroslongitudein;
	    eroshgt = eroshgtin;
		searchradius = searchradiusin;


 	  sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities1.dir");

  	  /* write closing clauses on html document */
	  if ( !( stream2 = fopen( filroot, "r")) )	   
	  {
		ErrorHandler( 1, 5);; //can't open cities1 file
		return -1;
	  } 
	  else //find English equivalent of Hebrew Metro name
	  {
		cnum = 0;
		while ( cnum < numUSAcities1 )
		{
			fgets(doclin, 255, stream2); //read in line of text
			cnum++;
		}
		fclose( stream2 );
	  }

      //this will define the eroscity name unless flagged by
      //yesmetro = 1 to use the inputed placename and coordinates
      strncpy( erosareabat, doclin, strlen(doclin) - 1 );
	  erosareabat[strlen(doclin) - 1] = 0;

	  //remove _area_country to form the english Metro area name
	  strcpy( engcityname, erosareabat );
	  engcityname[InStr( erosareabat, "_" ) - 1] = 0; //terminate string before "_"
  
      if (yesmetro == 0 ) // Then 'found the city in the city list, use it
	  {
        sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities2.dir");

	    if ( !(stream2 = fopen( filroot, "r")) ) 	   
		{
			ErrorHandler( 1, 6); //can't open cities2 file
			return -1;
		} 
	    else //find English equivalent of Hebrew Metro name
		{
			cnum = 0;
			while ( cnum < numUSAcities2 )
			{
			fgets(doclin, 255, stream2); //read in line of text
			cnum++;
			}
            fclose( stream2 );
		}
		
		if (InStr( eroscountry, "Israel")) //Israel neighborhoods -- check if Metro area corresponds to city area
		{
			if ( !InStr( doclin, engcityname) ) //they don't correspond, so exit with error message
			{
				ErrorHandler( 1, 7);
				return -1;
			}
			geo = false; //treat coordinates as ITM and not as geographical coordinates of longitude and latitude
		}

		pos = InStr( doclin, "/" );
		strcpy( eroscity, Mid( doclin, pos + 1, strlen(doclin) - pos - 1 ) );

		//now find eroslongitude, eroslatitude
		short foundcity = 0;
		short poscity1 = 0;
		short poscity2 = 0;
		char checkcity[50] = "";

        sprintf( filroot, "%s%s%s%s", drivcities, "eros/", &eroscountry[0], "cities3.dir");
  	    /* write closing clauses on html document */
	    if ( !(stream2 = fopen( filroot, "r")) ) 	   
		{
			ErrorHandler( 1, 8);; //can't open cities3 file
			return -1;
		} 
	    else //find eroscity 
		{
			while ( !feof(stream2) )
			{
               //fscanf( stream2, "%50c, %g, %g, %g\n", &doclin[0], &eroslatitude, &eroslongitude, &eroshgt );
			   fgets( doclin, 255, stream2 );

			   //parse out city name from "doclin"
			   poscity1 = InStr(doclin, "/");
			   poscity2 = InStr(doclin, "\\" );
			   if (poscity1 != 0 && poscity2 !=0 )
			   {
        		   strncpy( checkcity, Mid( doclin, poscity1 + 1, poscity2 - poscity1 - 1, buff ), poscity2 - poscity1 - 1 );

				   if ( strstr( doclin, erosareabat ) )
				   {
					   if (strstr( eroscity, checkcity ) )
					   {
						   foundcity = 1; //its the right place
						   //so parse out coordinates
						   pos = InStr( poscity2 + 2, doclin, "," );
						   eroslatitude = atof(Mid( doclin, poscity2 + 2, pos - poscity2 - 1, buff));
						   poscity1 = InStr( pos + 1, doclin, ",");
						   eroslongitude = atof(Mid( doclin, pos + 1, poscity1 - pos, buff));
						   eroshgt = atof(Mid( doclin, poscity1 + 1, strlen(doclin) - poscity1, buff));
						   break;
					   }
				   }
			   }
			}
			fclose( stream2 );
			if ( foundcity == 0 ) //didn't find appropriate city
			{
				ErrorHandler( 1, 9);; //return with error message that Metro area and city area don't match
				           //Note: this error should never occur if Javascript is working in php interface
				return -1;
			}
		}

	  }
	  else //use inputed coordinates through cgi
	  {
       strcpy( eroscity, eroscityin );
       eroslatitude = eroslatitudein;
       eroslongitude = eroslongitudein;
       eroshgt = eroshgtin;
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
		  break;
	  }

      strcat( erosareabat, "_" );
	  strcat( erosareabat, eroscountry );
      sprintf( currentdir, "%s%s%s%s", drivcities, "eros/", &erosareabat, "/" );

	}
	else if ( InStr( TableType, "Astr") ) //Astronomical calculations
	{
		strcpy( country, "המקום" );
		strcpy( astname, country );
		strcpy( hebcityname, country );
		strcpy( engcityname, country );
		geo = true;
		astronplace = true;
        lat = eroslatitudein;
        lon = eroslongitudein;
        
		if (ast) //ast = true, astronomical times
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

		if ( lon < -34.21 &&  lon> -35.8333 &&
             lat > 29.55 && lat < 33.3417 && geotz == 2 )
		{
         parshiotEY = true;  //'inside the borders of Eretz Yisroel
		}
        else
		{
         parshiotdiaspora = true; //somewhere in the diaspora
		}

		switch (types)
		{
		case 2:
          nplac[0] = 1;
          avekmxnetz = lon; //parameters used for z'manim tables
          avekmynetz = lat;
          avehgtnetz = hgt;
		  break;
		case 3:
     	  nplac[1] = 1;
          avekmxskiy = lon;
          avekmyskiy = lat;
          avehgtskiy = hgt;
		case 8:
          nplac[0] = 1;
     	  nplac[1] = 1;
          avekmxnetz = lon; //parameters used for z'manim tables
          avekmynetz = lat;
          avehgtnetz = hgt;
          avekmxskiy = lon;
          avekmyskiy = lat;
          avehgtskiy = hgt;
		  break;
		}

        aveusa = true; //invert coordinates in SunriseSunset due to
                       //silly nonstandard coordinate format

       sprintf( currentdir, "%s%s%s%s", drivcities, "ast/", Trim(astname), "/" );
	}

    //write the parameters to the userlog.log file
	if (WriteUserLog( UserNumber, zmanyes, typezman ) ) return -1;


////////////////////////<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>

	if (Caldirectories()) return -1;

	if (RdHebStrings()) return -1 ;

	if (LoadParshiotNames()) return -1;

	if (WriteHtmlHeader(stream2)) return -1; //write header info

	if (SunriseSunset(types, stream2)) return -1;

	if (WriteHtmlClosing(stream2)) return -1; //write the html closing info and close the stream

	return 0;
}

/////////////////////////hebnum//////////////////////////////////////
void hebnum(short k, char ch[])
///////////////////////////////////////////////////////////////////////
{
//PURPOSE: to generate Hebrew calendar date letters k=1,2,3 makes ד,ג,ב,א, etc.
//Revised Last: 04/30/2007
//Programmer: Chaim Keller
// no error code returns
////////////////////////////////////////////////////////////////////////////

	if (k <= 10)
	{
		ch[0] = (char)(k + 223);
		ch[1] = 0;
	}
	else if (k > 10 && k < 20)
	{
		ch[0] = (char)(233);
		ch[1] = (char)(k - 10 + 223);
		ch[2] = 0;
	}
	else if (k == 20)
	{
		ch[0] = (char)(235);
		ch[1] = 0;
	}
	else if (k > 20 && k < 30)
	{
		ch[0] = (char)(235);
		ch[1] = (char)(k - 20 + 223);
		ch[2] = 0;
	}
	else if (k == 30)
	{
		ch[0] = (char)(236);
		ch[1] = 0;
	}
	if (k == 15)
	{
		ch[0] = (char)(232);
		ch[1] = (char)(229);
		ch[2] = 0;
	}
	if (k == 16)
	{
		ch[0] = (char)(232);
		ch[1] = (char)(230);
		ch[2] = 0;
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

	short filtmp = 0;
	short flgheb = 0;
	short iheb = 0;
	short hebnumb = 0;
	short lendoc = 0;
	char buff[255] = "";
	char myfile[100] = "";
	char doclin[255]; //used for reading lines of text
	FILE *stream;

	if (optionheb) //read in hebrew strings from calhebrew.txt 
	{
		sprintf( myfile, "%s%s%s", drivjk, "calhebrew.txt", "\0" );

		if ( !( stream = fopen( myfile, "r")) ) 	   
		{
			//can't find hebrew strings, so use English strings as default
			optionheb = false;
			ErrorHandler(2, 1); //return with error flag set
			return -1;
		}
		else
		{
			iheb = 0;

			while( !feof( stream ) )
			{
				doclin[0] = 0; //clear the buffer
				
				fgets(doclin, 255, stream); //read in line of text

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
            
				if (strlen(doclin) > 1 ) //ignore blank lines containing just "\n"
				{
					//remove trailing carriage return character and terminate the string
					short lendoc = strlen(doclin) - 1;
					strncpy( buff, doclin, lendoc );
					buff[lendoc] = 0;
					sprintf( doclin, "%s%s", buff, "\0");

					if (Linux) //convert DOS Hebrew to ISO Hebrew
					{
						DosToIso( doclin, buff );
						strcpy( doclin, buff );
					}

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
					}

				iheb++;
				}
			}
			fclose(stream);
		}
	}

	//now read sponsorship information
	sprintf( myfile, "%s%s%s", drivweb, "SponsorLogo.txt", "\0" );

    if ( ( stream = fopen( myfile, "r")) ) 	   
	{
		while( !feof( stream ) )
		{
			doclin[0] = 0; //clear the string buffer

			fgets(doclin, 255, stream); //read in line of text

			if ((strstr( doclin,"[English]")) &&  !optionheb)
			{
				//english sponsor info
        		fgets(doclin, 255, stream); //read in line of text
				strncpy( &SponsorLine[1][0], doclin, strlen(doclin) - 1);
            	SponsorLine[1][strlen(doclin) - 1] = 0;
			}

			if ((strstr( doclin,"[titles]")) && !optionheb)
			{
				//english tables' title info
				if (!eros && !astronplace && !ast) //Eretz Yisroel tables
				{
                    fgets(doclin, 255, stream); //read in line of text
					strncpy( TitleLine, doclin, strlen(doclin) - 1);
					TitleLine[strlen(doclin) - 1] = 0;
				}
				else //world tables or astronomical calculations
				{
            		fgets(doclin, 255, stream); //read in line of text
			        fgets(doclin, 255, stream); //read in line of text
			        fgets(doclin, 255, stream); //read in line of text
					strncpy( TitleLine, doclin, strlen(doclin) - 1);
					TitleLine[strlen(doclin) - 1] = 0;
				}
			}

			if ((strstr( doclin,"[Hebrew]")) && optionheb)
			{
				//hebrew sponsor info
           		fgets(doclin, 255, stream); //read in line of text
				strncpy( &SponsorLine[0][0], doclin, strlen(doclin) - 1);
				SponsorLine[0][strlen(doclin) - 1] = 0;

				if (Linux) //convert DOS Hebrew to ISO Hebrew
				{
					DosToIso( &SponsorLine[0][0], buff );
					strcpy( &SponsorLine[0][0], buff );
					DosToIso( &SponsorLine[1][0], buff );
					strcpy( &SponsorLine[1][0], buff );
				}

			}

			if ((strstr( doclin,"[titles]")) && optionheb)
			{
				//hebrew tables' title info
				if (!eros && !astronplace && !ast) //Eretz Yisroel tables
				{
            		fgets(doclin, 255, stream); //read in line of text
            		fgets(doclin, 255, stream); //read in line of text
 				}
				else //tables for the world or astronomical tables
				{
            		fgets(doclin, 255, stream); //read in line of text
            		fgets(doclin, 255, stream); //read in line of text
            		fgets(doclin, 255, stream); //read in line of text
            		fgets(doclin, 255, stream); //read in line of text
				}
				strncpy( TitleLine, doclin, strlen(doclin) - 1);
				TitleLine[strlen(doclin) - 1] = 0;

				if (Linux) //convert DOS Hebrew to ISO Hebrew
				{
					DosToIso( TitleLine, buff );
					strcpy( TitleLine, buff );
				}

				strncpy(buff, &heb3[1][0], 1);
				buff[1] = 0;

				sprintf( &heb1[3][0],"%s%s%s%s%s%s",&heb1[1][0],"\"",TitleLine,"\""," ",buff);

				strncpy( &heb1[4][0], TitleLine, strlen(TitleLine) );

				strncpy(buff, &heb3[12][9], 16);
				buff[16] = 0;

				sprintf( &heb3[12][0],"%s%s%s%s%s%s",&heb1[1][0],"\"",TitleLine,"\""," ",buff);

			}

		}
		fclose(stream);
	}
	else //sponsor file not found, use defaults
	{
		if (optionheb)
		{
			if (eros)
			{
				strcpy( TitleLine, &heb1[1][0] );
				strcat( TitleLine, "\0");

				if (Linux) //convert DOS Hebrew to ISO Hebrew
				{
					DosToIso( TitleLine, buff );
					strcpy( TitleLine, buff );
				}

			}
			else
			{
				strcpy( TitleLine, &heb1[3][0] );
				strcat( TitleLine, "\0");

				if (Linux) //convert DOS Hebrew to ISO Hebrew
				{
					DosToIso( TitleLine, buff );
					strcpy( TitleLine, buff );
				}

			}
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
		}

	}

	return 0;
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
	
	char myfile[100] = "";
	char doclin[255]; //used for reading lines of text
	char buff[255] = "";
	short i = 0;
	short ier = 0;
	short errlog = 0;
	short lognum = 0;
	short nfound = 0;
	short numhol = 0;
	short nn = 0;
	short PMode = 0;
	short lendoc = 0;
	short parnum = 0;
	char Dparsha[16] = ""; // Torah reading for Diaspora
	char Iparsha[17] = ""; // Torah reading for Eretz Yisroel
	FILE *stream;

	//determine cycle name
	switch (hebleapyear)
	{
		case false: //Hebrew calendar non-leapyear
			strcpy( &Iparsha[0], "[0_" );
			break;
		case true: //Hebrew calendar leapyear
			strcpy( &Iparsha[0], "[1_" );
			break;
	}
	switch (dayRoshHashono) //day of the week of RoshHashono
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
	switch (yeartype) //chaser, kesidrah, shalem
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
			doclin[0] = 0; //clear the buffer
			
			fgets(doclin, 255, stream); //read in line of text

			if (strlen(Trim(doclin)) > 2)
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
							lendoc = strlen(doclin) - 1; //find_last_nonwhite_character(doclin, 0);
							strncpy( &arrStrParshiot[PMode][nn][0], doclin, lendoc );
            				arrStrParshiot[PMode][nn][lendoc] = 0; //add string termination

							if (Linux) //convert DOS Hebrew to ISO Hebrew
							{
								DosToIso( &arrStrParshiot[PMode][nn][0], buff );
								strcpy( &arrStrParshiot[PMode][nn][0], buff );
							}

							nn++;
							break;
						case 2:
						case 3:
							lendoc = strlen(doclin) - 1; //find_last_nonwhite_character(doclin, 0);
							strncpy( &holidays[PMode - 2][numhol][0], doclin, lendoc );
            				holidays[PMode - 2][numhol][lendoc] = 0; //add string termination

							if (Linux) //convert DOS Hebrew to ISO Hebrew
							{
								DosToIso( &holidays[PMode - 2][numhol][0], buff );
								strcpy( &holidays[PMode - 2][numhol][0], buff );
							}

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


	//load up the sedra for Eretz Yisroel and the diaspora
	sprintf( myfile, "%s%s%s", drivjk, "calparshiot_data.txt", "\0" );

    if ( (stream = fopen( myfile, "r")) ) 
	{
		PMode = -1;
		nfound = 0;
		while ( !feof(stream) )
		{
			doclin[0] = 0; //clear the buffer
			
			fgets(doclin, 255, stream); //read in line of text

			switch (PMode)
			{
				case -1:
                    if ((strstr( doclin, Iparsha)) &&
						(strstr( doclin, Dparsha))) //parse out both sedra
					{
						PMode = 2;
						nfound += 2;
					}
                    else if ((strstr( doclin, Iparsha)) &&
						(strstr( doclin, Dparsha) == NULL)) //parse out Eretz Yisroel sedra
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
					ParseSedra(doclin, PMode);
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
			ErrorHandler(3, 2);
			return -1;
		}
	}
	else
	{
		ErrorHandler(3, 3);
		return -1;
	}

	return 0;

}

///////////////////////////ParseSedra////////////////////////////////////
void ParseSedra(char doclin[255], short PMode)
/////////////////////////////////////////////////////////////////////////
//parses the parsha data to stuff the Sedra arrays with the Shabbosim and holidays
//////////////////////////////////////////////////////////////////////////////
{
	char myfile[100] = "";
	short errlog = 0;
	short mParsha = 0;
	short i = 0;
	short mParshaD = 0;
	short mParshaI = 0;
	short hc = 0;
	char ch[2] = " ";
	bool doubleparsha = false;
	char parsnum[255] = "";

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
	for (i = 0; i <= strlen(doclin)-1; i++)
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
                        sprintf( &arrStrSedra[PMode][mParsha][0],"%s%s%s",&arrStrParshiot[hc][atoi(parsnum)][0]
							                                    ,"-",&arrStrParshiot[hc][atoi(parsnum) + 1][0]);
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
					parsnum[0] = 0;

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
                        sprintf( &arrStrSedra[0][mParshaI][0],"%s%s%s",&arrStrParshiot[hc][atoi(parsnum)][0]
							                                 ,"-",&arrStrParshiot[hc][atoi(parsnum) + 1][0]);

						strcpy( &arrStrSedra[1][mParshaD][0], &arrStrSedra[0][mParshaI][0]);

						doubleparsha = false; //reset double parsha flag

						mParshaI++; //increment the parsha number
						mParshaD++; //increment the parsha number

					}
				}

				parsnum[0] = 0; //reset the sedra designation number

		}
        else if (strstr( ch, "D"))
		{
				parsnum[0] = 0;
				doubleparsha = true;
		}
        //else if (strstr( ch, " "))
		//{
			// do nothing
		//}
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
void InsertHolidays(char *calday, short i, short k)
//////////////////////////////////////////////////////////////////
//Inserts holiday names into date variable "calday"
////////////////////////////////////////////////////////////////
{
	short errlog = 0;
	short lendoc = 0;
	char l_buff[100] = "";
	bool RemoveUnderline = 0;
	short ir = 0;
	short yeartype = 0;
	char CalParsha[100] = "";
	char caldate[40] = "";
//	CVBString

	//passed variables are:
	//(1) the day field = calday$
	//(2) the increment variables i,k used to determine
	//    (a) the date field = CalDate$
	//    (b) the Shabbos Parsha field = CalParsha$
	//returned ier = -5 for error

    //================================================
    //add parshiot
   

	//Hebrew weekdays
	strcpy( calday, &stortim[4][i - 1][k - 1][0] );

	if ( !optionheb )
	//convert Hebrew weekdays to English weekdays
	{
        if (strstr(calday, &heb4[1][0]) )
		{
				strcpy( calday, "Sunday");
				goto ih50;
		}

        if (strstr(calday, &heb4[2][0]) )
		{
				strcpy( calday, "Monday");
				goto ih50;
		}

        if (strstr(calday, &heb4[3][0]) )
		{
				strcpy( calday, "Tuesday");
				goto ih50;
		}

        if (strstr(calday, &heb4[4][0]) )
		{
				strcpy( calday, "Wednesday");
				goto ih50;
		}

        if (strstr(calday, &heb4[5][0]) )
		{
				strcpy( calday, "Thursday");
				goto ih50;
		}

        if (strstr(calday, &heb4[6][0]) )
		{
				strcpy( calday, "Friday");
				goto ih50;
		}

        if (strstr(calday, &heb4[7][0]) )
		{
				strcpy( calday, "Shabbos");
				goto ih50;
		}
	}

ih50:
	if (parshiotEY)
	{
        if ( !strstr( &stortim[5][i - 1][k - 1][0], "-----") )
 		{
			if (optionheb)
			{
                if (strstr( &stortim[5][i - 1][k - 1][0], &heb4[6][0] ) )
				{
					strcpy( calday, &stortim[5][i - 1][k - 1][0]);
				}
				else
				{
                    sprintf( calday,"%s%s",&heb5[37][0],&stortim[5][i - 1][k - 1][0]);
				}
			}
			else if (! optionheb)
			{
				if (strstr( &stortim[5][i - 1][k - 1][0], "Shabbos" ) )
				{
					strcpy ( calday, &stortim[5][i - 1][k - 1][0] );
				}
				else
				{
                    sprintf( calday,"%s%s","Shabbos_",&stortim[5][i - 1][k - 1][0]);
				}
			}
		}
	}
	else if (parshiotdiaspora)
	{
		if ( !strstr( &stortim[6][i - 1][k - 1][0], "-----") )
 		{
			if (optionheb)
			{
                if (strstr( &stortim[6][i - 1][k - 1][0], &heb4[6][0] )  )
				{
					strcpy( calday, &stortim[6][i - 1][k - 1][0]);
				}
				else
				{
                    sprintf( calday,"%s%s",&heb5[37][0],&stortim[6][i - 1][k - 1][0]);
				}
			}
			else if (! optionheb)
			{
				if (strstr( &stortim[6][i - 1][k - 1][0], "Shabbos" ) )
				{
					strcpy ( calday, &stortim[6][i - 1][k - 1][0] );
				}
				else
				{
                    sprintf( calday,"%s%s","Shabbos_", &stortim[6][i - 1][k - 1][0]);
				}
			}
		}
	}


	//=========================================================
	//now add holidays to Shabbosim

	strcpy (caldate, Trim(&stortim[3][i - 1][k - 1][0]));

	if (parshiotEY)
	{
		strcpy( CalParsha, &stortim[5][i - 1][k - 1][0]);
	}
	else if (parshiotdiaspora)
	{
		strcpy( CalParsha, &stortim[6][i - 1][k - 1][0]);
	}

	if ( !strstr( CalParsha, "-----") )
	{
		//add in Shabbos Chanukah and Shabbos Shushan Purim when appropriate
		if (optionheb)
		{
			if ( strstr(Trim(caldate), &heb5[13][0]) ||
				 strstr(Trim(caldate), &heb5[14][0]) ||
				 strstr(Trim(caldate), &heb5[15][0]) ||
				 strstr(Trim(caldate), &heb5[16][0]) ||
				 strstr(Trim(caldate), &heb5[17][0]) ||
				 strstr(Trim(caldate), &heb5[18][0]) ||
				 strstr(Trim(caldate), &heb5[19][0]) ||
				 strstr(Trim(caldate), &heb5[20][0]) ||
				 strstr(Trim(caldate), &heb5[21][0]) )//Chanukah
			{
					if ((yeartype == 2 || yeartype == 3) && strstr(Trim(caldate), &heb5[21][0]) )
					{
						goto remU;
					}
 
					sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][6][0]);
					goto remU;
			}
			else if ( strstr(Trim(caldate), &heb5[24][0]) ||
				      strstr(Trim(caldate), &heb5[25][0]) )
			{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][8][0]);
					goto remU;
			}
			else
			{
					goto remU;
			}
		}
		else
		{
			if ( strstr(Trim(caldate), "25-Kislev")  ||
                 strstr(Trim(caldate), "26-Kislev")  ||
                 strstr(Trim(caldate), "27-Kislev")  ||
                 strstr(Trim(caldate), "28-Kislev")  ||
                 strstr(Trim(caldate), "29-Kislev")  ||
                 strstr(Trim(caldate), "30-Kislev")  ||
                 strstr(Trim(caldate), "1-Teves")  ||
                 strstr(Trim(caldate), "2-Teves")  ||
                 strstr(Trim(caldate), "3-Teves")  )
			{
					if ((yeartype == 2 || yeartype == 3) && strstr( Trim(caldate),"3-Teves") )
					{
						goto remU;
					}
					sprintf( l_buff, "%s%s%s", &holidays[1][6][0], ",_", calday );
					strcpy( calday, l_buff );
					goto remU;
			}
			else if (strstr(Trim(caldate), "15-Adar") ||
                     strstr(Trim(caldate), "15-Adar II") ) //Shushan Purim
			{
					sprintf( l_buff, "%s%s%s", &holidays[1][8][0], ",_", calday );
					strcpy( calday, l_buff );
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
		if (strstr(Trim(caldate), &heb5[1][0]) ||
		    strstr(Trim(caldate), &heb5[2][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][0][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[3][0]))
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][1][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[4][0]))
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][2][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[5][0]))
		{
				if (parshiotEY)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][3][0]);
				}
				else if (parshiotdiaspora)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][2][0]);
				}
		}
		else if ( strstr(Trim(caldate), &heb5[6][0]) || 
		          strstr(Trim(caldate), &heb5[7][0]) ||
		          strstr(Trim(caldate), &heb5[8][0]) ||
		          strstr(Trim(caldate), &heb5[9][0]) ||
		          strstr(Trim(caldate), &heb5[10][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][3][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[11][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][4][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[12][0]) )
		{
				if (parshiotdiaspora)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][5][0]);
				}
		}
		else if ( strstr(Trim(caldate), &heb5[13][0]) || 
		          strstr(Trim(caldate), &heb5[14][0]) ||
		          strstr(Trim(caldate), &heb5[15][0]) ||
		          strstr(Trim(caldate), &heb5[16][0]) ||
		          strstr(Trim(caldate), &heb5[17][0]) ||
		          strstr(Trim(caldate), &heb5[18][0]) ||
		          strstr(Trim(caldate), &heb5[19][0]) ||
		          strstr(Trim(caldate), &heb5[20][0]) ||
		          strstr(Trim(caldate), &heb5[21][0]) ) //Chanukah
		{
				if ((yeartype == 2 || yeartype == 3) && strstr( caldate, &heb5[21][0]) )
				{
					goto remU;
				}
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][6][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[22][0]) ||
		          strstr(Trim(caldate), &heb5[23][0]) ) 
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][7][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[24][0]) ||
		          strstr(Trim(caldate), &heb5[25][0]) ) 
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][8][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[26][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][9][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[27][0]) )
		{
				if (parshiotEY)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][10][0]);
				}
				else if (parshiotdiaspora)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][9][0]);
				}
		}
		else if ( strstr(Trim(caldate), &heb5[28][0]) || 
		          strstr(Trim(caldate), &heb5[29][0]) ||
		          strstr(Trim(caldate), &heb5[30][0]) ||
		          strstr(Trim(caldate), &heb5[31][0]) ) 
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][10][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[32][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][9][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[33][0]) )
		{
				if (parshiotdiaspora)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][9][0]);
				}
		}
		else if ( strstr(Trim(caldate), &heb5[34][0]) )
		{
                sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][11][0]);
		}
		else if ( strstr(Trim(caldate), &heb5[35][0]) )
		{
				if (parshiotdiaspora)
				{
                    sprintf( calday,"%s%s",&heb5[36][0], &holidays[0][11][0]);
				}
		}
		else
		{
		}
	}
	else
	{
		if ( strstr(Trim(caldate), "1-Tishrey") || 
             strstr(Trim(caldate), "2-Tishrey")) //Rosh Hoshono
		{
				sprintf( l_buff, "%s%s%s", &holidays[1][0][0], ",_", calday );
				strcpy( calday, l_buff );

		}
		else if ( strstr( Trim(caldate), "10-Tishrey") ) //Yom Hakipurim
		{
				sprintf( l_buff, "%s%s%s", &holidays[1][1][0], ",_", calday );
				strcpy( calday, l_buff );
		}
		else if ( strstr( Trim(caldate), "15-Tishrey") ) //Succos
		{
				sprintf( l_buff, "%s%s%s", &holidays[1][2][0], ",_", calday );
				strcpy( calday, l_buff );
		}
		else if ( strstr( Trim(caldate), "16-Tishrey") ) //Succos
		{
				if (parshiotEY) //Chol Hamoed Succos
				{
  	         		sprintf( l_buff, "%s%s%s", &holidays[1][3][0], ",_", calday );
			    	strcpy( calday, l_buff );
				}
				else if (parshiotdiaspora) //Succos
				{
  	         		sprintf( l_buff, "%s%s%s", &holidays[1][2][0], ",_", calday );
			    	strcpy( calday, l_buff );
				}
		}
		else if ( strstr(Trim(caldate), "17-Tishrey") || 
                  strstr(Trim(caldate), "18-Tishrey") ||
                  strstr(Trim(caldate), "19-Tishrey") ||
                  strstr(Trim(caldate), "20-Tishrey") ||
                  strstr(Trim(caldate), "21-Tishrey") ) //Chol Hamoed Succos
		{
         		sprintf( l_buff, "%s%s%s", &holidays[1][3][0], ",_", calday );
		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "22-Tishrey") ) //Shmini Azeres
		{
         		sprintf( l_buff, "%s%s%s", &holidays[1][4][0], ",_", calday );
		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "23-Tishrey") ) //Simchas Torah
		{
				if (parshiotdiaspora)
				{
           			sprintf( l_buff, "%s%s%s", &holidays[1][5][0], ",_", calday );
 		    		strcpy( calday, l_buff );
				}
		}
		else if ( strstr(Trim(caldate), "25-Kislev") || 
                  strstr(Trim(caldate), "26-Kislev") ||
                  strstr(Trim(caldate), "27-Kislev") ||
                  strstr(Trim(caldate), "28-Kislev") ||
                  strstr(Trim(caldate), "29-Kislev") ||
                  strstr(Trim(caldate), "30-Kislev") ||
                  strstr(Trim(caldate), "1-Teves") ||
                  strstr(Trim(caldate), "2-Teves") ||
                  strstr(Trim(caldate), "3-Teves") )
		{
				if ((yeartype == 2 || yeartype == 3) && strstr( Trim(caldate),"3-Teves") )
				{
					goto remU;
				}
        		sprintf( l_buff, "%s%s%s", &holidays[1][6][0], ",_", calday );
 		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "14-Adar") || 
                  strstr(Trim(caldate), "14-Adar II") ) //Purim
		{
        		sprintf( l_buff, "%s%s%s", &holidays[1][7][0], ",_", calday );
 		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "15-Adar") || 
                  strstr(Trim(caldate), "15-Adar II") ) //Shushan Purim
		{
        		sprintf( l_buff, "%s%s%s", &holidays[1][8][0], ",_", calday );
 		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "15-Nisan")) //Pesach
		{
        		sprintf( l_buff, "%s%s%s", &holidays[1][9][0], ",_", calday );
 		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "16-Nisan")) //Second day Pesach for diaspora
		{
				if (parshiotEY)
				{
		       		sprintf( l_buff, "%s%s%s", &holidays[1][10][0], ",_", calday );
 			    	strcpy( calday, l_buff );
				}
				else if (parshiotdiaspora)
				{
		       		sprintf( l_buff, "%s%s%s", &holidays[1][9][0], ",_", calday );
 			    	strcpy( calday, l_buff );
				}
		}
		else if ( strstr(Trim(caldate), "17-Nisan") || 
                  strstr(Trim(caldate), "18-Nisan") ||
                  strstr(Trim(caldate), "19-Nisan") ||
                  strstr(Trim(caldate), "20-Nisan") )
		{
	       		sprintf( l_buff, "%s%s%s", &holidays[1][10][0], ",_", calday );
		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "21-Nisan")) //last day of Pesach
		{
	       		sprintf( l_buff, "%s%s%s", &holidays[1][9][0], ",_", calday );
		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "22-Nisan")) //Second day last day of Pesach for diaspora
		{
				if (parshiotdiaspora)
				{
		       		sprintf( l_buff, "%s%s%s", &holidays[1][9][0], ",_", calday );
			    	strcpy( calday, l_buff );
				}
		}
		else if ( strstr(Trim(caldate), "6-Sivan") ) //Shavuos
		{
	       		sprintf( l_buff, "%s%s%s", &holidays[1][11][0], ",_", calday );
		    	strcpy( calday, l_buff );
		}
		else if ( strstr(Trim(caldate), "7-Sivan") ) //Second day of Shavuos for diaspora
		{
				if (parshiotdiaspora)
				{
		       		sprintf( l_buff, "%s%s%s", &holidays[1][11][0], ",_", calday );
			    	strcpy( calday, l_buff );
				}
		}
		else //don't change anything
		{
		}
	}

	//========================================================
	//remove "_" from calday$
	remU:
	if (RemoveUnderline)
	{
		for (ir = 1; ir <= strlen(calday); ir++)
		{
			if (strstr( &calday[ir - 1], "_") )
			{
				strncpy( &calday[ir -1], " ", 1);
			}
		}
	}  
	calday[strlen(calday)] = 0; //terminate calday string
	
}


//convert hebrew year in numbers to hebrew characters
// inpurt: hebrew year, yr, in numbers, e.g., 5767
// outpout: pointer to Hebrew string containing the Hebrew Year
char *hebyear(short yr)
{ 
	short hund = 0;
	short tens = 0;
	short ones = 0;
	char yrc[5] = "";
	char yrt[2] = "";
	char yron[2] = "";
    hund = (int)Fix((yr - 5000) * 0.01);
    tens = (int)Fix((yr - 5000 - hund * 100) * 0.1);
    ones = (int)Fix(yr - 5000 - hund * 100 - tens * 10);

   if ( hund == 0 )
   {
      //yrc = "";
   }
   else if ( hund == 1 )
   {
      yrc[0] = (char)(247); // "ק"
   }
   else if ( hund == 2 )
   {
      yrc[0] = (char)(248); //"ר"
   }
   else if ( hund == 3 )
   {
      yrc[0] = (char)(249); //"ש"
   }
   else if ( hund == 4 )
   {
      yrc[0] = (char)(250); //"ת"
   }
   else if ( hund == 5 )
   {
      //"תק"
      yrc[0] = (char)(248);
	  yrc[1] = (char)(247);
	  
   }
   else if ( hund == 6 )
   {
      //"תר"
      yrc[0] = (char)(250);
	  yrc[1] = (char)(248);
   }
   else if ( hund == 7 )
   {
      //"תש"
      yrc[0] = (char)(250);
	  yrc[1] = (char)(249);
   }
   else if ( hund == 8 )
   {
      //"תת"
      yrc[0] = (char)(250);
	  yrc[1] = (char)(250);
   }
   else if ( hund == 9 )
   {
      //"תתק"
      yrc[0] = (char)(250);
	  yrc[1] = (char)(250);
	  yrc[2] = (char)(247);
   }
   else if ( hund == 10 )
   {
      // "תת" + (char *)(34) + "ר"
      yrc[0] = (char)(250);
	  yrc[1] = (char)(250);
	  yrc[2] = (char)(34);
	  yrc[3] = (char)(248);
   }

   if ( tens == 1 )
   {
      yrt[0] = (char)(233); //"י"
   }
   else if ( tens == 2 )
   {
      yrt[0] = (char)(235); //"כ"
   }
   else if  ( tens == 3 )
   {
      yrt[0] = (char)(236); //"ל"
   }
   else if  ( tens == 4 )
   {
      yrt[0] = (char)(238); //"מ"
   }
   else if  ( tens == 5 )
   {
      yrt[0] = (char)(240); //"נ"
   }
   else if  ( tens == 6 )
   {
      yrt[0] = (char)(241); //"ס"
   }
   else if  ( tens == 7 )
   {
      yrt[0] = (char)(242); //"ע"
   }
   else if  ( tens == 8 )
   {
      yrt[0] = (char)(244); //"פ"
   }
   else if  ( tens == 9 )
   {
      yrt[0] = (char)(246); //"צ"
   }

   if ( ones != 0 )
   {
      yron[0] = (char)(ones + 223);
      sprintf( yrcal,"%s%s%s%s%s",&yrc[0],&yrt[0],"\"",&yron[0],"\0");
      return &yrcal[0];
   }
   else
   {
      sprintf( yrcal, "%s%s%s%s",&yrc[0],"\"",&yrt[0],"\0");
	  return &yrcal[0];
   }

}


//////////////////////////////////WriteTables////////////////////////////////
void WriteTables(FILE *stream2)
///////////////////////////////////////////////////////////////////////
//this routine writes the zemanim html file
//the console or internet modes
//filroot$ = output file name with directory
//root$ = output file name without directory
//ext$ = the file extension either: "htm" (not supported now: "zip", "csv", or "xml")

/////////////////////////////////////////////////////////////////////////
{
	short errlog = 0;
	short zmantype = 0;
	char delfile[255] = "";
	short lognum = 0;
	short nsloop = 0;
	short yrn = 0;
	short j = 0;
	short k = 0;
	short i = 0;
	short numday = 0;
	short linnum = 0;
	char totdoc[255] = "";
	short xslchild = 0;
	char ch[] = "123";
	char ch1[] = "123";
	char ch2[] = "123";
	char ch3[] = "123";
	char outdoc[255] = "";
	short m = 0;
	short nn = 0;
	bool reorder = false;
	short newnum = 0;
	short totnum = 0;
	short numsort = 0;
	short typezman = 0;
	char hdr[255] = "";
	char address[100] = "";
	char l_progvernum[4] = "";
	char l_datavernum[3] = "";
	char l_yrheb[5] = "";
	short yrheb = 0;
	short tmpn = 0;
	short xslnum = 0;
	short tmpnum = 0;
	short tmpfilxml = 0;
	char destinhtm[100] = "";
	char sourcehtm[100] = "";
	bool closerow = false;
	char l_buff[100] = "";
	char *yearheb = NULL;
	char *calday = "";

	//write html,xml version of sorted zmanim table

	//write html header

	fprintf(stream2,"%s\n", "<!d HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");

	if (! optionheb)
	{
		fprintf(stream2, "%s\n", "<HTML>");
	}
	else
	{
		sprintf( l_buff, "%s%s%s%s%s", "<HTML dir = ", "\"", "rtl", "\"", ">");
		fprintf(stream2, "%s\n", &l_buff[0] );
	}


	fprintf(stream2, "%s\n", "<head>");
	fprintf(stream2, "%s\n", "    <title>Your Chai Zemanim Table</title>");
	fprintf(stream2, "%s\n", "    <meta http-equiv=\"content-type\" content=\"text/html; charset=windows-1255\"/>");
	fprintf(stream2, "%s\n", "    <meta name=\"author\" content=\"Chaim Keller, Chai Tables, www.zemanim.org\"/>");
	fprintf(stream2, "%s\n", "</head>");
	fprintf(stream2, "%s\n", "<body>");

	if (optionheb) //convert hebrew year in numbers to hebrew characters
	{
		yearheb = hebyear(yrheb);
	}

    //convert Hebrew Year in English numbers to string
	 itoa(yrheb, l_yrheb, 10);

	if ( !optionheb )
	{
		if ( Trim(eroscity) == "" )
		{
			sprintf( l_buff, "%s%s%s%s%s%s%s", "<h2><center>", TitleLine, " Tables for ", citnamp, " for the year ", l_yrheb, "</center></h2>" );
		}
		else
		{
			sprintf( l_buff, "%s%s%s%s%s%s%s%s%s%s", "<h2><center>", TitleLine, " Tables for ", eroscity, " (", citnamp, ")", " for the year ", l_yrheb, "</center></h2>" );
		}
	}
	else
	{
		if (Trim(&eroscity[0]) == "")
		{
			sprintf( l_buff, "%s%s%s%s%s%s", "<h2><center>", &heb3[1][0], hebcityname, &heb3[2][0], yearheb, "</h2>");
		}
		else
		{
			sprintf( l_buff, "%s%s%s%s%s%s%s%s%s", "<h2><center>", " (", eroscity, ") ", &heb3[12][0], hebcityname, &heb3[2][0], yearheb, "</center></h2>");
		}
	}
	fprintf(stream2, "%s\n", &l_buff[0]);


	itoa( datavernum, l_datavernum, 10);
	itoa( progvernum, l_progvernum, 10);
	sprintf( address, "%s%s%s%s%s", &heb2[14][0], l_datavernum, ".", l_progvernum, " ©" );
	if (! optionheb )
	{
        sprintf( address, "%s%s%s%s%s", "© Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: ", l_datavernum, ".", l_progvernum );
	}


	//now write name of table

	if (typezman == 0)
	{
		if (optionheb )
		{
			strcpy( hdr, &heb3[3][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to the opinion of the Mogen Avrohom using a 16.1 degrees dawn and twilight (approx. 72 min.) ");
		}
	}
	else if (typezman == 1)
	{
		if (optionheb )
		{
			strcpy( hdr, &heb3[4][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to the opinion of the Mogen Avrohom using a 19.75 degrees dawn and twilight (approx. 90 min.) ");
		}
	}
	else if (typezman == 2)
	{
		if (optionheb)
		{
			sprintf( hdr, "%s%s%s%s", &heb3[5][0], "\"", heb3[6][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to the opinion of the Groh using the mishor sunrise and sunset");
		}
	}
	else if (typezman == 3)
	{
		if (optionheb)
		{
			sprintf( hdr, "%s%s%s%s", &heb3[5][0], "\"", heb3[7][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to the opinion of the Groh using the astronomical sunrise and sunset");
		}
	}
	else if (typezman == 4)
	{
		if (optionheb)
		{
			strcpy( hdr, &heb3[8][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to opinion of the Ben Ish Chai");
		}
	}
	else if (typezman == 5)
	{
		if (optionheb)
		{
			strcpy( hdr, &heb3[14][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to opinion of the Baal Hatanya");
		}
	}
	else if (typezman == 6)
	{
		if (optionheb)
		{
			strcpy( hdr, &heb3[15][0]);
		}
		else
		{
			sprintf( hdr, "%s", "Z'manim calculated acc. to opinion of Harav Zalman Baruch Melamid");
		}
	}
	else if (typezman == 7) //Yedidia's times
	{
		if (optionheb)
		{
			strcpy( hdr, &heb3[13][0] );
		}
		else
		{
			sprintf( hdr, "%s", "Astronomical Sunset");
		}
	}

	sprintf(l_buff, "%s%s%s", "<h3><center>", hdr, "</center></h3>");
	fprintf(stream2, "%s\n", l_buff);
	sprintf(l_buff, "%s%s%s", "<h3><center>", address, "</center></h3>"); 
	fprintf(stream2, "%s\n", l_buff);

	if (Trim(&SponsorLine[0][0]) != "" && optionheb )
	{
        sprintf(l_buff, "%s%s%s", "<h5><center>", &SponsorLine[0][0], "</center></h5>" );
		fprintf(stream2, l_buff);
	}
	else if (Trim(&SponsorLine[1][0]) != "" && !optionheb )
	{
        sprintf(l_buff, "%s%s%s", "<h5><center>", &SponsorLine[1][0], "</center></h5>" );
		fprintf(stream2, l_buff);
	}

	//generate table column legend
	if (reorder)
	{
		totnum = numsort;
	}
	else
	{
		totnum = newnum;
	}

	nn = -1;
	for (m = totnum; m >= 0; m--)
	{

		//attach letters or numbers to labels
		if (optionheb)
		{
			hebnum(m + 4, ch);
			strcpy( l_buff, outdoc );
			sprintf( outdoc, "%s%s%s%s%s", ch, " = ", &zmannames[m][0], " | ", l_buff );
		}
		else
		{
			nn++;
			itoa(nn + 4, ch, 10);
			strcpy( l_buff, outdoc );
			sprintf( outdoc, "%s%s%s%s%s", l_buff, " | ", ch, " = ", &zmannames[nn][0] );
		}
	}

	//now add the the dates, and days legends
	if (optionheb)
	{
		hebnum(1, ch1);
		hebnum(2, ch2);
		hebnum(3, ch3);
		strcpy( l_buff, outdoc );
		sprintf( outdoc, "%s%s%s%s%s%s%s%s%s%s%s%s%s%s", " | ", ch1, " = ", &heb3[11][0], " | ", 
			     ch2, " = ", heb3[10][0], " | ", ch3, " = ", &heb3[9][0], " | ", l_buff );
	}
	else
	{
		strcpy( l_buff, outdoc );
		sprintf( outdoc, "%s%s","| 1 = hebrew date | 2 = day | 3 = civil date ", l_buff );
	}

	fprintf(stream2, "%s\n", "<hr/>"); //spacers between the headers and column headers
	fprintf(stream2, "%s\n", "<br/>");


	fprintf(stream2, "%s\n", outdoc); //print the legend for the csv, htm, and zip files

	fprintf(stream2, "%s\n", "<br/>"); //spacers between columnheaders and time
	fprintf(stream2, "%s\n", "<br/>");
	fprintf(stream2, "%s\n", "<TABLE BORDER='1' CELLPADDING='1' CELLSPACING='1' ALIGN='CENTER' >");


	//create table's column headers
	linnum = 1; //column headers are first row in the htm document

	strcpy ( outdoc, "\0");

	//add new line tag for this first line
	fprintf(stream2, "%s\n", "    <TR>");

	nn = 0;
	for (m = 1; m <= totnum + 4; m++)
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
		
		strcpy( outdoc, ch);

		if (m == totnum + 4)
		{
			closerow = true;
		}
        parseit( outdoc, stream2, closerow, linnum);
		//outdoc$ = outdoc$ & comma$ & spacerb$ & ch$ & spacera$
	}

	//now create the actual table
	numday = -1;
	for (i = 1; i <= endyr; i++)
	{
		if (mmdate[1][i] > mmdate[0][i])
		{
			k = 0;
			strcpy( totdoc, "\0");

			for (j = mmdate[0][i - 1]; j <= mmdate[1][i - 1]; j++)
			{
				numday++;
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf(stream2, "%s\n", "    <TR>");

				strcpy( outdoc, Trim(&stortim[3][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				InsertHolidays(calday, i, k);
				strcpy( outdoc, calday);
                parseit( outdoc, stream2, closerow, linnum);
				strcpy ( outdoc, Trim(&stortim[2][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				for (m = 0; m <= totnum; m++)
				{
					if ( strstr( Mid(&zmantimes[m - 1][numday - 1][0], 0, 2, l_buff), "00") )
					{
						strcpy( &zmantimes[m - 1][numday - 1][0], "------");
					}
					strcpy( outdoc, Trim(&zmantimes[m - 1][numday - 1][0]) );
					if (m == totnum)
					{
						closerow = true;
					}
                    parseit( outdoc, stream2, closerow, linnum);
				}

			}
		}
		else if (mmdate[1][i - 1] < mmdate[0][i - 1])
		{
			k = 0;
			strcpy( totdoc, "\0" );

			for (j = mmdate[0][i - 1]; j <= yrend[0]; j++) //yl1%
			{
				numday++;
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf(stream2, "%s\n", "    <TR>");

				strcpy( outdoc, Trim(&stortim[3][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				InsertHolidays(calday, i, k);
				strcpy( outdoc, calday);
                parseit( outdoc, stream2, closerow, linnum);
				strcpy ( outdoc, Trim(&stortim[2][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				for (m = 0; m <= totnum; m++)
				{
					if ( strstr( Mid(&zmantimes[m - 1][numday - 1][0], 0, 2, l_buff), "00") )
					{
						strcpy( &zmantimes[m - 1][numday - 1][0], "------");
					}
					strcpy( outdoc, Trim(&zmantimes[m - 1][numday - 1][0]) );
					if (m == totnum)
					{
						closerow = true;
					}
                    parseit( outdoc, stream2, closerow, linnum);
				}
			}

			yrn++;

			for (j = 1; j <= mmdate[1][i - 1]; j++)
			{
				k++;

				//increment line number
				linnum++;

				//add new line tag for each new line
				fprintf(stream2, "%s\n", "    <TR>");

				numday++;
				strcpy( outdoc, Trim(&stortim[3][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				InsertHolidays(calday, i, k);
				strcpy( outdoc, calday);
                parseit( outdoc, stream2, closerow, linnum);
				strcpy ( outdoc, Trim(&stortim[2][i - 1][k - 1][0]) );
                parseit( outdoc, stream2, closerow, linnum);
				for (m = 0; m <= totnum; m++)
				{
					if ( strstr( Mid(&zmantimes[m - 1][numday - 1][0], 0, 2, l_buff), "00") )
					{
						strcpy( &zmantimes[m - 1][numday - 1][0], "------");
					}
					strcpy( outdoc, Trim(&zmantimes[m - 1][numday - 1][0]) );
					if (m == totnum)
					{
						closerow = true;
					}
                    parseit( outdoc, stream2, closerow, linnum);
				}
			}
		}
	}
	fprintf(stream2, "%s\n", "          </TABLE>");
	fprintf(stream2, "%s\n", "<P><BR/><BR/>");
	fprintf(stream2, "%s\n", "</P>");
	fprintf(stream2, "%s\n", "</BODY>");
	fprintf(stream2, "%s\n", "</HTML>");

	strcpy( eroscity, "\0");


}

////////////////parseit//////////////////////////////////////////////
void parseit( char outdoc[], FILE *stream2, bool closerow, short linnum)
{
	//parseit: adds tags to html, and accumlates text for the csv file

	char widthtim[4] = "";
	char timlet[255] = "";
	char l_buff[255] = "";

	strcpy( timlet, Trim(&outdoc[0]) );

	//parse lines looking for spaces
	itoa((int)(strlen(timlet) * 4.5), widthtim, 10);

	if (linnum % 2 == 0)
	{

		sprintf( l_buff, "%s%s%s", "        <TD WIDTH=\"", widthtim, "\" BGCOLOR=\"#ffff00\" >" );
		fprintf(stream2, "%s\n", l_buff );
	}
	else
	{
		sprintf( l_buff, "%s%s%s", "        <TD WIDTH=\"", widthtim, "\" >" );
		fprintf(stream2, "%s\n", l_buff );
	}

	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"3\" >", Trim(timlet), "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff);
	fprintf(stream2, "%s\n", "        </TD>");


	if (closerow) //end of row signaled, so print end of row tags
	{
		fprintf(stream2, "%s\n", "    </TR>");
		closerow = false;
	}


}


void heb12monthHTML(FILE *stream2, short setflag)
{
	short positnear = 0;
	short positunder = 0;
	char ch[4] = "";
	short j = 0;
	short i = 0;
	short iii = 0;
	short iheb = 0;
	short ii = 0;
	bool alsosunset = false;
	char l_buff[255] = "";

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
	
	/*
	fprintf(stream2, "%s\n", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
	fprintf(stream2, "%s\n", "<HTML>");
	fprintf(stream2, "%s\n", "<HEAD>");
	fprintf(stream2, "%s\n", "    <TITLE>Your Chai Table</TITLE>");
	fprintf(stream2, "%s\n", "    <META HTTP-EQUIV=\"CONTENT-TYPE\" CONTENT=\"text/html; charset=windows-1255\"/>");
	fprintf(stream2, "%s\n", "    <META NAME=\"AUTHOR\" CONTENT=\"Chaim Keller, Chai Tables, www.zemanim.org\"/>");
	fprintf(stream2, "%s\n", "</HEAD>");
	fprintf(stream2, "%s\n", "<BODY>");
	*/
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"670\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\" >");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"670\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"670\" VALIGN=\"TOP\">");
	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"4\" STYLE=\"font-size: 16pt\">", &storheader[ii][0][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"670\" VALIGN=\"TOP\">");
	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\">", &storheader[ii][1][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"670\" VALIGN=\"TOP\">");
	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\">", &storheader[ii][3][0], "</FONT></P>" );// &storheader[ii][3][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"670\" BORDER=\"1\" CELLPADDING=\"4\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"16\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"46\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"41\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"15\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"16\" VALIGN=\"TOP\">");
	fprintf(stream2, "%s\n", "            <P ALIGN=\"LEFT\"><BR/>");
	fprintf(stream2, "%s\n", "            </P>");
	fprintf(stream2, "%s\n", "        </TD>");

	if (optionheb)
	{
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][12][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"46\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][11][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][10][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][9][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][8][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][7][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][13][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][4][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][3][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][2][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][1][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][0][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"15\" VALIGN=\"BOTTOM\">");
		fprintf(stream2, "%s\n", "            <P ALIGN=\"CENTER\"><BR/>");
		fprintf(stream2, "%s\n", "            </P>");
		fprintf(stream2, "%s\n", "        </TD>");
	}
	else
	{
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][0][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"46\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][1][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][2][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][3][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][4][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][13][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][7][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][8][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][9][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][10][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][11][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" VALIGN=\"TOP\">");
    	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][12][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"15\" VALIGN=\"BOTTOM\">");
		fprintf(stream2, "%s\n", "            <P ALIGN=\"CENTER\"><BR/>");
		fprintf(stream2, "%s\n", "            </P>");
		fprintf(stream2, "%s\n", "        </TD>");
	}
	fprintf(stream2, "%s\n", "    </TR>");


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

			fprintf(stream2, "%s\n", "    <TR VALIGN=\"BOTTOM\">");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"16\" BGCOLOR=\"#ffff00\">");
        	l_buff[0] = 0; //clear buffer
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
			//now determine if this time is underlined or shabbos and add appropriate HTML ******
            timestring( ii, 0, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"46\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 1, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 2, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 3, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 4, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 5, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 6, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 7, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 8, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 9, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 10, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 11, iii, stream2 );

			fprintf(stream2, "%s\n", "        <TD WIDTH=\"16\" BGCOLOR=\"#ffff00\">");
        	l_buff[0] = 0; //clear buffer
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "    </TR>");
		}
		for (j = 1; j <= 3; j++)
		{
			iii = iii + 1;
			//Call hebnum(iii% + 1, ch$)
			//UPGRADE_WARNING: Couldn't resolve default property of object optionheb. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			if (optionheb == true)
			{
				hebnum(iii + 1, ch);
			}
			else
			{
				itoa( iii + 1, ch, 10);
			}
			fprintf(stream2, "%s\n", "    <TR VALIGN=\"BOTTOM\">");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"16\">");
        	l_buff[0] = 0; //clear buffer
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" >");
			//now determine if this time is underlined or shabbos and add appropriate HTML ******
            timestring( ii, 0, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"46\">");
            timestring( ii, 1, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 2, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 3, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 4, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 5, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
            timestring( ii, 6, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 7, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
            timestring( ii, 8, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
            timestring( ii, 9, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\">");
            timestring( ii, 10, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 11, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"16\">");
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "    </TR>");
		}
	}
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"670\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"670\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"670\" VALIGN=\"TOP\">");
   	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &storheader[ii][4][0], "</B></FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"670\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"670\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"670\" VALIGN=\"TOP\">");
	if (NearWarning[setflag]) //print near obstructions warning
	{
   		l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &storheader[ii][5][0], "</B></FONT></P>" );
		fprintf(stream2, "%s\n", l_buff );
	}
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<P><BR/><BR/>");
	fprintf(stream2, "%s\n", "</P>");
	/*
	fprintf(stream2, "%s\n", "</BODY>");
	fprintf(stream2, "%s\n", "</HTML>");
	*/

}

void timestring( short ii, short k, short iii, FILE *stream2 )
{

	char ntim[255] = "";
	char l_buff[255] = "";
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
		sprintf( l_buff, "%s%s%s", "<u>", ntim, "</u>" );
		strcpy( ntim, l_buff );
	}

	if (positnear != 0)
	{
		sprintf( l_buff, "%s%s%s", "<font color=\"#33cc66\">", ntim, "</font>" );
		strcpy( ntim, l_buff );
	}
	
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\">", ntim, "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");

}

void heb13monthHTML(FILE *stream2, short setflag)
{
	short positnear = 0;
	short positunder = 0;
	char ntim[10] = "";
	char ch[4] = "";
	short j = 0;
	short i = 0;
	short iii = 0;
	short iheb = 0;
	short ii = 0;
	bool alsosunset = false;
	char l_buff[100] = "";

	if (setflag == 0)
	{
		ii = 0;
	}
	else
	{
		ii = 1;
	}

	iheb = 1;
	if (!optionheb )
	{
		iheb = 0;
	}
	
/*
	fprintf(stream2, "%s\n", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
	fprintf(stream2, "%s\n", "<HTML>");
	fprintf(stream2, "%s\n", "<HEAD>");
	fprintf(stream2, "%s\n", "    <TITLE>Your Chai Table</TITLE>");
	fprintf(stream2, "%s\n", "    <META HTTP-EQUIV=\"CONTENT-TYPE\" CONTENT=\"text/html; charset=windows-1255\"/>");
	fprintf(stream2, "%s\n", "    <META NAME=\"AUTHOR\" CONTENT=\"Chaim Keller, Chai Tables, www.zemanim.org\"/>");
	fprintf(stream2, "%s\n", "</HEAD>");
	fprintf(stream2, "%s\n", "<BODY>");
*/

	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"696\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\" >");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"696\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"696\" VALIGN=\"TOP\" BGCOLOR=\"#ffffff\">");
   	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"4\" STYLE=\"font-size: 16pt\">", &storheader[ii][0][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"696\" VALIGN=\"TOP\">");
   	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"4\" STYLE=\"font-size: 8pt\">", &storheader[ii][1][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"696\" VALIGN=\"TOP\">");
   	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"4\" STYLE=\"font-size: 8pt\">", &storheader[ii][3][0], "</FONT></P>" ); //&storheader[ii][3][0], "</FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"696\" BORDER=\"1\" CELLPADDING=\"4\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"14\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"51\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"40\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"43\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"41\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"44\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"44\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"45\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"42\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"39\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"15\">");
	fprintf(stream2, "%s\n", "    <TR VALIGN=\"TOP\">");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"14\">");
	fprintf(stream2, "%s\n", "            <P ALIGN=\"CENTER\"><BR/>");
	fprintf(stream2, "%s\n", "            </P>");
	fprintf(stream2, "%s\n", "        </TD>");

	if (optionheb)
	{
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][12][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"51\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][11][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][10][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][9][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][8][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][7][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"43\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][6][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][5][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][4][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][3][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"45\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][2][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"42\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][1][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][0][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"15\">");
		fprintf(stream2, "%s\n", "            <P ALIGN=\"CENTER\"><BR/>");
		fprintf(stream2, "%s\n", "            </P>");
		fprintf(stream2, "%s\n", "        </TD>");
	}
	else
	{
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][0][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"51\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][1][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][2][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][3][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][4][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][5][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"43\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][6][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][7][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][8][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][9][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"45\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][10][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"42\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][11][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
	   	l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT FACE=\"Arial, sans-serif\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &monthh[iheb][12][0], "</B></FONT></P>");
		fprintf(stream2, "%s\n", l_buff );
		fprintf(stream2, "%s\n", "        </TD>");
		fprintf(stream2, "%s\n", "        <TD WIDTH=\"15\">");
		fprintf(stream2, "%s\n", "            <P ALIGN=\"CENTER\"><BR/>");
		fprintf(stream2, "%s\n", "            </P>");
		fprintf(stream2, "%s\n", "        </TD>");
	}
	fprintf(stream2, "%s\n", "    </TR>");


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

			fprintf(stream2, "%s\n", "    <TR VALIGN=\"BOTTOM\">");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"14\" BGCOLOR=\"#ffff00\">");
		   	l_buff[0] = 0; //clear buffer
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
			//now determine if this time is underlined or shabbos and add appropriate HTML ******
            timestring( ii, 0, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"51\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 1, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 2, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 3, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 4, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 5, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"43\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 6, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 7, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 8, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 9, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"45\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 10, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"42\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 11, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\" BGCOLOR=\"#ffff00\">");
            timestring( ii, 12, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"14\" BGCOLOR=\"#ffff00\">");
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "    </TR>");
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

			fprintf(stream2, "%s\n", "    <TR VALIGN=\"BOTTOM\">");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"14\">");
		   	l_buff[0] = 0; //clear buffer
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
			//now determine if this time is underlined or shabbos and add appropriate HTML ******
            timestring( ii, 0, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"51\">");
            timestring( ii, 1, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 2, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 3, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
            timestring( ii, 4, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"40\">");
            timestring( ii, 5, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"43\">");
            timestring( ii, 6, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"41\">");
            timestring( ii, 7, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
            timestring( ii, 8, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"44\">");
            timestring( ii, 9, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"45\">");
            timestring( ii, 10, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"42\">");
            timestring( ii, 11, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"39\">");
            timestring( ii, 12, iii, stream2 );
			fprintf(stream2, "%s\n", "        <TD WIDTH=\"14\">");
			sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 10pt\"><B>", ch, "</B></FONT></P>" );
			fprintf(stream2, "%s\n", l_buff );
			fprintf(stream2, "%s\n", "        </TD>");
			fprintf(stream2, "%s\n", "    </TR>");
		}
	}
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"696\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"696\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"696\" VALIGN=\"TOP\">");
   	l_buff[0] = 0; //clear buffer
	sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &storheader[iheb][4][0], "</B></FONT></P>" );
	fprintf(stream2, "%s\n", l_buff );
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<TABLE WIDTH=\"696\" BORDER=\"0\" CELLPADDING=\"0\" CELLSPACING=\"0\">");
	fprintf(stream2, "%s\n", "    <COL WIDTH=\"696\">");
	fprintf(stream2, "%s\n", "    <TR>");
	fprintf(stream2, "%s\n", "        <TD WIDTH=\"696\" VALIGN=\"TOP\">");
	if (NearWarning[setflag]) //print near obstruction warning
	{
   		l_buff[0] = 0; //clear buffer
		sprintf( l_buff, "%s%s%s", "            <P ALIGN=\"CENTER\"><FONT SIZE=\"1\" STYLE=\"font-size: 8pt\"><B>", &storheader[iheb][4][0], "</B></FONT></P>" );
		fprintf(stream2, "%s\n", l_buff );
	}
	fprintf(stream2, "%s\n", "        </TD>");
	fprintf(stream2, "%s\n", "    </TR>");
	fprintf(stream2, "%s\n", "</TABLE>");
	fprintf(stream2, "%s\n", "<P><BR/><BR/>");
	fprintf(stream2, "%s\n", "</P>");

	/*
	fprintf(stream2, "%s\n", "</BODY>");
	fprintf(stream2, "%s\n", "</HTML>");
	*/

}

 ////////////////////////////////////////////////////////////////////////////////////////////
short WriteUserLog( int UserNumber, short zmanyes, short typezman )
///////////////////////////////////////////////////////////////////////////////////////////
//WriteUserLog writes the user's inputed parameters to the userlog.log file
///////////////////////////////////////////////////////////////////////////////////////////
{

	  //write userlog info
	  char l_buff[100] = "";
      char filroot[100] = "";
      char userloglin[100] = "";
      FILE *stream;

	  sprintf( filroot, "%s%s", drivweb, "userlog.log");

  	  /* write closing clauses on html document */
	  if ( !( stream = fopen( filroot, "a")) ) 	   
	  {
	    ErrorHandler(4, 1); //can't open log file
		return -1;
	  }

	  sprintf( userloglin, "%s%s%s%s%s%s", eroscountry, ",", erosareabat, ",", eroscity, "\0");
	  strcpy( l_buff, userloglin );
	  sprintf( userloglin, "%ld%s%s%s%f%s%f%s%g%s%d", UserNumber, ",", l_buff, ",", 
		                    eroslatitude, ",", eroslongitude, ",", eroshgt, ",", yrheb  );
	  switch (types)
	  {
	    case 0:
		  strcat( userloglin, ",visible sunrise" );
		  break;
	    case 1:
		  strcat( userloglin, ",mishor sunrise" );
		  break;
	    case 2:
		  strcat( userloglin, ",astron. sunrise" );
		  break;
	    case 3:
		  strcat( userloglin, ",mishor sunset" );
		  break;
	    case 4:
		  strcat( userloglin, ",astron. sunset" );
		  break;
	    case 5:
		  strcat( userloglin, ",visible sunset" );
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
	    default:
		  {
		  ErrorHandler(4, 2); //error, unrecognizable types
		  return -1;
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
		  strcat( userloglin, itoa(typezman, l_buff, 10) );
	  }

	  if (optionheb)
	  {
		  strcat( userloglin, ",Hebrew" );
	  }
	  else
	  {
		  strcat( userloglin, ",English" );
	  }

	  //strcat( userloglin, "\n");
	  fprintf( stream, "%s\n", userloglin );
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
	switch (mode)
	{
	case 1: //main
		break;
	case 2: //RdHebStrings
		break;
	case 3:
		break;
	case 4:
		break;
	case 5:
		break;
	case 6:
		break;
	case 7:
		break;
	case 8:
		break;
	case 9:
		break;
	case 10:
		break;
	case 11:
		break;
	case 12:
		break;
	default:
		break;
	}
}
////////////////////////////////////////////////////////////////

///////////////////////////-////////////////////////////
short Caldirectories()
///////////////////////////////////////////////////////////////////
//reads in hebcitynames if required
///////////////////////////////////////////
{
   char *sun = "netz";
   char filroot[100] = "";
   char l_buff[100] = "";
   char batfile[100] = "";
   short ier = 0;
   DIR *pdir;
   struct dirent *pent;

	// search for near places for eros city

   switch (types)
   {
   case 0:
   case 2:
   case 4:
   case 6:
   case 7:
	   sun = "netz";

	   // need to calculate visible sunrise
	   // find "bat" file name of "netz" directory
	   if (!eros && !astronplace )
	   {
		   sprintf( filroot, "%s%s%s", currentdir, sun, "/" );
		   
		   if ( !(pdir = opendir( filroot )) )
		   {
			   ErrorHandler(7, 1);
			   return -1;
		   }
		   else
		   {
			  while ((pent=readdir(pdir)))
			  {
				 pent->d_name;
				 if ( strstr( strlwr(pent->d_name), ".bat" ) )
				 {
					 // found bat file, store its name
			         sprintf( batfile, "%s%s%s", filroot, pent->d_name, "\0" );
					 break;
				 }
			  }
		   }
		   closedir(pdir);

		   //open bat file and add places to the vantage point buffer
		   if ( netzskiy( batfile, sun ) ) return -1; //error returns -1

	   }
	   else if (eros && !astronplace) 
	   {
		   //bat file's name always consists of first 4 letters of currentdir
		   sprintf( batfile, "%s%s%s%s%s", currentdir, sun, "/"
		                     , Mid(erosareabat, 1, 4, l_buff), ".bat" );

		   //search for places within search radius and add vantage points to buffer
		   if ( calnearsearch( batfile, sun ) ) return -1; //error returns -1
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
	   sun = "skiy";

	   // need to calculate visible sunset
	   // find "bat" file name of "skiy" directory
	   if (!eros && !astronplace)
	   {
		   sprintf( filroot, "%s%s%s", currentdir, sun, "/" );
		   if ( !(pdir = opendir( filroot )) )
		   {
			   ErrorHandler(7, 1);
			   return -1;
		   }
		   else
		   {
			  while ((pent=readdir(pdir)))
			  {
				 pent->d_name;
				 if ( strstr( strlwr(pent->d_name), ".bat" ) )
				 {
					 // found bat file, store its name
					 sprintf( batfile, "%s%s%s", filroot, pent->d_name, "\0" );
					 break;
				 }
			  }
		   }
		   closedir(pdir);

		   //open bat file and add places to the vantage point buffer
		   if ( netzskiy( batfile, sun ) ) return -1; //error returns -1

	   }
	   else if (eros && !astronplace)
	   {
		   //bat file name always consists of first 4 letters of currentdir
		   sprintf( batfile, "%s%s%s%s%s", currentdir, sun, "/"
		                     , Mid(erosareabat, 1, 4, l_buff ), ".bat" );

		   //search for places within search radius and add vantage points to buffer
		   if (calnearsearch( batfile, sun )) return -1; //error returns -1
	   }
	   break;
   }

    //Hebrew calendar calculations///////////////////////////////////////

    HebCal(yrheb); //determine if it is hebrew leap year
                   //and determine the beginning and ending daynumbers of the Hebrew year
    FirstSecularYr = (yrheb - 5758) + 1997;
    SecondSecularYr = (yrheb - 5758) + 1998;
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
	char doclin[255] = "";
	char buff[255] = "";
	char citnam[255] = "";
	double lon, lat, hgt, dist;
	short pos = 0, pos1 = 0;

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

	if (!( stream = fopen( bat, "r" )) )
	{
		ErrorHandler(7,1);
		return -1;
	}

	ydegkm = 180.0/ (pi * 6371.315); //degrees latitude per kilometer
	xdegkm = 180.0/(pi * 6371.315 * cos(eroslatitude * pi/180.0)); //degrees longituder per kilometer

	//first read in Hebrew city name, and geotz
	fgets( doclin, 255, stream); //read in line of documentation

	//parse out the hebrew name and the time zone
	pos = InStr( doclin, "," );
	strcpy( hebcityname, Mid( doclin, 1, pos - 1, buff) );
	Trim(Replace( hebcityname, '\"', ' ' )); //get rid of VB style quotation marks
	//strcat( hebcityname, "\0" );
	//LTrim( hebcityname );

	if (Linux) //convert DOS Hebrew to ISO Hebrew
	{
		DosToIso( hebcityname, buff );
		strcpy( hebcityname, buff );
	}

	geotz = atof( Mid( doclin, pos + 1, strlen(doclin) - 1 - pos, buff ) );

	//search for vantage points within the search radius

	while( ! feof(stream) )
	{
		//the eros city string on Windows is always 27 characters long
		//this will be different for Linux, so even though that the following
		//will work for Windows, it will not work for Linux
		//fscanf( stream, "\"%27c\",%Lg,%Lg,%Lg", doclin, &lonbat, &latbat, &hgtbat );
		//so it is necessary to parse out the name and coordinates in order to make
		//this operation platform independent
		fgets( doclin, 255, stream );
		pos = InStr( 1, doclin, "," );
		Mid(doclin, pos - 13, 12, buff );
		strcat( buff, "\0"); //terminate the string
		//remove quotation marks if they exits
		Trim(Replace(buff, '"', ' ' ));
		sprintf( citnam, "%s%s%s%s", currentdir, sun, "/", buff );

		pos1 = InStr( pos + 1, doclin, "," );
		lat = atof( Mid(doclin, pos + 1, pos1 - pos - 1, buff ));

		pos = InStr( pos1 + 1, doclin, "," );
		lon = atof( Mid(doclin, pos1 + 1, pos - pos1 - 1, buff ));

		hgt = atof( Mid(doclin, pos + 1, strlen(doclin) - pos, buff ));

		if (InStr( strlwr(doclin), "version" )) //reached end of "bat" file, record version number
		{
			progvernum = lat;
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
           dist = sqrt( pow((lat - eroslongitude) / ydegkm , 2.0) + pow((lon - eroslatitude) / xdegkm, 2 ));
		}

		if ( dist <= searchradius )
		{
            if ( !foundvantage ) foundvantage = true; //flag that found a vantage point

			//add this vantage point to the time calculation place buffer
			if (InStr( sun, "netz" ))
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
			else if (InStr( sun, "skiy" ))
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
		ErrorHandler(7, 2); //print out that no vantage points were found
		                    //and may be necessary to increase search radius
        return -1;
	}
	else //finish averaging the coordinates
	{
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
	}

	return 0;
}

//////////////////////calnearsearch///////////////////////////
short netzskiy( char *bat, char *sun )
/////////////////////////////////////////////////////////////
//enters vantage points contained in Israel cities "bat" files to buffer
///////////////////////////////////////////////////////////////////////////
{

	char doclin[255] = "";
	char buff[255] = "";
	char citnam[255] = "";
	double kmx, kmy, hgt;
	short pos = 0, pos1 = 0;

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

	FILE *stream;

	if (!( stream = fopen( bat, "r" )) )
	{
		ErrorHandler(8, 1);
		return -1;
	}

	geotz = 2;

	//search for vantage points within the search radius

	while( ! feof(stream) )
	{
		//the eros city string on Windows is always 27 characters long
		//this will be different for Linux, so even though that the following
		//will work for Windows, it will not work for Linux
		//fscanf( stream, "\"%27c\",%Lg,%Lg,%Lg", doclin, &lonbat, &latbat, &hgtbat );
		//so it is necessary to parse out the name and coordinates in order to make
		//this operation platform independent
		fgets( doclin, 255, stream );
		pos = InStr( 1, doclin, "," );
		Mid(doclin, pos - 12, 12, buff );
		strcat( buff, "\0"); //terminate the string
		//remove quotation marks if they exits
		Replace(buff, '"', ' ' );
		sprintf( citnam, "%s%s%s%s", currentdir, sun, "/", Trim(buff) );

		pos1 = InStr( pos + 1, doclin, "," );
		kmx = atof( Mid(doclin, pos + 1, pos1 - pos - 1, buff ));

		if (InStr( strlwr(doclin), "version" )) //reached end of "bat" file, record version number
		{
			progvernum = kmx;
			break;
		}
		    
		pos = InStr( pos1 + 1, doclin, "," );
		kmy = atof( Mid(doclin, pos1 + 1, pos - pos1 - 1, buff ));

		hgt = atof( Mid(doclin, pos + 1, strlen(doclin) - pos, buff ));


		//add this vantage point to the time calculation place buffer
		if (InStr( sun, "netz" ))
		{
			
			nplac[0]++;
			strcpy( &fileo[0][nplac[0] - 1][0], citnam ); 
			latbat[0][nplac[0] - 1] = kmx;
			lonbat[0][nplac[0] - 1] = kmy;
			hgtbat[0][nplac[0] - 1] = hgt;
  		    
			avekmx[0] = avekmx[0] + kmx;
			avekmy[0] = avekmy[0] + kmy;
			avehgt[0] = avehgt[0] + hgt;

		}
		else if (InStr( sun, "skiy" ))
		{

			nplac[1]++;
			strcpy( &fileo[1][nplac[1] - 1][0], citnam ); 
			latbat[1][nplac[1] - 1] = kmx;
			lonbat[1][nplac[1] - 1] = kmy;
			hgtbat[1][nplac[1] - 1] = hgt;

			avekmx[1] = avekmx[1] + kmx;
			avekmy[1] = avekmy[1] + kmy;
			avehgt[1] = avehgt[1] + hgt;
		}
		
	}
	fclose( stream );

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
short SunriseSunset( short types, FILE *stream2 )
//////////////////////////////////////////////////////////////
//sets up headers, strings for tables
//////////////////////////////////////////////////////////////
{

	if (CreateHeaders( types )) return -1; //create headers for all the different table types
	                                //return -1 if error detected

	short ier, nsetflag, i,tabletyp;
    char fileon[255] = "";
	double kmxo, kmyo, hgt;

/*       Nsetflag = 0 for astr sunrise/visible sunrise/chazos */
/*       Nsetflag = 1 for astr sunset/visible sunset/chazos */
/*       Nsetflag = 2 for mishor sunrise if hgt = 0/chazos or astro. sunrise/chazos */
/*       Nsetflag = 3 for mishor sunset if hgt = 0/chazos or astro. sunset/chazos */
/*       Nsetflag = 4 for mishor sunrise/vis. sunrise/chazos */
/*       Nsetflag = 5 for mishor sunset/vis. sunset/chazos */

	///////////////////sunrises//////////////////////////////////////////

	for (i = 0; i < nplac[0]; i++ ) //sunrises
	{

		if (!astronplace)
		{
           strcpy( fileon, &fileo[0][i][0] ); 
           kmyo = latbat[0][i];
 		   kmxo = lonbat[0][i];
		   hgt = hgtbat[0][i];
		}
		else
		{
           kmyo = eroslatitude;
		   kmxo = eroslongitude;
		   hgt = eroshgt;
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
		}

		ier = netzski6( fileon, kmyo, kmxo, hgt, geotz, nsetflag, i);

		if (ier) //error detected
		{
			ErrorHandler(10, ier); //-1 = insufficient azimuthal range
			                       //-2 = can't open profile file
			                       //-3 = can't allocate memory for elev array
			return -1;
		}

	}

	//now convert calculated sunrise time in fractional hours to time string and
	//print out sunrise table if only single table of sunrise was requested
	switch (types)
	{
	case 0:
	case 4:
	  tabletyp = 0; //visible sunrise index of tims
	  PrintSingleTable( 0, tabletyp, stream2 ); //converts "tims" array in fractional hours to
	                      //time strings and prints out table to output
	  return 0;
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
	 
	  PrintSingleTable( 0, tabletyp, stream2 ); //converts "tims" array in fractional hours to
                    //time strings and prints out table to output
      return 0;
	  break;
	}

	////////////////////////////sunsets////////////////////////////////////////////////////////////////


	for (i = 0; i < nplac[1]; i++ ) //sunsets
	{

		if (!astronplace)
		{
           strcpy( fileon, &fileo[1][i][0] ); 
           kmyo = latbat[1][i];
 		   kmxo = lonbat[1][i];
		   hgt = hgtbat[1][i];
		}
		else
		{
           kmyo = eroslatitude;
		   kmxo = eroslongitude;
		   hgt = eroshgt;
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
		}

		ier = netzski6( fileon, kmyo, kmxo, hgt, geotz, nsetflag, i);

		if (ier) //error detected
		{
			ErrorHandler(10, ier);
			return -1;
		}

	}

	//now convert calculated sunset time in fractional hours to time string and
	//print out sunset table if only single table of sunset was requested
	switch (types)
	{
	case 1:
	case 5:
	  tabletyp = 1; //visible sunset index of tims
      PrintSingleTable( 1, tabletyp, stream2 ); //converts "tims" array in fractional hours to
                       //time strings and prints out table to output
      return 0;
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
	 
      PrintSingleTable( 1, tabletyp, stream2 ); //converts "tims" array in fractional hours to
                       //time strings and prints out table to output
      return 0;
	  break;
	}

	return 0;

}

////////////////////create headers for the different tables////////////////
short CreateHeaders( short types )
///////////////////////////////////////////////////////////////////////////
{

	char l_buff[255] = "";
	char l_buff1[255] = "";
	char l_buff2[255] = "";
	char l_buff3[255] = "";
	char l_buff4[255] = "";
	char l_buff5[255] = "";
	char l_buffcopy[255] = "";
	char l_buffwarning[255] = "";
	short ii = 25, iii = 31, iv = 29, iiv = 33; //astronomical sunrise captions
	char st0[255] = " of the Astronomical Sunrise of ";
	char st4[255] = " of the Astronomical Sunset of ";
	char st2[255] = "Based on the earliest astronomical sunrise for this place";
	char st3[255] = "Based on the latest astronomical sunset for this place";



	if (optionheb) // 'hebrew captions
	{
       if (eros || astronplace )
	   {
         sprintf( title, "%s%s%s%s", &heb1[1][0], "\"", TitleLine, "\"" );
	   }
       else
	   {
         sprintf( title, "%s%s%s%s", &heb1[1][0], "\"", TitleLine, "\"" );
	   }

	   //copyright and warning lines
	   sprintf( l_buffcopy, "%s%s%s%s%s%s%s%d%s%d", &heb1[24][0], " ", &heb2[2][0], "\""
			                    , &heb2[3][0], " ", &heb2[4][0], datavernum, ".", progvernum);
       sprintf( l_buffwarning, "%s%s%3.1f%s%s%s", &heb2[5][0], &heb2[6][0], distlim
			                                       , &heb2[7][0], "\"", &heb2[8][0] );
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
	   sprintf( l_buffcopy, "%s%d%s%d", "All times are according to Standard Time. Shabbosim are underlined. © Luchos Chai, Fahtal  56, Jerusalem  97430; Tel/Fax: +972-2-5713765. Version: "
			                      , datavernum, ".", progvernum );

       sprintf( l_buffwarning, "%s%3.1f%s", "The lighter colors denote times that may be inaccurate due to near obstructions (closer than "
                               , distlim, "km).  It is possible that such near obstructions should be ignored." );
	}

    if (optionheb && yrheb > 5810 || yrheb < 5710) 
	{
		ErrorHandler(8, 1);  //Warning: contact webmanager to update astronomical constants 
		return -1;
	}

    if ( astronplace ) //user inputed name as city's hebrew name
	{
      if ( optionheb && InStr(hebcityname, &heb1[5][0] ) ) //add " to USA's hebrew name
	  {
         sprintf( hebcityname, "%s%s%s%s%s", Mid(astname, 1, strlen(astname) - 4, l_buff)
		                                   , l_buff, &heb1[6][0], "\"", heb1[7][0] );
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
		   st0[0] = 0;	
		   strcpy( st0, " of the Mishor Sunrise of " );
		   st4[0] = 0;	
		   strcpy( st4, " of the Mishor Sunset of " );
		   st2[0] = 0;	
		   strcpy( st2, "Based on the earliest mishor sunrise for this place" );
		   st3[0] = 0;	
		   strcpy( st3, "Based on the latest mishor sunset for this place" );
		}

		if (optionheb) //Hebrew Eretz Yisroel cities
		{
			switch (types)
			{
				case 0:
				  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  }
				  else //use headers saved in the _port.sav file
				  {
					  sprintf( l_buff, "%s%s%s%s", &storheader[0][0][0], &heb2[1][0], " ", yrcal );
					  strcpy( &storheader[0][0][0], l_buff );
				  }
				  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
				  break;
				case 1:
				  //nsetflag = 1; //visible sunset/astr. sunset/chazos
					if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 )// use defulats
				  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
				  }
				  else //use headers stored in the _port.sav file
				  {
					  sprintf( l_buff, "%s%s%s%s", &storheader[1][0][0], &heb2[1][0], " ", yrcal );
					  strcpy( &storheader[1][0][0], l_buff );
				  }
				  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
				  break;
				case 2:
				  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
				  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[ii][0], hebcityname, &heb2[1][0], " ", yrcal );
				  sprintf( &storheader[0][1][0], "%s", &heb1[iv][0] );
				  sprintf( &storheader[0][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[0][5][0], "%s", "\0");
				  break;
				case 3:
				  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
				  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[iii][0], hebcityname, &heb2[1][0], " ", yrcal );
				  sprintf( &storheader[1][1][0], "%s", &heb1[iiv][0] );
				  sprintf( &storheader[1][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[1][5][0], "%s", "\0" );
				  break;
				case 4:
				  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 )//use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					  }
				  break;
				case 5:
				  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  }
				  break;
				case 6:
				  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					  }
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  }
				  break;
				case 7:
				  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
				  if ( InStr( &storheader[0][0][0], "NA") || strlen(Trim(&storheader[0][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[8][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[0][1][0], "%s", &heb1[16][0] );
					  sprintf( &storheader[0][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
					  }
				  if ( InStr( &storheader[1][0][0], "NA") || strlen(Trim(&storheader[1][0][0])) <= 2 ) //use defulats
					  {
					  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[9][0], hebcityname, &heb2[1][0], " ", yrcal );
					  sprintf( &storheader[1][1][0], "%s", &heb1[13][0] );
					  sprintf( &storheader[1][2][0], "%s", &heb1[21][0] );
					  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
					  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
					  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
					  }
				  break;
				case 8:
				  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
				  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[ii][0], hebcityname, &heb2[1][0], " ", yrcal );
				  sprintf( &storheader[0][1][0], "%s", &heb1[iv][0] );
				  sprintf( &storheader[0][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[0][5][0], "%s", "\0" );
				  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[iii][0], hebcityname, &heb2[1][0], " ", yrcal );
				  sprintf( &storheader[1][1][0], "%s", &heb1[iiv][0] );
				  sprintf( &storheader[1][2][0], "%s", &heb1[27][0] );
				  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
				  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
				  sprintf( &storheader[1][5][0], "%s", "\0" );
				  break;
			}
		}
		else //English Eretz Yisroel cities
		{
			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, st0, engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", st2 );
			  sprintf( &storheader[0][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, st4, engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", st3 );
			  sprintf( &storheader[1][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", "Based on the earliest sunrise that is seen anywhere within this place on any day" );
			  sprintf( &storheader[0][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset Times for ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", "Based on the latest sunset that is seen anywhere within this place on any day" );
			  sprintf( &storheader[1][2][0], "%s", "Terrain model: 25 meter DTM of Eretz Yisroel" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, st0, engcityname, " for the year ", yrheb  );
			  sprintf( &storheader[0][1][0], "%s", st2 );
			  sprintf( &storheader[0][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, st4, engcityname, " for the year ", yrheb  );
			  sprintf( &storheader[1][1][0], "%s", st3 );
			  sprintf( &storheader[1][2][0], "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  break;
			}
		}
	}
	else //eros or astronplace or both
	{
		ii = 41, iii = 42, iv = 43, iiv = 44; //astronomical sunrise/sunset captions
		st0[0] = 0;
		strcpy( st0, " of the Astronomical Sunrise of " );
		st4[0] = 0;
		strcpy( st4, " of the Astronomical Sunset of " );
		st2[0] = 0;
		strcpy( st2, "Based on the earliest astronomical sunrise for this place" );
		st3[0] = 0;
		strcpy( st3, "Based on the latest astronomical sunset for this place" );
		if ( !ast ) 
		{
		   ii = 43;
		   iii = 44;
		   iv = 36;
		   iiv = 40;
           st0[0] = 0;
		   strcpy( st0, " of the Mishor Sunrise of " );
           st4[0] = 0;
		   strcpy( st4, " of the Mishor Sunset of " );
           st2[0] = 0;
		   strcpy( st2, "Based on the earliest mishor sunrise for this place" );
           st3[0] = 0;
		   strcpy( st3, "Based on the latest mishor sunset for this place" );
		}

		if (optionheb)
		{
		    if (!astronplace && types !=2 && types != 3)
		    {
				switch( SRTMflag )
				{
				case 0:
				  sprintf( l_buff2, "%s%s%10.5lf%s%10.5lf%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslatitude
								 , &heb1[18][0], eroslongitude, " ) ", eroscity, &heb1[49][0]
								 , " ", searchradius, " ", &heb1[61][0] );
				  sprintf(l_buff1, "%s", &heb1[55][0] );
				  break;
				   
				case 1:
				  sprintf( l_buff2, "%s%s%10.5lf%s%10.5lf%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslatitude
								 , &heb1[18][0], eroslongitude, " ) ", eroscity, &heb1[51][0]
								 , " ", searchradius, " ", &heb1[61][0] );
				  sprintf(l_buff1, "%s", &heb1[56][0] );
				  break;

				case 2:
				  sprintf( l_buff2, "%s%s%10.5lf%s%10.5lf%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslatitude
								 , &heb1[18][0], eroslongitude, " ) ", eroscity, &heb1[53][0]
								 , " ", searchradius, " ", &heb1[61][0] );
				  sprintf(l_buff1, "%s", &heb1[55][0] );
				  break;

				case 9:
				  sprintf( l_buff2, "%s%s%7.3lf%s%7.3lf%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
								 , &heb1[18][0], eroslatitude, " ) ", eroscity, &heb1[58][0]
								 , " ", searchradius, " ", &heb1[61][0] );
				  sprintf(l_buff1, "%s", &heb1[21][0] );
				  break;

				}
			}
			else if (astronplace || types == 2 || types == 3)
			{
               if ( !ast ) //mishor
			   {
			       sprintf( l_buff4, "%s%s%s%s%s%s", title, &heb1[34][0], hebcityname, &heb2[1][0], " ", yrcal );
   			       sprintf( l_buff5, "%s%s%s%s%s%s", title, &heb1[37][0], hebcityname, &heb2[1][0], " ", yrcal );

        		   if (astronplace) //mishor calculation for a certain coordinate
					   {
					   sprintf( l_buff2, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[47][0], hebcityname, &heb1[18][0]
								, eroslatitude, " ,", &heb1[17][0], eroslongitude ); //sunrise
					   sprintf( l_buff2, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[48][0], hebcityname, &heb1[18][0]
								, eroslatitude, " ,", &heb1[17][0], eroslongitude ); //sunset
					   }
				   else //mishor calculation for a searchradius around a certain place
				   {
					   if (SRTMflag != 9)
					   {
						   sprintf( l_buff2, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
 				 						 , &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[35][0]
										 , " ", searchradius, " ", &heb1[66][0] );
						   sprintf( l_buff3, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
										 , &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[39][0]
										 , " ", searchradius, " ", &heb1[66][0] );
					   }
					   else if (SRTMflag == 9) //EY cities
					   {
						   sprintf( l_buff2, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
 				 						 , &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[35][0]
										 , " ", searchradius, " ", &heb1[66][0] );
						   sprintf( l_buff3, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
										 , &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[39][0]
										 , " ", searchradius, " ", &heb1[66][0] );
					   }
				   }

               }
			   else if (ast) // astronomical
			   {
			       sprintf( l_buff4, "%s%s%s%s%s%s", title, &heb1[25][0], hebcityname, &heb2[1][0], " ", yrcal ); //sunrise
  			       sprintf( l_buff5, "%s%s%s%s%s%s", title, &heb1[31][0], hebcityname, &heb2[1][0], " ", yrcal ); //sunset

                   if (astronplace) //astronomical calculation for a certain coordinate
				   {
					   sprintf( l_buff2, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[45][0], hebcityname, &heb1[18][0]
								, eroslongitude, " ,", &heb1[17][0], eroslatitude ); //sunrise
					   sprintf( l_buff2, "%s%s%s%10.5lf%s%s%10.5lf", &heb1[46][0], hebcityname, &heb1[18][0]
								, eroslongitude, " ,", &heb1[17][0], eroslatitude ); //sunset
				   }
				   else //astronomical calculation for a searchradius around a certain place
				   {
					   if (SRTMflag !=9 )
					   {
						   sprintf( l_buff2, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslatitude
									, &heb1[18][0], eroslongitude, " ) ", eroscity, " ", &heb1[28][0]
											 , " ", searchradius, " ", &heb1[66][0] );
						   sprintf( l_buff3, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslatitude
									, &heb1[18][0], eroslongitude, " ) ", eroscity, " ", &heb1[32][0]
											 , " ", searchradius, " ", &heb1[66][0] );
					   }
					   else if (SRTMflag == 9) //EY neighborhoods
					   {
						   sprintf( l_buff2, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
									, &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[28][0]
											 , " ", searchradius, " ", &heb1[66][0] );
						   sprintf( l_buff3, "%s%s%7.3lf%s%7.3lf%s%s%s%s%s%3.1f%s%s", "(", &heb1[17][0], eroslongitude
									, &heb1[18][0], eroslatitude, " ) ", eroscity, " ", &heb1[32][0]
											 , " ", searchradius, " ", &heb1[66][0] );
					   }
				   }
					
			   }

               sprintf(l_buff1, "%s", &heb1[26][0] );
			}

			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s", l_buff4 ); //, "%s", &heb1[ii][0] );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s", l_buff5 ); //, "%s", &heb1[iii][0] );
			  sprintf( &storheader[1][1][0], "%s", l_buff3 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%s%s", title, &heb1[10][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%s%s", title, &heb1[11][0], hebcityname, &heb2[1][0], " ", yrcal );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s", l_buff4 ); //, "%s", &heb1[ii][0] );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s", l_buff5 ); //, "%s", &heb1[iii][0] );
			  sprintf( &storheader[1][1][0], "%s", l_buff3 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[0][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  break;
			}
		}
		else
		{
            if (! astronplace)
			{
				switch( SRTMflag )
				{
				   case 0:
					  sprintf( l_buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "GTOPO30 DTM based calculations of the earliest sunrise within "
								   , searchradius, "km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( l_buff1, "%s", "Terrain model: USGS GTOPO30" );
					  break;

				   case 1:
					  sprintf( l_buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-2 DTM based calculations of the earliest sunrise within "
								   , searchradius, "km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( l_buff1, "%s", "Terrain model: SRTM-2" );
					  break;

				   case 2:
					  sprintf( l_buff2, "%s%3.1f%s%s%s%10.5lf%s%10.5f", "SRTM-1 DTM based calculations of the earliest sunrise within "
								   , searchradius, "km. around ", eroscity, "; longitude: "
								   , eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( l_buff1, "%s", "Terrain model: SRTM-1" );
					  break;

				   case 9:
					  sprintf( l_buff2, "%s%3.1f%s%s%s%7.3lf%s%7.3lf", "Israel DTM based calculations of the earliest sunrise within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude );
					  sprintf( l_buff1, "%s", "Terrain model: 25m DTM of Eretz Yisroel" );
					  break;
				}
			}
			else
			{
			   if ( !ast )
			   {
                  if (SRTMflag != 9 ) //not EY 
				  {
					  sprintf( l_buff3, "%s%s%s%s%4d", title, " of the Astronomical Sunrise for ", engcityname, " for the year ", yrheb );
					  sprintf( l_buff4, "%s%s%s%s%4d", title, " of the Astronomical Sunset for ", engcityname, " for the year ", yrheb );

					  sprintf( l_buff2, "%s%s%s%10.5lf%s%10.5f", "Based on the astronomical sunrise for ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( l_buff3, "%s%s%s%10.5lf%s%10.5f", "Based on the astronomical sunset for ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
				  }
				  else  //EY neighborhoods
				  {
					  sprintf( l_buff2, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Earliest mishor sunrise within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude );

					  sprintf( l_buff3, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Earliest mishor sunset within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude );
				  }
			   }
			   else
			   {
			      if (SRTMflag != 9)
				  {
					  sprintf( l_buff3, "%s%s%s%s%4d", title, " of the Mishor Sunrise for ", engcityname, " for the year ", yrheb );
					  sprintf( l_buff4, "%s%s%s%s%4d", title, " of the Mishor Sunset for ", engcityname, " for the year ", yrheb );

					  sprintf( l_buff2, "%s%s%s%10.5lf%s%10.5f", "Based on the mishor sunrise for ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
					  sprintf( l_buff3, "%s%s%s%10.5lf%s%10.5f", "Based on the mishor sunset for ", eroscity
									  , "; longitude: ", eroslongitude, ", latitude: ", eroslatitude );
				  }
				  else //EY neighborhoods
				  {
					  sprintf( l_buff2, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Latest astronomical sunrise within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude );

					  sprintf( l_buff3, "%s%3.1f%s%s%s%7.1lf%s%7.1lf", "Latest astronomical sunset within "
								   , searchradius, "km. around ", eroscity, "; ITMy: "
								   , eroslatitude, ", ITMx: ", eroslongitude );
				  }
			   }
			   
				  sprintf( l_buff1, "%s", "Based on refraction calculations for a combination of summer and winter mean mid-latitude atmospheres" );
			}

			switch (types)
			{
			case 0:
			  //nsetflag = 0; //visible sunrise/astr. sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 1:
			  //nsetflag = 1; //visible sunset/astr. sunset/chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 2:
			  //nsetflag = 2; //astronomical sunrise for hgt != 0, otherwise mishor sunrise/chazos
			  sprintf( &storheader[0][0][0], "%s", l_buff3 );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  break;
			case 3:
			  //nsetflag = 3; //astronomical sunset for hgt != 0, otherwise mishor sunset/chazos
			  sprintf( &storheader[1][0][0], "%s", l_buff4 );
			  sprintf( &storheader[1][1][0], "%s", l_buff3 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
			  break;
			case 4:
			  //nsetflag = 4; //vis sunrise/ mishor sunrise/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  break;
			case 5:
			  //nsetflag = 5; //visible sunset/ mishor sunset/ chazos
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 6:
			  //nsetflag = 6; //visible sunrise and visible sunset/ astr sunrise/astr sunset/ chazos
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 7:
			  //nsetflag = 7; //visible sunrise/ mishor sunrise and /visible sunset/ mishor sunset
			  sprintf( &storheader[0][0][0], "%s%s%s%s%4d", title, " of the Visible Sunrise for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", l_buffwarning );
			  sprintf( &storheader[1][0][0], "%s%s%s%s%4d", title, " of the Visible Sunset for the region of ", engcityname, " for the year ", yrheb );
			  sprintf( &storheader[1][1][0], "%s", l_buff2 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", l_buffwarning );
			  break;
			case 8:
			  //nsetflag = 8; //astronomical sunrise and sunset/ mishor sunrise and sunset if hgt = 0
			  sprintf( &storheader[0][0][0], "%s", l_buff3 );
			  sprintf( &storheader[0][1][0], "%s", l_buff2 );
			  sprintf( &storheader[0][2][0], "%s", l_buff1 );
			  sprintf( &storheader[0][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[0][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[0][5][0], "%s", "\0" );
			  sprintf( &storheader[1][0][0], "%s", l_buff4 );
			  sprintf( &storheader[1][1][0], "%s", l_buff3 );
			  sprintf( &storheader[1][2][0], "%s", l_buff1 );
			  sprintf( &storheader[1][3][0], "%s", &SponsorLine[1][0] );
			  sprintf( &storheader[1][4][0], "%s", l_buffcopy );
			  sprintf( &storheader[1][5][0], "%s", "\0" );
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
	char filroot[255] = "";
	char doclin[255] = "";
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
				 ErrorHandler(9, 2);
				 return -1;
			 }
			 else // read contents
			 {
				 while ( !feof(stream) )
				 {

					 fgets( doclin, 255, stream );

					 if (nline <= 4 && strlen(doclin) > 2 ) //sunrise and sunset info
					 {
					 strncpy( &storheader[tmpsetflag][nline][0], doclin, strlen(doclin) - 1 );
					 storheader[tmpsetflag][nline][strlen(doclin) - 1] = 0;

					 if (tmpsetflag == 1 && nline == 1 && strstr( &storheader[tmpsetflag][nline][0], &heb1[12][0] ) )
						 strcpy( &storheader[tmpsetflag][nline][0], &heb1[13][0] ); //typo fix for old sav files

					 }
					 else if (nline == 5) //steps
					 {
						 if ( !strstr( doclin, "NA") )
						 {
							 steps[tmpsetflag] = atoi( doclin ); //round to steps seconds
						 }
						 else //use default steps
						 {
							 steps[tmpsetflag] = RoundSeconds;
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
					 doclin[0] = 0; //clear the buffer

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

/////////////////////////check inputs//////////////////////
short CheckInputs(char *strlat, char *strlon, char *strhgt)
//////////////////////////////////////////////////////////
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
		if ( !strstr( "0123456789.", Mid(strlat, i, 1, ch) ) ) return -1; //found non numerical character
	}
	
	for (i = 0; i < lonlen; i++)
	{
		if ( !strstr( "0123456789.", Mid(strlon, i, 1, ch) ) ) return -1; //found non numerical character
	}

	for (i = 0; i < hgtlen; i++)
	{
		if ( !strstr( "0123456789.", Mid(strhgt, i, 1, ch) ) ) return -1; //found non numerical character
	}

	//now convert
	lat = atof( strlat );
	lon = atof( strlon );
	hgt = atof( strhgt );

    if (lat + 66 < 0 || lat - 66 > 0 ) return -1;
    if (lon + 180 < 0 || lat - 180 > 0 ) return -1;
    if ((lon > 0 && geotz > 0) || (lon < 0 && geotz < 0)) return -1;

	//if got here, then everything is OK!

	return 0;
}



/////////////////////////Hebrew Calendar/////////////////////////////
void HebCal(short yrheb)
/////////////////////////////////////////////////////////////////////
//calculates Hebrew year months for selected year and other Hebrew Year values
//////////////////////////////////////////////////////////////////////////
{

/****************declare  functions*************************/

    void newdate(short nyear, short nnew, short flag);

    void lpyrcivil();

	void engdate();

	void dmh(short n2rh);



//declare local variables	
	
	short n1day = 0;
	double cal1 = 0;
	double cal2 = 0;
	double cal3 = 0;
	short k = 0;
	short leapyr = 0;
	short iheb = 0;
	short constdif = 0;
	short nnew = 0;
	short kyr = 0;
	short flag = 0;
	short nyear = 0;
	short yrstep = 0;
	short difdec = 0;
	char buff[10];
	char monthhh[8] = "";
	

//	lpyr:

	//difdec = 0;
	//stdyy = 299; //English date paramters of Rosh Hashanoh year 1
	stdyy = 275;   //'English date paramters of Rosh Hashanoh 5758
	//styr = -3760;
	short styr = 1997;
	//yl = 366;
	yl = 365;
	//rh2 = 2; //Rosh Hashanoh year 1 is on day 2 (Monday)
	rh2 = 5;      //'Rosh Hashanoh 5758 is on day 5 (Thursday)
	yrr = 2; //5758 is a regular kesidrah year of 354 days
	leapyear = 0;
	//ncal1 = 2; //molad of Tishri 1 year 1-- day;hour;chelakim
	//ncal2 = 5;
	//ncal3 = 204;
	ncal1 = 5; //molad of Tishri 1 5758 day;hour;chelakim
	ncal2 = 4;
	ncal3 = 129; 
	nt1 = ncal1;
	nt2 = ncal2;
	nt3 = ncal3;
	n1rhoo = nt1;
	leapyr2 = leapyear;
	n1yreg = 4; //change in molad after 12 lunations of reg. year
	n2yreg = 8;
	n3yreg = 876;
	n1ylp = 5; //change in molad after 13 lunations of leap year
	n2ylp = 21;
	n3ylp = 589;
	n1mon = 1; //monthly change in molad after 1 lunation
	n2mon = 12;
	n3mon = 793;
    n11 = ncal1; //initialize molad
	n22 = ncal2;
	n33 = ncal3;  

	//Cls
	//chosen year to calculate monthly moladim
	//yrstep = yrheb - 1;
	yrstep = yrheb - 5758;

	nyear = 0;
	flag = 0;
	for (kyr = 1; kyr <= yrstep; kyr++)
	{
		nyear = nyear + 1;
		nnew = 1;
  		newdate( nyear, nnew, flag);
		n1rhooo = n1rho;
	}
	//now calculate molad of Tishri 1 of next year in order to
	//determine if the desired year is choser,kesidrah, or sholem
	leapyr2 = leapyear;
	nyear = nyear + 1;
	flag = 1;
	nnew = 0;
	newdate(nyear, nnew, flag);
	//now calculate english date and molad of each rosh chodesh of desired year, yr%
	n1rh = n1rhoo;
	n2rh = nt2;
	n3rh = nt3;

	constdif = n1rhooo - rh2;
	rhday = rh2;
	if (rhday == 0)
	{
		rhday = 7;
	}

	//record year datum to be used for parshiot and holidays
	dayRoshHashono = rhday;
	yeartype = yrr;
	hebleapyear = false;
	if (leapyear == 1)
	{
		hebleapyear = true;
	}

	dmh(n2rh);

	dyy = stdyy; //- difdec%
	//strncpy( hdryr, "-", 1 );
    //convert stdyy to string and concatenate it to hdryr
	itoa(styr, buff, 10);
	strncpy( hdryr + 1, buff, 4);

	engdate();
	iheb = 1;
	strcpy( &mdates[0][0][0], &monthh[1][0][0]);
	strncpy( &mdates[1][0][0], dates, 11);
	mdates[1][0][12] = 0;
	//mdates(1, 1) = monthh(iheb, 1);
	//mdates(2, 1) = dates;
	mmdate[0][0] = dyy;

	//start and end day numbers of the first year
	yrstrt[0] = dyy;
	yrend[0] = 365;
	if (leapyr == 1)
	{
		yrend[0] = 366;
	}

	//If newhebcalfm.Check4.Value = vbChecked Then
	//calculate dyy% for first shabbos
	fshabos0 = 7 - rhday + dyy;
	//   End If

	//now calculate other molados and their english date
	endyr = 12;
	if (leapyear == 1)
	{
		endyr = 13;
	}
	//If magnify = True Then GoTo 250
	for (k = 1; k <= endyr - 1; k++)
	{

		if (k < 5)
		{
			strcpy( monthhh, &monthh[1][k-1][0] );
		}
		else if (k == 5)
		{
			if (leapyear == 0)
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
			if (leapyear == 0)
			{
				strcpy( monthhh, &monthh[1][k][0] );
			}
			else
			{
				strcpy( monthhh, &monthh[1][k-1][0] );
			}
		}

		n33 = n3rh + n3mon;
		cal3 = (float)n33 / 1080;
		ncal3 = Cint((cal3 - Fix(cal3)) * 1080.0);
		n22 = n2rh + n2mon;
		cal2 = ((float)n22 + Fix(cal3)) / 24;
		ncal2 = Cint((cal2 - Fix(cal2)) * 24.0);
		n11 = n1rh + n1mon;
		cal1 = ((float)n11 + Fix(cal2)) / 7;
		ncal1 = Cint((cal1 - Fix(cal1)) * 7.0);
		n1rh = ncal1;
		n2rh = ncal2;
		n3rh = ncal3;

		dmh(n2rh);
		n1day = n1rh;
		if (n1day == 0)
		{
			n1day = 7;
		}
		if (k == 1)
		{
			dyy = dyy + 30;
		}
		else if (k == 2)
		{
			if (yrr != 3)
			{
				dyy = dyy + 29;
			}
			if (yrr == 3)
			{
				dyy = dyy + 30;
			}
		}
		else if (k == 3)
		{
			if (yrr == 1)
			{
				dyy = dyy + 29;
			}
			if (yrr != 1)
			{
				dyy = dyy + 30;
			}
		}
		else if (k == 4)
		{
			dyy = dyy + 29;
		}
		else if (k == 5)
		{
			dyy = dyy + 30;
		}
		else if (k >= 6 && leapyear == 0)
		{
			if (k == 6)
			{
				dyy = dyy + 29;
			}
			if (k == 7)
			{
				dyy = dyy + 30;
			}
			if (k == 8)
			{
				dyy = dyy + 29;
			}
			if (k == 9)
			{
				dyy = dyy + 30;
			}
			if (k == 10)
			{
				dyy = dyy + 29;
			}
			if (k == 11)
			{
				dyy = dyy + 30;
			}
		}
		else if (k >= 6 && leapyear == 1)
		{
			if (k == 6)
			{
				dyy = dyy + 30;
			}
			if (k == 7)
			{
				dyy = dyy + 29;
			}
			if (k == 8)
			{
				dyy = dyy + 30;
			}
			if (k == 9)
			{
				dyy = dyy + 29;
			}
			if (k == 10)
			{
				dyy = dyy + 30;
			}
			if (k == 11)
			{
				dyy = dyy + 29;
			}
			if (k == 12)
			{
				dyy = dyy + 30;
			}
		}
        //strncpy( hdryr, "-", 1 );
        //convert styr to string and concatenate it to hdryr
		itoa(styr, buff, 10);
	    strncpy( hdryr + 1, buff, 4);

		dyy = dyy - 1;
		engdate();
		strncpy( &mdates[2][k - 1][0], dates, 11 );
		mmdate[1][k - 1] = dyy;
		dyy = dyy + 1;
		engdate();
		strcpy( &mdates[0][k][0], monthhh );
		strncpy( &mdates[1][k][0], dates, 11 );
		mdates[1][k][12] = 0;
		mmdate[0][k] = dyy;
	}

	250;
	if (styr == 1997)
	{
        styr = 1998;
	}
	dyy = dyy + 28;
    //strncpy( hdryr, "-", 1 );
    //convert styr to string and concatenate it to hdryr
	itoa(styr, buff, 10);
    strncpy( hdryr + 1, buff, 4);

	engdate(); //: LOCATE 13, 5: Print "end of year: "; dates$
	strncpy( &mdates[1][endyr - 1][0], dates, 11 );
	mdates[1][endyr - 1][12] = 0;
	mmdate[1][endyr - 1] = dyy;
	yrend[1] = dyy;
}

void newdate(short nyear, short nnew, short flag)
{
	short leapyr2 = 0;
	short yd = 0;
	short difdyy = 0;
	short difrh = 0;
	short yrstep = 0;
	short kyr = 0;
	double cal1 = 0;
	double cal2 = 0;
	double cal3 = 0;
	short n111 = 0;
	short n222 = 0;
	short n333 = 0;

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
			leapyear = 1;
			n111 = n1ylp;
			n222 = n2ylp;
			n333 = n3ylp;
			break;
		default:
			leapyear = 0;
			n111 = n1yreg;
			n222 = n2yreg;
			n333 = n3yreg;
			break;
	}
	n33 = n33 + n333;
	cal3 = (float)n33 / 1080;
	ncal3 = Cint((cal3 - Fix(cal3)) * 1080.0);
	n22 = n22 + n222;
	cal2 = ((float)n22 + Fix(cal3)) / 24;
	ncal2 = Cint((cal2 - Fix(cal2)) * 24.0);
	n11 = n11 + n111;
	cal1 = ((float)n11 + Fix(cal2)) / 7;
	ncal1 = Cint((cal1 - Fix(cal1)) * 7.0);
	n11 = ncal1;
	n22 = ncal2;
	n33 = ncal3; //molad of Tishri 1 of this iteration
	//now use dechiyos to determine which day of week Rosh Hashanoh falls on
	n1rh = n11;
	n2rh = n22;
	n3rh = n33;
	//difd% = 0
	n1rho = n1rh;
	switch (nyear)
	{
		case 3:
		case 6:
		case 8:
		case 11:
		case 14:
		case 17:
		case 19:
			if (n11 == 2 && n22 + n33 / 1080 > 15.545)
			{
				n1rh = n1rh + 1;
				//difd% = 1
				goto HC500;
			}
			break;
	}
	if (n2rh >= 18)
	{
		n1rh = n1rh + 1;
		//difd% = 1
		if (n1rh == 8)
		{
			n1rh = 1;
		}
		if (n1rh == 1 || n1rh == 4 || n1rh == 6)
		{
			n1rh = n1rh + 1;
		}
		goto HC500;
	}
	if (n1rh == 1 || n1rh == 4 || n1rh == 6)
	{
		n1rh = n1rh + 1;
		//difd% = 1
	}
	//GOTO 500

	//    IF (leapyear% <> 1 AND flag% = 0) AND n1rh% = 3 AND n2rh% + n3rh% / 1080 > 9.188 THEN
	//       n1rh% = 5
	//       END IF
	if ((flag == 0 || flag == 1 || kyr == yrstep) && n1rh == 3 && n2rh + n3rh / 1080 > 9.188)
	{
		switch (nyear + 1)
		{
			case 20:
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
				n1rh = 5;
				break;
		}
	}
	if (n1rh == 0)
	{
		n1rh = 7;
	}

HC500:
	if (rh2 >= n1rh)
	{
		difrh = 7 - rh2 + n1rh;
	}
	if (rh2 < n1rh)
	{
		difrh = n1rh - rh2;
	}
	if (nnew == 1)
	{
		n1rhoo = n1rh;
	}
	if ((leapyear == 0 && difrh == 3) || (leapyear == 1 && difrh == 5))
	{
		yrr = 1;
	}
	else if ((leapyear == 0 && difrh == 4) || (leapyear == 1 && difrh == 6))
	{
		yrr = 2;
	}
	else if ((leapyear == 0 && difrh == 5) || (leapyear == 1 && difrh == 7))
	{
		yrr = 3;
	}
	if (leapyear == 0)
	{
		if (yrr == 1)
		{
			difdyy = 353;
		}
		if (yrr == 2)
		{
			difdyy = 354;
		}
		if (yrr == 3)
		{
			difdyy = 355;
		}
	}
	else if (leapyear == 1)
	{
		if (yrr == 1)
		{
			difdyy = 383;
		}
		if (yrr == 2)
		{
			difdyy = 384;
		}
		if (yrr == 3)
		{
			difdyy = 385;
		}
		//INPUT "cr", crr$
	}
	if (flag != 1)
	{
		dyy = stdyy + difdyy - yl; //- difdec%
		stdyy = dyy;
		styr = styr + 1;
		yd = styr - 1988;
		yl = 365;
		//LOCATE 18, 1: PRINT "styr%,dy;difdyy%="; styr%; dyy%; difdyy%
		//LOCATE 19, 1: PRINT "rh2%,n1rh%;leapyear;leapyr2%="; rh2%; n1rh%; leapyear%; leapyr2%
		//LOCATE 20, 1: PRINT "leapyear%;leapyr2%;year%;kyr%"; leapyear%; leapyr2%; kyr% - 1 + 5758
		//INPUT "cr", crr$
		if (yd % 4 == 0)
		{
			yl = 366;
		}
		if (yd % 4 == 0 && styr % 100 == 0 && styr % 400 != 0)
		{
			yl = 365;
		}
		rh2 = n1rh; //rh2% is the day of the week (1-7) of the Rosh Hashonoh
		leapyr2 = leapyear;
		nt1 = n11;
		nt2 = n22;
		nt3 = n33;
	}
}

void lpyrcivil()
{
	short styr = 0;
	short yl = 0;
	short yd = 0;
	short yrheb = 0;
	yd = yrheb - 1988;
	yl = 365;
	if (yd % 4 == 0)
	{
		yl = 366;
	}
	if (yd % 4 == 0 && styr % 100 == 0 && styr % 400 != 0)
	{
		yl = 365;
	}
}

void engdate()
{
//	char monthe[12][5];
	short leapyr = 0;
	short yl = 0;
	short myear0 = 0;
	short myear = 0;
	short yleng = 0;
	short yreng = 0;
	short ydeng = 0;
	short newyear = 0;
	short dyyn = 0;
	char buff[10];

	newyear = 0;
	ydeng = styr - 1988;
	yreng = styr;
	yleng = 365;
	if (ydeng % 4 == 0)
	{
		yleng = 366;
	}
	if (ydeng % 4 == 0 && yreng % 100 == 0 && yreng % 400 != 0)
	{
		yleng = 365;
	}
	if (dyy > yleng )
	{
		myear = yreng;
		myear0 = myear;
		yreng = yreng + 1;
        //strncpy( hdryr, "-", 1 );
        //convert styr to string and concatenate it to hdryr
        itoa(yreng, buff, 10);
        strncpy( hdryr + 1, buff, 4);

		dyy = dyy - yleng;
		ydeng = yreng - 1988;
		yleng = 365;
		if (ydeng % 4 == 0)
		{
			yleng = 366;
		}
		if (ydeng % 4 == 0 && yreng % 100 == 0 && yreng % 400 != 0)
		{
			yleng = 365;
		}
		styr = yreng;
		yl = yleng;
		newyear = 1;
	}
	leapyr = 0;
	if (yl == 366) //leap years
	{
		leapyr = 1;
	}

	//determine the day of the month, and the month
	if (dyy >= 1 && dyy < 32)
	{
		dyyn = dyy;
		strncpy( dates, "Jan-", 4 );
	}
	if (dyy >= 32 && dyy < 60 + leapyr)
	{
		dyyn = dyy - 31;
		strncpy( dates, "Feb-", 4 );
	}
	if (dyy >= 60 + leapyr && dyy < 91 + leapyr)
	{
		dyyn = dyy - 59 - leapyr;
		strncpy( dates, "Mar-", 4 );
	}
	if (dyy >= 91 + leapyr && dyy < 121 + leapyr)
	{
		dyyn = dyy - 90 - leapyr;
		strncpy( dates, "Apr-", 4 );
	}
	if (dyy >= 121 + leapyr && dyy < 152 + leapyr)
	{
		dyyn = dyy - 120 - leapyr;
		strncpy( dates, "May-", 4 );
	}
	if (dyy >= 152 + leapyr && dyy < 182 + leapyr)
	{
		dyyn = dyy - 151 - leapyr;
		strncpy( dates, "Jun-", 4 );
	}
	if (dyy >= 182 + leapyr && dyy < 213 + leapyr)
	{
		dyyn = dyy - 181 - leapyr;
		strncpy( dates, "Jul-", 4 );
	}
	if (dyy >= 213 + leapyr && dyy < 244 + leapyr)
	{
		dyyn = dyy - 212 - leapyr;
		strncpy( dates, "Aug-", 4 );
	}
	if (dyy >= 244 + leapyr && dyy < 274 + leapyr)
	{
		dyyn = dyy - 243 - leapyr;
		strncpy( dates, "Sep-", 4 );
	}
	if (dyy >= 274 + leapyr && dyy < 305 + leapyr)
	{
		dyyn = dyy - 273 - leapyr;
		strncpy( dates, "Oct-", 4 );
	}
	if (dyy >= 305 + leapyr && dyy < 335 + leapyr)
	{
		dyyn = dyy - 304 - leapyr;
		strncpy( dates, "Nov-", 4 );
	}
	if (dyy >= 335 + leapyr && dyy < 365 + leapyr)
	{
		dyyn = dyy - 334 - leapyr;
		strncpy( dates, "Dec-", 4 );
	}

	//now concatenate the day and year to the month to generate "dates"
	itoa(dyyn, buff, 10);

	if (dyy < 10 )
	{
		strncat( dates, "0", 1);
        strncat( dates, buff, 1);
	}
	else
	{
	strncpy( dates + 4, buff, 2);
	}
	strncpy( dates + 6, hdryr, 5);

	//IF newyear% = 1 AND yl% = 366 THEN dyy% = dyy% - 1
}

void dmh(short n2rh)
{
	float minc = 0;

	Hourr = n2rh + 6;
	if (Hourr < 12)
	{
		tm = " PM night";
	}
	else if (Hourr >= 12)
	{
		Hourr = Hourr - 12;
		if (Hourr < 12)
		{
			if (Hourr == 0)
			{
				Hourr = 12;
			}
			tm = " AM";
		}
		else if (Hourr >= 12)
		{
			Hourr = Hourr - 12;
			if (Hourr == 0)
			{
				Hourr = 12;
			}
			tm = " PM afternoon";
		}
	}
	minc = ((float)n3rh * 60.0 ) / 1080.0;
	Min = Fix(minc);
	hel = (short)((minc - Min) * 18);

}

/////////////////////Prints single table/////////////////////
short PrintSingleTable(short setflag, short tbltypes, FILE *stream2)
/////////////////////////////////////////////////////////
{
	char t3subb[12] = "12:00:00 PM";
	short sp = 0;
	short j = 0;
	short addmon = 0;
	short dayweek = 0;
	short yrn = 0;
	short i = 0;
	short yl = 0;
	short yd = 0;
	short myear1 = 0;
	short k = 0;
	short ntotshabos = 0;
	short nshabos = 0;
	short nadd = 0;
	short myear = 0;
	short myear0 = 0;
	short rdflag = 0;
	short katznum = 0;
	short fshabos = 0;
	short iheb = 0;
	NearWarning[0] = false;
	NearWarning[1] = false;

	yrn = 1;
	nshabos = 0;
	dayweek = rhday - 1;
	addmon = 0;
	ntotshabos = 0;
	fshabos = fshabos0;
	if (!(optionheb)) iheb = 1;

   	myear = yrheb - 3761;

	yd = myear - 1988;
    yl = 365;
    if (yd % 4 == 0) yl = 366;
    if ((yd % 4 == 0) && (myear % 100 == 0) && (myear % 400 != 0)) yl = 365;

	sp = 1; //round up
	if (setflag == 1) sp = 2; //round down

	for (i = 1; i <= endyr; i++)
	{

		if (mmdate[1][i - 1] > mmdate[0][i - 1] )
		{

			if (i > 1 && mmdate[0][i - 1] == 1)
			{
				//Rosh Chodesh is on January 1, so increment the year
				//(this happens in the year 2006, 2036)
				yrn++;
			}

			strcpy( &stortim[setflag][i - 1][29][0], "\0" ); //zero time

			k = 0;
			for (j = mmdate[0][i - 1]; j <= mmdate[1][i - 1]; j++)
			{
				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}
		}
		else if (mmdate[1][i - 1] < mmdate[0][i - 1])
		{

			strcpy( &stortim[setflag][i - 1][29][0], "\0" ); //zero time


			k = 0;

			for (j = mmdate[0][i - 1]; j <= yrend[0]; j++)
			{
				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}


			yrn++;

			for (j = 1; j <= mmdate[1][i - 1]; j++)
			{
				PST_iterate( &k, &j, &i, &ntotshabos, &dayweek, &yrn, &fshabos,
					         &nshabos, &addmon, setflag, sp, tbltypes, yl, iheb );
			}


		}
	}

	//print out single tables to stdout or to html
	if (endyr == 12)
	{
		heb12monthHTML(stream2, setflag);
	}
	else if (endyr == 13)
	{
		heb13monthHTML(stream2, setflag);
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

//       convert dy at particular yr to calendar date
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
	char ch[3] = "";
	char caldate[12] = "";

	*k += 1;

	round( tims[*yrn - 1][tbltypes][*j - 1], setflag, sp, t3subb );
	if (tbltypes == 0 || tbltypes == 1) //check for near obstructions
	{
		if ( dobs[*yrn - 1][setflag][*j - 1] < distlim ) //add obstruction flag to string
		{
			strcat( t3subb, "*");
			NearWarning[setflag] = true;
		}
	}

	if (*fshabos + *nshabos * 7 == *j) //this is shabbos
	{
		*nshabos += 1; //<<<2--->>>
		*ntotshabos += 1;
		//add sedra info
		strcpy( &stortim[5][*i - 1][*k - 1][0], &arrStrSedra[0][*ntotshabos - 1][0] );
		strcpy( &stortim[6][*i - 1][*k - 1][0], &arrStrSedra[1][*ntotshabos - 1][0] );

		//add marker that it is Shabbos
		strcat( t3subb, "_");

		//check for end of year
		if (*fshabos + *nshabos * 7 > yrend[0])
		{
			*fshabos = 7 - (yrend[0] - (*fshabos + (*nshabos - 1) * 7));
			*nshabos = 0;
		}
	}
	else
	{
		strcpy( &stortim[5][*i - 1][*k - 1][0], "-----" );
		strcpy( &stortim[6][*i - 1][*k - 1][0], "-----" );
	}

	//now actually store the time string
	strcpy( &stortim[setflag][*i - 1][*k - 1][0], t3subb );

	//store the civil date
	EngCalDate( *j, FirstSecularYr + *yrn - 1, yl, caldate );
	strcpy( &stortim[2][*i - 1][*k - 1][0], caldate );

	if (optionheb == true)
	{
		hebnum(*k, ch);
	}
	else
	{
		itoa(*k, ch, 10);
	}

	//store Hebrew date
	sprintf( &stortim[3][*i - 1][*k - 1][0], "%s%s%s", ch, "-",  &monthh[iheb][*i + *addmon - 1][0] );

	if (endyr == 12 && *i == 6)
	{
		sprintf( &stortim[3][*i - 1][*k - 1][0], "%s%s%s", ch, "-", &monthh[iheb][14][0] );
		*addmon = 1;
	}

	*dayweek += 1;
	if (*dayweek == 8)
	{
		*dayweek = 1;
	}

	strcpy( &stortim[4][*i - 1][*k - 1][0], hebweek(*dayweek, ch) );

}



///////////////////////function round///////////////////////////
char *round( double tim, short setflag, short sp, char t3subb[] )
////////////////////////////////////////////////////////////////
//converts fractional hour into time string
//rounds to requested step size and adds cushions
////////////////////////////////////////////////////////////////
//Note: "setflag" does not determine direction of rounding for this program!!
//It is determined solely by the value of "sp".
//setflag = 0 for sunrise
//setflag = 1 for sunset
/////////////////////////////////////////////////////////////
//sp = 1 forces rounding up
//sp = 2 forces rounding down
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
	short naccur = 0;
	short thr = 0, tmin = 0, tsec = 0, secmin = 0, hrhr = 0;
	short minad = 0, secmins = 0, minmin = 0, hrad = 0;
	double t1, ssec;
	char ch[3] = "";


	thr = (short)(tim);
	tmin = (short)((tim - thr) * 60.0 );
	t1 = (double)tmin/60.0;
	tsec = (short)((tim - thr - t1) * 3600.0 + 0.5) + accur[setflag];

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
	if (thr > 12) thr = thr - 12;

	/////////////seconds/////////////////////////////// 

	//round up or down
	secmin = tsec;

	minad = 0;
	if (secmin % steps[setflag] == 0) //no need for rounding
	{
		goto rnd50; 
	}

	if (sp == 1) //round up
	{
		ssec = (double)secmin / (double)steps[setflag];
		secmins = Cint(Fix(ssec + 0.499999));
		if (ssec - (short)Fix(ssec) + 0.000001 < 0.1)
		{
			secmins++;
		}
		secmin = steps[setflag] * secmins;
		if (secmin == 60)
		{
			secmin = 0;
			minad = 1;
		}
	}
	else if (sp == 2) //round down
	{
		secmin = steps[setflag] * (short)(Fix((secmin / steps[setflag]) * 10) / 10);
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


	////////////////write time string//////////////////////

	sprintf( t3subb, "%d:%002d:%002d%s", hrhr, minmin, secmin, "\0" );
	//sprintf( t3subb, "%002d:%002d:%002d%s", hrhr, minmin, secmin, "\0" );

//	if (InStr(Mid(t3subb, 1, 1, ch ), "0" ) )
//		{
//			strncpy( t3subb, t3subb + 1, strlen(t3subb));
//		}


	return t3subb;

} 


/*
/////////////////////////string routine finding size of strings///////////////
//routines using for find first and last non white string (precursors to "Trim")
//taken from Wikipedia entry for "Trim" 
*/

/*
int isspace ( char c )
{
   return c == ' '  || c == '\t' ||
          c == '\v' || c == '\f'  ;  
}

char *LTrim ( char * p )
{
    while ( p && isspace(*p) ) ++p ;
	strcat( p, "\0" ); //terminate the string
    return p ;
}

char *RTrim ( char * p )
{
    char * temp;
    if (!p) return p;
    temp = (char *)(p + strlen(p)-1);
    while ( (temp>=p) && isspace(*temp) ) --temp;
    *(temp+1) = '\0';
    return p;
}

char *Trim ( char * p )
{
    return strcat( LTrim(RTrim(p)), "\0");
}
*/


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

 ///////////////emulation of Trim functions///////////////////////
 //LTrim
 char *LTrim(char * str)
 {
	 short itmp;
	 itmp = find_first_nonwhite_character(str, 0);
	 return strcpy( str, &str[itmp] );
	 //return MyStrcpy( str, &str[itmp]);
 }
 //RTrim
char *RTrim(char * str)
 {
	 short itmp;
	 itmp = find_last_nonwhite_character(str, 0);
	 str[itmp] = 0;
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
    if ( nstart + nsize > lenstr ) nsize = lenstr - nstart; //nstart too big, so reduce it

	char l_buff[255]="";
	strncpy( l_buff, str + nstart - 1, nsize );
	l_buff[nsize] = 0;
	strcpy( buff, l_buff);
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
    if ( nstart + nsize > lenstr ) nsize = lenstr - nstart; //nstart too big, so reduce it

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
	//works like VB, i.e., substrings begin at position 1 and NOT at 0
	char l_buff[255] = "";
	short lenstr1 = 0;
	short lenstr2 = 0;
	char *pos;
	lenstr1 = strlen(str);
	lenstr2 = strlen(str2);
	if (lenstr2 > lenstr1) return 0; //subdtring bigger than string
	strncpy( l_buff, str + nstart - 1, lenstr1 - nstart + 1);
	l_buff[lenstr1 - nstart + 1] = 0;
	pos = strstr( l_buff, str2) ;
    if (pos == NULL )
	{
		return 0; //substring not found
	}
	else //substring found
	{
		return pos - l_buff + nstart; //like VB position
	}
}

//////////////////emulation of InStr(2 parameters) /////////////////////
int InStr( char * str, char * str2 )
{
	//works like VB, i.e., substrings begin at position 1 and NOT at 0
	char l_buff[255] = "";
	short lenstr1 = 0;
	short lenstr2 = 0;
	char *pos;
	lenstr1 = strlen(str);
	lenstr2 = strlen(str2);
	if (lenstr2 > lenstr1) return 0; //subdtring bigger than string
	strncpy( l_buff, str, lenstr1 );
	l_buff[lenstr1] = 0;
	pos = strstr( l_buff, str2) ;
    if (pos == NULL )
	{
		return 0; //substring not found
	}
	else //substring found
	{
		return pos - l_buff + 1; //like VB position
	}

}
/*
/////////////////emulation of Replace function////////////////////////////
char *Replace( char *str, char *Old, char *New, char buff[] )
//replaces any character "Old" with charcater "New" in the string "str"
//returns the new "str", Note: must declare: char buff[size]
//Example: remove "_" from string "random_place"
//
//strcpy( filroot, "random_place" );
//Replace( filroot, "_", " ", buff );
//returns "random place"
//
{
	char l_buff[1] = "";
	strcpy( buff, str);

	for (short i= 1; i <= strlen(str); i++) //starts at i=1 like VB
	{
		if (InStr( Mid( str, i, 1, l_buff ), Old ))
		{
			strncpy( buff + i - 1,  New, 1);
		}
	}
	return buff;
}
*/

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
char *LCase(char *str, char buff[])
//example: LCase(citnam, buff)

{
   strcpy( buff, str );

   for (short i=0; i<strlen(buff); i++)
   {
      if (buff[i] >= 0x41 && buff[i] <= 0x5A)
      buff[i] = buff[i] + 0x20;
  }
  return buff;
}

/*
////////////////MyStrcpy/////////////////////////////////
char *MyStrcpy(char *des,char *src)
{
	while(*src)
	{
		*des++=*src++;
	}
	*des=0;
	return des;
}
*/

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

/*
//////////////////convert from DOS Hebrew to ISO Hebrew//////////
char *DosToIso( char *str, char buff[] )
/////////////////////////////////////////////////
//converts Windows Hebrew (DOS) strings into
// standard (ISO) Hebrew that can be
//read by Linux, etc.
///////////////////////////////////////////////////////
{
	char ch[2] = "";

	strcpy( buff, str );

	for (short i=0; i < strlen(str); i++ )
	{
		Mid( str, i, 1, ch );
		if ( (char)ch >= (char)128 )
		{
			buff[i] = buff[i] + 96; //Aleph is (char)224 instead of (char)128 )
		}
	}

	return buff;
}
*/
//////////////////////////LoadConstants//////////////////////////
short LoadConstants()
///////////////////loads up constanst//////////////////////
//the weather constants must be changed as Global warming
//1. changes the atmosphere.  In partiuclar, the temperature
//zones will increase in latitude
//2. menat's program constants should be pretty constant barring
//new modeling of the atmosphere.
//3. astronomical constants should be updated before 2050
{

	char doclin[255] = "";
	char buff[255] = "";
	short ipath = 0, i = 0;
	FILE *stream;

	//load math constants

	pi = acos( -1 );
	cd = pi / 180.;
	pi2 = pi * 2.0;
	ch = 12. / pi;
	hr = 60.;


//load up paths, refraction, atmoshperic, and astronomical constants

//file must reside in directory that this program runs in, i.e., the cgi-bin directory

	if (stream = fopen( "PathsZones.txt", "r" ) )
	{
		while (!feof(stream))
		{
			fgets( doclin, 255, stream ); //look for headers
			
			if (InStr( doclin, "[paths]" ))
			{
				ipath = 1;
			}
			else if (InStr( doclin, "[world zones]" ))
			{
				ipath = 2;
			}
			else if (InStr( doclin, "[Eretz Yisroel zones]" ))
			{
				ipath = 3;
			}
			else if (InStr( doclin, "[adhoc shift]" ))
			{
				ipath = 4;
			}
			else if (InStr( doclin, "[menatsum/win.ren]" ))
			{
				ipath = 5;
			}
			else if (InStr( doclin, "[astron. constants]" ))
			{
				ipath = 6;
			}
			else if (InStr( doclin, "[time cushions]" ))
			{
				ipath = 7;
			}
			else if (InStr( doclin, "[obstruction limit]" ))
			{
				ipath = 8;
			}
			else
			{
				ipath = 0; //blank detected
			}

			//read constants
			switch (ipath)
			{
			case 1:

				/////////read paths///////////////////////////
				fgets( doclin, 255, stream);
				strncpy( buff, doclin, strlen(doclin) - 1 );
				buff[strlen(doclin) - 1] = 0;
				strcpy( drivjk, buff);

				fgets( doclin, 255, stream);
				strncpy( buff, doclin, strlen(doclin) - 1 );
				buff[strlen(doclin) - 1] = 0;
				strcpy( drivweb, buff);

				fgets( doclin, 255, stream);
				strncpy( buff, doclin, strlen(doclin) - 1 );
				buff[strlen(doclin) - 1] = 0;
				strcpy( drivfordtm, buff);

				fgets( doclin, 255, stream);
				strncpy( buff, doclin, strlen(doclin) - 1 );
				buff[strlen(doclin) - 1] = 0;
				strcpy( drivcities, buff);

				ipath = 0;
				break;

			case 2:

				///////////read temperature zones for the rest of the world///////////////////
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[0], &sumwinzone[0][0], &sumwinzone[0][1], &sumwinzone[0][2], &sumwinzone[0][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[1], &sumwinzone[1][0], &sumwinzone[1][1], &sumwinzone[1][2], &sumwinzone[1][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[2], &sumwinzone[2][0], &sumwinzone[2][1], &sumwinzone[2][2], &sumwinzone[2][3] );
				fscanf( stream, "%f,%d,%d,%d,%d\n", &latzone[3], &sumwinzone[3][0], &sumwinzone[3][1], &sumwinzone[3][2], &sumwinzone[3][3] );
				ipath = 0;
				break;

			case 3:

				/////////read temperature zone constants for Eretz Yisroel//////////////////
				fscanf( stream, "%d,%d,%d,%d\n", &sumwinzone[4][0], &sumwinzone[4][1], &sumwinzone[4][2], &sumwinzone[4][3] );
				ipath = 0;
				break;

			case 4:

				//read adhoc fix to times (secs)
				fscanf( stream, "%d\n", &nacurrFix );
				ipath = 0;
				break;

			case 5:

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

			case 6:

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

				fscanf( stream, "%lg\n", &ap0 );//mean anomaly for Jan 0, RefYear 00:00 at Greenwich meridian
				ap0 = cd * ap0;

				fscanf( stream, "%lg\n", &mp0 );//mean longitude of sun Jan 0, RefYear 00:00 at Greenwich meridian
				mp0 = cd * mp0;

				fscanf( stream, "%lg\n", &ob0 );//ecliptic angle for Jan 0, RefYear 00:00 at Greenwich meridian
				ob0 = cd * ob0;

				fscanf( stream, "%d\n", &RefYear);//Reference year of constants
				ipath = 0;
				break;

			case 7:

				//read cushions for Eretz Yisroel (25 m), GTOPO30 (1 km), SRTM-1 (90 m), SRTM-2 (30 m)
				fscanf( stream, "%f,%f,%f,%f,%f\n", &cushion[0], &cushion[1], &cushion[2], &cushion[3], &cushion[4] );
				ipath = 0;
				break;

			case 8:

				//read minimum distance to obstruction not to be considered a near obstruction
				fscanf( stream, "%g,%g,%g,%g\n", &distlims[0], &distlims[1], &distlims[2], &distlims[3] );
				ipath = 0;
				break;
			}
		}
		fclose( stream ); //nothing else to read

	}
	else //use defaults
	{

		strcpy( drivjk, "e:/jk/" ); //path to hebrew strings
		strcpy( drivweb, "c:/inetpub/webpub/" ); //path to web files
		strcpy( drivfordtm, "e:/fordtm/" ); //path to fordtm (temp directory)
		strcpy( drivcities, "e:/cities/" ); //path to cities directory

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

		//////////////////astronomical constants///////////////////
		ac0 = cd * .9856003; //change in the mean anamolay per day 
		mc0 = cd * .9856474; //change in the mean longitude per day
		ec0 = cd * 1.915; //constants of the Ecliptic longitude
		e2c0 = cd * .02;
		ob10 = cd * -4e-7; //change in the obliquity of ecliptic per day
	/*  mean anomaly for Jan 0, RefYear 00:00 at Greenwich meridian */
		ap0 = cd * 357.528;
	/*  mean longitude of sun Jan 0, RefYear 00:00 at Greenwich meridian */
		mp0 = cd * 280.461;
	/*  ecliptic angle for Jan 0, RefYear 00:00 at Greenwich meridian */
		ob0 = cd * 23.439;
	   /////////////////////////////////////////////////////////////
		RefYear = 1996; //reference year for ephemerel calculations
		////////////////////////////////////////////////////////////

		///////////////cushions////////////////////
		cushion[0] = 15; //EY (25 m)
		cushion[1] = 45; //GTOPO30/SRTM30 (1 km)
		cushion[2] = 35; //SRTM-1 (90 m)
		cushion[3] = 20; //SRTM-2 (30 m)

		distlims[0] = 5; //EY (25 m)
		distlims[1] = 30; //GTOPO30/SRTM30 (1 km)
		distlims[2] = 10; //SRTM-1 (90 m)
		distlims[3] = 6;  //SRTM-2 (30 m)

	}

//////////Hebrew Months in Hebrew/English////////////
	sprintf( doclin, "%s%s%s", drivjk, "Calhebmonths.txt", "\0" );

	if (stream = fopen( doclin, "r" ) )
	{
		for( i = 0; i < 14; i++ )
		{
			fgets( doclin, 255, stream );
			strncpy( &monthh[0][i][0], doclin, strlen(doclin) - 1);
			strcat( &monthh[0][i][0], "\0" );
			
			if (Linux) //convert to ISO Hebrew string
			{
				DosToIso( &monthh[0][i][0], doclin );
				strcpy( &monthh[0][i][0], doclin );
			}
		
		}
		fclose( stream );
	}
	else //use defaults
	{
		strcpy( &monthh[0][0][0], "תשרי");
		strcpy( &monthh[0][1][0], "מרחשון");
		strcpy( &monthh[0][2][0], "כסלו");
		strcpy( &monthh[0][3][0], "טבת");
		strcpy( &monthh[0][4][0], "שבט");
		strcpy( &monthh[0][5][0], "'אדר א");
		strcpy( &monthh[0][6][0], "'אדר ב");
		strcpy( &monthh[0][7][0], "ניסן");
		strcpy( &monthh[0][8][0], "אייר");
		strcpy( &monthh[0][9][0], "סיון");
		strcpy( &monthh[0][10][0], "תמוז");
		strcpy( &monthh[0][11][0], "אב");
		strcpy( &monthh[0][12][0], "אלול");
		strcpy( &monthh[0][13][0], "אדר");

		//convert DOS Hebrew to ISO
		for ( i = 0; i < 14; i++ )
		{
			DosToIso( &monthh[0][i][0], doclin );
			strcpy( &monthh[0][i][0], doclin );
		}
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

	return 0;
}


////////////convert from DOS Hebrew to ISO Hebrew////////////
char *DosToIso( char *str, char *buff )
/////////////////////////////////////////////////////////////
//converts Windows Hebrew (DOS) strings into
// standard (ISO) Hebrew that can be
//read by Linux, etc.
/////////////////////////////////////////////////////////////
{
	int strLen = strlen(str);
	strcpy( buff, str);

	for(int i=0;i<strLen;i++)
	{
		if( buff[i] >= (char)128 && buff[i] < (char)224 )
		{
			buff[i] += 96;//'א'-128
		}
	}
	return buff;
}


/////////////////////////HTML HEADER////////////////////////
short WriteHtmlHeader( FILE *stream2 )
////////////////writes HTML header of output///////////////////
{
	char filroot[100] = "";

	//open html file (will be eventually stdout)
    sprintf( filroot, "%s%s%s", drivweb, "data/ChaiTable.html", "\0");

  	/* open html document */
    if (!( stream2 = fopen( filroot, "w")) ) 	   
    {
		ErrorHandler(11, 1); //can't open file
		return -1; 
	} 

	fprintf(stream2, "%s\n", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
	fprintf(stream2, "%s\n", "<HTML>");
	fprintf(stream2, "%s\n", "<HEAD>");
	fprintf(stream2, "%s\n", "    <TITLE>Your Chai Table</TITLE>");
	fprintf(stream2, "%s\n", "    <META HTTP-EQUIV=\"CONTENT-TYPE\" CONTENT=\"text/html; charset=windows-1255\"/>");
	fprintf(stream2, "%s\n", "    <META NAME=\"AUTHOR\" CONTENT=\"Chaim Keller, Chai Tables, www.zemanim.org\"/>");
	fprintf(stream2, "%s\n", "</HEAD>");
	fprintf(stream2, "%s\n", "<BODY>");

	return 0;

}

/////////////////////////HTML closing/////////////////////////
short WriteHtmlClosing( FILE *stream2 )
//////////////////////wrties HTML closing lines//////////////
{
	fprintf(stream2, "%s\n", "</BODY>");
	fprintf(stream2, "%s\n", "</HTML>");

	if (!fclose(stream2)) 
	{
		ErrorHandler(12, 1); //can't close file
		return -1;
	}

	return 0;

}


/////////////////////netzski6////////////////////////////////////////
short netzski6(char *fileo, double kmxo, double kmyo, double hgt
			              , float geotd, short nsetflag, short nfil)
//////////////////////////////////////////////////////////////////////
/* netzski7.f -- translated by f2c (version 20021022).
   You must link the resulting object file with the libraries:
	-lf2c -lm   (in that order)
*/
//this is the subroutine version of netzski6
//outputs times for inputed place
///////////////////////////////////////////////////////////////////////
{

/****************local external subroutines****************/
    extern /* Subroutine */ short decl_(double *, double *, double *
	    , double *, double *, double *, double *, 
	    double *, double *, double *);

    extern /* Subroutine */ short casgeo_(double *, double *, 
	    double *, double *);

    extern /* Subroutine */ short refract_(double *, double *, 
	    double *, double *, int *, double *, double *,
	     int *, int *);


/************ System generated locals ******************/
    int i__1, i__2, i__3;
    double d__1;


/************* Local variables ******************/

    bool adhocset;
    int nweather;
    double elevintx, elevinty, d__, p, t, z__;
    int nplaceflg;
    bool adhocrise;
    double a1, e2, e3, e4, t3, t6, lg, ra, df;
    double ac, mc, ec, e2c, ob1, ap, mp, ob;

    int ne;
    double td, tf, es, lr, dy, lt, ms, et;
    int nn;
    double pt, sr, al1, dy1;
    int ns1, ns2, ns3, ns4;
    double sr1, sr2;
    int nac;
    double aas;
    double air, ref;
    double fdy, azi, eps;
    int nyd, nyf, nyl, nyr;
    double tss, al1o, alt1, azi1, azi2; //, chd62;

	//char CHD46[130] = "";
	char doclin[255] = "";
	double h1,h2,h3,h4,h5,h6,h7,h8;

    //double dobs[1464]	/* was [2][2][366] */;
    double hran;
    //double elev[4803]	/* was [3][1601] */;
	//double elev[3][1];
    double hgto;
    double azio;
    int nazn;
    //double tims[5124]	/* was [2][7][366] */;
    int nyri;
    double t3sub;

    double dobss, lnhgt;
    int nyear;
    double dobsy;
    //double azils[1464]	/* was [2][2][366] */;
    int niter, nloop;
    bool astro;

    int ndirec, nabove;
    double maxang;
    int nlpend, nfound;
    int nloops, ndystp;
    double refrac1;
    int nyltst, nyrstp, nyrtst;
    double stepsec, elevint;
    int nacurr;
	double tmp1,tmp2,tmp3;


	FILE *stream;

/**************************************************************/

/*       Version 16 (of netzski6) */

/*       Subroutine version of netzski6.  It is called by main program to */
/*       calculate a maximum of 7 zemanim for each place. */
/*       I.e., the mishor sunrise/sunset, the astronomical sunrise/sunset, */
/*       the visible sunrise/sunset, and chazos. */
/*       The results of the calculations are expressed as an array of */
/*       7 numbers vs day iteration corresponding to the days of the Hebrew year. */

/* ----------------------------------------------------- */
/*            = 20 * maxang + 1 ;number of azimuth points separated by .1 degrees */
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
//
    nweather = 3;
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
    p = 1013.;
/*                                  'mean atmospheric pressure (mb) */
    t = 27.;
/*                                  'mean atmospheric temperature (degrees Celsius) */
    stepsec = 2.5f;
/*                                  'step size in seconds */
/*       'N.B. astronomical constants are accurate to 10 seconds only to the year 2050 */
/*       '(Astronomical Almanac, 1996, p. C-24) */
/*       Addhoc fixes: (15 second fixes based on netz observations at Neveh Yaakov) */
    adhocrise = true;
    adhocset = false;
/* -------------------------------------------------------------------------------------- */

/*    geo = false;

/*       read in elevations from Profile file */
/*       which is a list of the view angle (corrected for refraction) */
/*       as a function of azimuth for the particular observation */
/*       point at kmxo,kmyo,hgt */

//    s_copy(fileo, "e:/cities/acco/netz/1134-aco.pro", (ftnlen)32, (ftnlen)32);
//    strncpy(fileo, "e:/cities/acco/skiy/1134-aco.pro", strlen("e:/cities/acco/skiy/1134-aco.pro"));

    astro = false;

	if (nsetflag != 2 && nsetflag != 3)
	{

    if ( stream = fopen( fileo, "r" ) )
		{
			fgets( doclin, 255, stream ); //read line of text, and then scan info
			fscanf(stream, "%lg,%lg,%lg,%lg,%lg,%lg,%lg,%lg\n", &h1,&h2,&h3,&h4,&h5,&h6,&h7,&h8);

			nazn = 0;
			while ( !feof(stream) )
			{
				if (nazn == 0 )
				{
					fscanf(stream,"%lg,%lg,%lg,%lg,%lg,%lg\n", &tmp1,&tmp2,&e2,&e3,&tmp3,&e4);

					maxang = (int)fabs(tmp1);
					
					//free the old elev memory if already allocated and if this if first place
					if (!nfil && elev) free( elev );
					//reallocate the memory for the elev array
					elev = (ELEVAT*) malloc((20 * (unsigned long int)maxang + 1) * sizeof(ELEVAT));
					if (! elev ) //error in allocation memory
					{
						return -3;
					}
					else  //stuff the 0 element with the values already read
					{
						elev[0].vert[0] = tmp1; //azimuth
						elev[0].vert[1] = tmp2; //view angle
						elev[0].vert[2] = tmp3; //distance (km) to obstruction
					}
				}
				else //keep on reading
				{
					fscanf(stream,"%lg,%lg,%lg,%lg,%lg,%lg\n", &elev[nazn].vert[0]
						                                     , &elev[nazn].vert[1]
															  ,&e2,&e3
															  ,&elev[nazn].vert[2]
															  ,&e4);
				}

				nazn++;
			}

			fclose(stream);
		}
		else
		{
			return -2; //goto L9000;
		}
	}
	else //astronomical/mishor times
	{
		astro = true; //no DTM file to read
	}

/* ----ad hoc summer and winter ranges for different latitudes----- */
//n60:
    if (geo) {
	lt = kmxo;
	lg = kmyo;
	if (fabs(lt) < latzone[0]) {
/*             tropical region, so use tropical atmosphere which */
/*             is approximated well by the summer mid-latitude */
/*             atmosphere for the entire year */
	    nweather = 0;
/*          define days of the year when winter and summer begin */
	} else if (fabs(lt) >= latzone[0] && fabs(lt) < latzone[1]) {
	    ns1 = sumwinzone[0][0];
	    ns2 = sumwinzone[0][1];
/*             no adhoc fixes  i.e., = 0,360 */
	    ns3 = sumwinzone[0][2];
	    ns4 = sumwinzone[0][3];
	} else if (fabs(lt) >= latzone[1] && fabs(lt) < latzone[2]) {
/*             weather similar to Eretz Israel */
	    ns1 = sumwinzone[1][0];
	    ns2 = sumwinzone[1][1];
/*             Times for ad-hoc fixes to the visible and astr. sunrise */
/*             (to fit observations of the winter netz in Neve Yaakov). */
/*             This should affect  sunrise and sunset equally. */
/*             However, sunset hasn't been observed, and since it would */
/*             make the sunset times later, it's best not to add it to */
/*             the sunset times as a chumrah. */
	    ns3 = sumwinzone[1][2];
	    ns4 = sumwinzone[1][3];
	} else if (fabs(lt) >= latzone[2] && fabs(lt) < latzone[3]) {
/*             the winter lasts two months longer than in Eretz Israel */
	    ns1 = sumwinzone[2][0];
	    ns2 = sumwinzone[2][1];
	    ns3 = sumwinzone[2][2];
	    ns4 = sumwinzone[2][3];
	} else if (fabs(lt) >= latzone[4]) {
/*             the winter lasts four months longer than in Eretz Israel */
	    ns1 = sumwinzone[3][0];
	    ns2 = sumwinzone[3][1];
	    ns3 = sumwinzone[3][2];
	    ns4 = sumwinzone[3][3];
	}
    } else {  //Eretz Yisroel
	casgeo_(&kmyo, &kmxo, &lt, &lg);
	ns1 = sumwinzone[4][0];
	ns2 = sumwinzone[4][1];
	ns3 = sumwinzone[4][2];
	ns4 = sumwinzone[4][3];
    }

    td = geotd;
/*       time difference from Greenwich England */
    if (nplaceflg == 2) {
	td = -8.; //Los Angeles
    }
    lr = cd * lt;
    tf = td * 60. + lg * 4.;

/* ----------------begin astronomical calculations--------------------------- */
    i__1 = SecondSecularYr;
    for (nyear = FirstSecularYr; nyear <= i__1; ++nyear) {
/* 	     define year iteration number */
	nyrstp = nyear - FirstSecularYr + 1;
	nyr = nyear;
/*       year for calculation */
	nyd = nyr - RefYear;

/* -------------------initialize astronomical constants--------------------------- */
	//constansts taken from p. C24 of 1996 Astronomical Almananc, accurate to 6 seconds
	//from 1950 to 2050 
	//the initial values are loaded in the function "LoadConstants"
	ac = ac0; //change in the mean anamolay per day 
    mc = mc0; //change in the mean longitude per day
    ec = ec0; //constants of the Ecliptic longitude
    e2c = e2c0;
    ob1 = ob10; //change in the obliquity of ecliptic per day
    ap = ap0; //mean anomaly for Jan 0, RefYear 00:00 at Greenwich meridian
    mp = mp0; //  mean longitude of sun Jan 0, RefYear 00:00 at Greenwich meridian 
    ob = ob0; //  ecliptic angle for Jan 0, 1996 00:00 at Greenwich meridian 
/*-----------------------------------------------------------------------------------*/

//  shift the ephemerials for the time zone
    ap -= td / 24. * ac;
    mp -= td / 24. * mc;
    ob -= td / 24. * ob1;

	nyf = 0;
	if (nyd < 0) {
	    i__2 = nyr;
	    for (nyri = RefYear - 1; nyri >= i__2; --nyri) {
		nyrtst = nyri;
		nyltst = 365;
		if (nyrtst % 4 == 0) {
		    nyltst = 366;
		}
		if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
		    nyltst = 365;
		}
		nyf -= nyltst;
		}
	} else if (nyd >= 0) {
	    i__2 = nyr - 1;
	    for (nyri = RefYear; nyri <= i__2; ++nyri) {
		nyrtst = nyri;
		nyltst = 365;
		if (nyrtst % 4 == 0) {
		    nyltst = 366;
		}
		if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
		    nyltst = 365;
		}
		nyf += nyltst;
		}
	}
	nyl = 365;
	if (nyr % 4 == 0) {
	    nyl = 366;
	}
	if (nyr % 100 == 0 && nyr % 400 != 0) {
	    nyl = 365;
	}
	nyf += -1462;
	ob += ob1 * nyf;
	mp += nyf * mc;

	while (mp < 0.)
	{
	    mp += pi2;
	}

	while (mp > pi2)
	{
	    mp -= pi2;
	}

	ap += nyf * ac;

	while (ap < 0.)
	{
	    ap += pi2;
	}

	while (ap > pi2)
	{
	    ap -= pi2;
	}

	d__1 = (double) yrend[nyear - FirstSecularYr];
	for (dy = (double) yrstrt[nyear - FirstSecularYr]; dy <= d__1; dy++) 
	{
/*          day iteration */
	    ndystp = (int)dy;
	    dy1 = dy;
/*          change ecliptic angle from noon of previous day to noon today */
	    ob += ob1;
	    decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
	    ra = atan(cos(ob) * tan(es));
	    if (ra < 0.) {
		ra += pi;
	    }
	    df = ms - ra;

		while ( fabs(df) > pi * .5 )
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

		et = df / cd * 4.;
/*          equation of time (minutes) */
	    t6 = tf + 720. - et;
/*          astronomical noon in minutes */
        //////////////record astronomical noon in hours/////////////
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

	    fdy = acos(-tan(d__) * tan(lr)) * ch / 24.;
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
	    if (hgt > 0.) {
		lnhgt = log(hgt * .001);
	    }
	    if (lt >= 0. && nweather == 3 && (dy <= (double) ns1 || dy >= 
		    (double) ns2) || nweather == 1 || lt < 0. && nweather 
		    == 3 && (dy > (double) ns1 && dy < (double) ns2)) 
		    {
/*              winter refraction */
		if (hgt <= 0.) {
		    goto L690;
		}
		ref = exp(winref[4] + winref[5] * lnhgt + winref[6] * lnhgt * 
			lnhgt + winref[7] * lnhgt * lnhgt * lnhgt);
		eps = exp(winref[0] + winref[1] * lnhgt + winref[2] * lnhgt * 
			lnhgt + winref[3] * lnhgt * lnhgt * lnhgt);
L690:	air = cd * 90. + (eps + ref + winrefo) / 1e3;
		a1 = (ref + winrefo) / (cd * 1e3);
	    } else if (lt >= 0. && nweather == 3 && (dy > (double) ns1 && 
		    dy < (double) ns2) || nweather == 0 || lt < 0. && 
		    nweather == 3 && (dy <= (double) ns1 || dy >= (
		    double) ns2)) {
/*              summer refraction */
		if (hgt <= 0.) {
		    goto L691;
		}
        ref = exp(sumref[4] + sumref[5] * lnhgt + sumref[6] * lnhgt *
			lnhgt	+ sumref[7] * lnhgt * lnhgt * lnhgt);
		eps = exp(sumref[0] + sumref[1] * lnhgt + sumref[2] * lnhgt * 
			lnhgt + sumref[3] * lnhgt * lnhgt * lnhgt);
L691:	air = cd * 90. + (eps + ref + sumrefo) / 1e3;
		a1 = (ref + sumrefo) / (cd * 1e3);
	    }
	    if (nsetflag % 2 != 0) {
/*              calculate astron./mishor sunsets */
/*              Nsetflag = 1 or Nsetflag = 3 OR Nsetflag = 5 */
		ndirec = -1;
		dy1 = dy - ndirec * fdy;
/*              calculate sunset/sunrise */
		decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
		air += (1 - cos(aas) * .017) * .2667 * cd;
		sr1 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * cos(
			d__))) * ch;
		sr = sr1 * hr;
/*              hour angle in minutes */
		t3 = (t6 - ndirec * sr) / hr;
		nacurr = 0;
/* *************************BEGIN SUNSET WINTER ADHOC FIX************************ */
		if (lt >= 0. && nweather == 3 && (dy <= (double) ns3 || 
			dy >= (double) ns4) && adhocset) {
/*                 add winter addhoc fix */
		    nacurr = nacurrFix; //15;
		}
/* *************************END SUNSET WINTER ADHOC FIX************************ */
		t3 += nacurr / 3600.;
/*               t3sub = t3 */
/*              astronomical sunrise/sunset in hours.frac hours */
		if (astro) {
		    azi1 = cos(d__) * sin(sr1 / ch);
		    azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(sr1 
			    / ch);
		    azio = atan(azi1 / azi2);
		    if (azio > 0.) {
			azi1 = 90. - azio / cd;
		    }
		    if (azio < 0.) {
			azi1 = -(azio / cd + 90.);
		    }
		}
//              record chazos 
//		tims[nyrstp - 1][6][ndystp - 1] = t6 / hr;
//              write astronomical sunset times to file and loop in days 
		if (astro || nsetflag == 5) //astro == true is same as nsetflag == 3
		{
 			if ( !ast )
			{
				//////////////record mishor sunset /////////////
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
				if (nfil == 0)
				{
					tims[nyrstp - 1][3][ndystp - 1] = t3;
				}
				else
				{
					if (t3 > tims[nyrstp - 1][3][ndystp - 1] )
						tims[nyrstp - 1][3][ndystp - 1] = t3;
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
/*             go back to calculate visible sunset */
			   goto L688;
		    }
		}
		else
		{
 
            //////////////record astro sunset /////////////
			if (nfil == 0)
			{
	            tims[nyrstp - 1][3][ndystp - 1] = t3;
			}
			else
			{
				if (t3 > tims[nyrstp - 1][3][ndystp - 1] )
					tims[nyrstp - 1][3][ndystp - 1] = t3;
			}
			////////////////////////////////////////////////
/*                  continue on to calculate vis. sunset */
		    nsetflag = 1;
		    nloop = 370;
		    if (nplaceflg == 0) {
			nloop = 75;
		    }
		    nlpend = -10;
		}
	    } else if (nsetflag % 2 == 0) {
/*              Nsetflag = 0 or Nsetflag = 2 OR Nsetflag = 4 */
/* 	         calculate astron./mishor sunrise */
		ndirec = 1;
		dy1 = dy + fdy;
		decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
		air += (1 - cos(aas) * .017) * .2667 * cd;
/*              calculate sunset */
		sr2 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * cos(
			d__))) * ch;
		sr2 *= hr;
		tss = (t6 + sr2) / hr;
/*              sunset in hours.fraction of hours */
		dy1 = dy - fdy;
/*              calculate sunrise */
		decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
		sr1 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * cos(
			d__))) * ch;
		sr = sr1 * hr;
/*              hour angle in minutes */
		t3 = (t6 - sr) / hr;
		nacurr = 0;
/* *************************BEGIN SUNRISE WINTER ADHOC FIX************************ */
		if (lt >= 0. && nweather == 3 && (dy <= (double) ns3 || 
			dy >= (double) ns4) && adhocrise) {
/*                 add winter addhoc fix */
		    nacurr = -nacurrFix; //-15;
		}
/* *************************END SUNRISE WINTER ADHOC FIX************************ */
		t3 += nacurr / 3600.;
/*               t3sub = t3 */
/*              astronomical sunrise in hours.frac hours */
/*              sr1= hour angle in hours at astronomical sunrise */
		if (astro) {
		    azi1 = cos(d__) * sin(sr1 / ch);
		    azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(sr1 
			    / ch);
		    azio = atan(azi1 / azi2);
		    if (azio > 0.) {
			azi1 = 90. - azio / cd;
		    }
		    if (azio < 0.) {
			azi1 = -(azio / cd + 90.);
		    }
		}
//              write astronomical sunrise time to file and loop in days 
//              record chazos 
//		tims[nyrstp - 1][6][ndystp - 1] = t6 / hr;

//              write astronomical sunrise times to file and loop in days 
		if (astro || nsetflag == 4) //astro == true is same as nsetflag == 2
		{

            if ( !ast )
			{
				//////////////record mishor sunrise /////////////
				if (nfil == 0)
				{
					tims[nyrstp - 1][4][ndystp - 1] = t3;
				}
				else
				{
					if (t3 < tims[nyrstp - 1][4][ndystp - 1] )
						tims[nyrstp - 1][4][ndystp - 1] = t3;
				}
				////////////////////////////////////////////////
			}
			else
			{
				//////////////record astro sunrise /////////////
				if (nfil == 0)
				{
					tims[nyrstp - 1][2][ndystp - 1] = t3;
				}
				else
				{
					if (t3 < tims[nyrstp - 1][2][ndystp - 1] )
						tims[nyrstp - 1][2][ndystp - 1] = t3;
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
/*             go back to calculate astro sunrise, vis sunrise */
			   goto L688;
		    }
		}
		else
		{

            //////////////record astro sunrise /////////////
			if (nfil == 0)
			{
	            tims[nyrstp - 1][2][ndystp - 1] = t3;
			}
			else
			{
				if (t3 < tims[nyrstp - 1][2][ndystp - 1] )
					tims[nyrstp - 1][2][ndystp - 1] = t3;
			}
			////////////////////////////////////////////////

/*                  continue on to calculate vis. sunrise */
		    nsetflag = 0;
		    nloop = 72;
/*                  start search 2 minutes above astronomical horizon */
		    nlpend = 2000;
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
	/*      step in hour angle by stepsec sec */
			dy1 = dy - ndirec * hran / 24.;
			decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
	/*      recalculate declination at each new hour angle */
			z__ = sin(lr) * sin(d__) + cos(lr) * cos(d__) * cos(hran / ch);
			z__ = acos(z__);
			alt1 = pi * .5 - z__;
	/*      altitude in radians of sun's center */
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
				//azimuthal range insufficient
				return -1; //goto L9050;
			}

			azi1 = ndirec * azi1;

			if (nsetflag == 1) 
			{
	/*          to first order, the top and bottom of the sun have the */
	/*          following altitudes near the sunset */
				al1 = alt1 / cd + a1;
	/*          first guess for apparent top of sun */
				pt = 1.;
				refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &
					ns1, &ns2);
	/*          first iteration to find position of upper limb of sun */
				al1 = alt1 / cd + .2666f + pt * refrac1;
	/*          top of sun with refraction */
				if (niter == 1) 
				{
					for (nac = 1; nac <= 10; ++nac)
					{
						al1o = al1;
						refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &ns1, &ns2);
	/*                  subsequent iterations */
						al1 = alt1 / cd + .2666f + pt * refrac1;
	/*                  top of sun with refraction */
						if (fabs(al1 - al1o) < .004f) break;
	/* L693: */
					}
				}
	//L695:
	/*          now search digital elevation data */
				ne = (int) ((azi1 + maxang) * 10);
				elevinty = elev[ne + 1].vert[1] - elev[ne].vert[1];
				elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
				elevint = elevinty / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[1];
				if ((elevint >= al1 || nloops <= 0) && nabove == 1) 
				{
	/*              reached sunset */
					nfound = 1;
	/*              margin of error */
					nacurr = 0;
					t3sub = t3 + ndirec * (stepsec * nloops + nacurr) / 3600.;
					if (nloops < 0)
					{
		/*              visible sunset can't be latter then astronomical sunset! */
						t3sub = t3;
					}
	/*                 <<<<<<<<<distance to obstruction>>>>>> */
					dobsy = elev[ne + 1].vert[2] - elev[ne].vert[2];
					elevintx = elev[ne + 1].vert[0] - elev[ne].vert[0];
					dobss = dobsy / elevintx * (azi1 - elev[ne].vert[0]) + elev[ne].vert[2];
	/*              record visual sunset results for time,azimuth,dist. to obstructions */

            		/////only record times if they are earlier (sunrise) or latter (sunset)
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

/*              finished for this day, loop to next day */
				//goto L925;
				}
				else if (elevint < al1)
				{
					nabove = 1;
				}
			}
			else if (nsetflag == 0) 
			{
				al1 = alt1 / cd + a1;
/*              first guess for apparent top of sun */
				pt = 1.;
				refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &ns1, &ns2);
/*              first iteration to find top of sun */
				al1 = alt1 / cd + .2666f + pt * refrac1;
/*              top of sun with refraction */
				if (niter == 1)
				{
					for (nac = 1; nac <= 10; ++nac)
					{
						al1o = al1;
						refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &ns1, &ns2);
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
				ne = (int) ((azi1 + maxang) * 10);
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
						t3sub = t3 + (stepsec * nloops + nacurr) / 3600.;
						if (nloops < 0)
						{
/*                         sunrise can't be before astronomical sunrise */
/*                         so set it equal to astronomical sunrise time */
/*                         minus the adhoc winter fix if any */
	                       t3sub = t3;
						}

/*                      record visual sunrise results for time,azimuth,dist. to obstructions */
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
	/* L850: */
		}
		}
/* L900: */

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
				goto L692;
			}
			else if (nloop >= -100 && nsetflag == 0)
			{
			   if (nsetflag == 0)
			   {
				  nloop += -24;
	/*            IF (nloop .LT. 0) nloop = 0 */
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
		}

//L925: //next day
	    ;
	}
/* L975: */
	//next year
    }

	if (elev)
	{
		free( elev ); //free elev memory
		elev = 0;
	}

    return 0;
} /* MAIN__ */

/* Subroutine */ short decl_(double *dy1, double *mp, double *mc, 
	double *ap, double *ac, double *ms, double *aas, 
	double *es, double *ob, double *d__)
{
    // Builtin functions 
    //double acos(double), sin(double), asin(double);

    /* Local variables */
    

	double ec, e2c;

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
    *d__ = asin(sin(*ob) * sin(*es));
    return 0;
} /* decl_ */

/* Subroutine */ short refwin1_(double *hgt, double *x, double *
	frwin, double *p, double *t)
{
    double x1;

/*       refwin1:  'calculated from MENAT.FOR/FITP using x = view angle */
    if (*hgt >= 1150.) {
	x1 = *hgt * -1.4893e-5f + .16878f;
	*frwin = *p / (*t + 273) * (x1 + *x * .0148f + *x * 1e-4f * *x);
	*frwin /= *x * .47806f + 1 + *x * .07447f * *x;
    } else if (*hgt >= 1050. && *hgt < 1150.) {
	x1 = *hgt * -1.5118e-5f + .16901f;
	*frwin = *p / (*t + 273) * (x1 + *x * .01575f + *x * 7e-5f * *x);
	*frwin /= *x * .4834f + 1 + *x * .07641f * *x;
    } else if (*hgt >= 950. && *hgt < 1050.) {
	x1 = *hgt * -1.5058e-5f + .16896f;
	*frwin = *p / (*t + 273) * (x1 + *x * .02169f - *x * 1.4e-4f * *x);
	*frwin /= *x * .52128f + 1 + *x * .09025f * *x;
    } else if (*hgt >= 850. && *hgt < 950.) {
	x1 = *hgt * -1.445e-5f + .16838f;
	*frwin = *p / (*t + 273) * (x1 + *x * .02111f - *x * 1.2e-4f * *x);
	*frwin /= *x * .5161f + 1 + *x * .08825f * *x;
    } else if (*hgt >= 750. && *hgt < 850.) {
	x1 = *hgt * -1.4876e-5f + .16876f;
	*frwin = *p / (*t + 273) * (x1 + *x * .0232f - *x * 1.9e-4f * *x);
	*frwin /= *x * .52786f + 1 + *x * .09266f * *x;
    } else if (*hgt >= 650. && *hgt < 750.) {
	x1 = *hgt * -1.5295e-5f + .16907f;
	*frwin = *p / (*t + 273) * (x1 + *x * .02383f - *x * 2.1e-4f * *x);
	*frwin /= *x * .5303f + 1 + *x * .09358f * *x;
    } else if (*hgt >= 550. && *hgt < 650.) {
	x1 = *hgt * -1.5256e-5f + .1691f;
	*frwin = *p / (*t + 273) * (x1 + *x * .02976f - *x * 4.2e-4f * *x);
	*frwin /= *x * .56622f + 1 + *x * .10671f * *x;
    } else if (*hgt >= 450. && *hgt < 550.) {
	x1 = *hgt * -1.6268e-5f + .1696f;
	*frwin = *p / (*t + 273) * (x1 + *x * .02905f - *x * 3.9e-4f * *x);
	*frwin /= *x * .56005f + 1 + *x * .10422f * *x;
    } else if (*hgt >= 350. && *hgt < 450.) {
	x1 = *hgt * -1.5319e-5f + .1691f;
	*frwin = *p / (*t + 273) * (x1 + *x * .01416f + *x * 1.1e-4f * *x);
	*frwin /= *x * .46823f + 1 + *x * .06829f * *x;
    } else if (*hgt >= 250. && *hgt < 350.) {
	x1 = *hgt * -1.5492e-5f + .16916f;
	*frwin = *p / (*t + 273) * (x1 + *x * .01168f + *x * 2e-4f * *x);
	*frwin /= *x * .45241f + 1 + *x * .06206f * *x;
    } else if (*hgt >= 150. && *hgt < 250.) {
	x1 = *hgt * -1.6191e-5f + .16941f;
	*frwin = *p / (*t + 273) * (x1 - *x * .01107f + *x * 8.3e-4f * *x);
	*frwin /= *x * .32072f + 1 + *x * .0059f * *x;
    } else if (*hgt >= 50. && *hgt < 150.) {
	x1 = *hgt * -1.5257e-5f + .16924f;
	*frwin = *p / (*t + 273) * (x1 - *x * .00922f + *x * 8e-4f * *x);
	*frwin /= *x * .33125f + 1 + *x * .01087f * *x;
    } else if (*hgt >= -50. && *hgt < 50.) {
	x1 = *hgt * -1.5938e-5f + .16872f;
	*frwin = *p / (*t + 273) * (x1 + *x * .0305f - *x * 2.1e-4f * *x);
	*frwin /= *x * .55065f + 1 + *x * .10908f * *x;
    } else if (*hgt >= -150. && *hgt < -50.) {
	x1 = *hgt * -1.6064e-5f + .16872f;
	*frwin = *p / (*t + 273) * (x1 + *x * .03083f - *x * 2.2e-4f * *x);
	*frwin /= *x * .55083f + 1 + *x * .10908f * *x;
    } else if (*hgt < -150.) {
	x1 = *hgt * -1.6195e-5f + .1687f;
	*frwin = *p / (*t + 273) * (x1 + *x * .03119f - *x * 2.2e-4f * *x);
	*frwin /= *x * .55118f + 1 + *x * .10915f * *x;
    }
    return 0;
} /* refwin1_ */

/* Subroutine */ short refsum1_(double *hgt, double *x, double *
	frsum, double *p, double *t)
{
    double x1;

/*       refsum1:  'calculated from MENAT.FOR/FITP using x = view angle */
    if (*hgt >= 1150.) {
	x1 = *hgt * -1.2745e-5f + .15274f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01047f + *x * 1.7e-4f * *x);
	*frsum /= *x * .43732f + 1 + *x * .06206f * *x;
    } else if (*hgt >= 1050. && *hgt < 1150.) {
	x1 = *hgt * -1.232e-5f + .15225f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01207f + *x * 1.1e-4f * *x);
	*frsum /= *x * .44836f + 1 + *x * .06573f * *x;
    } else if (*hgt >= 950. && *hgt < 1050.) {
	x1 = *hgt * -1.2599e-5f + .15256f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01589f - *x * 3e-5f * *x);
	*frsum /= *x * .47504f + 1 + *x * .07487f * *x;
    } else if (*hgt >= 850. && *hgt < 950.) {
	x1 = *hgt * -1.2876e-5f + .15279f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01631f - *x * 4e-5f * *x);
	*frsum /= *x * .47688f + 1 + *x * .07551f * *x;
    } else if (*hgt >= 750. && *hgt < 850.) {
	x1 = *hgt * -1.2262e-5f + .15229f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02014f - *x * 1.7e-4f * *x);
	*frsum /= *x * .50277f + 1 + *x * .08453f * *x;
    } else if (*hgt >= 650. && *hgt < 750.) {
	x1 = *hgt * -1.2343e-5f + .15239f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02058f - *x * 1.8e-4f * *x);
	*frsum /= *x * .50458f + 1 + *x * .0851f * *x;
    } else if (*hgt >= 550. && *hgt < 650.) {
	x1 = *hgt * -1.3012e-5f + .15286f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02555f - *x * 3.6e-4f * *x);
	*frsum /= *x * .53764f + 1 + *x * .09663f * *x;
    } else if (*hgt >= 450. && *hgt < 550.) {
	x1 = *hgt * -1.2399e-5f + .15248f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02054f - *x * 1.8e-4f * *x);
	*frsum /= *x * .50165f + 1 + *x * .08389f * *x;
    } else if (*hgt >= 350. && *hgt < 450.) {
	x1 = *hgt * -1.2714e-5f + .15264f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01256f + *x * 9e-5f * *x);
	*frsum /= *x * .44703f + 1 + *x * .06355f * *x;
    } else if (*hgt >= 250. && *hgt < 350.) {
	x1 = *hgt * -1.332e-5f + .15287f;
	*frsum = *p / (*t + 273) * (x1 + *x * .01515f);
	*frsum /= *x * .50458f + 1 + *x * .0851f * *x;
    } else if (*hgt >= 150. && *hgt < 250.) {
	x1 = *hgt * -1.2843e-5f + .15277f;
	*frsum = *p / (*t + 273) * (x1 - *x * .0075f + *x * 6.7e-4f * *x);
	*frsum /= *x * .31655f + 1 + *x * .01201f * *x;
    } else if (*hgt >= 50. && *hgt < 150.) {
	x1 = *hgt * -1.2951e-5f + .15281f;
	*frsum = *p / (*t + 273) * (x1 - *x * .00838f + *x * 7e-4f * *x);
	*frsum /= *x * .3113f + 1 + *x * .00986f * *x;
    } else if (*hgt >= -50. && *hgt < 50.) {
	x1 = *hgt * -1.3233e-5f + .15237f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02692f - *x * 2.1e-4f * *x);
	*frsum /= *x * .52849f + 1 + *x * .10086f * *x;
    } else if (*hgt >= -150. && *hgt < -50.) {
	x1 = *hgt * -1.3322e-5f + .15236f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02718f - *x * 2.1e-4f * *x);
	*frsum /= *x * .52854f + 1 + *x * .10083f * *x;
    } else if (*hgt < -150.) {
	x1 = *hgt * -1.3424e-5f + .15234f;
	*frsum = *p / (*t + 273) * (x1 + *x * .02744f - *x * 2.2e-4f * *x);
	*frsum /= *x * .52857f + 1 + *x * .10079f * *x;
    }
    return 0;
} /* refsum1_ */

/* Subroutine */ short refract_(double *al1, double *p, double *t, 
	double *hgt, int *nweather, double *dy, double *
	refrac1, int *ns1, int *ns2)
{
    double x, frwin, frsum;
    extern /* Subroutine */ short refwin1_(double *, double *, 
	    double *, double *, double *), refsum1_(double *, 
	    double *, double *, double *, double *);

    if (*nweather == 0) {
	x = *al1;
	refsum1_(hgt, &x, &frsum, p, t);
	*refrac1 = frsum;
    } else if (*nweather == 1) {
	x = *al1;
	refwin1_(hgt, &x, &frwin, p, t);
	*refrac1 = frwin;
/*       ELSE IF (nweather .EQ. 2) THEN */
/*          refrac1 = FNref(al1) */
    } else if (*nweather == 3) {
	x = *al1;
	if (*dy <= (double) (*ns1) || *dy >= (double) (*ns2)) {
	    refwin1_(hgt, &x, &frwin, p, t);
	    *refrac1 = frwin;
	} else if (*dy > (double) (*ns1) && *dy < (double) (*ns2)) {
	    refsum1_(hgt, &x, &frsum, p, t);
	    *refrac1 = frsum;
	}
    }
    return 0;
} /* refract_ */

/* Subroutine */ short casgeo_(double *gn, double *ge, double *l1, 
	double *l2)
{
    /* System generated locals */
    double d__1;

    /* Builtin functions */
    //double sin(double), sqrt(double), tan(double), cos(double)
	    

    /* Local variables */
    double r__, a2, b2, c1, c2, f1, g1, g2, e4, c3, d1, a3, o1, o2,
	     a4, b3, s1, s2, b4, c4, c5, x1, y1, y2, x2, c6, d2, r3, r4, r2, 
	    a5, d3;

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
    *l1 = r__ * *l1;
    a5 = sqrt(a2) * c3;
    d3 = x1 * a5 / cos(r2);
/*       LON */
    *l2 = r__ * ((s1 + d3) / f1);
/*       THIS IS THE EASTERN HEMISPHERE! */
    *l2 = -(*l2);
    return 0;
} /* casgeo_ */


/* Main program alias  int netzski7_ () { MAIN__ (); return 0; } */
