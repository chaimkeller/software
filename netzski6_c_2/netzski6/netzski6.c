/* netzski6_ver17.f -- translated by f2c (version 20021022).
   You must link the resulting object file with the libraries:
	-lf2c -lm   (in that order)
*/
//////////////////////////////////////////////////////////////////
//			Version 20 02/15/2021
///////////////////////////////////////////////////////////////////
//added: ability to read external temperature, pressure file
//then to sutiably modify the refraction based on those values
//the refraction scaling rules are now a function of height and view angle
///////////////////////////////////////////////////////////////

//uncomment header files for getch debugging
#include <math.h>
#include <stdio.h>
#include <string.h>
#include <ctype.h>
//#include <conio.h> //uncomment for getch() debugging
#include "f2c.h"
#include <conio.h>
#include <malloc.h>

/* Common Block Declarations */

struct {
    char t3subc[8], caldayc[11];
} t3s_;

#define t3s_1 t3s_
#define MAXDBLSIZE 100 //maximum character size of data elements to be read using ParseString

/* Table of constant values */

static integer c__1 = 1;
static integer c__3 = 3;
static integer c__5 = 5;
static integer c__0 = 0;
static integer c__9 = 9;
static integer c__4 = 4;
static doublereal ErrorFudge = 0; //0.59;
static doublereal c_b175 = 1.687; //2.2727; //1.68271708343173; //1.687; //1.7081; //1.686701538132792; //1.7081; //<-------
static doublereal c_b176 = .69;
static doublereal c_b177 = 73.;
static doublereal c_b178 = 9.56267125268496f; //9.572702286470884f; //9.56267125268496f; <------
static doublereal c_b179 = 2.18; //exponent for ref  (need expln why it turned out different than c_b075 ????) <-----
								 //but makes dif for Jerusalem astronomical is no more than 2 seconds.
static doublereal c_b180 = 0.5; //reduced vdw exponent for astronomical altitudes
static doublereal c_b181 = -0.2; //exponent for eps <----
static doublereal c_press = 1.0856; //vdw pressure exponent for view angle = 0 degrees

short Temperatures(double lt, double lg, integer MinTemp[], integer AvgTemp[], integer MaxTemp[] );
short lenMonth( short x);

////////////functions that emulate MS VB 6.0 functions//////////
int InStr( short nstart, char * str, char * str2 );
int InStr( char * str, char * str2 );

short ParseString( char *doclin, char *sep, char arr[][100] );
char *fgets_CR( char *str, short strsize, FILE *stream );
short Cint( double x );

char strarr[4][100];
double cd, pi;

logical FixProfHgtBug = FALSE_; //added 101520 to fix the bug in newreadDTM that recorded the 
						   //ground height and not the calculation height with the added observer height
						   //030320 -- found that bat file is most reliable recording to height + observer height
						   //          and even if it errs, it errs to lower height which is better


/* Main program */ int MAIN__(void)
{
    /* Format strings */
    //static char fmt_710[] = "(a11,3x,a8,2x,a12,3x,f7.3,3x,f6.2)"
    static char fmt_710[] = "(a11,3x,a8,2x,a12,3x,f7.3,3x,f8.4,3x,f6.2)"; //081021 added ,3x,f8.4 for viewangle
    //static char fmt_785[] = "(a11,2x,a8,3x,a12,3x,f7.3,3x,f6.2)"; 
    static char fmt_785[] = "(a11,2x,a8,3x,a12,3x,f7.3,3x,f8.4,3x,f6.2)"; //same
    static char fmt_700[] = "(1x,a11,3x,a8,3x,a8,3x,a9,3x,i1)";
    static char fmt_705[] = "(a11,3x,a8,3x,a8,3x,a9)";
    static char fmt_770[] = "(1x,a11,3x,a8,3x,a8,3x,a9,3x,i1)";
    static char fmt_775[] = "(a11,3x,a8,3x,a8,3x,a9)";
    //static char fmt_1305[] = "(a11,3x,a8,6x,a12,3x,f7.3,3x,f6.2)";
    static char fmt_1305[] = "(a11,3x,a8,6x,a12,3x,f7.3,3x,f8.4,3x,f6.2)"; //same

    /* System generated locals */
    integer i__1, i__2, i__3, i__4;
    real r__1;
    doublereal d__1, d__2, d__3, d__4;
    olist o__1;
    cllist cl__1;
    inlist ioin__1;

    /* Builtin functions */
    double acos(doublereal);
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer f_inqu(inlist *), f_open(olist *), s_rsfe(cilist *), do_fio(
	    integer *, char *, ftnlen), e_rsfe(void), f_clos(cllist *), 
	    s_rsle(cilist *), do_lio(integer *, integer *, char *, ftnlen), 
	    e_rsle(void), s_wsle(cilist *), e_wsle(void);
    /* Subroutine */ int s_paus(char *, ftnlen);
    integer s_wsfi(icilist *), e_wsfi(void), i_indx(char *, char *, ftnlen, 
	    ftnlen);
    double sqrt(doublereal), pow_dd(doublereal *, doublereal *);
    integer i_dnnt(doublereal *);
    double d_mod(doublereal *, doublereal *);
    integer s_wsfe(cilist *), e_wsfe(void);
    double cos(doublereal), tan(doublereal), atan(doublereal), log(doublereal)
	    , exp(doublereal), sin(doublereal);

    /* Local variables */
    static integer nstrtyr;
    static logical adhocset;
    static integer nsetflag;
    static doublereal meantemp;
    static integer nweather;
    static doublereal trrbasis;
    static char comentss[12];
    static doublereal elevintx, elevinty, exponent;
    static char testdriv[8];
    static doublereal d__;
    static integer i__;
    static doublereal k, m, p, t, z__;
    static integer nplaceflg;
    static logical adhocrise;
    static doublereal a1, e2;
    static char f2[27];
    static doublereal e3;
    static integer i1, i2;
    static doublereal t3, t6;
    static integer nopendruk;
    static doublereal ac, ec, df, ch, mc, ap, ob, lg, ra;
    static integer ne;
    static doublereal td, hr, tf, es, mp, lr, dy, lt;
    static integer mint[12], avgt[12], maxt[12];
    static doublereal tk, tkmin, tkmax, tkavg, ms, et, sr;
    static integer nn;
    static doublereal pt, yr, e2c;
    static char ch1[1], ch2[2], ch3[3], ch4[4];
    static doublereal al1, ob1;
    static char cn4[4];
    static doublereal pi2, pathlength, dy1;
    static integer ns1, ns2, ns3, ns4, pos; //, np1, np2, np3;
    static doublereal sr1, sr2;
    static logical nointernet;
    static integer nac;
    static doublereal aas;
    static char dat[11];
    extern /* Subroutine */ int t1500_(doublereal *, integer *);
    static logical geo;
    static doublereal air, ref;
    extern /* Subroutine */ int t1750_(doublereal *, integer *, integer *);
    static doublereal hgt, fdy, azi;
    static integer itk;
    static doublereal eps;
    static integer nyd, nyf, nyl, nyr;
    static doublereal tss, al1o, alt1, azi1, azi2;
    static char ts1c[8], ts2c[9], ext1[4*2*100], ext3[4*2*100];
	char buff[50] = "";
    static integer nyr1;
    static doublereal defb;
    static char chd62[62];
    extern /* Subroutine */ int decl_(doublereal *, doublereal *, doublereal *
	    , doublereal *, doublereal *, doublereal *, doublereal *, 
	    doublereal *, doublereal *, doublereal *);
    static char chd46[46];
	static char chd48[48];
    static integer nadd;
    static doublereal defm;
    static char name__[8];
    static doublereal dobs;
    static integer nfil;
    static doublereal hran;
    //static real elev[6404]	/* was [4][1601] */;
	real *elev;
    static integer ndtm;
    static doublereal dist, azio;
    static integer nazn;
    static doublereal stat;
    static char tssc[8];
    static doublereal kmxo, kmyo;
    static integer nyri, yrst[2], nset1, nset2;
    static doublereal stat2, t3sub;
    static char place[8*2];
    static integer negch;
    static doublereal deltd;
    static char placn[30*2*100];
    static integer nplac[2];
    static char fileo[27];
    static doublereal dobsf, geotd, avref,viewangf; //////viewangf added 081021 ///////////////
    static char namet[23];
    static doublereal distd, dobss[1464]	/* was [2][2][366] */, lnhgt;
	/////////////added 081021 to keep track of viewangle at sunries/sunset/////////////
	static doublereal viewang[1464], viewangle;
	/////////////////////////////////////////////////////////////////
    static integer nyear;
    static doublereal dobsy;
    static integer yrend[2];
    static doublereal azils[1464]	/* was [2][2][366] */;
    static integer niter, ndruk, nloop;
    static logical astro;
    static integer nstat;
    static doublereal vdwsf, vdwalt, vdwrefr, distx, vbwast, vbweps;
    static char timss[8];
    static doublereal tsscs[1464]	/* was [2][2][366] */;
    static integer nyear1;
    static doublereal dcmrad;
    extern /* Subroutine */ int casgeo_(doublereal *, doublereal *, 
	    doublereal *, doublereal *);
    static integer ndirec;
    static doublereal hgtdif;
    static char plcfil[27];
    static integer nabove;
    static doublereal maxang, maxang0;
    static integer nlpend, nmenat;
    static char coment[8];
    static doublereal rearth;
    static integer nyrheb, nfound, numfil;
    static char filtmp[18], tmpfil[18];
    static doublereal vdwref[6];
    extern /* Subroutine */ int refvdw_(doublereal *, doublereal *, 
	    doublereal *, const short *, const double *);
    static integer nendyr;
	static doublereal nacurr;
    static char filout[27];
    static real azinss;
    static doublereal vdweps[5];
    static integer nloops;
    static char yesint[22];
    static logical exists;
    static integer ndystp;
    static doublereal refrac1;
    static integer nyltst, nyrstp, nyrtst;
    static char locname[40];
    static integer intflag, nplaces, ntrcalc;
	static doublereal Tground;
    static doublereal hgtdifx;
    static logical nofiles;
    static char filstat[18];
    static doublereal stepsec;
    static char coments[12*2*2*366], drivlet[1];
    static doublereal trbasis, elevint;
    static integer nleapyr, nzskflg;
	static doublereal hgt0;
	static doublereal hgtobs = 1.8; //maximum height of the observer
	double check = 0;


	double MaxMinTemp = -999999;
	short WinTemp = 0;
	short DayCount = 0;
	short iTemp,kTemp,mTemp = 0;
	double DayLength = 0.0;
	double MaxWinLen = 11.1;
	char fnam[255] = "";
	short nweatherflag = 0;
	short adhocflag = 0;
	logical FindWinter = TRUE_; //version 18 way of determining the winter months
	logical EnableSunriseInv = FALSE_; //flag to subtract time for suspected inversion at this day
	logical EnableSunsetInv = FALSE_; //flag to subtract time for suspected inversion at this day

	logical ReadTemps = TRUE_;  //true to read from netzskiy.tme, false to read directly from WorldClim binary files
	doublereal AddTemp = 0;
	doublereal TRfudge = 1.0; //0.95;//increasing positive lapse rate will lower the TR as well as the total atm refraction
	doublereal TRfudgeVis = 1.0;
	doublereal ViewAdd = 0.0;
	doublereal HorizonInv = 0.01;
	doublereal WinterPercent = 1.0;
	static doublereal altmin;
	static doublereal vbwexp;
	static char fndynum[255] = "";
	static char doclin[255] = "";
	static logical flagdynum = FALSE_;	//used for special case of Edmonton observations
	static logical showCalc = FALSE_;  //diagnostic variable to show details of refraction iterations
	static short yrtst;
	static short dynum; //used for Edmonton observations
	static double Tobs; //uwed for Edmonton observations, but can be extended to user inputed ground temperature
	static double Pobs; //used for Edmonton observations, but can be extended to user inputed pressure
	//more variables to compare details of refraction iterations
	static doublereal vdwsf_2;
	static doublereal al1_2;
	static doublereal refrac1_2;
	static doublereal al1o_2;
	static doublereal alt1_2;
	static integer nloops_2;
	static logical Diagnostics = FALSE_;

	//static short DTMType = -1;
	static integer distlim = 5;
	static integer AddCushions = 0;
	static doublereal AddedTime = 0;
	static doublereal accobs = 0;
	static integer CushionAmount = 0;

	int ErrorCode = 0;  //EK added 061922

	//doublereal Min_Temp_adjust = 0.9;//range (1 - 0) adjust minimum temperature to be average minimum temperature
	//for the total atmospheric refraction, use pt
	short ier = 0;

	FILE *stream;

    /* Fortran I/O blocks */
    static cilist io___18 = { 0, 1, 0, "(I1)", 0 };
    static cilist io___20 = { 0, 10, 0, 0, 0 };
    static cilist io___22 = { 0, 10, 0, 0, 0 };
    static cilist io___25 = { 0, 10, 0, 0, 0 };
    static cilist io___27 = { 0, 10, 0, "(A)", 0 };
    static cilist io___29 = { 0, 10, 0, "(A)", 0 };
    static cilist io___35 = { 0, 1, 0, 0, 0 };
    static cilist io___38 = { 0, 10, 0, "(A)", 0 };
    static cilist io___40 = { 0, 10, 0, 0, 0 };
    static cilist io___49 = { 0, 10, 0, 0, 0 };
    static cilist io___51 = { 0, 10, 0, 0, 0 };
    static cilist io___55 = { 0, 6, 0, 0, 0 };
    static icilist io___58 = { 0, ch1, 0, "(I1)", 1, 1 };
    static icilist io___61 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___63 = { 0, ch3, 0, "(I3)", 3, 1 };
    static cilist io___65 = { 0, 6, 0, 0, 0 };
    static icilist io___66 = { 0, ch1, 0, "(I1)", 1, 1 };
    static icilist io___67 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___68 = { 0, ch3, 0, "(I3)", 3, 1 };
    static cilist io___90 = { 0, 6, 0, 0, 0 };
    static cilist io___93 = { 0, 2, 0, "(A)", 0 };
    static cilist io___98 = { 0, 2, 0, "(A)", 0 };
    static cilist io___101 = { 0, 2, 1, 0, 0 };
    static icilist io___121 = { 0, ch4, 0, "(I4)", 4, 1 };
    static icilist io___151 = { 0, ch1, 0, "(I1)", 1, 1 };
    static icilist io___152 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___153 = { 0, ch3, 0, "(I3)", 3, 1 };
    static icilist io___154 = { 0, ch4, 0, "(I4)", 4, 1 };
    static cilist io___155 = { 0, 3, 0, "(F8.4)", 0 };
    static cilist io___183 = { 0, 6, 0, 0, 0 };
    static cilist io___187 = { 0, 6, 0, "(1X,A11,3X,A8)", 0 };
    static cilist io___188 = { 0, 1, 0, "(A11,3X,A8)", 0 };
    static cilist io___190 = { 0, 2, 0, fmt_710, 0 };
    static cilist io___193 = { 0, 6, 0, 0, 0 };
    static cilist io___195 = { 0, 6, 0, "(1X,A11,3X,A8)", 0 };
    static cilist io___196 = { 0, 1, 0, "(A11,3X,A8)", 0 };
    static cilist io___197 = { 0, 2, 0, fmt_785, 0 };
    static cilist io___206 = { 0, 6, 0, 0, 0 };
    static cilist io___207 = { 0, 6, 0, 0, 0 };
    static cilist io___208 = { 0, 6, 0, 0, 0 };
    static cilist io___219 = { 0, 4, 0, 0, 0 };
    static cilist io___222 = { 0, 6, 0, 0, 0 };
    static cilist io___223 = { 0, 6, 0, fmt_700, 0 };
    static cilist io___224 = { 0, 1, 0, fmt_705, 0 };
    static cilist io___225 = { 0, 2, 0, fmt_710, 0 };
    static cilist io___226 = { 0, 6, 0, 0, 0 };
    static cilist io___227 = { 0, 4, 0, 0, 0 };
    static cilist io___228 = { 0, 6, 0, fmt_770, 0 };
    static cilist io___229 = { 0, 1, 0, fmt_775, 0 };
    static cilist io___230 = { 0, 2, 0, fmt_785, 0 };
    static cilist io___231 = { 0, 6, 0, 0, 0 };
    static cilist io___232 = { 0, 6, 0, 0, 0 };
    static cilist io___233 = { 0, 6, 0, 0, 0 };
    static icilist io___243 = { 0, ch4, 0, "(I4)", 4, 1 };
    static cilist io___251 = { 0, 1, 0, fmt_1305, 0 };
    static cilist io___252 = { 0, 3, 0, "(I1)", 0 };
    static cilist io___253 = { 0, 3, 0, "(I2)", 0 };
    static icilist io___254 = { 0, ch4, 0, "(I4)", 4, 1 };
    static cilist io___255 = { 0, 3, 0, "(A1,A27,A1)", 0 };
    static cilist io___256 = { 0, 3, 0, "(A1,A27,A1)", 0 };


/*       Version 21 //20 //19 //18 //17 */
/*       this program calculates sunrise/sunset tables for lattest format */
/*       it is called by the WINDOWS 95 program CAL PROGRAM */
/*       the output filenames reside on the file NETZSKIY.tm3 */
/*       This version allows for arbitrary number of decimal places */
/*       on coordinates.  It adds an adhoc fix for Eretz Yisroel's */
/*       sunrise and sunet during two months of the winter (based */
/*       on observations at Neveh Yaakov. The adhoc fixes are controlled */
/*       by the logical variables ADHOCRISE (sunrise), ADHOCSET (sunset) */

/*       -------Version Info------------ */
/*       Version 14 only calculates the days that are included in the Hebrew Year */

/*       Version 15 doesn't use COMPARE, rather compares the times in memory */
/*       and doesn't write itermediate time files for each neighborhood.  Version 15 */
/*       only supports the calculation of one Hebrew Year at a time. */
/*       Also permits turning off writing any intermediate neighborhood files */

/* 		Version 16 adds temperature scaling for Menat atmospheres for a proof of principle test */
/* 		A latter version will use van der Werf's atmospheric model for the entire world simply */
/* 		modified by the appropriate average ground temperature for either sunrise or sunset supplied */
/* 		via the netzskiy.tm3 file either calculated by the Cal Program and based either on the WorldClim model */
/* 		or provided by user input for a certain day whenever that option becomes viable. */
/* 	    For this version, minimum and average temperatures for each month are based on the WorldClim */
/* 		model for the profile file's coordinates and are calculated by Cal Program and are written to a */
/* 		new line in the netzskiy.tm3 file. */
/* 	    For this version, only Menat's winter atmosphere is used modified by temperature */
/* 		scaling laws as determined by Van der Werf's 2003 paper for his own ray tracing calculations. */
/* 		So even though they don't apply to Menat's modeling, there are used as a proof of principle test. */
/* 		Both the total atmospheric refraction and refraction at a certain view anglea and height are */
/* 		according to the scaling laws that Van der Werf published in his paper. */

/* 		From version 16, netzski6 is now given the task of calculating terrestrial refraction, TR, */
/* 		using the modified Wikipedia expression for both sunrise and sunset. */
/* 		I.e., TR (deg.) = 0.0083 * Pathlength (km)**Exponent * K * Ground Pressure (mb)/ (Ground Temperature (K) 
** 2). */
/* 		The Ground Temperature is modeled to be different at sunrise and at sunset. */
/* 		For this version, ground temperature at sunrise is apprximated by the minimum monthly WorldClim */
/* 		temperature for the profile file's coordinates, and the ground temperature at sunset is approximated */
/* 		by the WorldClim average monthly temperature for the profile file's coordinates. This will be tweeked in 
*/
/* 		minor version updates to improve the fit to Rabbi Druk's measurements. */
/* 		Both temepratures are determined for the coordinates using the WorldClim version 2 model and are provided
 */
/* 		via the netzskiy.tm3 file from the Cal Program, and if they are zeros, then they are determined by this r
outine. */

/* 		The adhoc sunrise and sunset fixes are NO LONGER USED from version 16, since it is assumed that they were
 */
/* 		needed due to lack of temperature modeling. */

/* 		As of version 17, Menat's refraction is no longer used, and is replaced by van der Werf's */
/*		As of version 18, no terrestrial refraction contribution is added or subtracted
/*		Winter atmosphere for removing the 15 seconds added cushion is defined by ave min monthly temp < 1/2 * max min temperature
/* 		There is also no longer any mention of winter and summer, so the nweather parameter is */
/* 		currently no longer in use */
/*
/*		Version 18: removed redundant use of VDW rescaling
/*		Version 19: uses maximum WorldClim temperatures for VDW calcuations
/*		Version 20: added option to add cushion instead of printing in green
/*		Version 21: add final viewangle output
/*
/* ----------------------------------------------------- */
/*            = 20 * maxang + 1 ;number of azimuth points separated by .1 degrees */
/*            = maximum number of places that can be analyzed at once */
/* -------------------------------------------------------------------------------- */
/* ///////////////////commented out for version >= 15 ////////////////////////// */
/*        REAL*4 azin1(366),times1(366), disob(366),DAT1(366)*11 */
/*        CHARACTER filin1(366)*12, tim1(366)*8, tim(nplace)*8 */
/*        CHARACTER TIMR1*8, timr*36, filin*27, CH80*80 */
/*        REAL*4 azin(nplace) */
/*        LOGICAL ,SUMMER,WINTER */
/* ////////////////////////////////////////////////////////////////////////////// */
/*       Version 15-etc. parameters */
/* 		MT is Minimum Temperature Array, AT is Average Temperature Array */
/*	    VERSION 20 -- add option to add additional cushion to times (handling problem won't print green times) */
    pi = acos(-1.);
    cd = pi / 180.;
    dcmrad = 1 / (cd * 1e3); //converts mrads to degrees
    pi2 = pi * 2.;
    ch = 12. / pi;
    hr = 60.;
    geo = FALSE_;
    nointernet = TRUE_;
/* 		Average radius of Earth in km */
    rearth = 6356.766f;
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
/*       read in scan parameters */
/*       first find the netzskiy.tm3 file--this determines the directory */
    s_copy(testdriv, "cdefghij", (ftnlen)8, (ftnlen)8);
    for (i__ = 1; i__ <= 8; ++i__) {
	*(unsigned char *)drivlet = *(unsigned char *)&testdriv[i__ - 1];
	s_copy(tmpfil + 1, ":/jk/netzskiy.tm3", (ftnlen)17, (ftnlen)17);
	*(unsigned char *)tmpfil = *(unsigned char *)drivlet;
	ioin__1.inerr = 0;
	ioin__1.infilen = 18;
	ioin__1.infile = tmpfil;
	ioin__1.inex = &exists;
	ioin__1.inopen = 0;
	ioin__1.innum = 0;
	ioin__1.innamed = 0;
	ioin__1.inname = 0;
	ioin__1.inacc = 0;
	ioin__1.inseq = 0;
	ioin__1.indir = 0;
	ioin__1.infmt = 0;
	ioin__1.inform = 0;
	ioin__1.inunf = 0;
	ioin__1.inrecl = 0;
	ioin__1.innrec = 0;
	ioin__1.inblank = 0;
	f_inqu(&ioin__1);
	if (exists) {
/*             found default drive */
	    *(unsigned char *)yesint = *(unsigned char *)drivlet;
	    goto L7;
	}
/* L5: */
    }
/*       now read in the internet.yes file */
/*       if operating internet mode, don't write stat files */
L7:
    s_copy(yesint + 1, ":/cities/internet.yes", (ftnlen)21, (ftnlen)21);
    ioin__1.inerr = 0;
    ioin__1.infilen = 22;
    ioin__1.infile = yesint;
    ioin__1.inex = &exists;
    ioin__1.inopen = 0;
    ioin__1.innum = 0;
    ioin__1.innamed = 0;
    ioin__1.inname = 0;
    ioin__1.inacc = 0;
    ioin__1.inseq = 0;
    ioin__1.indir = 0;
    ioin__1.infmt = 0;
    ioin__1.inform = 0;
    ioin__1.inunf = 0;
    ioin__1.inrecl = 0;
    ioin__1.innrec = 0;
    ioin__1.inblank = 0;
    f_inqu(&ioin__1);
    if (exists) {
/*          found internet.yes file */
	o__1.oerr = 0;
	o__1.ounit = 1;
	o__1.ofnmlen = 22;
	o__1.ofnm = yesint;
	o__1.orl = 0;
	o__1.osta = "old";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
	s_rsfe(&io___18);
	do_fio(&c__1, (char *)&intflag, (ftnlen)sizeof(integer));
	e_rsfe();
	if (intflag == 1) {
	    nointernet = FALSE_;
	}
	cl__1.cerr = 0;
	cl__1.cunit = 1;
	cl__1.csta = 0;
	f_clos(&cl__1);
    }

	//now look for refraction flag file used for comparing 
	//ray tracing refraction values to standard refraction values 
	//appearing in the ESAA 9.331
	sprintf(fnam,"%s%s", drivlet, ":/jk/refflag.tmp");
	if (stream = fopen( fnam, "r"))
	{
		fscanf(stream,"%d\n", &nweatherflag);
		fclose (stream);
	}

	//check for file containg name of daynumber, temperatures, pressures
	sprintf(fnam,"%s%s", drivlet, ":\\jk\\daynumflag.tmp");
	if (stream = fopen( fnam, "r"))
	{
		fgets_CR(doclin, 255, stream); //read in line of text containing the file name of the daynumber file
		strcpy( fndynum, doclin );
		flagdynum = TRUE_;
		fclose (stream);
	}

	//adhoc flags: subtracts 15 seconds under special conditions for sunrises
	// adds 15 seconds under special conditions for sunsets

	adhocrise = TRUE_;
	adhocset = FALSE_;
	//read file that determines the sunrise adhoc fix value
	//if file doesn't exist, then use initialized value for adhocrise
	sprintf(fnam,"%s%s", drivlet, ":/jk/adhocflag.tmp");
	if (stream = fopen( fnam, "r"))
	{
		fscanf(stream,"%d\n", &adhocflag);
		if (adhocflag == 1)
		{
			adhocrise = TRUE_;
		}
		else
		{
			adhocrise = FALSE_;
		}
		fclose (stream);
	}

	/*
	sprintf(tmpfil, "%s%s", drivlet, ":/jk/netzskiy.tm3");
	if (stream = fopen( fnam, "r") )
	{

	}
	fclose( stream );
	*/
/* 'summer and winter refraction values for heights 1 m to 999 m */
/* 'as calculated by MENAT.FOR */
/*       opening file netzskiy.tm3 */
    o__1.oerr = 0;
    o__1.ounit = 10;
    o__1.ofnmlen = 18;
    o__1.ofnm = tmpfil;
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
    s_rsle(&io___20);
    do_lio(&c__3, &c__1, (char *)&nyrheb, (ftnlen)sizeof(integer));
    e_rsle();
    s_rsle(&io___22);
    do_lio(&c__3, &c__1, (char *)&nzskflg, (ftnlen)sizeof(integer));
    do_lio(&c__5, &c__1, (char *)&geotd, (ftnlen)sizeof(doublereal));
    e_rsle();
    if (nzskflg <= -4) {
	geo = TRUE_;
    } else {
	geotd = 2.;
    }
    s_rsle(&io___25);
    do_lio(&c__3, &c__1, (char *)&numfil, (ftnlen)sizeof(integer));
    e_rsle();
    s_rsfe(&io___27);
    do_fio(&c__1, cn4, (ftnlen)4);
    e_rsfe();
    s_rsfe(&io___29);
    do_fio(&c__1, locname, (ftnlen)40);
    e_rsfe();

	/*
	if (strstr(locname, 'eros'))
	{
		if (strstr(locname, "USA"))
		{
			DTMType = 1; //USGS 30 m DED
			distlim = 5;
			CushionAmount = 20;
		}
		else if (strstr(locname, "Israel"))
		{
			DTMType = 0;//EY DTM
			distlim = 4;
			CushionAmount = 15;
		}
		else if (strstr(locname, "England"))
		{
			DTMType = 2; //50 m Survey Ordinance DTM
			distlim = 8;
			CushionAmount = 1.67 * 20;
		}
		else
		{
			DTMType = 3; //SRTM based DTM
			distlim = 10;
			CushionAmount = 35;
		}
	}
	else
	{
		DTMType = 0; //EY DTM
		distlim = 4;
		CushionAmount = 15;
	}
	*/

    stat = 20.f / numfil;
    stat2 = 0.f;
    nstat = 1;
/*       write or zero status file */
    if (nointernet) {
	s_copy(filstat + 1, ":/jk/stat0001.tmp", (ftnlen)17, (ftnlen)17);
	*(unsigned char *)filstat = *(unsigned char *)drivlet;
	ioin__1.inerr = 0;
	ioin__1.infilen = 18;
	ioin__1.infile = filstat;
	ioin__1.inex = &exists;
	ioin__1.inopen = 0;
	ioin__1.innum = 0;
	ioin__1.innamed = 0;
	ioin__1.inname = 0;
	ioin__1.inacc = 0;
	ioin__1.inseq = 0;
	ioin__1.indir = 0;
	ioin__1.infmt = 0;
	ioin__1.inform = 0;
	ioin__1.inunf = 0;
	ioin__1.inrecl = 0;
	ioin__1.innrec = 0;
	ioin__1.inblank = 0;
	f_inqu(&ioin__1);
	if (exists) {
	    o__1.oerr = 0;
	    o__1.ounit = 1;
	    o__1.ofnmlen = 18;
	    o__1.ofnm = filstat;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 1;
	    o__1.ofnmlen = 18;
	    o__1.ofnm = filstat;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
	s_wsle(&io___35);
	do_lio(&c__3, &c__1, (char *)&c__0, (ftnlen)sizeof(integer));
	e_wsle();
	cl__1.cerr = 0;
	cl__1.cunit = 1;
	cl__1.csta = 0;
	f_clos(&cl__1);
    }
    nplac[0] = 0;
    nplac[1] = 0;
    i__1 = numfil;
    for (nfil = 1; nfil <= i__1; ++nfil) {
	s_rsfe(&io___38);
	do_fio(&c__1, fileo, (ftnlen)27);
	e_rsfe();
	s_rsle(&io___40);
	do_lio(&c__5, &c__1, (char *)&kmxo, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&kmyo, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&hgt, (ftnlen)sizeof(doublereal));
	do_lio(&c__3, &c__1, (char *)&nstrtyr, (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&yrst[0], (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&yrend[0], (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&nendyr, (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&yrst[1], (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&yrend[1], (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&nsetflag, (ftnlen)sizeof(integer));
	e_rsle();

	if (fabs(kmxo) > 80 && fabs(kmyo) > 80 && geotd == 0)
	{
		//these are ITM coordinates for EY location
		geo = FALSE_;
		geotd = 2;
	}

	if (!ReadTemps) {
		if (!geo) {
			casgeo_(&kmyo, &kmxo, &lt, &lg);
		}else{
			lt = kmyo;
			lg = kmxo;
		}

		ier = Temperatures(lt,-lg, mint, avgt, maxt);

	}else{ //read the temmperatures from the netzskiy.tm3 file
		s_rsle(&io___49);
		do_lio(&c__3, &c__1, (char *)&mint[0], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[1], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[2], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[3], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[4], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[5], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[6], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[7], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[8], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[9], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[10], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&mint[11], (ftnlen)sizeof(integer));
		e_rsle();
		s_rsle(&io___51);
		do_lio(&c__3, &c__1, (char *)&avgt[0], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[1], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[2], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[3], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[4], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[5], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[6], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[7], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[8], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[9], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[10], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&avgt[11], (ftnlen)sizeof(integer));
		e_rsle();
		s_rsle(&io___51);
		do_lio(&c__3, &c__1, (char *)&maxt[0], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[1], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[2], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[3], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[4], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[5], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[6], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[7], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[8], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[9], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[10], (ftnlen)sizeof(integer));
		do_lio(&c__3, &c__1, (char *)&maxt[11], (ftnlen)sizeof(integer));
		e_rsle();

		//check for errors in the netzskiy.tm3 file vis a vis the temperatures
		if (mint[0] == 0 && avgt[0] == 0 && maxt[0] == 0 && mint[11] == 0 && avgt[11] == 0 && maxt[11] == 0) 
		{
			if (!geo) {
				casgeo_(&kmyo, &kmxo, &lt, &lg);
			}else{
				lt = kmyo;
				lg = kmxo;
			}

			ier = Temperatures(lt,-lg, mint, avgt, maxt);
		}
	}

	/////////////////////addded parameters from netzskiy.tm3 ///////070521///////////////////
	//using currentdir string to dertimne DTM type failed
	//instead read in flags from netzskiy.tm3 for which type of way to handle near obstructions
	s_rsle(&io___51);
	do_lio(&c__3, &c__1, (char *)&AddCushions, (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&distlim, (ftnlen)sizeof(integer));
	do_lio(&c__3, &c__1, (char *)&CushionAmount, (ftnlen)sizeof(integer));
	e_rsle();
	///////////////////////////////////////////////////////////////////////////////////////

	astro = FALSE_;
	if (nsetflag == 2 || nsetflag == 3) {
	    astro = TRUE_;
	    if (nsetflag == 2) {
		nsetflag = 0;
	    }
	    if (nsetflag == 3) {
		nsetflag = 1;
	    }
	}
	if (nsetflag == 0) {
	    ++nplac[0];                     
	    if (nplac[0] > 100) {
		if (nointernet && ! nofiles) {
		    s_wsle(&io___55);
		    do_lio(&c__9, &c__1, " ERROR--exceeded max. num. of plac"
			    "es!", (ftnlen)37);
		    e_wsle();
		    s_paus("continue", (ftnlen)8);
		}
		goto L1000;
	    }
	    if (nointernet && ! nofiles) {
/*                prepare names of intermediate neighborhood files */
		s_copy(placn + (nsetflag + 1 + (nplac[0] << 1) - 3) * 30, 
			fileo + 15, (ftnlen)30, (ftnlen)8);
/*                names used for COMPARE for versions <= 14 */
		 if (nplac[0] < 10) {
		    s_wsfi(&io___58);
		    do_fio(&c__1, (char *)&nplac[0], (ftnlen)sizeof(integer));
		    e_wsfi();
		    s_copy(ext1 + (nsetflag + 1 + (nplac[0] << 1) - 3 << 2), 
			    ".pl", (ftnlen)3, (ftnlen)3);
		    *(unsigned char *)&ext1[(nsetflag + 1 + (nplac[0] << 1) - 
			    3 << 2) + 3] = *(unsigned char *)ch1;
		} else if (nplac[0] >= 10 && nplac[0] <= 99) {
		    s_wsfi(&io___61);
		    do_fio(&c__1, (char *)&nplac[0], (ftnlen)sizeof(integer));
		    e_wsfi();
		    s_copy(ext1 + (nsetflag + 1 + (nplac[0] << 1) - 3 << 2), 
			    ".p", (ftnlen)2, (ftnlen)2);
		    s_copy(ext1 + ((nsetflag + 1 + (nplac[0] << 1) - 3 << 2) 
			    + 2), ch2, (ftnlen)2, (ftnlen)2);
		} else if (nplac[0] >= 100) {
		    s_wsfi(&io___63);
		    do_fio(&c__1, (char *)&nplac[0], (ftnlen)sizeof(integer));
		    e_wsfi();
		    *(unsigned char *)&ext1[(nsetflag + 1 + (nplac[0] << 1) - 
			    3) * 4] = '.';
		    s_copy(ext1 + ((nsetflag + 1 + (nplac[0] << 1) - 3 << 2) 
			    + 1), ch3, (ftnlen)3, (ftnlen)3);
		}
		s_copy(ext3 + (nsetflag + 1 + (nplac[0] << 1) - 3 << 2), 
			fileo + 23, (ftnlen)4, (ftnlen)4);
	    }
	} else if (nsetflag == 1) {
	    ++nplac[1];
	    if (nplac[1] > 100) {
		if (! nofiles && nointernet) {
		    s_wsle(&io___65);
		    do_lio(&c__9, &c__1, " ERROR--exceeded max. num. of plac"
			    "es!", (ftnlen)37);
		    e_wsle();
		    s_paus("continue", (ftnlen)8);
		}
		goto L1000;
	    }
	    if (! nofiles && nointernet) {
/*                prepare names of intermediate neighborhood files */
		s_copy(placn + (nsetflag + 1 + (nplac[1] << 1) - 3) * 30, 
			fileo + 15, (ftnlen)30, (ftnlen)8);
/*                names used for COMPARE */
		if (nplac[1] < 10) {
		    s_wsfi(&io___66);
		    do_fio(&c__1, (char *)&nplac[1], (ftnlen)sizeof(integer));
		    e_wsfi();
		    s_copy(ext1 + (nsetflag + 1 + (nplac[1] << 1) - 3 << 2), 
			    ".pl", (ftnlen)3, (ftnlen)3);
		    *(unsigned char *)&ext1[(nsetflag + 1 + (nplac[1] << 1) - 
			    3 << 2) + 3] = *(unsigned char *)ch1;
		} else if (nplac[1] >= 10 && nplac[1] <= 99) {
		    s_wsfi(&io___67);
		    do_fio(&c__1, (char *)&nplac[1], (ftnlen)sizeof(integer));
		    e_wsfi();
		    s_copy(ext1 + (nsetflag + 1 + (nplac[1] << 1) - 3 << 2), 
			    ".p", (ftnlen)2, (ftnlen)2);
		    s_copy(ext1 + ((nsetflag + 1 + (nplac[1] << 1) - 3 << 2) 
			    + 2), ch2, (ftnlen)2, (ftnlen)2);
		} else if (nplac[1] >= 100) {
		    s_wsfi(&io___68);
		    do_fio(&c__1, (char *)&nplac[0], (ftnlen)sizeof(integer));
		    e_wsfi();
		    *(unsigned char *)&ext1[(nsetflag + 1 + (nplac[1] << 1) - 
			    3) * 4] = '.';
		    s_copy(ext1 + ((nsetflag + 1 + (nplac[1] << 1) - 3 << 2) 
			    + 1), ch3, (ftnlen)3, (ftnlen)3);
		}
		s_copy(ext3 + (nsetflag + 1 + (nplac[1] << 1) - 3 << 2), 
			fileo + 23, (ftnlen)4, (ftnlen)4);
	    }
	}
/*       Nstrtyr                    'beginning year to calculate */
/*       Nendyr                     'last year to calculate */
/* --------------------------<<<<USER INPUT>>>---------------------------------- */
	nplaceflg = 1;
/*                                  '=0 calculates earliest sunrise/latest sunset for selected places in Jerus
alem */
/*                                  '=1 calculation for a specific place--must input place number */
/*                                  '=2 for calculations for LA */
/*      'Nsetflag                   '=0 for calculating sunrises */
/*                                  '=1 for calculating sunsets */
	nweather = 0;
/* 								   'this parameter is now deprecated */
/* 								   'since there is only one atmosphere */
	ndtm = 5;
	nmenat = 1;
/*                                  '=0 uses tp */
/*                                  '=1 calculates sunrises over Harei Moav using theo. refrac. */
	niter = 1;
/*                                  '=0 doesn't iterate the sunrise times */
/*                                  '=1 iterates the sunrise times */
	ndruk = 0;
/*                                  '=0 gives normal output */
/*                                  '= 1 use to calculate files to compare to druk output/vangeld output */
	p = 1013.25f;
/*                                  'mean atmospheric pressure (mb) */
	t = 27.;
/*                                  'mean atmospheric temperature (degrees Celsius) */
	stepsec = 2.5f;
/*                                  'step size in seconds */
	nofiles = FALSE_;
/*                                   =.TRUE. to turn off neighborhood files (automatic for internet calculatio
ns) */
/*                                   =.FALSE. to turn on writing of neighborhood files */
	nplaces = 1;
/*       'N.B. astronomical constants are accurate to 10 seconds only to the year 2050 */
/*       '(Astronomical Almanac, 1996, p. C-24) */
/*       Addhoc fixes: (15 second fixes based on netz observations at Neveh Yaakov) */
/* 		As of Version 16, ADHOCRISE = .FALSE. */
//	adhocrise = TRUE_; //TRUE_;
//	adhocset = FALSE_;
/* -------------------------------------------------------------------------------------- */
	if (nsetflag == 0) {
	    nloop = 72;
/*          start search 2 minutes above astronomical horizon */
	    nlpend = 2000;
	} else if (nsetflag == 1) {
/*          sunset */
	    nloop = 370;
	    if (nplaceflg == 0) {
		nloop = 75;
	    }
	    nlpend = -10;
	}
	nopendruk = 0;
	i__2 = nendyr;
	for (nyear = nstrtyr; nyear <= i__2; ++nyear) {
/* 	     define year iteration number */
	    nyrstp = nyear - nstrtyr + 1;
	    if (nointernet && ! nofiles) {
		s_copy(place + (nsetflag << 3), fileo + 15, (ftnlen)8, (
			ftnlen)8);
		if (ndruk == 1 && nopendruk == 0) {
/*             open temporary file */
		    *(unsigned char *)filtmp = *(unsigned char *)drivlet;
		    s_copy(filtmp + 1, ":/JK/", (ftnlen)5, (ftnlen)5);
		    s_copy(filtmp + 6, place + (nsetflag << 3), (ftnlen)8, (
			    ftnlen)8);
		    s_copy(filtmp + 14, ".TMP", (ftnlen)4, (ftnlen)4);
		    o__1.oerr = 0;
		    o__1.ounit = 4;
		    o__1.ofnmlen = 18;
		    o__1.ofnm = filtmp;
		    o__1.orl = 0;
		    o__1.osta = "NEW";
		    o__1.oacc = 0;
		    o__1.ofm = 0;
		    o__1.oblnk = 0;
		    f_open(&o__1);
		    nopendruk = 1;
		}
	    }


		if (FindWinter) {
			///////////////version 18 way of determining winter months prone to strong near ground inversion layers/////////////
			//in the past, ns1, ns2 were used to determine the summer and winter refraction
			//             ns3, ns4 were used to determine the adhoc cushion removal/addition
			//after advent of van der Werf raytracing, ns1, nd2 are no longer in use, since winter/summer refraction doesn't exist
			//adhoc regions are now used for dealing with likely Inversion layers in the winter months, and they use ns1, ns2

			//determine highest min temperature
			
			for (iTemp = 0; iTemp < 12; iTemp++) {
				if (MaxMinTemp < mint[iTemp] ) MaxMinTemp = mint[iTemp];
			}
			WinTemp = WinterPercent * MaxMinTemp; //winter is defined as months whose mean min temperature less than 80% of maximium min temp of the year

			//corrected model now only subtracts time for horizon or below horizon views during deep winter months
			/*
			ns1 = 0;
			ns2 = 367;
			ns3 = 0;
			ns4 = 367;
			//now determine which months have mean minimum temperatures below 1/2 * MaxMinTemp
			if (lt >= 0) {
				for (iTemp = 0; iTemp < 12; iTemp++)
				{
					if (mt[iTemp] > WinTemp && ns3 == 0) {
						ns3 = DayCount; //end of witner
					}
					if (mt[iTemp] < WinTemp && ns3 != 0 && ns4 == 367) {
						ns4 = DayCount + 1; //beginning of winter
					}
					DayCount += lenMonth(iTemp + 1);
				}
				ns3 = 40;
				if (ns4 == 367) ns4 = 310;
			}else if (lt < 0) {
				for (iTemp = 0; iTemp < 12; iTemp++)
				{
					if (mt[iTemp] < WinTemp && ns3 == 0) {
						ns3 = DayCount + 1; //end of witner
					}
					if (mt[iTemp] > WinTemp && ns3 != 0 && ns4 == 367) {
						ns4 = DayCount; //beginning of winter
					}
					DayCount += lenMonth(iTemp + 1);
				}
			}
			*/

			//determine geographic coordinates if not yet available, e.g.,for EY profile files
			if (geo) {
				lt = kmxo;
				lg = kmyo;
			}else{
				casgeo_(&kmyo, &kmxo, &lt, &lg);
			}

		} else {
			//old latitude determination of winter percentage -- deoesn't take height innto consideration

			if (geo) {
			lt = kmxo;
			lg = kmyo;
			if (abs(lt) < 20.) {
	/*             tropical region, so use tropical atmosphere which */
	/*             is approximated well by the summer mid-latitude */
	/*             atmosphere for the entire year */
				ns1 = 0;
				ns2 = 0;
				ns3 = 367;
				ns4 = 367;
	/*          define days of the year when winter and summer begin */
			} else if (abs(lt) >= 20. && abs(lt) < 30.) {
				ns1 = 55;
				ns2 = 320;
	/*             no adhoc sunset fixes */
				ns3 = 0;
				ns4 = 366;
			} else if (abs(lt) >= 30. && abs(lt) < 40.) {
	/*             weather similar to Eretz Israel */
				ns1 = 85;
				ns2 = 290;
	/*             Times for ad-hoc fixes to the visible and astr. sunrise */
	/*             (to fit observations of the winter netz in Neve Yaakov). */
	/*             This should affect  sunrise and sunset equally. */
	/*             However, sunset hasn't been observed, and since it would */
	/*             make the sunset times later, it's best not to add it to */
	/*             the sunset times as a chumrah. */
				ns3 = 30;
				ns4 = 330;
			} else if (abs(lt) >= 40. && abs(lt) < 50.) {
	/*             the winter lasts two months longer than in Eretz Israel */
				ns1 = 115;
				ns2 = 260;
				ns3 = 60;
				ns4 = 300;
			} else if (abs(lt) >= 50.) {
	/*             the winter lasts four months longer than in Eretz Israel */
				ns1 = 145;
				ns2 = 230;
				ns3 = 90;
				ns4 = 270;
			}
			} else {
			casgeo_(&kmyo, &kmxo, &lt, &lg);
			ns1 = 85;
			ns2 = 290;
			ns3 = 30;
			ns4 = 330;
			}

			/* version 17
			if (geo) {
			lt = kmxo;
			lg = kmyo;
	// 		   in this version the above parameters are no longer used, see comment below 
			} else {
			casgeo_(&kmyo, &kmxo, &lt, &lg);
	// 		   for this version 16, only use van der Werf's atmosphere which is very 
	// 		   similar to the US Standard Atmosphere, and scale the refraction 
	// 		   according to van der Werf scaling laws, using the minimum temperature 
			}
			*/
		}

/*       read in elevations from Profile file */
/*       which is a list of the view angle (corrected for refraction) */
/*       as a function of azimuth for the particular observation */
/*       point at kmxo,kmyo,hgt */
	    if (nointernet) {
		s_wsle(&io___90);
		do_lio(&c__9, &c__1, "                  PLEASE WAIT--READING"
			" DTM FILE", (ftnlen)47);
		e_wsle();
	    }
/* >>>>>>>>>>>>>>>>>>>>>>>>>>> */
	    s_copy(f2, fileo, (ftnlen)27, (ftnlen)27);
	    s_copy(comentss, fileo + 15, (ftnlen)12, (ftnlen)12);
	    o__1.oerr = 0;
	    o__1.ounit = 2;
	    o__1.ofnmlen = 27;
	    o__1.ofnm = f2;
	    o__1.orl = 0;
	    o__1.osta = "old";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	    s_rsfe(&io___93);
	    do_fio(&c__1, chd48, (ftnlen)48);
	    e_rsfe();
/* 		check if this is old format profile file, if so, set flag to remove */
/* 		old form of terrestrial refraction contribution */
		//first look for word "APPRNR"

		pos = InStr(1, chd48, "APPRNR");
		if (pos > 0) {
			//can either be old form of TR, or no TR added or unknown mean added (depricated, but treat anyways)
			pos = InStr(1, chd48, ",0");
			if (pos > 0 ) {
				//new format without any TR added, or TR was subtracted out
				ntrcalc = 0;
			} else {
				pos = InStr(1, chd48, ",1");
				if (pos > 0) {
					//old depricated format using TR with undeclared mean ground temperature, so calculate it using the WorldClim
					ntrcalc = 1;
				} else {
					//old type rofile file using Menat based TR
					ntrcalc = 2;
				}
			}
		} else {
			//TR calculated with decalred ground temparature written into the header as the last number, so extract it
			char Tgroundch[7] = "\0";
			strncpy(Tgroundch, chd48 + 40, 6);
			Tground = atof(Tgroundch);
			if (Tground == 0.0)
			{
				//failed to extract the ground temperature
				//must be some old format type used, e.g., for astronomical places
				//ground temperature not set, so calculate it using WorldClim
				ntrcalc = 1;;
			}
			else
			{
				ntrcalc = 3;
			}

		}

		/*
		//old way of doing it, updated with the code in the lines above on 6/26/20
		np1 = i_indx(chd48, ",0", (ftnlen)48, (ftnlen)2);
		np2 = i_indx(chd48, ",1", (ftnlen)48, (ftnlen)2);

	    if (np1 > 0 && np2 == 0) {
			//new format w/o TR calculated, so calculate it 
		ntrcalc = 0;
	    } else if (np1 == 0 && np2 > 0) {
			//average temp TR contrib, remove it and recalculte it using month's average temperature
			//this assumes that the change in the horizon profile will be small when changing the temperature
		ntrcalc = 1;
	    } else {
			//old profile file with old TR calculation, so remove it and recalculate using new expression
			//However, assumption that the horizon profile will not change is more iffy in this case
		ntrcalc = 2;
	    }
		*/

	    s_rsfe(&io___98);
	    do_fio(&c__1, chd62, (ftnlen)62);
	    e_rsfe();

		if (hgt != 0) //if hgt = 0, then this is mishor, so no need to compare heights with the one in the profile
		{
			//read more accurate record of the height from this doc line in the profile
			if ( ParseString( chd62, ",", strarr, 4 ) )
			{
				//reading more accurate height directly from file failed
				//must be some old file format, or the astr files generated by that option
				//so exit out of caution with warning printing on console
				printf("********ERROR detected********\n");
				printf("File format error detected in the profile's coordinate line\n");
				printf("**********ABORTING*******\n");
				return -1; 
				;
			}
			else //extract text
			{
				hgt0 = hgt;
				hgt = atof( &strarr[2][0] );

				if (FixProfHgtBug && ntrcalc != 0) //some new format files have added heights already, so skip
				{
					//bat and profile heights are identical
					//usually means that newreadDTM recorded the ground height
					//instead of the calculation height after adding the observer height
					//this was bug that was fixed on 101520
					//so fix it here
					if (hgt0 == hgt) hgt += hgtobs;
				}
			}
		}

	    nazn = 0;
L40:
	    ++nazn;

/*       stuff elev array with:  azi,vang,kmx,kmy,dist,hgt */
	    i__3 = s_rsle(&io___101);
	    if (i__3 != 0) {
		goto L50;
	    }

		if (nazn == 1)
		{
			//read in first azimuth and determine the angular range and reallocate memory for it
			real elevazi = 0;
			i__3 = do_lio(&c__4, &c__1, (char *)&elevazi, (
				ftnlen)sizeof(real));
			if (i__3 != 0) {
			goto L50;
			}

			//determine the maximum azimuth range from the first read azimuth value
			maxang = (doublereal)fabs(elevazi);

			if (maxang0 == 0) 
			{
				//first allocation of memory

				elev = (real *) calloc(80 * maxang + 1, sizeof(real));

				if (elev == NULL) //memory allocation failes
				{
					printf("Memory allocation failed, aborting....\n");
					return -1;
				}
				else
				{
					//record the azimuth range

					maxang0 = maxang;
				}

			}
			else
			{

				if (maxang != maxang0)  //reallocate the memory
				{

					free(elev); //release the allocated memory for elev and reallocate

					elev = (real *) calloc(80 * maxang + 1, sizeof(real));

					if (elev == NULL) //memory allocation failes
					{
						printf("Memory reallocation failed, aborting....\n");
						return -1;
					}
					else
					{
						//record the new azimuth range

						maxang0 = maxang;
					}

				}
			}

			//transfer the azimuth value already read to the allocated array
			elev[(nazn << 2) - 4] = elevazi;

		}
		else
		{
			i__3 = do_lio(&c__4, &c__1, (char *)&elev[(nazn << 2) - 4], (
				ftnlen)sizeof(real));
			if (i__3 != 0) {
			goto L50;
			}
		}
	    i__3 = do_lio(&c__4, &c__1, (char *)&elev[(nazn << 2) - 3], (
		    ftnlen)sizeof(real));
	    if (i__3 != 0) {
		goto L50;
	    }
	    i__3 = do_lio(&c__5, &c__1, (char *)&e2, (ftnlen)sizeof(
		    doublereal));
	    if (i__3 != 0) {
		goto L50;
	    }
	    i__3 = do_lio(&c__5, &c__1, (char *)&e3, (ftnlen)sizeof(
		    doublereal));
	    if (i__3 != 0) {
		goto L50;
	    }
	    i__3 = do_lio(&c__4, &c__1, (char *)&elev[(nazn << 2) - 2], (
		    ftnlen)sizeof(real));
	    if (i__3 != 0) {
		goto L50;
	    }
	    i__3 = do_lio(&c__4, &c__1, (char *)&elev[(nazn << 2) - 1], (
		    ftnlen)sizeof(real));
	    if (i__3 != 0) {
		goto L50;
	    }
	    i__3 = e_rsle();
	    if (i__3 != 0) {
		goto L50;
	    }

		//remove terrestrial refraction additon used by rdhalba4 to create the profile
	    if (ntrcalc == 2) {
			//old TR based on Menat ray tracing, for old profile files using old versions 
			deltd = hgt - elev[(nazn << 2) - 1];
			distd = elev[(nazn << 2) - 2];
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
			elev[(nazn << 2) - 3] -= avref;
	    } else if (ntrcalc == 1 || ntrcalc == 3) {
			//remove terrestrial refraction addition used by program rdhalba4 to create the profile
			//rdhalbat now uses the average temperature both for sunrise and sunset profiles

			if (ntrcalc == 1) {
				meantemp = 0.f;
				for (itk = 1; itk <= 12; ++itk) {
				meantemp += avgt[itk - 1];
	/* L43: */
				}
				meantemp = meantemp / 12.f + 273.15f;
			} else {
				meantemp = Tground;
			}

			/*
			if (nsetflag == 0) {
				//earlier version of rdhalba4 used minimum temperatures for sunrise profiles
				meantemp = 0.f;
				for (itk = 1; itk <= 12; ++itk) {
				meantemp += mint[itk - 1];
				}
				meantemp = meantemp / 12.f + 273.15f;
			} else if (nsetflag == 1) {
				//earlier version of rdhalba4 used average temperature for sunset profiles
				meantemp = 0.f;
				for (itk = 1; itk <= 12; ++itk) {
				meantemp += avgt[itk - 1];
			    }
				meantemp = meantemp / 12.f + 273.15f;
			}
			*/

			deltd = hgt - elev[(nazn << 2) - 1];
			//convert height difference to kms
			deltd *= .001f;
			//distance to obstruction 
			distd = elev[(nazn << 2) - 2]; //units of kms
	/* 		calculate conserved portion of Wikipedia's terrestrial refraction expression */
	/* 		Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere */
			//so have p * 8.15 * 1000 * (0.0343 - 0.0065)/(T * T * 3600) degrees
			trrbasis = p * 8.15f * 1e3f * .0277f / (meantemp * meantemp * 3600);
			//in the winter, dt/dh will be positive for inversion layer and refraction is increased
	/* 	    now calculate pathlength of curved ray using Lehn's parabolic approximation */
	/* Computing 2nd power */
			d__1 = distd;
	/* Computing 2nd power */
			d__3 = distd;
	/* Computing 2nd power */
			d__2 = abs(deltd) - d__3 * d__3 * .5f / rearth; //units of kms
			pathlength = sqrt(d__1 * d__1 + d__2 * d__2); //units of kms
	/* 		    now add terrestrial refraction based on modified Wikipedia expression */
			if (abs(hgtdif) > 1e3f) {
				exponent = .99f;
			} else {
				exponent = .9965f;
			}
			//ViewAdd = TRfudge * trrbasis * pow_dd(&pathlength, &exponent);
			elev[(nazn << 2) - 3] -= TRfudge * trrbasis * pow_dd(&pathlength, &exponent);
	    }

	    //if (nazn == 1) {
		//maxang = (r__1 = elev[(nazn << 2) - 4], dabs(r__1));
	    //}
	    goto L40;
L50:
	    --nazn;
	    cl__1.cerr = 0;
	    cl__1.cunit = 2;
	    cl__1.csta = 0;
	    f_clos(&cl__1);
	    s_copy(coment, place + (nsetflag << 3), (ftnlen)8, (ftnlen)8);
	    if (! nofiles && nointernet) {
/*          prepare filenames of output files */
		nyear1 = nyear;
		if (abs(nyear1) < 1000) {
		    nyear1 = abs(nyear1) + 1000;
		}
		s_wsfi(&io___121);
		i__3 = abs(nyear1);
		do_fio(&c__1, (char *)&i__3, (ftnlen)sizeof(integer));
		e_wsfi();
		if (nplaceflg == 1 || nplaceflg == 2) {
		    s_copy(name__, place + (nsetflag << 3), (ftnlen)4, (
			    ftnlen)4);
		    s_copy(name__ + 4, ch4, (ftnlen)4, (ftnlen)4);
		}
		if (nsetflag == 0) {
		    *(unsigned char *)namet = *(unsigned char *)drivlet;
		    s_copy(namet + 1, ":/fordtm/netz/", (ftnlen)14, (ftnlen)
			    14);
		    s_copy(namet + 15, name__, (ftnlen)8, (ftnlen)8);
		} else if (nsetflag == 1) {
		    *(unsigned char *)namet = *(unsigned char *)drivlet;
		    s_copy(namet + 1, ":/fordtm/skiy/", (ftnlen)14, (ftnlen)
			    14);
		    s_copy(namet + 15, name__, (ftnlen)8, (ftnlen)8);
		}
		s_copy(plcfil, namet, (ftnlen)23, (ftnlen)23);
		s_copy(plcfil + 23, ext1 + (nsetflag + 1 + (nplac[nsetflag] <<
			 1) - 3 << 2), (ftnlen)4, (ftnlen)4);
		ioin__1.inerr = 0;
		ioin__1.infilen = 23;
		ioin__1.infile = namet;
		ioin__1.inex = &exists;
		ioin__1.inopen = 0;
		ioin__1.innum = 0;
		ioin__1.innamed = 0;
		ioin__1.inname = 0;
		ioin__1.inacc = 0;
		ioin__1.inseq = 0;
		ioin__1.indir = 0;
		ioin__1.infmt = 0;
		ioin__1.inform = 0;
		ioin__1.inunf = 0;
		ioin__1.inrecl = 0;
		ioin__1.innrec = 0;
		ioin__1.inblank = 0;
		f_inqu(&ioin__1);
		if (exists) {
		    o__1.oerr = 0;
		    o__1.ounit = 1;
		    o__1.ofnmlen = 23;
		    o__1.ofnm = namet;
		    o__1.orl = 0;
		    o__1.osta = "OLD";
		    o__1.oacc = 0;
		    o__1.ofm = 0;
		    o__1.oblnk = 0;
		    f_open(&o__1);
		} else {
		    o__1.oerr = 0;
		    o__1.ounit = 1;
		    o__1.ofnmlen = 23;
		    o__1.ofnm = namet;
		    o__1.orl = 0;
		    o__1.osta = "NEW";
		    o__1.oacc = 0;
		    o__1.ofm = 0;
		    o__1.oblnk = 0;
		    f_open(&o__1);
		}
		ioin__1.inerr = 0;
		ioin__1.infilen = 27;
		ioin__1.infile = plcfil;
		ioin__1.inex = &exists;
		ioin__1.inopen = 0;
		ioin__1.innum = 0;
		ioin__1.innamed = 0;
		ioin__1.inname = 0;
		ioin__1.inacc = 0;
		ioin__1.inseq = 0;
		ioin__1.indir = 0;
		ioin__1.infmt = 0;
		ioin__1.inform = 0;
		ioin__1.inunf = 0;
		ioin__1.inrecl = 0;
		ioin__1.innrec = 0;
		ioin__1.inblank = 0;
		f_inqu(&ioin__1);
		if (exists) {
		    o__1.oerr = 0;
		    o__1.ounit = 2;
		    o__1.ofnmlen = 27;
		    o__1.ofnm = plcfil;
		    o__1.orl = 0;
		    o__1.osta = "OLD";
		    o__1.oacc = 0;
		    o__1.ofm = 0;
		    o__1.oblnk = 0;
		    f_open(&o__1);
		} else {
		    o__1.oerr = 0;
		    o__1.ounit = 2;
		    o__1.ofnmlen = 27;
		    o__1.ofnm = plcfil;
		    o__1.orl = 0;
		    o__1.osta = "NEW";
		    o__1.oacc = 0;
		    o__1.ofm = 0;
		    o__1.oblnk = 0;
		    f_open(&o__1);
		}
	    }
	    td = geotd;
/*       time difference from Greenwich England */
	    if (nplaceflg == 2) {
		td = -8.;
	    }
	    lr = cd * lt;
	    tf = td * 60. + lg * 4.;
	    ac = cd * .9856003;
	    mc = cd * .9856474;
	    ec = cd * 1.915;
	    e2c = cd * .02;
	    ob1 = cd * -4e-7;
/*       change of ecliptic angle per day */
/*       mean anomaly for Jan 0, 1996 00:00 at Greenwich meridian */
	    ap = cd * 357.528;
/*       mean longitude of sun Jan 0, 1996 00:00 at Greenwich meridian */
	    mp = cd * 280.461;
/*       ecliptic angle for Jan 0, 1996 00:00 at Greenwich meridian */
	    ob = cd * 23.439;
/*       shift the ephemerials for the time zone */
	    ap -= td / 24. * ac;
	    mp -= td / 24. * mc;
	    ob -= td / 24. * ob1;
	    nyr = nyear;
/*       year for calculation */
	    nyd = nyr - 1996;
	    nyf = 0;
	    if (nyd < 0) {
		i__3 = nyr;
		for (nyri = 1995; nyri >= i__3; --nyri) {
		    nyrtst = nyri;
		    nyltst = 365;
		    if (nyrtst % 4 == 0) {
			nyltst = 366;
		    }
		    if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
			nyltst = 365;
		    }
		    nyf -= nyltst;
/* L100: */
		}
	    } else if (nyd >= 0) {
		i__3 = nyr - 1;
		for (nyri = 1996; nyri <= i__3; ++nyri) {
		    nyrtst = nyri;
		    nyltst = 365;
		    if (nyrtst % 4 == 0) {
			nyltst = 366;
		    }
		    if (nyrtst % 100 == 0 && nyrtst % 400 != 0) {
			nyltst = 365;
		    }
		    nyf += nyltst;
/* L150: */
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
L600:
	    if (mp < 0.) {
		mp += pi * 2.;
	    }
/* L610: */
	    if (mp < 0.) {
		goto L600;
	    }
L620:
	    if (mp > pi * 2.) {
		mp -= pi * 2.;
	    }
/* L630: */
	    if (mp > pi * 2.) {
		goto L620;
	    }
/* L640: */
	    ap += nyf * ac;
L650:
	    if (ap < 0.) {
		ap += pi * 2.;
	    }
/* L660: */
	    if (ap < 0.) {
		goto L650;
	    }
L670:
	    if (ap > pi * 2.) {
		ap -= pi * 2.;
	    }
/* L680: */
	    if (ap > pi * 2.) {
		goto L670;
	    }
	    d__1 = (doublereal) yrend[nyear - nstrtyr];
	    for (dy = (doublereal) yrst[nyear - nstrtyr]; dy <= d__1; dy += 
		    1.) {

			if (dy == 341)
			{
				int cc;
				cc = 1;
			}
/* 		    determine the minimum and average temperature for this day for current place */
/* 			use Meeus's forumula p. 66 to convert daynumber to month, */
/* 			no need to interpolate between temepratures -- that is overkill */

			//this is how to debug for a certain year and daynumber
			//set breakpoint on ccc = 1
			/*
			if (nyear == 2022 && dy == 2)
			{
				short ccc;
				ccc = 1;
			}
			*/

		k = 2.;
		if (nyl == 366) {
		    k = 1.;
		}
		//model temperature variation that fits the data the 
		m = (doublereal) ((integer) ((k + dy) * 9 / 275 + .98f));
		/*
		//tweek the WorldClim ground temperatures
		switch ((int)m) {
		case 1:
			AddTemp = 30;
			break;
		case 2:
			AddTemp = 10;
			break;
		case 3:
			AddTemp = 10;
			break;
		case 4:
			AddTemp = 10;
			break;
		case 5:
			AddTemp = 10;
			break;
		case 6:
			AddTemp = 20;
			break;
		case 7:
			AddTemp = 20;
			break;
		case 8:
			AddTemp = 20;
			break;
		case 9:
			AddTemp = 10;
			break;
		case 10:
			AddTemp = 10;
			break;
		case 11:
			AddTemp = 10;
			break;
		case 12:
			AddTemp = 30;
			break;
		}
		*/

		if (flagdynum) //use the actual ground temperatures and pressures instead
		{
			if (stream = fopen( fndynum, "r"))
			{
				char astarr[4][100];

				while ( !feof(stream) )
				{
					tk = -9999;
					p = 1013.25;
					showCalc = FALSE_;

					//read in line of data, if year and daynumber match, use the temperature and pressure
					fgets_CR( doclin, 255, stream );

					if ( ParseString( doclin, ",", astarr, 4 ) )
					{
						//blank line, or something else wrong, so close file
						fclose (stream);
						Tobs = -9999;
						break;
					}
					else //extract text
					{
						yrtst = atoi( &astarr[0][0] );
						dynum = atoi( &astarr[1][0] );
						Tobs = atof( &astarr[2][0] );
						Pobs = atof( &astarr[3][0] );

						if (yrtst == nyear && dynum == dy)
						{
							tk = Tobs;
							tkmin = mint[(integer) m - 1] + 273.15f;
							p = Pobs;
							if (Diagnostics) 
							{
								printf("found temperature dy, yr, Tobs, p, tmin = %f, %d, %f, %f, %f\n", dy, nyear, tk, p, tkmin);
								getch();
							    showCalc = TRUE_; //flag to show the details of the calculation and compare for using just tk
								nloops_2 = 0;
							}
							break;
						}

					}					
				}
				fclose (stream);
				if (tk != -9999) goto L500; //skip temperature definition via WorldClim temperatures
			}
		}

		AddTemp = 0; //30; //for now don'tuse different temperatures than the WorldClim ones
/* 			convert minimum and average temperatures to Kelvin scale */
		//if (nsetflag == 0) {
/* 				sunrise, model temp. as minimum diurnal temperatures */
			tkmin = mint[(integer) m - 1] + 273.15f + AddTemp;
			//tk = Min_Temp_adjust * tk + (1 - Min_Temp_adjust) * (at[(integer) m - 1] + 273.15f);
/* 				now find average minimum temperature to use in removing profile's TR */
		//} else if (nsetflag == 1) {
/* 				sunset, model temp. as average diurnal temperature */
			tkavg = avgt[(integer) m - 1] + 273.15f + AddTemp;
			tkmax = maxt[(integer) m - 1] + 273.15f + AddTemp;
		//}
		if (nsetflag == 0) {
			//approximate average minimum temperature for sunrise
			tk = tkmin; //tkmax; //0.8 * tkmin + 0.2 * tkavg;
		} else {
			//use average temperature for sunset
			tk = tkavg; //tkmin; //tkmax;
		}
			

L500:
/* 			calculate van der Werf temperature scaling factor for refraction */
		d__2 = 288.15f / tk;
		vdwsf = pow_dd(&d__2, &c_b175); //<-------

		/////////////022321 edit - add pressure scaling law to vbwsf////////////////////
		d__2 = p/1013.25;
		vdwsf *= pow_dd(&d__2, &c_press); //pressure scalintg law VDW graph 5a
		//////////////////////////////////////////////////////////////////////////

/* 			calculate van der Werf scaling factor for view angles */
		d__2 = 288.15f / tk;
		vdwalt = pow_dd(&d__2, &c_b176);
		//scaling law for ref
		d__2 = 288.15f / tk;
		vdwrefr = pow_dd(&d__2, &c_b179);//<-------
		//scaling for astronomical times
		d__2 = 288.15f / tk;
		vbwast = pow_dd(&d__2, &c_b180);

		//scaling for local ray altitude (compliment of zenith angle)
		d__2 = 288.15f / tk;
		vbweps = pow_dd(&d__2, &c_b181); //<-----
/* 			calculate conserved portion of Wikipedia's terrestrial refraction expression */
/* 			Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere */
		//trbasis = p * 8.15f * 1e3f * .0277f / (tk * tk * 3600);
		trbasis = vdwsf * p * 8.15f * 1e3f * .0277f / (288.15 * 288.15 * 3600);
		//trbasis = p * 8.15f * 1e3f * .0277f / (tk * tk * 3600);
		//trbasis = pow(288.15/tk, 2.5) * p * 8.15f * 1e3f * .0277f / (288.15 * 288.15 * 3600);

/* 	     day iteration */
		ndystp = i_dnnt(&dy);
		if (nointernet) {
/*             update status file */
		    if (d_mod(&dy, &c_b177) == 0.) {
			stat2 += stat;
			*(unsigned char *)filstat = *(unsigned char *)drivlet;
			s_copy(filstat + 1, ":/jk/stat", (ftnlen)9, (ftnlen)9)
				;
			s_copy(filstat + 14, ".tmp", (ftnlen)4, (ftnlen)4);
			++nstat;
			if (nstat <= 9) {
			    s_copy(filstat + 10, "000", (ftnlen)3, (ftnlen)3);
			    s_wsfi(&io___151);
			    do_fio(&c__1, (char *)&nstat, (ftnlen)sizeof(
				    integer));
			    e_wsfi();
			    *(unsigned char *)&filstat[13] = *(unsigned char *
				    )ch1;
			} else if (nstat <= 99 && nstat >= 10) {
			    s_copy(filstat + 10, "00", (ftnlen)2, (ftnlen)2);
			    s_wsfi(&io___152);
			    do_fio(&c__1, (char *)&nstat, (ftnlen)sizeof(
				    integer));
			    e_wsfi();
			    s_copy(filstat + 12, ch2, (ftnlen)2, (ftnlen)2);
			} else if (nstat <= 999 && nstat >= 100) {
			    s_wsfi(&io___153);
			    do_fio(&c__1, (char *)&nstat, (ftnlen)sizeof(
				    integer));
			    e_wsfi();
			    s_copy(filstat + 11, ch3, (ftnlen)3, (ftnlen)3);
			} else if (nstat <= 9999 && nstat >= 1000) {
			    s_wsfi(&io___154);
			    do_fio(&c__1, (char *)&nstat, (ftnlen)sizeof(
				    integer));
			    e_wsfi();
			    s_copy(filstat + 10, ch4, (ftnlen)4, (ftnlen)4);
			}
			ioin__1.inerr = 0;
			ioin__1.infilen = 18;
			ioin__1.infile = filstat;
			ioin__1.inex = &exists;
			ioin__1.inopen = 0;
			ioin__1.innum = 0;
			ioin__1.innamed = 0;
			ioin__1.inname = 0;
			ioin__1.inacc = 0;
			ioin__1.inseq = 0;
			ioin__1.indir = 0;
			ioin__1.infmt = 0;
			ioin__1.inform = 0;
			ioin__1.inunf = 0;
			ioin__1.inrecl = 0;
			ioin__1.innrec = 0;
			ioin__1.inblank = 0;
			f_inqu(&ioin__1);
			if (exists) {
			    o__1.oerr = 0;
			    o__1.ounit = 3;
			    o__1.ofnmlen = 18;
			    o__1.ofnm = filstat;
			    o__1.orl = 0;
			    o__1.osta = "OLD";
			    o__1.oacc = 0;
			    o__1.ofm = 0;
			    o__1.oblnk = 0;
			    f_open(&o__1);
			} else {
			    o__1.oerr = 0;
			    o__1.ounit = 3;
			    o__1.ofnmlen = 18;
			    o__1.ofnm = filstat;
			    o__1.orl = 0;
			    o__1.osta = "NEW";
			    o__1.oacc = 0;
			    o__1.ofm = 0;
			    o__1.oblnk = 0;
			    f_open(&o__1);
			}
			s_wsfe(&io___155);
			do_fio(&c__1, (char *)&stat2, (ftnlen)sizeof(
				doublereal));
			e_wsfe();
			cl__1.cerr = 0;
			cl__1.cunit = 3;
			cl__1.csta = 0;
			f_clos(&cl__1);
		    }
		}
		dy1 = dy;
/*          change ecliptic angle from noon of previous day to noon today */
		ob += ob1;
		decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
		ra = atan(cos(ob) * tan(es));
		if (ra < 0.) {
		    ra += pi;
		}
		df = ms - ra;
L685:
		if (abs(df) > pi * .5) {
		    if (df > 0.) {
			df -= df / abs(df) * pi;
		    } else if (df < 0.) {
			df += df / abs(df) * pi;
		    }
		} else {
		    goto L687;
		}
		goto L685;
L687:
		et = df / cd * 4.;
/*          equation of time (minutes) */
		t6 = tf + 720. - et;
/*          astronomical noon in minutes */
		check = -tan(d__) * tan(lr);
		if (fabs(check) > 1)
		{
			fdy = 0;
		}
		else
		{
			fdy = acos(-tan(d__) * tan(lr)) * ch / 24.;
		}
/*          astronomical refraction calculations, hgt in meters */
		ref = 0.;
		eps = 0.;
		if (hgt > 0.) {
		    lnhgt = log(hgt * .001);
		}
/* 		   calculate total atmospheric refraction from the observer's height */
/* 		   to the horizon and then to the end of the atmosphere */
/* 		   All refraction terms have units of mrad */
		if (hgt <= 0.) {
		    goto L690;
		}
		ref = exp(vdwref[0] + vdwref[1] * lnhgt + vdwref[2] * lnhgt * 
			lnhgt + vdwref[3] * lnhgt * lnhgt * lnhgt + vdwref[4] 
			* lnhgt * lnhgt * lnhgt * lnhgt + vdwref[5] * lnhgt * 
			lnhgt * lnhgt * lnhgt * lnhgt);
		eps = exp(vdweps[0] + vdweps[1] * lnhgt + vdweps[2] * lnhgt * 
			lnhgt + vdweps[3] * lnhgt * lnhgt * lnhgt + vdweps[4] 
			* lnhgt * lnhgt * lnhgt * lnhgt);

		if (nweatherflag == 1) //using standard refraction ESAA 9.331
		{
			ref = sqrt(hgt) * 1.75/60.0;  //standard refraction from observer to horizon
			eps = sqrt(hgt) * 0.37/60.0; //geometrical dip angle from observer to horizon
			refrac1 = 34.0/60.0;   //standard refraction from horizon to inifinity	
			//air = (90 + eps + ref) * cd; //air in radians
			vdwsf = p  / tk; //pressure and temeprature dependence of ESAA 3.283 
			vdwrefr = 1.0;
			vbweps = 1.0;
			adhocrise = FALSE_;
			a1 = (288.15/tk) * (ref + refrac1) * cd; //a1 in radians
			air = (90 + eps) * cd + a1; //air in radians
			trbasis = (288.15/tk) * p * 8.15f * 1e3f * .0277f / (288.15 * 288.15 * 3600);
			a1 = a1 / cd;  //convert a1 back to degrees to be consistent with the nweatherflag == 0 option
			goto L700;
		}
/* 	       now add the all the contributions together due to the observer's height */
/* 		   along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0 */
/* 		   for this calculation, leave the refraction in units of mrad */
L690:
		a1 = 0.;
		refrac1 = c_b178; //vdW 288.15 degress K total atmospheric refraction from zero to infinity (mrad)
		//air = cd * 90. + (vbweps * eps + vdwrefr * ref + vdwsf * refrac1) / 1e3;
		air = cd * 90. + (eps + vdwsf * (ref + refrac1)) / 1e3;
		//air = cd * 90. + (eps + ref + refrac1) / 1e3;
		//proper temperature dependence is to use the scaling law for solar altitudes
		//all refraction terms are in units of mrad so convert them to radians by multiplying by 1e-3
		//air = cd * 90. + eps * 1e-3 + atan(tan((ref + refrac1) * 1e-3) * vbwast);

		//air = cd * 90. + (eps + ref + vdwsf * refrac1) / 1e3; //diagnostics

		//a1 = (vdwrefr * ref + vdwsf * refrac1) / 1e3; //redundant
		//a1 = (ref + refrac1) / 1e3;  //radians
		//a1 = vdwsf * (ref + refrac1) * 1e-3;  //convert to radians from mrad

		a1 = (ref + refrac1) / (cd * 1e3);  //convert to degrees from mrad
L700:
/* 		   leave a1 in radians */
		//a1 = atan(tan(a1) * vdwalt);  //scale for temperature //wrong - <-redundant
		if (nsetflag == 1) {
/*              astron. sunset */
		    ndirec = -1;
		    dy1 = dy - ndirec * fdy;
/*              calculate sunset/sunrise */
		    decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__)
			    ;
		    air += (1 - cos(aas) * .0167) * .2667 * cd;

			//////////changes EK 061922/////////////////////////////////
			//check for perpetual day/night
			check = tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			    cos(d__));

			if (fabs(check) > 1)
			{
				//perpetual day/night, no sunset, print out error and write flags for Cal Program
				//then skip to next day
				ErrorCode = -2;
				goto ErrorHandler;
			}
			else
			{
				sr1 = acos(check) * ch;
			}
		    //sr1 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			//    cos(d__))) * ch;
			/////////////////////////////////////////////////////////////////

		    sr = sr1 * hr;
/*              hour angle in minutes */
		    t3 = (t6 - ndirec * sr) / hr;
		    nacurr = 0;

/* *************************BEGIN SUNSET WINTER ADHOC FIX************************ */
			//new of version 18 -- determine portion of year prone to inversion layers using daylength
			if (adhocset) {
				EnableSunsetInv = FALSE_;
				if (FindWinter) {
					DayLength = 2 * sr1;
					//make sure that it is cold by checking miniimum temperature for the relevant month
					//use Meeus's forumula p. 66 to convert daynumber to month,
					//no need to interpolate between temepratures -- that is overkill
					kTemp = 2;
					if (nyl == 366) kTemp = 1;
					mTemp = floor(9 * (kTemp + dy) / 275 + 0.98);
					if (DayLength <= MaxWinLen) { //&& mint[mTemp-1] <= WinTemp) {
						//nacurr = 15;
						EnableSunsetInv = TRUE_;
					}
				}else{
					if (lt >= 0. && nweather == 0 && (dy <= (doublereal) ns3 
						|| dy >= (doublereal) ns4) ) {
		/*                 add winter addhoc fix					*/
						EnableSunsetInv = TRUE_;
						//nacurr = 15;
					}
					if (lt < 0. && nweather == 0 && (dy >= (doublereal) ns3 
						|| dy <= (doublereal) ns4) ) {
		/*                 add winter addhoc fix					*/
						EnableSunsetInv = TRUE_;
						//nacurr = 15;
					}
				}
			}
/* *************************END SUNSET WINTER ADHOC FIX************************ */
		    //t3 += nacurr / 3600.;
		    //t3sub = t3;
/*              astronomical sunrise/sunset in hours.frac hours */
		    if (astro) {
			azi1 = cos(d__) * sin(sr1 / ch);
			azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(
				sr1 / ch);
			azio = atan(azi1 / azi2);
			if (azio > 0.) {
			    azi1 = 90. - azio / cd;
			}
			if (azio < 0.) {
			    azi1 = -(azio / cd + 90.);
			}
/*                 write astronomical sunset times to file and loop in days */
			if (! nofiles && nointernet) {
/*                    write output to intermediate files */
			    if (dy == 1.) {
				s_wsle(&io___183);
				do_lio(&c__9, &c__1, "sunset for year ", (
					ftnlen)16);
				do_lio(&c__5, &c__1, (char *)&yr, (ftnlen)
					sizeof(doublereal));
				do_lio(&c__9, &c__1, "place: ", (ftnlen)7);
				do_lio(&c__9, &c__1, fileo, (ftnlen)27);
				e_wsle();
			    }

				nacurr = 0;
				if (EnableSunsetInv) {
					//add adhoc inversion fix to mishor/ast sunset
					nacurr = 15;
				    //t3 += nacurr / 3600.;
					EnableSunsetInv = FALSE_; //reset flag for next day
				}

				t3sub = t3 + nacurr/3600;

			    t1500_(&t3sub, &negch);
			    s_copy(tssc, t3s_1.t3subc, (ftnlen)8, (ftnlen)8);
/*                    convert dy to calendar date */
			    t1750_(&dy, &nyl, &nyr);
			    s_wsfe(&io___187);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, tssc, (ftnlen)8);
			    e_wsfe();
			    s_wsfe(&io___188);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, tssc, (ftnlen)8);
			    e_wsfe();
/*                    add dist to obst to be consistent w/visible formats */
			    dobs = 120.;
			    s_wsfe(&io___190);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, tssc, (ftnlen)8);
			    do_fio(&c__1, comentss, (ftnlen)12);
			    do_fio(&c__1, (char *)&azi1, (ftnlen)sizeof(
				    doublereal));
				////////////////081021/////add viewangle/////////////
				viewangle = 0.0;
			    do_fio(&c__1, (char *)&viewangle, (ftnlen)sizeof(
				    doublereal));
				//////////////////////////////////////////////
			    do_fio(&c__1, (char *)&dobs, (ftnlen)sizeof(
				    doublereal));
			    e_wsfe();
			}
			goto L920;
		    }
		} else if (nsetflag == 0) {
/* 	         astron. sunrise */
		    ndirec = 1;
		    dy1 = dy + fdy;
		    decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__)
			    ;
		    air += (1 - cos(aas) * .0167) * .2667 * cd;
/*              calculate sunset */

			//////////////changes EK 061922///////////////////////////
			//check for perpetual night or day
			check = -tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			    cos(d__));

			if (fabs(check) > 1)
			{
				//no sunrise, print out error messages and write flags and goto next day
				ErrorCode = -1;
				goto ErrorHandler;
			}
			else
			{
				sr2 = acos(check) * ch;
			}
		    //sr2 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			//    cos(d__))) * ch;
			///////////////////////////////////////////////////

		    sr2 *= hr;
		    tss = (t6 + sr2) / hr;
/*              sunset in hours.fraction of hours */
		    dy1 = dy - fdy;
/*              calculate sunrise */
		    decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__)
			    ;
			//////////////////chages EK 061922//////////////////////////////////////
			//check for perpetual night or day
			check = -tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			    cos(d__));
			if (fabs(check) > 1) 
			{
				//no sunrsie, perpertual day/night?
				ErrorCode = -1;
				goto ErrorHandler;
			}
			else
			{
				sr1 = acos(check) * ch;
			}
		    //sr1 = acos(-tan(lr) * tan(d__) + cos(air) / (cos(lr) * 
			//    cos(d__))) * ch;
			///////////////////////////////////////////////////////

		    sr = sr1 * hr;
/*              hour angle in minutes */
		    t3 = (t6 - sr) / hr;
		    nacurr = 0;
/* *************************BEGIN SUNRISE WINTER ADHOC FIX************************ */
			//new of version 18 -- determine portion of year prone to inversion layers using daylength
			if (adhocrise) {
				EnableSunriseInv = FALSE_;
				if (FindWinter) {
					DayLength = 2 * sr1;
					//make sure that it is cold by checking miniimum temperature for the relevant month
					//use Meeus's forumula p. 66 to convert daynumber to month,
					//no need to interpolate between temepratures -- that is overkill
					kTemp = 2;
					if (nyl == 366) kTemp = 1;
					mTemp = floor(9 * (kTemp + dy) / 275 + 0.98);
					if (DayLength <= MaxWinLen) { //&& mint[mTemp-1] <= WinTemp) {
						//nacurr = -15;
						EnableSunriseInv = TRUE_;
					}
				}else{
					if (lt >= 0. && nweather == 0 && (dy <= (doublereal) ns3 
						|| dy >= (doublereal) ns4) ) {
		/*                 add winter addhoc fix */
						EnableSunriseInv = TRUE_;
						//nacurr = -15;
					}
					if (lt < 0. && nweather == 0 && (dy >= (doublereal) ns3 
						|| dy <= (doublereal) ns4) ) {
		/*                 add winter addhoc fix */
						EnableSunriseInv = TRUE_;
						//nacurr = -15;
					}
				}
			}
/* *************************END SUNRISE WINTER ADHOC FIX************************ */
		    //t3 += nacurr / 3600.;
		    //t3sub = t3;
/*              astronomical sunrise in hours.frac hours */
/*              sr1= hour angle in hours at astronomical sunrise */
		    if (astro) {
			azi1 = cos(d__) * sin(sr1 / ch);
			azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(
				sr1 / ch);
			azio = atan(azi1 / azi2);
			if (azio > 0.) {
			    azi1 = 90. - azio / cd;
			}
			if (azio < 0.) {
			    azi1 = -(azio / cd + 90.);
			}
/*                 write astronomical sunrise time to file and loop in days */
			if (! nofiles && nointernet) {
/*                    write output to intermediate files */
			    if (dy == 1.) {
				s_wsle(&io___193);
				do_lio(&c__9, &c__1, "sunrise for year ", (
					ftnlen)17);
				do_lio(&c__5, &c__1, (char *)&yr, (ftnlen)
					sizeof(doublereal));
				do_lio(&c__9, &c__1, "place: ", (ftnlen)7);
				do_lio(&c__9, &c__1, fileo, (ftnlen)27);
				e_wsle();
			    }

				nacurr = 0;
				if (EnableSunriseInv) {
					//add adhoc inversion fix to mishor/ast sunrise
					nacurr = -15;
				    //t3 += nacurr / 3600.;
					EnableSunriseInv = FALSE_; //reset flag for next day
				}

				t3sub = t3 + nacurr/3600;

			    t1500_(&t3sub, &negch);
			    s_copy(ts1c, t3s_1.t3subc, (ftnlen)8, (ftnlen)8);
/*                    convert dy to calendar date */
			    t1750_(&dy, &nyl, &nyr);
			    s_wsfe(&io___195);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, ts1c, (ftnlen)8);
			    e_wsfe();
			    s_wsfe(&io___196);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, ts1c, (ftnlen)8);
			    e_wsfe();
/*                    add dist to obst to be consistent w/visible formats */
			    dobs = 120.;
			    s_wsfe(&io___197);
			    do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
			    do_fio(&c__1, ts1c, (ftnlen)8);
			    do_fio(&c__1, comentss, (ftnlen)12);
			    do_fio(&c__1, (char *)&azi1, (ftnlen)sizeof(
				    doublereal));
				//////////////081021/////////add viewangle//////
				viewangle = 0.0;
			    do_fio(&c__1, (char *)&viewangle, (ftnlen)sizeof(
				    doublereal));
				/////////////////////////////////////////////
			    do_fio(&c__1, (char *)&dobs, (ftnlen)sizeof(
				    doublereal));
			    e_wsfe();
			}
			goto L920;
		    }
		}
L692:
		nn = 0;
		nfound = 0;
		nabove = 0;
		i__3 = nlpend;
		i__4 = ndirec;
		for (nloops = nloop; i__4 < 0 ? nloops >= i__3 : nloops <= 
			i__3; nloops += i__4) {
		    ++nn;
		    hran = sr1 - stepsec * nloops / 3600.;
/*           step in hour angle by stepsec sec */
		    dy1 = dy - ndirec * hran / 24.;
		    decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__)
			    ;
/*           recalculate declination at each new hour angle */
			//calculate the zenith angle for each hour angle
		    z__ = sin(lr) * sin(d__) + cos(lr) * cos(d__) * cos(hran /
			     ch);  //zenith angle
		    z__ = acos(z__);
		    alt1 = pi * .5 - z__; //solar altitude = complement of zenith angle
/*           altitude in radians of sun's center */
		    azi1 = cos(d__) * sin(hran / ch);
		    azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(
			    hran / ch);
		    azi = atan(azi1 / azi2) / cd;
		    if (azi > 0.) {
			azi1 = 90. - azi;
		    }
		    if (azi < 0.) {
			azi1 = -(azi + 90.);
		    }

			//////////////changes EK 061922/////////////////////////////
		    if (abs(azi1) > maxang && nointernet) {
				//passed maximum azimuth of profile file,
				//print error message, write flags, and go to next day
				ErrorCode = -3;
				goto ErrorHandler;
		    }
			///////////////////////////////////////////////////////


			//*******************************************************************************
			//032722 --- here is where to start loop to check for sunrise occuring first on a part of the sun
			//other than the upper limb -- loop in azi by width of sun, and corresponding non-refracted altitude
			//for each step need to determine the refraction altitude and to check if at that azimuth the sun can be seen
			//over the terrain -- step in 3 steps on either side of sun
			//*********************************************************************
		    azi1 = ndirec * azi1;
			/////////////////////visible sunset////////////////////////////////////////
		    if (nsetflag == 1) {
/*              to first order, the top and bottom of the sun have the */
/*              following altitudes near the sunset */
			//al1 = atan(tan(alt1 + a1) * vdwalt) / cd;  //scale for temperature
			al1 = alt1 / cd + a1;
			//al1 = alt1 / cd;
/*              first guess for apparent top of sun */
			pt = 1.0; //1.06; //1.0;
			if (!nweatherflag) {
				//recalculate the temperature scaling factor, vdwsf, for the observer height and view angle
				//to decent approximation, only adjust once and not for each iteration
				altmin = al1 * 60.0;//convert degrees to arcminutes
				vbwexp = 1.68271708343173 -1.04055307822257E-04 * hgt 
					   - 4.65108719859762E-03 * altmin
					   + 1.47797248203399E-05 * altmin * altmin
					   - 2.37213585737878E-08 * altmin * altmin * altmin
                       + 1.48109107804038E-11 * altmin * altmin * altmin * altmin;
				vbwexp += ErrorFudge;
				//value for hgt = 750 meters
				//vdwsf = 1.60518041607982 -4.24521138064861E-03 * altmin; //use linear approximation to first order
					        //+  1.28352879338537E-05 * altmin * altmin
					        //-1.93196363948442E-08 * altmin * altmin * altmin
							//+ 1.13058575397583E-11 * altmin * altmin * altmin * altmin;
				d__2 = 288.15f / tk;
				vdwsf = pow_dd(&d__2, &vbwexp); 
			}
			refvdw_(&hgt, &al1, &refrac1, &nweatherflag, &vdwsf);
/*              first iteration to find position of upper limb of sun */
			//al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
			//al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
			if (nweatherflag) 
			{
				al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
			}
			else
			{
				al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
			}
			//scale by vdw temperature scaling
			//al1 = atan(tan(al1 * cd) * vdwalt) / cd;
/*              top of sun with refraction */
			if (niter == 1) {
			    for (nac = 1; nac <= 10; ++nac) {

					al1o = al1;
					refvdw_(&hgt, &al1, &refrac1, &nweatherflag, &vdwsf);
	/*                    subsequent iterations */
					//al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
					if (nweatherflag) 
					{
						al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
					}
					else
					{
						al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
					}

					if ((d__2 = al1 - al1o, abs(d__2)) < .004f) {
						//this is visible sunset
						goto L695;
					}

/* L693: */
			    }
			}
L695:
			ne = (integer) ((azi1 + maxang) * 10) + 1;
			elevinty = elev[(ne + 1 << 2) - 3] - elev[(ne << 2) - 
				3];
			elevintx = elev[(ne + 1 << 2) - 4] - elev[(ne << 2) - 
				4];
			elevint = elevinty / elevintx * (azi1 - elev[(ne << 2)
				 - 4]) + elev[(ne << 2) - 3];
/* ////////////// now modify apparent elevention, elevint, with terrestrial refraction///// */
/* 			   first calculate interpolated difference of heights */
			hgtdifx = elev[(ne + 1 << 2) - 1] - elev[(ne << 2) - 
				1];
			hgtdif = hgtdifx / elevintx * (azi1 - elev[(ne << 2) 
				- 4]) + elev[(ne << 2) - 1];
			hgtdif -= hgt;
			//convert to kms
			hgtdif *= .001f;
/* 			   now calculate interpolated distance to obstruction along earth */
			distx = elev[(ne + 1 << 2) - 2] - elev[(ne << 2) - 2];
			dist = distx / elevintx * (azi1 - elev[(ne << 2) - 4])
				 + elev[(ne << 2) - 2];
/* 			   now calculate pathlength of curved ray using Lehn's parabolic approximation */
/* Computing 2nd power */
			d__2 = dist;
/* Computing 2nd power */
			d__4 = dist;
/* Computing 2nd power */
			d__3 = abs(hgtdif) - d__4 * d__4 * .5f / 
				rearth;
			pathlength = sqrt(d__2 * d__2 + d__3 * d__3);
/* 			   now add terrestrial refraction based on modified Wikipedia expression */
			if (abs(hgtdif) > 1e3f) {
			    exponent = .99f;
			} else {
			    exponent = .9965f;
			}
			//ViewAdd = TRfudgeVis * trbasis * pow_dd(&pathlength, &exponent);
			elevint += TRfudgeVis * trbasis * pow_dd(&pathlength, &exponent);
/* //////////////////////////////////////////////////////////////////////////////////////// */
			if ((elevint >= al1 || nloops <= 0) && nabove == 1) {
/*                 reached sunset */
			    nfound = 1;
/*                 margin of error */
			    nacurr = 0;

				if (EnableSunsetInv && elevint <= HorizonInv) {
					//add adhoc inversion fix to mishor/ast sunset for view angles <= 0
					nacurr = 15; //this is minimum adhoc inversion fix (hours); can be a lot later
				    //t3 += nacurr / 3600.;
					EnableSunsetInv = FALSE_; //reset flag for this day
				}

			    if (nointernet && ! nofiles) {
				t3sub = t3 + nacurr/3600; //astronomical sunset with adhoc fix (since also affects ast sunset)
				t1500_(&t3sub, &negch);
				s_copy(ts1c, t3s_1.t3subc, (ftnlen)8, (ftnlen)
					8);
				t3sub = (d__2 = (stepsec * nloops - nacurr) / 
					3600., abs(d__2)); //difference in time using ast sunset without adhoc fix as reference
/*                    visible sunset can't be later then astronomical sunset! */
				if (nloops <= 0) {
				    t3sub = 0.;
				}
				if (nointernet && ndruk == 1) {
				    s_wsle(&io___219);
				    do_lio(&c__5, &c__1, (char *)&dy, (ftnlen)
					    sizeof(doublereal));
				    do_lio(&c__5, &c__1, (char *)&t3sub, (
					    ftnlen)sizeof(doublereal));
				    e_wsle();
				}
				t1500_(&t3sub, &negch);
				s_copy(ts2c, "         ", (ftnlen)9, (ftnlen)
					9);
				s_copy(ts2c + 1, t3s_1.t3subc, (ftnlen)8, (
					ftnlen)8);
				if (nloops > 0) {
				    *(unsigned char *)ts2c = '-';
				}
			    }
			    t3sub = t3 + ndirec * (stepsec * nloops + nacurr) 
				    / 3600.; //this is visible sunset with adhoc fix
			    if (nloops < 0) {
/*                     visible sunset can't be later then astronomical sunset! */
				t3sub = t3 + nacurr/3600;
			    }
/*                 <<<<<<<<<distance to obstruction>>>>>> */
			    dobsy = elev[(ne + 1 << 2) - 2] - elev[(ne << 2) 
				    - 2];
			    elevintx = elev[(ne + 1 << 2) - 4] - elev[(ne << 
				    2) - 4];
			    dobs = dobsy / elevintx * (azi1 - elev[(ne << 2) 
				    - 4]) + elev[(ne << 2) - 2];
				////////////////added 081021///recrod viewangle at sunset/////////////
				viewangle = al1;
				//////////////////////////////////////////////////////////////

				//////////////option to add cushion for near obstructions/////070521/////////////////
				AddedTime = 0.0;
				if (AddCushions)
				{
					//check if this is considered near obstruction, if so add time cushion
					if (dobs > 0 && lt < 90)
					{
						accobs = 300 * (cos(31.0 * cd)/cos(lt * cd)) * atan(0.001 * distlim/dobs)/cd;
						AddedTime = accobs - CushionAmount;
						if (AddedTime < 0) AddedTime = 0.0;
						
					}
					else
					{
						//add 1 hour
						AddedTime = 3600;
					}

					t3sub = t3sub - AddedTime/3600;
					
				}
				/////////////////////////////////////////////////////////////////////////////////

			    if (nointernet && ! nofiles) {
				t1500_(&t3sub, &negch);
				s_copy(tssc, t3s_1.t3subc, (ftnlen)8, (ftnlen)
					8);
/*                    sunset */
/*                    convert dy to calendar date */
				t1750_(&dy, &nyl, &nyr);
				if (dy == 1.) {
				    s_wsle(&io___222);
				    do_lio(&c__9, &c__1, "sunset for year ", (
					    ftnlen)16);
				    do_lio(&c__5, &c__1, (char *)&yr, (ftnlen)
					    sizeof(doublereal));
				    do_lio(&c__9, &c__1, "place: ", (ftnlen)7)
					    ;
				    do_lio(&c__9, &c__1, fileo, (ftnlen)27);
				    e_wsle();
				}
				s_wsfe(&io___223);
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, ts1c, (ftnlen)8);
				do_fio(&c__1, ts2c, (ftnlen)9);
				do_fio(&c__1, (char *)&nac, (ftnlen)sizeof(
					integer));
				e_wsfe();
				s_wsfe(&io___224);
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, ts1c, (ftnlen)8);
				do_fio(&c__1, ts2c, (ftnlen)9);
				e_wsfe();
				s_wsfe(&io___225);
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, comentss, (ftnlen)12);
				do_fio(&c__1, (char *)&azi1, (ftnlen)sizeof(
					doublereal));
				///////////////081021 add viewangle ////////////
				do_fio(&c__1, (char *)&viewangle, (ftnlen)sizeof(
					doublereal));
				/////////////////////////////////////////////////
				do_fio(&c__1, (char *)&dobs, (ftnlen)sizeof(
					doublereal));
				e_wsfe();
			    }
			    goto L900;
			} else if (elevint < al1) {
			    nabove = 1;
			}
			/////////////////visible sunrise/////////////////////////////////////////////
			} else if (nsetflag == 0) { 

				//diagnostics
				/*
				if (strstr(comentss, ".p14") && nyear == 2022 && dy == 2)
				{
					short aaaa;
					aaaa = 1;
				}
				*/

			al1 = alt1 / cd  + a1;  //convert altitude from radians to degrees
			//al1 = alt1 / cd;
/*                 first guess for apparent top of sun */
			pt = 1.0; //1.06; //1.0;

			if (!nweatherflag) {
				//recalculate the temperature scaling factor, vdwsf, for the observer height and view angle
				//as preliminary test, use the fit to the exponent at height = 750 meters
				altmin = al1 * 60.0;//convert degrees to arcminutes
				vbwexp = 1.68271708343173 -1.04055307822257E-04 * hgt 
					   - 4.65108719859762E-03 * altmin
					   + 1.47797248203399E-05 * altmin * altmin
					   - 2.37213585737878E-08 * altmin * altmin * altmin
                       + 1.48109107804038E-11 * altmin * altmin * altmin * altmin;
				vbwexp += ErrorFudge;
				//vdwsf = 1.60518041607982 -4.24521138064861E-03 * altmin; //use linear approximation to first order
					        //+  1.28352879338537E-05 * altmin * altmin
					        //-1.93196363948442E-08 * altmin * altmin * altmin
							//+ 1.13058575397583E-11 * altmin * altmin * altmin * altmin;
				d__2 = 288.15f / tk;
				vdwsf = pow_dd(&d__2, &vbwexp);

				if (showCalc) //diagnostics
				{
					d__2 = 288.15f / tkmin;
					vdwsf_2 = pow_dd(&d__2, &vbwexp);
				}
				//now multiply by pressure scaling law
			}
			refvdw_(&hgt, &al1, &refrac1, &nweatherflag, &vdwsf);
			if (showCalc) //diagnostics
			{
				refvdw_(&hgt, &al1_2, &refrac1_2, &nweatherflag, &vdwsf_2);
			}
/*                 first iteration to find top of sun */
			//al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * 
			//	dcmrad; //redundant
			//al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
			//sale solar altitude with vdw temp scaling relationship
			//al1 = atan(tan(al1 * cd) * vdwalt) / cd;
/*                 top of sun with refraction */
			if (nweatherflag) 
			{
				al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
			}
			else
			{
				al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
				if (showCalc) //diagnostics
				{
					al1_2 = alt1 / cd + .2667f + pt * vdwsf_2 * refrac1_2 * dcmrad;
				}
			}
			if (niter == 1) {
			    for (nac = 1; nac <= 10; ++nac) {
					al1o = al1;
					refvdw_(&hgt, &al1, &refrac1, &nweatherflag, &vdwsf);

					if (showCalc) //diagnostics
					{
						al1o_2 = al1_2;
						refvdw_(&hgt, &al1_2, &refrac1_2, &nweatherflag, &vdwsf_2);
					}
	/*                       subsequent iterations */

					if (nweatherflag) 
					{
						al1 = alt1 / cd + .2667f + pt * refrac1 * dcmrad;
					}
					else
					{
						al1 = alt1 / cd + .2667f + pt * vdwsf * refrac1 * dcmrad;
						if (showCalc) //diagnostics
						{
							al1_2 = alt1 / cd + .2667f + pt * vdwsf_2 * refrac1_2 * dcmrad;
						}
					}
					//scale solar altitude with vdw temp scaling law to add temp dependence
					//al1 = atan(tan(al1 * cd) * vdwalt) / cd;
	/*                       top of sun with refraction */
					if ((d__2 = al1 - al1o, abs(d__2)) < .004f) {
						//this is sunrise

						/*
						if (!weatherflag && showCalc)
						{
							//show details
							
							printf("Num iteration, vbwsf, refrac1, difference = %d, %f, %f, %f\n", nac, vdwsf, refrac1, abs(d__2));
							printf("Num iter orig, vbswf_2, refrac1_2, difference = %d, %f, %f, %f\n", nac, vdwsf_2, refrac1_2, fabs(al1_2 - al1o_2));
							getch();
						}
						*/
						goto L760;
					}
/* L750: */
			    }
			}
L760:
			alt1 = al1;
			if (showCalc) 
			{
				alt1_2 = al1_2;
			}
/*              now search digital elevation data */
			ne = (integer) ((azi1 + maxang) * 10) + 1;
			elevinty = elev[(ne + 1 << 2) - 3] - elev[(ne << 2) - 
				3];
			elevintx = elev[(ne + 1 << 2) - 4] - elev[(ne << 2) - 
				4];
			elevint = elevinty / elevintx * (azi1 - elev[(ne << 2)
				 - 4]) + elev[(ne << 2) - 3];
/* ////////////// now modify apparent elevention, elevint, with terrestrial refraction///// */
/* 			   first calculate interpolated difference of heights */
			hgtdifx = elev[(ne + 1 << 2) - 1] - elev[(ne << 2) - 
				1];
			hgtdif = hgtdifx / elevintx * (azi1 - elev[(ne << 2) 
				- 4]) + elev[(ne << 2) - 1];
			hgtdif -= hgt;
/* 			   convert to kms */
			hgtdif *= .001f;
/* 			   now calculate interpolated distance to obstruction along earth */
			distx = elev[(ne + 1 << 2) - 2] - elev[(ne << 2) - 2];
			dist = distx / elevintx * (azi1 - elev[(ne << 2) - 4])
				 + elev[(ne << 2) - 2];
/* 			   now calculate pathlength of curved ray using Lehn's parabolic approximation */
/* Computing 2nd power */
			d__2 = dist;
/* Computing 2nd power */
			d__4 = dist;
/* Computing 2nd power */
			d__3 = abs(hgtdif) - d__4 * d__4 * .5f / 
				rearth;
			pathlength = sqrt(d__2 * d__2 + d__3 * d__3);
/* 			   now add terrestrial refraction based on modified Wikipedia expression */
			if (abs(hgtdif) > 1e3f) {
			    exponent = .99f;
			} else {
			    exponent = .9965f;
			}
			//ViewAdd = TRfudgeVis * trbasis * pow_dd(&pathlength, &exponent)
			elevint += TRfudgeVis * trbasis * pow_dd(&pathlength, &exponent);
/* //////////////////////////////////////////////////////////////////////////////////////// */
			if (elevint <= alt1_2 && elevint > alt1 && showCalc)
			{
				if (nloops_2 == 0)
				{
					nloops_2 = nloops;
					printf("Sunrise shows up first with observed temperatures\n");
					printf("solar altitude is at: %f degrees\n", alt1_2);
					printf("nloops seconds = %d\n", nloops_2);
				}
			}
			if (elevint <= alt1) {
			    if (nn > 1) {
				nfound = 1;
			    }
			    if (nfound == 0) {
				goto L900;
			    }
			    if (dy == 1. && nointernet) {
				s_wsle(&io___226);
				do_lio(&c__9, &c__1, "sunrise for year: ", (
					ftnlen)18);
				do_lio(&c__3, &c__1, (char *)&nyr, (ftnlen)
					sizeof(integer));
				do_lio(&c__9, &c__1, " place: ", (ftnlen)8);
				do_lio(&c__9, &c__1, fileo, (ftnlen)27);
				e_wsle();
			    }
/*                 <<<<<<<<<distance to obstruction>>>>>> */
			    dobsy = elev[(ne + 1 << 2) - 2] - elev[(ne << 2) 
				    - 2];
			    elevintx = elev[(ne + 1 << 2) - 4] - elev[(ne << 
				    2) - 4];
			    dobs = dobsy / elevintx * (azi1 - elev[(ne << 2) 
				    - 4]) + elev[(ne << 2) - 2];
/*                 sunrise */
				/////////////////added 081021///record viewangle at sunrise/////////
				viewangle = alt1;
				////////////////////////////////////////////////

				if (elevint > alt1_2 && showCalc)
				{
					printf("Sunrise didn't yet occur with observed temperature\n");
				}

				nacurr = 0;
				if (EnableSunriseInv && elevint <= HorizonInv) {
					//add adhoc inversion fix to mishor/ast sunrise for view angles <= 0
					nacurr = -15; //this is minimum adhoc sunrise inversion fix,(hours); can be a lot earlier
				    //t3 += nacurr / 3600.;
					EnableSunriseInv = FALSE_; //reset flag for this day
				}

			    if (nointernet && ! nofiles) {
				t3sub = t3 + nacurr/3600; //astronomical sunrise will also be earlier due to inversion

				t1500_(&t3sub, &negch);
				s_copy(ts1c, t3s_1.t3subc, (ftnlen)8, (ftnlen)
					8);
/*                    astronomical time */
				t3sub = (stepsec * nloops + nacurr) / 3600.; //difference in time without the inversion fix

				if (showCalc)
				{
					if (nloops_2 < nloops && nloops_2 != 0)
					{
						printf("nloops = %d\n",nloops);
						printf("difference in time is: %d seconds\n", nloops - nloops_2);
						getch();
					}
				}

				if (nloops < 0) {
				    t3sub = 0.;
				}
				if (ndruk == 1) {
				    s_wsle(&io___227);
				    do_lio(&c__5, &c__1, (char *)&dy, (ftnlen)
					    sizeof(doublereal));
				    do_lio(&c__5, &c__1, (char *)&t3sub, (
					    ftnlen)sizeof(doublereal));
				    e_wsle();
				}
				t1500_(&t3sub, &negch);
				s_copy(ts2c, "         ", (ftnlen)9, (ftnlen)
					9);
				s_copy(ts2c + 1, t3s_1.t3subc, (ftnlen)8, (
					ftnlen)8);
/*                    difference between astronom to sunrise */
			    }
			    //nacurr = 0;
/*                 margin of error (seconds) to take into */
/*                 account differences between observations and the DTM `` */
				//add adhoc fix for inversion to the visible sunrise time
			    t3sub = t3 + (stepsec * nloops + nacurr) / 3600.;
			    if (nloops < 0) {
/*                    sunrise can't be before astronomical sunrise */
/*                    so set it equal to astronomical sunrise time */
/*                    minus the adhoc winter fix if any */
				t3sub = t3 + nacurr/3600;
			    }

				//diagnostics
				/*
				if (nyear == 2022 && dy == 2) //t3sub < 5.29334 && nyear == 2021 && dy == 250) //5.29134 && nyear == 2021 && dy == 250)
				{
					short vvvv;
					vvvv = 1;
				}
				*/

				/////////////////070521 option to add cushion for near obstructions//////////////////
				AddedTime = 0.0;
				if (AddCushions)
				{
					//check if this is considered near obstruction, if so add time cushion
					if (dobs > 0 && lt < 90)
					{
						accobs = 300 * (cos(31.0 * cd)/cos(lt * cd)) * atan(.001 * distlim/dobs)/cd;
						AddedTime = accobs - CushionAmount;
						if (AddedTime < 0) AddedTime = 0.0;
						
					}
					else
					{
						//add 1 hour
						AddedTime = 3600;
					}

					t3sub = t3sub + AddedTime/3600;
					
				}
				//////////////////////////////////////////////////////////////////////////////

			    if (nointernet && ! nofiles) {
				t1500_(&t3sub, &negch);
				s_copy(tssc, t3s_1.t3subc, (ftnlen)8, (ftnlen)
					8);
/*                    convert dy to calendar date */
				t1750_(&dy, &nyl, &nyr);
				s_wsfe(&io___228);
				//display the results on the console and write to the time difference and pl place file for this particular place
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, ts1c, (ftnlen)8);
				do_fio(&c__1, ts2c, (ftnlen)9);
				do_fio(&c__1, (char *)&nac, (ftnlen)sizeof(
					integer));
				e_wsfe();
				//write to pl and 
				s_wsfe(&io___229);
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, ts1c, (ftnlen)8);
				do_fio(&c__1, ts2c, (ftnlen)9);
				e_wsfe();
				s_wsfe(&io___230);
				do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
				do_fio(&c__1, tssc, (ftnlen)8);
				do_fio(&c__1, comentss, (ftnlen)12);
				do_fio(&c__1, (char *)&azi1, (ftnlen)sizeof(
					doublereal));
				/////////////added 081021 -- record viewangle at sunrise/sunset/////
				do_fio(&c__1, (char *)&viewangle, (ftnlen)sizeof(
					doublereal));
				/////////////////////////////////////////////
				do_fio(&c__1, (char *)&dobs, (ftnlen)sizeof(
					doublereal));
				e_wsfe();
			    }
			    goto L900;
			}
		    }
/* L850: */
		}
L900:
		if (nfound == 0 && nloop > 0 && nsetflag == 1) {
		    if (nointernet) {
			s_wsle(&io___231);
			do_lio(&c__9, &c__1, "*****search error-changing nlo"
				"op*****", (ftnlen)37);
			e_wsle();
		    }
		    if (nsetflag == 0) {
			nloop += -24;
			if (nloop < 0) {
			    nloop = 0;
			}
		    } else if (nsetflag == 1) {
			nloop += 24;
		    }
		    goto L692;
		} else if (nfound == 0 && nloop >= -100 && nsetflag == 0) {
		    if (nointernet) {
			s_wsle(&io___232);
			do_lio(&c__9, &c__1, "*****search error-changing nlo"
				"op*****", (ftnlen)37);
			e_wsle();
		    }
		    if (nsetflag == 0) {
			nloop += -24;
/*               IF (nloop .LT. 0) nloop = 0 */
		    } else if (nsetflag == 1) {
			nloop += 24;
		    }
		    goto L692;
		} else if (nfound == 0 && nloop < -100 && nsetflag == 0) {
		    if (nointernet) {
			s_wsle(&io___233);
			do_lio(&c__9, &c__1, "***search not successful, incr"
				"ease nlpend***", (ftnlen)44);
			e_wsle();
		    }
		    nlpend += 500;
		    nloop = 72;
		    goto L692;
		}
		if (nsetflag == 1) {
		    if (nloops + 30 < nloop || nloops + 30 > nloop) {
			nloop = nloops + 30;
		    }
		} else if (nsetflag == 0) {
		    if (nloops + 30 < nloop) {
			nloop = nloops + 30;
		    } else if (nloops + 30 > nloop) {
			nloop = nloops - 30;
		    }
		}
/*        //////////////calculations complete///////////////////// */
/*        check for latest sunset times and earliest sunrise times */
/*        among all the places (this replaces time comparisons of COMPARE) */
L920:
		if (nfil == 1) {

			//diagnostics
			/*
			if (nyrstp == 2 && ndystp == 2)
			{
				short ccc;
				ccc = 1;
			}
			*/

/*           first place, so intialize latest sunset arrays */
		    tsscs[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			    t3sub;
		    s_copy(coments + (nsetflag + 1 + (nyrstp + (ndystp << 1) 
			    << 1) - 7) * 12, comentss, (ftnlen)12, (ftnlen)12)
			    ;
		    azils[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			    azi1;
		    dobss[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			    dobs;
			viewang[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] =
				viewangle;
		} else {
/*           check if sunrise/sunset for new place is earlier/later */
		    if (nsetflag == 0) {
/*              sunrise */
			if (t3sub < tsscs[nsetflag + 1 + (nyrstp + (ndystp << 
				1) << 1) - 7]) {

				//diagnostics
				/*
				if (nyrstp == 2 && ndystp == 2)
				{
					short cccc;
					cccc = 1;
				}
				*/

			    tsscs[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = t3sub;
			    s_copy(coments + (nsetflag + 1 + (nyrstp + (
				    ndystp << 1) << 1) - 7) * 12, comentss, (
				    ftnlen)12, (ftnlen)12);
			    azils[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = azi1;
			    dobss[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = dobs;
			    viewang[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = viewangle;
			}
		    } else if (nsetflag == 1) {
/*              sunset */
			if (t3sub > tsscs[nsetflag + 1 + (nyrstp + (ndystp << 
				1) << 1) - 7]) {
			    tsscs[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = t3sub;
			    s_copy(coments + (nsetflag + 1 + (nyrstp + (
				    ndystp << 1) << 1) - 7) * 12, comentss, (
				    ftnlen)12, (ftnlen)12);
			    azils[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = azi1;
			    dobss[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = dobs;
			    viewang[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1)
				     - 7] = viewangle;
			}
		    }
		}
L925:;
	    }
/* L975: */
	}
	cl__1.cerr = 0;
	cl__1.cunit = 1;
	cl__1.csta = 0;
	f_clos(&cl__1);
	cl__1.cerr = 0;
	cl__1.cunit = 2;
	cl__1.csta = 0;
	f_clos(&cl__1);
	if (ndruk == 1) {
	    cl__1.cerr = 0;
	    cl__1.cunit = 4;
	    cl__1.csta = 0;
	    f_clos(&cl__1);
	}
L1000:
	;
    }
    cl__1.cerr = 0;
    cl__1.cunit = 10;
    cl__1.csta = 0;
    f_clos(&cl__1);
/*       PROGRAM COMPARE --finds earliest/latest times of nplac .PLC files */
/*       number of places to search */
    s_copy(filout + 15, placn + 4, (ftnlen)4, (ftnlen)4);
    if (nplac[0] == 0) {
	s_copy(filout + 15, placn + 34, (ftnlen)4, (ftnlen)4);
    }
    nset1 = 0;
    nset2 = 1;
    if (nplac[0] == 0) {
	nset1 = 1;
    }
    if (nplac[1] == 0) {
	nset2 = 0;
    }
    i__1 = nset2;
    for (nsetflag = nset1; nsetflag <= i__1; ++nsetflag) {
	if (nsetflag == 0) {
	    *(unsigned char *)filout = *(unsigned char *)drivlet;
	    s_copy(filout + 1, ":/fordtm/netz/", (ftnlen)14, (ftnlen)14);
	    nadd = 0;
	} else if (nsetflag == 1) {
	    *(unsigned char *)filout = *(unsigned char *)drivlet;
	    s_copy(filout + 1, ":/fordtm/skiy/", (ftnlen)14, (ftnlen)14);
	    nadd = 1;
	}
	if (nplac[nsetflag] == 1 && (nofiles || nointernet)) {
	    goto L2500;
	}

	i__2 = nendyr;
	for (nyr = nstrtyr; nyr <= i__2; ++nyr) {
	    nyrstp = nyr - nstrtyr + 1;
/*             determine number of days in this year */
	    nyl = 365;
	    if (nyr % 4 == 0) {
		nyl = 366;
	    }
	    if (nyr % 100 == 0 && nyr % 400 != 0) {
		nyl = 365;
	    }
	    nyr1 = nyr;
	    if (abs(nyr1) < 1000) {
		nyr1 = abs(nyr1) + 1000;
	    }
	    s_wsfi(&io___243);
	    i__4 = abs(nyr1);
	    do_fio(&c__1, (char *)&i__4, (ftnlen)sizeof(integer));
	    e_wsfi();
	    s_copy(filout + 15, cn4, (ftnlen)4, (ftnlen)4);
	    s_copy(filout + 19, ch4, (ftnlen)4, (ftnlen)4);
	    nleapyr = 0;
	    if (nyl == 366) {
		nleapyr = 1;
	    }
	    s_copy(filout + 23, ".com", (ftnlen)4, (ftnlen)4);
	    ioin__1.inerr = 0;
	    ioin__1.infilen = 27;
	    ioin__1.infile = filout;
	    ioin__1.inex = &exists;
	    ioin__1.inopen = 0;
	    ioin__1.innum = 0;
	    ioin__1.innamed = 0;
	    ioin__1.inname = 0;
	    ioin__1.inacc = 0;
	    ioin__1.inseq = 0;
	    ioin__1.indir = 0;
	    ioin__1.infmt = 0;
	    ioin__1.inform = 0;
	    ioin__1.inunf = 0;
	    ioin__1.inrecl = 0;
	    ioin__1.innrec = 0;
	    ioin__1.inblank = 0;
	    f_inqu(&ioin__1);
	    if (exists) {
		o__1.oerr = 0;
		o__1.ounit = 1;
		o__1.ofnmlen = 27;
		o__1.ofnm = filout;
		o__1.orl = 0;
		o__1.osta = "OLD";
		o__1.oacc = 0;
		o__1.ofm = 0;
		o__1.oblnk = 0;
		f_open(&o__1);
	    } else {
		o__1.oerr = 0;
		o__1.ounit = 1;
		o__1.ofnmlen = 27;
		o__1.ofnm = filout;
		o__1.orl = 0;
		o__1.osta = "NEW";
		o__1.oacc = 0;
		o__1.ofm = 0;
		o__1.oblnk = 0;
		f_open(&o__1);
	    }
	    i1 = yrst[nyr - nstrtyr];
	    i2 = yrend[nyr - nstrtyr];
	    i__4 = i2;
	    for (i__ = i1; i__ <= i__4; ++i__) {
/*              convert daynumber to date */
		d__1 = (doublereal) i__;
		t1750_(&d__1, &nyl, &nyr);
		s_copy(dat, t3s_1.caldayc, (ftnlen)11, (ftnlen)11);
/*              convert fractional hour to time */
		t3sub = tsscs[nsetflag + 1 + (nyrstp + (i__ << 1) << 1) - 7];
		t1500_(&t3sub, &negch);
		s_copy(timss, t3s_1.t3subc, (ftnlen)8, (ftnlen)8);
		s_copy(comentss, coments + (nsetflag + 1 + (nyrstp + (i__ << 
			1) << 1) - 7) * 12, (ftnlen)12, (ftnlen)12);
		azinss = azils[nsetflag + 1 + (nyrstp + (i__ << 1) << 1) - 7];
		dobsf = dobss[nsetflag + 1 + (nyrstp + (i__ << 1) << 1) - 7];
		viewangf = viewang[nsetflag + 1 + (nyrstp + (i__ << 1) << 1) - 7];
		s_wsfe(&io___251);
		do_fio(&c__1, dat, (ftnlen)11);
		do_fio(&c__1, timss, (ftnlen)8);
		do_fio(&c__1, comentss, (ftnlen)12);
		do_fio(&c__1, (char *)&azinss, (ftnlen)sizeof(real));
		do_fio(&c__1, (char *)&viewangf, (ftnlen)sizeof(doublereal)); //////added 081021 /////////////
		do_fio(&c__1, (char *)&dobsf, (ftnlen)sizeof(doublereal));
		e_wsfe();
/* L1500: */
	    }
/* L2000: */
	}
	cl__1.cerr = 0;
	cl__1.cunit = 1;
	cl__1.csta = 0;
	f_clos(&cl__1);
	goto L2500;
L2500:
	s_copy(filout + 15, "NETZSKIY.TM2", (ftnlen)12, (ftnlen)12);
	ioin__1.inerr = 0;
	ioin__1.infilen = 27;
	ioin__1.infile = filout;
	ioin__1.inex = &exists;
	ioin__1.inopen = 0;
	ioin__1.innum = 0;
	ioin__1.innamed = 0;
	ioin__1.inname = 0;
	ioin__1.inacc = 0;
	ioin__1.inseq = 0;
	ioin__1.indir = 0;
	ioin__1.infmt = 0;
	ioin__1.inform = 0;
	ioin__1.inunf = 0;
	ioin__1.inrecl = 0;
	ioin__1.innrec = 0;
	ioin__1.inblank = 0;
	f_inqu(&ioin__1);
	if (exists) {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 27;
	    o__1.ofnm = filout;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 27;
	    o__1.ofnm = filout;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
	if (nplac[nsetflag] < 10) {
	    s_wsfe(&io___252);
	    do_fio(&c__1, (char *)&nplac[nsetflag], (ftnlen)sizeof(integer));
	    e_wsfe();
	} else {
	    s_wsfe(&io___253);
	    do_fio(&c__1, (char *)&nplac[nsetflag], (ftnlen)sizeof(integer));
	    e_wsfe();
	}
	i__2 = nendyr;
	for (nyr = nstrtyr; nyr <= i__2; ++nyr) {
	    s_copy(filout + 15, cn4, (ftnlen)4, (ftnlen)4);
	    s_wsfi(&io___254);
	    do_fio(&c__1, (char *)&nyr, (ftnlen)sizeof(integer));
	    e_wsfi();
	    if (nplac[nsetflag] == 1 && (nofiles || nointernet)) {
		goto L2600;
	    }
	    s_copy(filout + 19, ch4, (ftnlen)4, (ftnlen)4);
	    s_copy(filout + 23, ".com", (ftnlen)4, (ftnlen)4);
	    s_wsfe(&io___255);
	    do_fio(&c__1, "\"", (ftnlen)1);
	    do_fio(&c__1, filout, (ftnlen)27);
	    do_fio(&c__1, "\"", (ftnlen)1);
	    e_wsfe();
	    goto L2700;
L2600:
	    s_copy(plcfil, filout, (ftnlen)15, (ftnlen)15);
	    s_copy(plcfil + 19, ch4, (ftnlen)4, (ftnlen)4);
	    s_copy(plcfil + 23, ext1 + (nsetflag + 1 + (nplac[nsetflag] << 1) 
		    - 3 << 2), (ftnlen)4, (ftnlen)4);
	    s_copy(plcfil + 15, place + (nsetflag << 3), (ftnlen)4, (ftnlen)4)
		    ;
	    s_wsfe(&io___256);
	    do_fio(&c__1, "\"", (ftnlen)1);
	    do_fio(&c__1, plcfil, (ftnlen)27);
	    do_fio(&c__1, "\"", (ftnlen)1);
	    e_wsfe();
L2700:
	    ;
	}
	cl__1.cerr = 0;
	cl__1.cunit = 3;
	cl__1.csta = 0;
	f_clos(&cl__1);
/* L3000: */
    }
/*       now write closing file to tell Cal Program that all */
/*       calculations have ended */
	*filstat = "";
    *(unsigned char *)filstat = *(unsigned char *)drivlet;
    s_copy(filstat + 1, ":/jk/netzend.tmp\0", (ftnlen)17, (ftnlen)17);
    ioin__1.inerr = 0;
    ioin__1.infilen = 18;
    ioin__1.infile = filstat;
    ioin__1.inex = &exists;
    ioin__1.inopen = 0;
    ioin__1.innum = 0;
    ioin__1.innamed = 0;
    ioin__1.inname = 0;
    ioin__1.inacc = 0;
    ioin__1.inseq = 0;
    ioin__1.indir = 0;
    ioin__1.infmt = 0;
    ioin__1.inform = 0;
    ioin__1.inunf = 0;
    ioin__1.inrecl = 0;
    ioin__1.innrec = 0;
    ioin__1.inblank = 0;
    f_inqu(&ioin__1);
    if (exists) {
	o__1.oerr = 0;
	o__1.ounit = 1;
	o__1.ofnmlen = 18;
	o__1.ofnm = filstat;
	o__1.orl = 0;
	o__1.osta = "OLD";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    } else {
	o__1.oerr = 0;
	o__1.ounit = 1;
	o__1.ofnmlen = 18;
	o__1.ofnm = filstat;
	o__1.orl = 0;
	o__1.osta = "NEW";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    }
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
L9999:
    return 0;

/////////////////////////changes EK 061922 ////////////////////////////////////////////
//inline errorhandler
ErrorHandler:
;
	//convert day and year to calendar date in civil calendar
	t1750_(&dy, &nyl, &nyr);

	azi1 = 0;
	dobs = 0;

	switch (ErrorCode)
	{
	case -1://perpetual day/night no sunrise
		//set flags for Cal Program
		t3sub = -99;
		viewangle = 0;
		strcpy(tssc, "NoSunrise", 8);
		strcpy(ts1c, "    NA   ", 8);
		strcpy(ts2c, "    NA   ", 8);
		nac = 1;
		azi1 = maxang;
		viewangle = 0.0;
		dobs = 0.0;

		buff[0] = 0;
		if (nointernet) {
			sprintf(buff, " %s %s\n", t3s_1.caldayc, "No sunrise. Probably perpetual day or night");
			printf(buff);}
		break; 
	case -2://perpetual day/night no sunset
		//set flags for Cal Program
		t3sub = 99;
		viewangle = 0;
		strcpy(tssc, "NoSunset", 8);
		strcpy(ts1c, "    NA   ", 8);
		strcpy(ts2c, "    NA   ", 8);
		nac = 2;
		azi1 = maxang;
		viewangle = 0.0;
		dobs = 0.0;

		buff[0] = 0;
		if (nointernet) {
			sprintf(buff, " %s %s\n", t3s_1.caldayc, "No sunset. Probably perpetual day or night");
			printf(buff);}
		break;
	case -3://passed azimuth range of profile file
		//set flags for Cal Program
		t3sub = 55;
		viewangle = 0;
		strcpy(tssc, ">max azi", 8);
		strcpy(ts1c, "    NA   ", 8);
		strcpy(ts2c, "    NA   ", 8);
		nac = 3;
		azi1 = maxang;
		viewangle = 0.0;
		dobs = 0.0;

		buff[0] = 0;
		if (nointernet) {
			sprintf(buff, " %s %s%5.1f %s\n", t3s_1.caldayc, "Maximum Azimuth: +/-", maxang, "passed!");
			printf(buff);}
		break;
	}

	s_wsfe(&io___228);
	//display the results on the console and write to the time difference and pl place file for this particular place
	do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
	do_fio(&c__1, tssc, (ftnlen)8);
	do_fio(&c__1, ts1c, (ftnlen)8);
	do_fio(&c__1, ts2c, (ftnlen)9);
	do_fio(&c__1, (char *)&nac, (ftnlen)sizeof(
		integer));
	e_wsfe();

	if (nointernet && ! nofiles) {

		//write to pl and 
		s_wsfe(&io___229);
		do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
		do_fio(&c__1, tssc, (ftnlen)8);
		do_fio(&c__1, ts1c, (ftnlen)8);
		do_fio(&c__1, ts2c, (ftnlen)9);
		e_wsfe();
		s_wsfe(&io___230);
		do_fio(&c__1, t3s_1.caldayc, (ftnlen)11);
		do_fio(&c__1, tssc, (ftnlen)8);
		do_fio(&c__1, comentss, (ftnlen)12);
		do_fio(&c__1, (char *)&azi1, (ftnlen)sizeof(
			doublereal));
		/////////////added 081021 -- record viewangle at sunrise/sunset/////
		do_fio(&c__1, (char *)&viewangle, (ftnlen)sizeof(
			doublereal));
		/////////////////////////////////////////////
		do_fio(&c__1, (char *)&dobs, (ftnlen)sizeof(
			doublereal));
		e_wsfe();
		}


	//if this is the first place in a seriers of places, add fake information so it will be skipped
	//when determing the earliest sunrise and the latest sunset
	if (nfil == 1) {

		tsscs[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			t3sub;
		s_copy(coments + (nsetflag + 1 + (nyrstp + (ndystp << 1) 
			<< 1) - 7) * 12, comentss, (ftnlen)12, (ftnlen)12)
			;
		azils[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			azi1;
		dobss[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] = 
			dobs;
		viewang[nsetflag + 1 + (nyrstp + (ndystp << 1) << 1) - 7] =
			viewangle;
	}

	//reset loop parameters
	nfound = -1;
	azi1 = 0;
	if (nsetflag == 0) {
		nloop = 72;
		// start search 2 minutes above astronomical horizon */
		nlpend = 2000;
	} else if (nsetflag == 1) {
		// sunset
		nloop = 370;
		if (nplaceflg == 0) {
		nloop = 75;
		}
		nlpend = -10;
	}

	//skip to next day
	goto L925;
//////////////////end inline ErrorHandler//////////////////////////////////////////

} /* MAIN__ */

/* Subroutine */ int decl_(doublereal *dy1, doublereal *mp, doublereal *mc, 
	doublereal *ap, doublereal *ac, doublereal *ms, doublereal *aas, 
	doublereal *es, doublereal *ob, doublereal *d__)
{
    /* Builtin functions */
    double acos(doublereal), sin(doublereal), asin(doublereal);

    /* Local variables */
    static doublereal cd, ec, pi, e2c, pi2;

/*       calculates new declination */
    pi = acos(-1.);
    pi2 = pi * 2.;
    cd = pi / 180.;
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

/* Subroutine */ int t1500_(doublereal *t3sub, integer *negch)
{
    /* Builtin functions */
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer s_wsfi(icilist *), do_fio(integer *, char *, ftnlen), e_wsfi(void)
	    ;

    /* Local variables */
    static doublereal t1;
    static char tshr[2];
    static integer nt3hr;
    static char tssec[2], tsmin[2];
    static integer nt3sec, nt3min;

    /* Fortran I/O blocks */
    static icilist io___269 = { 0, tshr, 0, "(I2)", 2, 1 };
    static icilist io___270 = { 0, tsmin, 0, "(I2)", 2, 1 };
    static icilist io___271 = { 0, tssec, 0, "(I2)", 2, 1 };


/*       convert hour.frac hour into hrs:min:sec */
/*       seconds are rounded UP to nearest second */
    s_copy(tshr, "  ", (ftnlen)2, (ftnlen)2);
    s_copy(tsmin, "  ", (ftnlen)2, (ftnlen)2);
    s_copy(tssec, "  ", (ftnlen)2, (ftnlen)2);
/* 1500 */
    nt3hr = (integer) (*t3sub);
    nt3min = (integer) ((*t3sub - nt3hr) * 60.f);
    t1 = nt3min / 60.f;
    nt3sec = (integer) ((*t3sub - nt3hr - t1) * 3600 + .5f);
    if (nt3sec < 0) {
	*negch = -1;
	nt3sec = abs(nt3sec);
    } else {
	*negch = 1;
    }
    if (nt3sec == 60) {
	++nt3min;
	nt3sec = 0;
    }
    if (nt3min == 60) {
	++nt3hr;
	nt3min = 0;
    }
    s_wsfi(&io___269);
    do_fio(&c__1, (char *)&nt3hr, (ftnlen)sizeof(integer));
    e_wsfi();
    s_wsfi(&io___270);
    do_fio(&c__1, (char *)&nt3min, (ftnlen)sizeof(integer));
    e_wsfi();
    if (*(unsigned char *)tsmin == ' ') {
	*(unsigned char *)tsmin = '0';
    }
    s_wsfi(&io___271);
    do_fio(&c__1, (char *)&nt3sec, (ftnlen)sizeof(integer));
    e_wsfi();
    if (*(unsigned char *)tssec == ' ') {
	*(unsigned char *)tssec = '0';
    }
    s_copy(t3s_1.t3subc, tshr, (ftnlen)2, (ftnlen)2);
    *(unsigned char *)&t3s_1.t3subc[2] = ':';
    s_copy(t3s_1.t3subc + 3, tsmin, (ftnlen)2, (ftnlen)2);
    *(unsigned char *)&t3s_1.t3subc[5] = ':';
    s_copy(t3s_1.t3subc + 6, tssec, (ftnlen)2, (ftnlen)2);
    return 0;
} /* t1500_ */

/* Subroutine */ int t1750_(doublereal *dy, integer *nyl, integer *nyr)
{
    /* System generated locals */
    integer i__1;

    /* Builtin functions */
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer s_wsfi(icilist *), do_fio(integer *, char *, ftnlen), e_wsfi(void)
	    ;

    /* Local variables */
    static char ch2[2], ch4[4];
    static integer ndy, nyr1, nleap;

    /* Fortran I/O blocks */
    static icilist io___275 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___276 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___277 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___278 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___279 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___280 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___281 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___282 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___283 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___284 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___285 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___286 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___287 = { 0, ch2, 0, "(I2)", 2, 1 };
    static icilist io___290 = { 0, ch4, 0, "(I4)", 4, 1 };


/*       'convert dy at particular yr to calendar date */
/* 1750 */
    s_copy(t3s_1.caldayc, "           ", (ftnlen)11, (ftnlen)11);
    nleap = 0;
    ndy = (integer) (*dy);
    if (*nyl == 366) {
	nleap = 1;
    }
    if (ndy <= 31) {
	s_copy(t3s_1.caldayc, "Jan-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___275);
	do_fio(&c__1, (char *)&ndy, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (*nyl == 365 && ndy > 31 && ndy <= 59) {
	s_copy(t3s_1.caldayc, "Feb-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___276);
	i__1 = ndy - 31;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
	nleap = 0;
    } else if (*nyl == 366 && ndy > 31 && ndy <= 60) {
	s_copy(t3s_1.caldayc, "Feb-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___277);
	i__1 = ndy - 31;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
	nleap = 1;
    } else if (ndy > nleap + 59 && ndy <= nleap + 90) {
	s_copy(t3s_1.caldayc, "Mar-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___278);
	i__1 = ndy - 59 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 90 && ndy <= nleap + 120) {
	s_copy(t3s_1.caldayc, "Apr-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___279);
	i__1 = ndy - 90 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 120 && ndy <= nleap + 151) {
	s_copy(t3s_1.caldayc, "May-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___280);
	i__1 = ndy - 120 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 151 && ndy <= nleap + 181) {
	s_copy(t3s_1.caldayc, "Jun-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___281);
	i__1 = ndy - 151 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 181 && ndy <= nleap + 212) {
	s_copy(t3s_1.caldayc, "Jul-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___282);
	i__1 = ndy - 181 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 212 && ndy <= nleap + 243) {
	s_copy(t3s_1.caldayc, "Aug-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___283);
	i__1 = ndy - 212 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 243 && ndy <= nleap + 273) {
	s_copy(t3s_1.caldayc, "Sep-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___284);
	i__1 = ndy - 243 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 273 && ndy <= nleap + 304) {
	s_copy(t3s_1.caldayc, "Oct-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___285);
	i__1 = ndy - 273 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 304 && ndy <= nleap + 334) {
	s_copy(t3s_1.caldayc, "Nov-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___286);
	i__1 = ndy - 304 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    } else if (ndy > nleap + 334) {
	s_copy(t3s_1.caldayc, "Dec-", (ftnlen)4, (ftnlen)4);
	s_wsfi(&io___287);
	i__1 = ndy - 334 - nleap;
	do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
	e_wsfi();
    }
    s_copy(t3s_1.caldayc + 4, ch2, (ftnlen)2, (ftnlen)2);
    if (*(unsigned char *)&t3s_1.caldayc[4] == ' ') {
	*(unsigned char *)&t3s_1.caldayc[4] = '0';
    }
    *(unsigned char *)&t3s_1.caldayc[6] = '-';
    nyr1 = *nyr;
    if (abs(nyr1) < 1000) {
	nyr1 = abs(nyr1) + 1000;
    }
    s_wsfi(&io___290);
    i__1 = abs(nyr1);
    do_fio(&c__1, (char *)&i__1, (ftnlen)sizeof(integer));
    e_wsfi();
    s_copy(t3s_1.caldayc + 7, ch4, (ftnlen)4, (ftnlen)4);
    return 0;
} /* t1750_ */

/* Subroutine */ int refvdw_(doublereal *hgt, doublereal *x, doublereal *fref, const short *nweatherflag, const double *vdwsf)
{

    static doublereal ca[11], coef[55];	/* was [5][11] */


	//(1.0e3 * cd) * converts to mrad
	if (*nweatherflag == 1) //use ESAA 3.283-1/-2 (low precision refraction as test of need for VDW refraction)
	{
		//*fref =  *vdwsf * (0.004676 * 1.0e3 * cd)  / tan(cd * (*x + 7.32 / (*x + 4.32)));
		
		if (*x >= 15.0)
		{
			//*fref =  *vdwsf * (0.004676 * 1.0e3 * cd)  / tan(cd * (*x + 7.31 / (*x + 4.4))); 
			*fref =  *vdwsf * (0.004676 * 1.0e3 * cd)  / tan(cd * (*x + 7.32 / (*x + 4.32)));
		}
		else
		{
			//this is G.G. Bennett formula
			//Bennett, G.G.
			//The calculation of astronomical refraction in marine navigation,
			//Journal of Inst. navigation, No. 35, page 255-259, 1982  p. 257
			*fref =  *vdwsf * (1.0e3 * cd) * (0.1594 + 0.0196 * *x + 0.00002 * *x * *x) / (1 + 0.505 * *x + 0.0845 * *x * *x);
		}
		
		return 0;
	}

	//else -- use VDW refraction
	/*
	else
	{
		double am;
		am = *x * 60.0; //convert degrees to arcminues
		*fref = 8.98755719362534
			    - 5.58561739295863E-02 * am
				+ 2.66134862130963E-04 * am * am
				- 1.06415699029082E-06 * am * am * am
				+ 3.64369315470219E-09 * am * am * am * am
				- 1.02333444831416E-11 * am * am * am * am * am
				+ 2.19141928981676E-14 * am * am * am * am * am * am
				- 3.32469432541882E-17 * am * am * am * am * am * am * am
				+ 3.3052943791392E-20 * am * am * am * am * am * am * am * am
				- 1.91266829900908E-23 * am * am * am * am * am * am * am * am * am
				+ 4.85844438121201E-27 * am * am * am * am * am * am * am * am * am * am;
		return 0;
	}
	*/

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
    return 0;
} /* refvdw_ */

/* Subroutine */ int casgeo_(doublereal *gn, doublereal *ge, doublereal *l1, 
	doublereal *l2)
{
    /* System generated locals */
    doublereal d__1;

    /* Builtin functions */
    double sin(doublereal), sqrt(doublereal), tan(doublereal), cos(doublereal);

    /* Local variables */
    static doublereal r__, a2, b2, c1, c2, f1, g1, g2, e4, c3, d1, a3, o1, o2,
	     a4, b3, s1, s2, b4, c4, c5, x1, y1, y2, x2, c6, d2, r3, r4, r2, 
	    a5, d3;
	static logical ggpscorrection = TRUE_;

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
    if (ggpscorrection) {
        //Use the approximate correction factor
        //in order to agree with GPS.
        *l2 -= 0.0013;
        *l1 += 0.0013;
	}
    return 0;
} /* casgeo_ */

/* Main program alias */ int netzski6_ () { MAIN__ (); return 0; }

short Temperatures(double lt, double lg, integer MinTemp[], integer AvgTemp[], integer MaxTemp[] )
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
    
			IKMY = Cint((ULYMAP - Y) / YDIM) + 1;
			IKMX = Cint((X - ULXMAP) / XDIM) + 1;
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

//////////////lenMonth(x) determines number of days in month x, 1<=x<=12////////////////////
//source: https://cmcenroe.me/2014/12/05/days-in-month-formula.html
short lenMonth( short x) {
	return 28 + (short)(x + floor(x/8)) % 2 + 2 % x + 2 * floor(1/x);
}

/////////////////CNK C eumulation library for VB6 functions/////////////////////

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


//////////////////////////ParseString///////////////////////////
short ParseString( char *doclin, char *sep, char arr[][100] )
///////////////////////////////////////////////////////////////////////////////////////////
//Parses a line of text which is composed of tokens separated by "sep" (e.g., ",", " ", etc)
//stuffs the extracted strings into the array, arr
//number of tokens entries to extract = NUMENTRIES
//each token can be a maximum MAXDBLSIZE characters long
//maximum number of data entries that can be extracted is defined by the array size = MAXARRSIZE
//////////////////////////////////////////////////////////////////
{

	char *token;
	short numtoken = 0;

	//get first token
	token = strtok( doclin, sep );
	strcpy( &arr[0][0], token );

	//get other tokens
    while( token != NULL )
    {
       if (numtoken + 1 >= 4) break;
	   token = strtok( NULL, sep );
	   if (token == NULL) return -1; //can't find token in the string
	   numtoken++;
	   strcpy( &arr[numtoken][0], token );
    }

	return 0;
}


///////////////RemoveCRLF/////////////////
char *RemoveCRLF( char *str )
///////////////////////////////////////
//remove the carriage return and line feed that typically end DOS files
///////////////////////////////////////////////
{
	/*
	//everything commented out is deprecated
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
	*/

	//now check last character
	if (!isalnum(str[strlen(str)-1]))
	{
		str[strlen(str)-1] = 0; //terminate the string
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

