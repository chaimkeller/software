/* rdhalba6.f -- translated by f2c (version 20021022).
   You must link the resulting object file with the libraries:
	-lf2c -lm   (in that order)

  VERSION 19 061120 -- added read Maps & More TerrestrialRefraction flag on scanlist.txt line
  VERSION 20 050622 -- added support for running the program on other drive letters
*/

#include <math.h>
#include <stdio.h>
#include <string.h>
//#include <conio.h> -- used for getch diagnostics, uncomment for diagnostics
#include "f2c.h"

//missing intrinisic function from the f2c library
#define hfix_(x) ((shortint)x) //short integer cast
#define jfix_(x) ((integer)x) //long integer cast

short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] );
int WhichJKC();
int WhichVB();

//globals
char testdrv[13] = "cdefghijklmn";
char drivlet[2] = "c";
int WhereJKC,WhereWorldClim;


/* Table of constant values */

static integer c__3 = 3;
static integer c__1 = 1;
static integer c__9 = 9;
static integer c__5 = 5;
static integer c__801 = 801;
static integer c__512 = 512;

/* Main program */ int MAIN__(void)
{
    /* Format strings */
    static char fmt_155[] = "(a,3x,i4,3x,f9.4,3x,f9.4)";
    static char fmt_260[] = "(f8.3,\002,\002,f8.3,\002,\002,f8.2,4(\002,\002"
	    ",f8.3),\002,\002,f6.3)";
    static char fmt_302[] = "(f5.1,\002,\002,f7.4,\002,\002,f9.4,\002,\002,f"
	    "9.4,\002,\002,f7.3,\002,\002,f6.1)";

    /* System generated locals */
    integer i__1, i__2;
    doublereal d__1, d__2, d__3;
    olist o__1;
    cllist cl__1;
    inlist ioin__1;

    /* Builtin functions */
    double acos(doublereal);
    integer f_open(olist *), s_rsle(cilist *), do_lio(integer *, integer *, 
	    char *, ftnlen), e_rsle(void), s_wsle(cilist *), e_wsle(void), 
	    f_clos(cllist *);
    /* Subroutine */ int s_paus(char *, ftnlen);
    double sqrt(doublereal);
    integer i_dnnt(doublereal *);
    double tan(doublereal), cos(doublereal), sin(doublereal), atan(doublereal);
    integer s_wsfe(cilist *), do_fio(integer *, char *, ftnlen), e_wsfe(void),
	     f_inqu(inlist *);
	short Temperatures(double lt, double lg, short MinTemp[], short AvgTemp[], short MaxTemp[] );
    /* Local variables */
    static doublereal startkmx, d__;
    static integer i__, j;
    static doublereal d1, d2, s1, s2, x1, x2, y1, y2, z1, z2, cd, dstpointx, 
	    dstpointy, lg;
    static integer ne;
    static doublereal re, pi;
    static integer nk, nn;
    static doublereal lt, lg1, lg2, re1, re2, az1, x1d, y1d, z1d, az2, lt1, 
	    lt2, x1p, y1p, z1p, x1s, y1s, z1s, bne;
    static logical sea;
    static doublereal hgt, azi;
    static logical new__;
    static doublereal kmx, kmy, bne1, bne2, hgt2, kmy1, kmy2, defb, defm, 
	    mang, anav, hgt21, hgt22;
    static integer ncol;
    static doublereal vang, rekm, aprn, dist, dkmy;
    static real prof[2703]	/* was [3][901] */, dtms[42000]	/* was [3][
	    14000] */;
    static integer ikmx, ikmy;
    extern /* Subroutine */ int getz_(integer *, integer *, doublereal *, 
	    integer *);
    static doublereal kmxo, kmyo;
    static logical netz;
    static doublereal kmxp, kmyp;
    static integer nrow;
    static doublereal vang1, vang2, dist1, dist2;
    static integer nkmx1, ndela;
    static char fnam32[32], fnam33[33], fnam34[34], fnam35[35], fnam36[36];
    static doublereal fudge, deltd, angle, avref, kmycc, distd;
    static logical found;
    static integer nntil, nloop;
    static doublereal kmxoo;
    static integer nnetz;
    static doublereal kmyoo;
    static logical gridcd;
    extern /* Subroutine */ int casgeo_(doublereal *, doublereal *, 
	    doublereal *, doublereal *);
    static char placen[36];
    static integer nocean;
    static doublereal azibeg, delazi, aziend, maxang, begkmx, kmybeg;
    static integer nlenfn, nkmycc;
    static doublereal delkmy, hgtobs, endkmx, kmyend, azicos, azisin, apprnr;
    static integer numazi, numdtm;
    static doublereal sofkmx;
    static integer nskipy;
    static logical exists;
    static integer numkmx, numkmy, numtot;
    static doublereal xentry;
    static integer istart1, nstart2, ndelazi, nbegkmx;
    static doublereal viewang, treehgt, minview;
    static integer nomentr;
    static doublereal skipkmx, skipkmy;
    static integer numskip;
    static doublereal proftmp;

	short MinT[12],AvgT[12],MaxT[12],ier = 0, iTK;
	float MeanTemp = 0;
	double PATHLENGTH = 0.0;
	double RETH = 6356.766;
	double TRpart = 0.0;	
	double Exponent = 1.0;
	logical TerrestrialRefraction = FALSE_; //TRUE_; //set to true to model the terrestrial refraction (default)
	logical TROld_Type = FALSE_; //True to use old terrestrial refraction, False to use new wikipedia expression
	static integer TemperatureModel = 1;  // flag for terrestrial refraction type written to end of each
										  // scanlist.txt line by Maps & MOre /////added 061120///////////
	char Header[255] = "";

    /* Fortran I/O blocks */
    static cilist io___16 = { 0, 1, 0, 0, 0 };
    static cilist io___19 = { 0, 6, 0, 0, 0 };
    static cilist io___21 = { 0, 1, 0, 0, 0 };
    static cilist io___39 = { 0, 6, 0, 0, 0 };
    static cilist io___40 = { 0, 6, 0, 0, 0 };
    static cilist io___131 = { 0, 6, 0, fmt_155, 0 };
    static cilist io___133 = { 0, 3, 0, "(A)", 0 };
    static cilist io___138 = { 0, 3, 0, fmt_260, 0 };
    static cilist io___146 = { 0, 3, 0, fmt_302, 0 };
    static cilist io___147 = { 0, 6, 0, 0, 0 };
    static cilist io___148 = { 0, 1, 0, 0, 0 };


/*       RDHALBAT calculates terrain profiles using J.K.Hall's DTM */
/*       for Windows 2000.  It doesn't use Ram drives, rather */
/*       reads the DTM tiles directly from the hard disk. */
/*       It adds a correction for refraction during the calculation. */
/*       Place information is read from SCANLIST.TXT file */
/*       information on the RAMDRIVE files is read from LDRAM.TXT */
/*       LAST REVISED 05/10/2005 */
/* ----------------------------------------------------- */
/* -------------------ARRAYS----------------------------- */
/*       NOMENT = 20 * MAXANG + 1   <<<<<< */
/*        CHARACTER*2 CHMAP(14,26)*2, CHMNE, CHMNEO */
/*        CHARACTER SF*10, CH*1,FNAM2*36, CH18*18, CH39*39 */
/*        INTEGER*2 IO(512) */
/* ------------------PARAMETERS-------------------------- */
/*       distance between search points (meters) in kmx, kmy */
    dstpointx = 25.;
    dstpointy = 25.;
/*       nearest approach to observation point (kms) for 1,2 search */
    apprnr = .5f;
/*       start search from this kmx */
    startkmx = 210.;
/*        STARTKMX = 153 */
/*       end search at this kmx */
    sofkmx = 235.;
    maxang = 30.;
/* >>>>>> remember to change noment if change maxang <<<<<<<<<< */
    hgtobs = 1.8f;
/*       maxang is angle range, hgtobs is height of instrument or observer */
    treehgt = 0.;
/*       treehgt is height of trees on mountains measured by the DTM */
    netz = TRUE_;
/*       netz is the flag used to calculate netz times */
/*       conv deg to rad */
    gridcd = TRUE_;
/*       gridcd flags rdhall to convert kmxo,kmyo to lg,lt */
/*       if gridcd is false then use listed values of lg,lt */
    pi = acos(-1.);
    cd = pi / 180;
    nkmx1 = 1;
    new__ = TRUE_;
    minview = 10.;
/*       READ IN STATUS FILE, RAM NAMES, INPUT PARAMETERS */
L8:
	////////////////////////////////////////////////////////
	/*
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
	/*
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
	s_rsfe(&io___16);
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
	*/

	WhereJKC = WhichJKC();
	if (WhereJKC == 0)
	{
		printf("Can't find the jk_c directory, aborting...");
		exit(1);
	}
	else
	{
		drivlet[0] = testdrv[WhereJKC - 1];
		drivlet[1] = 0;
	}

	//////////////////////////////////////////////////////
    o__1.oerr = 0;
    o__1.ounit = 1;
    o__1.ofnmlen = 18;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":\\jk_c\\STATUS.TXT" );
    //o__1.ofnm = "C:\\jk_c\\STATUS.TXT";
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
    s_rsle(&io___16);
    do_lio(&c__3, &c__1, (char *)&nloop, (ftnlen)sizeof(integer));
    do_lio(&c__3, &c__1, (char *)&numtot, (ftnlen)sizeof(integer));
    e_rsle();
    if (nloop > numtot) {
	s_wsle(&io___19);
	do_lio(&c__9, &c__1, "SCAN FINISHED", (ftnlen)13);
	e_wsle();
	goto L999;
    }
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
    o__1.oerr = 0;
    o__1.ounit = 1;
    o__1.ofnmlen = 20;
    //o__1.ofnm = "C:\\jk_c\\SCANLIST.TXT";
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":\\jk_c\\SCANLIST.TXT" );
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
    i__1 = nloop;
    for (i__ = 1; i__ <= i__1; ++i__) {
	s_rsle(&io___21);
	do_lio(&c__3, &c__1, (char *)&nn, (ftnlen)sizeof(integer));
	do_lio(&c__9, &c__1, placen, (ftnlen)36);
	do_lio(&c__5, &c__1, (char *)&kmxo, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&kmyo, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&hgt, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&startkmx, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&sofkmx, (ftnlen)sizeof(doublereal));
	do_lio(&c__3, &c__1, (char *)&nnetz, (ftnlen)sizeof(integer));
	do_lio(&c__5, &c__1, (char *)&mang, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&aprn, (ftnlen)sizeof(doublereal));
	do_lio(&c__3, &c__1, (char *)&TemperatureModel, (ftnlen)sizeof(integer));

	switch (TemperatureModel) 
	{
	case -1:
		//old terrestrial refraction model
		TerrestrialRefraction = TRUE_;
		TROld_Type = TRUE_;
		break;
	case 0:
		//no terrestrial refraction
		TerrestrialRefraction = FALSE_;
		TROld_Type = FALSE_;
		break;
	case 1:
		//new terrestrial refraction model of Bomford - Wikipedia model
		//using mean temperature
		TerrestrialRefraction = TRUE_;
		TROld_Type = FALSE_;
		break;
	case 2:
		//new terrestrial refraction model of Bomford but remove from profile file
		TerrestrialRefraction = TRUE_;
		TROld_Type = FALSE_;
		break;
	default:
		TerrestrialRefraction = TRUE_;
		TROld_Type = FALSE_;
		break;
	}

/*           READ(1,27)NN,PLACEN,KMXO,KMYO,HGT,STARTKMX,SOFKMX,NNETZ, */
/*     *              MANG,APRN */
/* L25: */
    }
/* 27      FORMAT(I4,2X,A36,1X,2(F8.3,1X),F7.1,1X,2(F9.3,1X),I1,1X, */
/*     *        F4.1,1X,F6.3) */
    if (mang != 0.) {
	maxang = mang;
    }
    if (maxang > 45.) {
	maxang = 45.;
    }
    mang = maxang;
    nomentr = (integer) (maxang * 20 + 1);
    if (aprn > 0.) {
	apprnr = aprn;
    }
    if (nnetz == 0) {
	netz = FALSE_;
    }
    if (nnetz == 1) {
	netz = TRUE_;
    }
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
    gridcd = TRUE_;
/*       initialize interpolation array */
    for (i__ = 1; i__ <= 3; ++i__) {
	i__1 = nomentr;
	for (j = 1; j <= i__1; ++j) {
	    prof[i__ + j * 3 - 4] = -100.f;
/* L28: */
	}
/* L29: */
    }
    nlenfn = 36;
    for (i__ = 1; i__ <= 36; ++i__) {
	if (i__ <= 32) {
	    *(unsigned char *)&fnam32[i__ - 1] = *(unsigned char *)&placen[
		    i__ - 1];
	}
	if (i__ <= 33) {
	    *(unsigned char *)&fnam33[i__ - 1] = *(unsigned char *)&placen[
		    i__ - 1];
	}
	if (i__ <= 34) {
	    *(unsigned char *)&fnam34[i__ - 1] = *(unsigned char *)&placen[
		    i__ - 1];
	}
	if (i__ <= 35) {
	    *(unsigned char *)&fnam35[i__ - 1] = *(unsigned char *)&placen[
		    i__ - 1];
	}
	if (i__ <= 36) {
	    *(unsigned char *)&fnam36[i__ - 1] = *(unsigned char *)&placen[
		    i__ - 1];
	}
	if (*(unsigned char *)&placen[i__ - 1] == 32) {
	    nlenfn = i__;
	    if (i__ < 32) {
		s_wsle(&io___39);
		do_lio(&c__9, &c__1, "FILE: ", (ftnlen)6);
		do_lio(&c__9, &c__1, placen, (ftnlen)36);
		do_lio(&c__9, &c__1, "HAS BLANK!--SKIPPING IT", (ftnlen)23);
		e_wsle();
		s_paus("PRESS ANY KEY TO CONTINUE TO NEXT FILE", (ftnlen)38);
		goto L500;
	    }
	    goto L32;
	}
/* L30: */
    }
L32:
    s_wsle(&io___40);
    do_lio(&c__9, &c__1, "SCAN: ", (ftnlen)6);
    do_lio(&c__9, &c__1, placen, (ftnlen)36);
    e_wsle();
    kmyoo = kmyo;
    kmxoo = kmxo;
    if (gridcd) {
		casgeo_(&kmyoo, &kmxoo, &lt, &lg);
	/*          this subroutine converts the ISREAL grid coordinates to lt,lg */
	/*          using the Cassini (Transverse Cylindrical) Projection */

		//now determine the mean average temperatures
		ier = Temperatures(lt,-lg, MinT, AvgT, MaxT);

		if (ier == 0) {

			MeanTemp = 0.0;
			for (iTK = 0; iTK < 12; iTK++) {
				MeanTemp += AvgT[iTK];

			/*
			if (nnetz == 1) { 
				//sunrise horizon, use minimum temperatures
				for (iTK = 0; iTK < 12; iTK++) {
					MeanTemp += MinT[iTK];
				}
			}else{
				//sunset horizon, use average temperatures
				for (iTK = 0; iTK < 12; iTK++) {
					MeanTemp += AvgT[iTK];
				}
			*/
			}

			MeanTemp /= 12.0; //take average over year
			MeanTemp += 273.15; //convert to degrees Kelvin
			/* diagnostics
			printf("latitude, longitude, Mean Temperatures = %lg,%lg,%lg\n",lt,-lg,MeanTemp);
			getch(); */

		} else {
			//use default mean temperatures of 15 degrees Celsius
			MeanTemp = 288.15;
		}

		//calculate conserved portion of terrain refraction based on Wikipedia's expression
		//lapse rate is set at -0.0065 K/m, Ground Pressure = 1013.25 mb
		TRpart = 8.15 * 1013.25 * 1000.0 * (0.0342 - 0.0065)/(MeanTemp * MeanTemp * 3600); //units of deg/km

		//diagnostics
		/*
		printf("MeanTemp,TRpart = %lg,%lg\n",MeanTemp, TRpart);
		getch();*/

    }
    hgt += hgtobs;
    if (new__) {
/*          NEW SCAN SIGNLED--SET DEFAULT VIEW ANGLES */
	if (hgt >= 0.) {
	    anav = sqrt(hgt) * -.0348f;
	}
	if (hgt < 0.) {
	    anav = -1.f;
	}
	i__1 = nomentr;
	for (i__ = 1; i__ <= i__1; ++i__) {
	    prof[i__ * 3 - 3] = anav - minview;
/* L35: */
	}
	new__ = FALSE_;
    }
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
    d__1 = begkmx * 40.f;
    nbegkmx = i_dnnt(&d__1);
    begkmx = nbegkmx * .025f;
    skipkmx = dstpointx * .001f;
    skipkmy = dstpointy * .001f;
    d__2 = (d__1 = endkmx - begkmx, abs(d__1)) / skipkmx;
    numkmx = i_dnnt(&d__2) + 1;
    numskip = 1;
    i__ = nkmx1;
L40:
    kmx = begkmx + (i__ - 1) * skipkmx;
    ikmx = (integer) ((kmx + 20) * 40) + 1;
/*          now determine desired range in kmy */
    if (endkmx > kmxo + 10) {
/*             if distance >= 20 km, then narrow azimuthal range to MAXANG = 30 */
	if ((d__1 = kmx - begkmx, abs(d__1)) >= 20.) {
	    maxang = 30.;
	}
/*             if distance >= 40 km, then only 50m grid */
	if ((d__1 = kmx - begkmx, abs(d__1)) >= 40.) {
	    numskip = 2;
	}
    } else if (begkmx < kmxo - 10) {
/*             if distance >= 20 km, then narrow azimuthal range to MAXANG = 30 */
	if ((d__1 = kmx - endkmx, abs(d__1)) > 20.) {
	    maxang = 30.;
	}
/*             if distance >= 40 km, then only 50m grid */
	if ((d__1 = kmx - endkmx, abs(d__1)) > 40.) {
	    numskip = 2;
	}
/*             if distance <= 20 then azimuthal range = 45 */
	if ((d__1 = kmx - endkmx, abs(d__1)) <= 20.) {
	    maxang = 45.;
	}
/*             if distance <= 40 km, then only 50m grid */
	if ((d__1 = kmx - endkmx, abs(d__1)) <= 40.) {
	    numskip = 1;
	}
    }
    dkmy = (d__1 = (kmx - kmxoo) * tan((maxang + 2) * cd), abs(d__1));
    numkmy = (integer) (dkmy * 2 / skipkmy + 1);
    d__1 = (kmyoo - dkmy) * 40;
    nkmycc = i_dnnt(&d__1);
    kmycc = nkmycc * .025f;
    numdtm = (integer) ((mang - maxang) * 10);
    sea = FALSE_;
    nocean = 0;
    i__1 = numkmy;
    i__2 = numskip;
    for (j = 1; i__2 < 0 ? j >= i__1 : j <= i__1; j += i__2) {
	kmy = kmycc + skipkmy * (j - 1);
/*             this subroutine converts kmx,kmy to lat,long */
	if (sea) {
	    hgt2 = 0.f;
	    ++nskipy;
	    kmy = kmybeg + delkmy * nskipy;
	    casgeo_(&kmy, &kmx, &lt2, &lg2);
	    if (kmy <= kmyend) {
		goto L45;
	    } else if (kmy > kmyend) {
		goto L125;
	    }
	}
	casgeo_(&kmy, &kmx, &lt2, &lg2);
	ikmy = (integer) ((380 - kmy) * 40) + 1;
/*              IKMY = NINT((380 - KMY)*40) + 1 */
/*             this subroutine reads J.K. Hall's DTM at any (kmy,kmx) and returns the altitude */
	nrow = ikmy;
	ncol = ikmx;
	getz_(&ikmy, &ikmx, &hgt2, &nntil);
/*             ignore ocean depths for sunset calculations */
	if (hgt2 < 0. && (kmx < 160. || kmy >= 260.)) {
	    if (! netz) {
		hgt2 = 0.f;
		++nocean;
		if (nocean > 80) {
		    kmybeg = kmy;
		    kmyend = kmycc + skipkmy * (numkmy - 1);
		    numazi = numkmy - j - 1;
		    azibeg = dtms[(numdtm - 1) * 3 - 3];
		    aziend = maxang + 2;
/*                       numazi=(aziend-azibeg)*20 */
		    delkmy = (kmyend - kmybeg) * 2.2f / numazi;
		    if (numazi <= 30) {
			nocean = 0;
			goto L45;
		    }
		    nskipy = 0;
		    sea = TRUE_;
		}
	    }
	}
L45:
	fudge = treehgt;
	hgt2 += fudge;
	lt1 = lt;
	lg1 = lg;
	x1 = cos(lt * cd) * cos(lg * cd);
	x2 = cos(lt2 * cd) * cos(lg2 * cd);
	y1 = cos(lt * cd) * sin(lg * cd);
	y2 = cos(lt2 * cd) * sin(lg2 * cd);
	z1 = sin(lt * cd);
	z2 = sin(lt2 * cd);
	re = 6371315.;
	rekm = 6371.315f;
/* Computing 2nd power */
	d__1 = x1 - x2;
/* Computing 2nd power */
	d__2 = y1 - y2;
/* Computing 2nd power */
	d__3 = z1 - z2;
	distd = rekm * sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
/*              ra = 1.00336414 */
/*              ra = 0.99664714 */
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
	viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
	d__ = (dist1 - dist2 * cos(angle)) / dist1;
	x1d = x1 * (1 - d__) - x2;
	y1d = y1 * (1 - d__) - y2;
	z1d = z1 * (1 - d__) - z2;
	x1p = -sin(lg * cd);
	y1p = cos(lg * cd);
	z1p = 0.;
	azicos = x1p * x1d + y1p * y1d;
	x1s = -cos(lg * cd) * sin(lt * cd);
	y1s = -sin(lg * cd) * sin(lt * cd);
	z1s = cos(lt * cd);
	azisin = x1s * x1d + y1s * y1d + z1s * z1d;
	azi = atan(azisin / azicos);
/*             add contribution of atmospheric refraction to view angle */

	if (TerrestrialRefraction) {

		if (TROld_Type) {
			//model terrestrial refraction
			//as of this version, use new calculation of terrestrial refraction
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

		}else{

			//use terrestrial refraction formula from Wikipedia, suitably modified to fit van der Werf ray tracing
			PATHLENGTH = sqrt(distd*distd + pow(fabs(deltd)*0.001 - 0.5*(distd*distd/RETH),2.0));
			if (fabs(deltd) > 1000.0) {
				Exponent = 0.99;
			}else{
				Exponent = 0.9965;
			}

			avref = TRpart * pow(PATHLENGTH, Exponent); //degrees

			//diagnostics
			/*
			if (avref > 0.03) {
				printf("TRpart,deltd,distd,pathlength,exponent,avref=%lg,%lg,%lg,%lg,%lg,%lg\n",TRpart,deltd,distd,PATHLENGTH,Exponent,avref);
				getch();
			}
			*/
		}
	}else{
		avref = 0.0;
	}

	++numdtm;
	dtms[numdtm * 3 - 3] = azi / cd;
	dtms[numdtm * 3 - 2] = viewang / cd + avref;
	dtms[numdtm * 3 - 1] = kmy;
/* L120: */
    }
/*       interpolate into azimuth file: prof() */
L125:
    istart1 = (integer) ((mang - maxang) * 10 + 1);
L150:
    if (netz) {
	az2 = dtms[istart1 * 3 - 3];
	vang2 = dtms[istart1 * 3 - 2];
	kmy2 = dtms[istart1 * 3 - 1];
    } else {
	az1 = dtms[istart1 * 3 - 3];
	vang1 = dtms[istart1 * 3 - 2];
	kmy1 = dtms[istart1 * 3 - 1];
    }
    nstart2 = istart1 + 1;
L160:
    if (netz) {
	az1 = dtms[nstart2 * 3 - 3];
	vang1 = dtms[nstart2 * 3 - 2];
	kmy1 = dtms[nstart2 * 3 - 1];
    } else {
	az2 = dtms[nstart2 * 3 - 3];
	vang2 = dtms[nstart2 * 3 - 2];
	kmy2 = dtms[nstart2 * 3 - 1];
    }
/* L162: */
    if (az2 >= -mang && az1 <= mang) {
	bne1 = (integer) (az1 * 10) / 10.f - .1f;
	bne2 = (integer) (az2 * 10) / 10.f + .1f;
	found = FALSE_;
	ndela = (integer) ((bne2 - bne1) / .1f + 1);
	i__2 = ndela;
	for (ndelazi = 1; ndelazi <= i__2; ++ndelazi) {
	    delazi = bne1 + (ndelazi - 1) * .1f;
	    if (delazi < -mang) {
		goto L165;
	    }
	    if (delazi > mang + .1f) {
		goto L170;
	    }
	    if (az1 <= delazi && az2 >= delazi) {
		bne = (delazi + mang) * 10.f + 1.f;
		ne = i_dnnt(&bne);
		if (ne >= 0 && ne <= nomentr) {
		    vang = (vang2 - vang1) / (az2 - az1) * (delazi - az1) + 
			    vang1;
		    proftmp = prof[ne * 3 - 3];
		    if (vang >= proftmp) {
			prof[ne * 3 - 3] = vang;
			prof[ne * 3 - 2] = begkmx + (i__ - 1) * skipkmx;
			prof[ne * 3 - 1] = (kmy2 - kmy1) / (az2 - az1) * (
				delazi - az1) + kmy1;
		    }
		    found = TRUE_;
		}
	    } else if (az1 < delazi && az2 < delazi) {
		goto L170;
	    }
L165:
	    ;
	}
L170:
	if (found) {
	    goto L180;
	}
	++nstart2;
	if (nstart2 <= numdtm) {
	    goto L160;
	} else {
	    goto L180;
	}
    }
L180:
    ++istart1;
    if (istart1 < numdtm) {
	goto L150;
    }
    s_wsfe(&io___131);
    do_fio(&c__1, " completed interpreting: ", (ftnlen)25);
    do_fio(&c__1, (char *)&i__, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&kmx, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&kmy, (ftnlen)sizeof(doublereal));
    e_wsfe();
/* L550: */
    i__ += numskip;
    if (i__ <= numkmx) {
	goto L40;
    }
    if (nlenfn == 32) {
	ioin__1.inerr = 0;
	ioin__1.infilen = 32;
	ioin__1.infile = fnam32;
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
	    o__1.ofnmlen = 32;
	    o__1.ofnm = fnam32;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 32;
	    o__1.ofnm = fnam32;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
    } else if (nlenfn == 33) {
	ioin__1.inerr = 0;
	ioin__1.infilen = 33;
	ioin__1.infile = fnam33;
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
	    o__1.ofnmlen = 33;
	    o__1.ofnm = fnam33;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 33;
	    o__1.ofnm = fnam33;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
    } else if (nlenfn == 34) {
	ioin__1.inerr = 0;
	ioin__1.infilen = 34;
	ioin__1.infile = fnam34;
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
	    o__1.ofnmlen = 34;
	    o__1.ofnm = fnam34;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 34;
	    o__1.ofnm = fnam34;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
    } else if (nlenfn == 35) {
	ioin__1.inerr = 0;
	ioin__1.infilen = 35;
	ioin__1.infile = fnam35;
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
	    o__1.ofnmlen = 35;
	    o__1.ofnm = fnam35;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 35;
	    o__1.ofnm = fnam35;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
    } else if (nlenfn == 36) {
	ioin__1.inerr = 0;
	ioin__1.infilen = 36;
	ioin__1.infile = fnam36;
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
	    o__1.ofnmlen = 36;
	    o__1.ofnm = fnam36;
	    o__1.orl = 0;
	    o__1.osta = "OLD";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	} else {
	    o__1.oerr = 0;
	    o__1.ounit = 3;
	    o__1.ofnmlen = 36;
	    o__1.ofnm = fnam36;
	    o__1.orl = 0;
	    o__1.osta = "NEW";
	    o__1.oacc = 0;
	    o__1.ofm = 0;
	    o__1.oblnk = 0;
	    f_open(&o__1);
	}
    }

    s_wsfe(&io___133);

	/////////////////////////////////////comment 101520/////////////////////////////////////
	//			rdhalba4 correctly records the calculation hgt and not the ground height
	///////////////////////////////////////////////////////////////////////////////////////

	if (TerrestrialRefraction) {
		if (TROld_Type) { //old version of refraction
		    do_fio(&c__1, "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR", (ftnlen)46);
		}else{ //new Wikipedia version
			if (TemperatureModel == 0 || TemperatureModel == 2) {
				do_fio(&c__1, "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR,0", (ftnlen)48);
			} else if (TemperatureModel == 1) {
				sprintf( Header, "%s,%6.2f","kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy", MeanTemp );
				do_fio(&c__1, Header, (ftnlen)46 );
			}
		}
	}else{ //no terrestrial refraction added
		    do_fio(&c__1, "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR,0", (ftnlen)48);
	}

    e_wsfe();
    d1 = dstpointx;
    d2 = dstpointy;
    s1 = startkmx;
    s2 = sofkmx;
    s_wsfe(&io___138);
    do_fio(&c__1, (char *)&kmxo, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&kmyo, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&hgt, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&s1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&s2, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&d1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&d2, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&apprnr, (ftnlen)sizeof(doublereal));
    e_wsfe();
    i__2 = nomentr;
    for (nk = 1; nk <= i__2; ++nk) {
	xentry = -mang + (nk - 1) / 10.f;
/*          find distances and heights to peaks */
	kmxp = prof[nk * 3 - 2];
	kmyp = prof[nk * 3 - 1];
	ikmx = (integer) ((kmxp + 20) * 40) + 1;
	ikmy = (integer) ((380 - kmyp) * 40) + 1;
	nrow = ikmy;
	ncol = ikmx;
	getz_(&ikmy, &ikmx, &hgt2, &nntil);
	if (hgt2 < 0. && (kmxp < 185. || kmyp >= 260.)) {
	    if (! netz) {
		hgt2 = 0.f;
	    }
	}
	if (hgt2 < 0. && netz) {
	    i__1 = ikmy - 1;
	    getz_(&i__1, &ikmx, &hgt21, &nntil);
	    i__1 = ikmy + 1;
	    getz_(&i__1, &ikmx, &hgt22, &nntil);
	    hgt2 = (hgt21 + hgt22) / 2;
	}
/* Computing 2nd power */
	d__1 = kmxo - kmxp;
/* Computing 2nd power */
	d__2 = kmyo - kmyp;
	dist = sqrt(d__1 * d__1 + d__2 * d__2);
	distd = hgt - hgt2;

	if (TerrestrialRefraction && TemperatureModel == 2) { //remove the added terrestrial refraction
		PATHLENGTH = sqrt(dist*dist + pow(fabs(deltd)*0.001 - 0.5*(dist*dist/RETH),2.0));
		if (fabs(deltd) > 1000.0) {
			Exponent = 0.99;
		}else{
			Exponent = 0.9965;
		}

		avref = TRpart * pow(PATHLENGTH, Exponent); //degrees
		prof[nk * 3 - 3] -= avref;
	}

	s_wsfe(&io___146);
	do_fio(&c__1, (char *)&xentry, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&prof[nk * 3 - 3], (ftnlen)sizeof(real));
	do_fio(&c__1, (char *)&prof[nk * 3 - 2], (ftnlen)sizeof(real));
	do_fio(&c__1, (char *)&prof[nk * 3 - 1], (ftnlen)sizeof(real));
	do_fio(&c__1, (char *)&dist, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&hgt2, (ftnlen)sizeof(doublereal));
	e_wsfe();
/* L300: */
    }
    cl__1.cerr = 0;
    cl__1.cunit = 3;
    cl__1.csta = 0;
    f_clos(&cl__1);
    s_wsle(&io___147);
    do_lio(&c__9, &c__1, " profile copied to permanent file: ", (ftnlen)35);
    do_lio(&c__9, &c__1, placen, (ftnlen)36);
    e_wsle();
/*       Update STATUS.TXT file */
L500:
    o__1.oerr = 0;
    o__1.ounit = 1;
    o__1.ofnmlen = 18;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":\\jk_c\\STATUS.TXT" );
    //o__1.ofnm = "C:\\jk_c\\STATUS.TXT";
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
    ++nloop;
    s_wsle(&io___148);
    do_lio(&c__3, &c__1, (char *)&nloop, (ftnlen)sizeof(integer));
    do_lio(&c__3, &c__1, (char *)&numtot, (ftnlen)sizeof(integer));
    e_wsle();
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
    if (nloop > numtot) {
	s_paus("!!CALCULATIONS COMPLETED!!", (ftnlen)26);
    } else {
/*          go back to do next place */
	goto L8;
    }
L999:
    return 0;
} /* MAIN__ */

/* Subroutine */ int getz_(integer *nrow, integer *ncol, doublereal *zm, 
	integer *nntil)
{
    /* Initialized data */

    static char chmneo[2] = "XX";
    static integer ib = 0;
    static char dtmfol[1] = "X";

    /* Format strings */
    static char fmt_13[] = "(a78)";
    static char fmt_30[] = "(\002 ** ERROR IN GETZ ** \002,a2,8i6)";
    static char fmt_23[] = "(1x,a2,f8.1,2f8.3,6i7)";

	static integer t1,t2,t3;

    /* System generated locals */
    integer i__1;
    olist o__1;
    cllist cl__1;
    inlist ioin__1;

    /* Builtin functions */
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer f_inqu(inlist *), f_open(olist *), s_rsfe(cilist *), do_fio(
	    integer *, char *, ftnlen), e_rsfe(void), f_clos(cllist *), s_cmp(
	    char *, char *, ftnlen, ftnlen), s_rdue(cilist *), do_uio(integer 
	    *, char *, ftnlen), e_rdue(void), s_wsfe(cilist *), e_wsfe(void);
    /* Subroutine */ int s_paus(char *, ftnlen);
    integer s_wsle(cilist *), do_lio(integer *, integer *, char *, ftnlen), 
	    e_wsle(void);

    /* Local variables */
    static char testdriv[10];
    static integer i__, j, n, ic;
    static doublereal ge, gn;
    static integer ni;
    static shortint io[512];
    static char sf[9];
    static integer ir, iz;
    static doublereal kx, ky;
    static integer nn1, nn2, ifn, ios, nrec;
    //extern doublereal hfix_(integer *);
    //extern integer jfix_(integer *);
    static char chmap[2*14*26], chmne[2], dtmfn[18], doclin[78], dtmdrv[4];
    static logical exists;

    /* Fortran I/O blocks */
    static cilist io___154 = { 0, 5, 0, "(A4)", 0 };
    static cilist io___158 = { 0, 5, 1, fmt_13, 0 };
    static cilist io___161 = { 0, 5, 1, fmt_13, 0 };
    static cilist io___174 = { 1, 5, 1, 0, 0 };
    static cilist io___177 = { 0, 6, 0, fmt_30, 0 };
    static cilist io___180 = { 0, 6, 0, fmt_23, 0 };
    static cilist io___183 = { 0, 6, 0, 0, 0 };


    if (ib != 0) {
	goto L15;
    }
/*       PULL IN THE DTM TILE SCHEMATIC */
    s_copy(dtmfn + 1, ":\\dtm\\dtm-map.loc", (ftnlen)17, (ftnlen)17);
/*       Its location should be in c:\jk_c\mapcdi~1.sav */
    ioin__1.inerr = 0;
    ioin__1.infilen = 20;
    ioin__1.infile = "C:\\jk_c\\MAPCDI~1.SAV";
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
	o__1.ounit = 5;
	o__1.ofnmlen = 20;
	o__1.ofnm = "C:\\jk_c\\MAPCDI~1.SAV";
	o__1.orl = 0;
	o__1.osta = "OLD";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
	s_rsfe(&io___154);
	do_fio(&c__1, dtmdrv, (ftnlen)4);
	e_rsfe();
	*(unsigned char *)dtmfn = *(unsigned char *)dtmdrv;
	*(unsigned char *)&dtmfol[0] = *(unsigned char *)dtmdrv;
	cl__1.cerr = 0;
	cl__1.cunit = 5;
	cl__1.csta = 0;
	f_clos(&cl__1);
    } else {
/*          search for it (assumes contiguous acessible drives) */
	s_copy(testdriv, "cdefghijkl", (ftnlen)10, (ftnlen)10);
	for (i__ = 1; i__ <= 10; ++i__) {
	    *(unsigned char *)dtmfn = *(unsigned char *)&testdriv[i__ - 1];
	    *(unsigned char *)&dtmfol[0] = *(unsigned char *)dtmfn;
	    ioin__1.inerr = 0;
	    ioin__1.infilen = 18;
	    ioin__1.infile = dtmfn;
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
/*                found default drive */
		goto L3;
	    }
/* L2: */
	}
    }
L3:
    o__1.oerr = 0;
    o__1.ounit = 5;
    o__1.ofnmlen = 18;
    o__1.ofnm = dtmfn;
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
/*       READ THE HEADER LINES */
    for (i__ = 1; i__ <= 3; ++i__) {
	i__1 = s_rsfe(&io___158);
	if (i__1 != 0) {
	    goto L14;
	}
	i__1 = do_fio(&c__1, doclin, (ftnlen)78);
	if (i__1 != 0) {
	    goto L14;
	}
	i__1 = e_rsfe();
	if (i__1 != 0) {
	    goto L14;
	}
/* L5: */
    }
    n = 0;
    for (i__ = 4; i__ <= 54; ++i__) {
	i__1 = s_rsfe(&io___161);
	if (i__1 != 0) {
	    goto L14;
	}
	i__1 = do_fio(&c__1, doclin, (ftnlen)78);
	if (i__1 != 0) {
	    goto L14;
	}
	i__1 = e_rsfe();
	if (i__1 != 0) {
	    goto L14;
	}
	if (i__ % 2 == 0) {
	    ++n;
	    for (j = 1; j <= 14; ++j) {
		nn1 = (j - 1) * 5 + 6;
		nn2 = nn1 + 1;
		s_copy(chmap + (j + n * 14 - 15 << 1), doclin + (nn1 - 1), (
			ftnlen)2, nn2 - (nn1 - 1));
/*            PRINT *,'CHMAP(', J,N,' ) = ', CHMAP(J,N) */
/*            PAUSE 'continue-->' */
/* L7: */
	    }
	}
/* L9: */
    }
L14:
    cl__1.cerr = 0;
    cl__1.cunit = 5;
    cl__1.csta = 0;
    f_clos(&cl__1);
    ib = 1;
/*       GETZ FINDS THE HEIGHT OF A POINT AT THE NORW AND NCOL FROM 380N */
/*       AND -20E WHERE 1,1 IS THAT CORNER POINT */
/*       FIND THE PROPER INDICES I,J OF THE PROPER .SUM FILE */
L15:
    j = (*nrow - 2) / 800 + 1;
    i__ = (*ncol - 2) / 800 + 1;
    s_copy(chmne, chmap + (i__ + j * 14 - 15 << 1), (ftnlen)2, (ftnlen)2);
    if (s_cmp(chmne, "  ", (ftnlen)2, (ftnlen)2) == 0) {
	goto L35;
    }
    if (s_cmp(chmne, chmneo, (ftnlen)2, (ftnlen)2) == 0) {
	goto L21;
    }
    cl__1.cerr = 0;
    cl__1.cunit = 5;
    cl__1.csta = 0;
    f_clos(&cl__1);
    s_copy(sf, dtmfn, (ftnlen)7, (ftnlen)7);
    s_copy(sf + 7, chmne, (ftnlen)2, (ftnlen)2);
    o__1.oerr = 1;
    o__1.ounit = 5;
    o__1.ofnmlen = 9;
    o__1.ofnm = sf;
    o__1.orl = 1024;
    o__1.osta = "OLD";
    o__1.oacc = "DIRECT";
    o__1.ofm = 0;
    o__1.oblnk = 0;
    i__1 = f_open(&o__1);
    if (i__1 != 0) {
	goto L35;
    }
    s_copy(chmneo, chmne, (ftnlen)2, (ftnlen)2);
/*       CONVERT TO GRID LOCATION IN .SUM FILE */
L21:
    ir = *nrow - (j - 1) * 800;
    ic = *ncol - (i__ - 1) * 800;
    i__1 = ir - 1;
	//t1 = jfix_(i__1);
	//t2 = jfix_(c__801);
	//t3 = jfix_(ic);
    ifn = jfix_(i__1) * jfix_(c__801) + jfix_(ic);
    i__1 = (ifn - 1) / 512;
    nrec = (integer) (hfix_(i__1) + 1);
    i__1 = (ifn - 1) % 512;
    ni = (integer) (hfix_(i__1) + 1);
    io___174.cirec = nrec;
    ios = s_rdue(&io___174);
    if (ios != 0) {
	goto L100001;
    }
    ios = do_uio(&c__512, (char *)&io[0], (ftnlen)sizeof(shortint));
    if (ios != 0) {
	goto L100001;
    }
    ios = e_rdue();
L100001:
    if (ios > 0) {
	goto L25;
    }
/*        NREC=NREC+1 */
    iz = io[ni - 1];
    *zm = (real) iz / 10.f;
    return 0;
/*       ERROR CATCHING SECTION */
L25:
    s_wsfe(&io___177);
    do_fio(&c__1, chmne, (ftnlen)2);
    do_fio(&c__1, (char *)&nrec, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&(*nrow), (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&(*ncol), (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&i__, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&j, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&ir, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&ic, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&ios, (ftnlen)sizeof(integer));
    e_wsfe();
    gn = 380.f - (*nrow - 1) * .025f;
    ge = (*ncol - 1) * .025f - 20.f;
    s_wsfe(&io___180);
    do_fio(&c__1, chmne, (ftnlen)2);
    do_fio(&c__1, (char *)&(*zm), (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&gn, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&ge, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&i__, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&j, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&ir, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&ic, (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&(*nrow), (ftnlen)sizeof(integer));
    do_fio(&c__1, (char *)&(*ncol), (ftnlen)sizeof(integer));
    e_wsfe();
    s_paus("continue-->", (ftnlen)11);
    return 0;
/*       NO SUCH MAP SHEET */
L35:
    *zm = -3276.7f;
    kx = (*ncol - 1) * .025f - 20;
    ky = 380 - (*nrow - 1) * .025f;
/*       check if landed on blank tiles in DTM */
    if (kx >= 240. && ky <= 20.) {
    } else if (kx <= 60.1f) {
    } else if (kx >= 160. && ky <= -119.99f) {
    } else if (kx <= 80.1f && ky <= -59.99f) {
    } else if (ky <= -140. || ky >= 380.) {
    } else if (kx <= 80.1f && ky >= 139.99f) {
    } else if (kx <= 120.1f && ky >= 179.99f) {
    } else if (kx <= 140.1f && ky >= 259.99f) {
    } else if (kx <= 160.1f && ky >= 299.99f) {
    } else if (kx <= 180.1f && ky >= 319.99f) {
    } else {
	s_wsle(&io___183);
	do_lio(&c__9, &c__1, "kmx,kmy,TILE-NAME=", (ftnlen)18);
	do_lio(&c__5, &c__1, (char *)&kx, (ftnlen)sizeof(doublereal));
	do_lio(&c__5, &c__1, (char *)&ky, (ftnlen)sizeof(doublereal));
	do_lio(&c__9, &c__1, " ", (ftnlen)1);
	do_lio(&c__9, &c__1, sf, (ftnlen)9);
	e_wsle();
	s_paus("!!!no such map sheet!!!", (ftnlen)23);
    }
    return 0;
} /* getz_ */

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


/*        REAL*4 GN,GE */
/*        PRINT *,'cas',GN,GE */
    g1 = *gn * 1e3;
    g2 = *ge * 1e3;
    r__ = 57.2957795131;
    b2 = .03246816;
    f1 = 206264.806247096;
    s1 = 126763.49;
    s2 = 114242.75;
    e4 = .006803480836;
    c1 = .0325600414007;
    c2 = 2.55240717534e-9;
    c3 = .032338519783;
    x1 = 1170251.56;
    y1 = 1126867.91;
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
    y1 += -1e6;
L10:
    x1 = x2 - x1;
    y1 = y2 - y1;
    d1 = y1 * b2 / 2.;
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
    c6 /= 3.;
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

/* Main program alias */ int rdhalba6_ () { MAIN__ (); return 0; }


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

	WhereWorldClim = WhichVB();
	if (WhereWorldClim != 0)
	{
		drivlet[1] = testdrv[WhereWorldClim - 1];
		drivlet[2] = 0;
		strcpy(FilePathBil, drivlet[0]);
		strcat(FilePathBil,":\\DevStudio\\Vb\\WorldClim_bil");
	}
	else
	{
		printf("Can't find the WorldClim files");
		exit(1);
	}

	//strcpy(FilePathBil,"c:\\DevStudio\\Vb\\WorldClim_bil");

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
	//find drive letter where the WorldClim_bil folder resides

	if (Tempmode == 0) { // read the minimum temperatures for this lat,lon
		strcat(FilePathBil,"//min_");
	}
	else if ( Tempmode == 1) { // read the average temperatures for this lat, lon
		strcat(FilePathBil,"//avg_");
	}
	else if ( Tempmode == 2) { // read the maximum temperatures for this lat, lon
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

//////////////////////////////////////////////////
//function determine on which drive letter jk_c directory resides
int WhichJKC()
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
		strcat(jkdir, ":\\jk_c\\mapcdinfo.sav" );
		if ( (stream = fopen( jkdir, "r" )) != NULL )
		   fclose(stream);
		   return i;
	}
	return 0; //return 0 if didn't find anything
}
////////////////////////////////////////////////////////////


//////////////////////////////////////////////////
//function determine on which drive letter the /devstudio/vb/WorldClim directory resides
int WhichVB()
{
	FILE *stream;
	char drivlet[2] = "c";
	char climdir[255] = "";
	char *Climdir;
	int i;
	for (i=1; i<= strlen(testdrv); i++)
	{
		drivlet[0] = testdrv[i - 1];
		drivlet[1] = 0;
		strcpy(Climdir, (const char *)drivlet);
		strcat(Climdir, ":\\devstudio\\vb\\WorldClim_bil\\avg_Apr.bil" );
		if ( (stream = fopen( Climdir, "r" )) != NULL )
		   fclose(stream);
		   return i;
	}
	return 0; //return 0 if didn't find anything
}
////////////////////////////////////////////////////////////