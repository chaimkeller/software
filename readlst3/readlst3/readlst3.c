/* readlst4.f -- translated by f2c (version 20021022).
   You must link the resulting object file with the libraries:
	-lf2c -lm   (in that order)
*/

#include "f2c.h"
#include <stdio.h>

//missing intrinisic function from the f2c library
#define hfix_(x) ((shortint)x) //short integer cast
#define jfix_(x) ((integer)x) //long integer cast

/* Table of constant values */

static integer c__9 = 9;
static integer c__1 = 1;
static integer c__3 = 3;
static integer c__5 = 5;
static integer c__0 = 0;
static integer c__801 = 801;
static integer c__512 = 512;

int WhichJKC();
int WhichDTM();

//globals
char testdrv[13] = "cdefghijklmn";
char drivlet[2] = "c";
int WhereJKC,WhereDTM;

/* Main program */ int MAIN__(void)
{
    /* Format strings */
    static char fmt_30[] = "(i5,3x,a,2(3x,f5.0),3x,f5.1,3x,i1)";
    static char fmt_80[] = "(i4,\002,\002,a38,\002,\002,2(f8.3,\002,\002),f7"
	    ".1,\002,\002,2(f9.3,\002,\002),i1,\002,\002,f4.1,\002,\002,f6.3)";

    /* System generated locals */
    integer i__1, i__2;
    doublereal d__1, d__2;
    olist o__1;
    cllist cl__1;
    inlist ioin__1;

    /* Builtin functions */
    double acos(doublereal), tan(doublereal);
    integer f_open(olist *), f_inqu(inlist *), s_rsle(cilist *), do_lio(
	    integer *, integer *, char *, ftnlen), e_rsle(void);
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer s_wsfe(cilist *), do_fio(integer *, char *, ftnlen), e_wsfe(void),
	     s_wsle(cilist *), e_wsle(void), i_dnnt(doublereal *), f_clos(
	    cllist *);

    /* Local variables */
    static doublereal startkmx;
    static integer i__, j;
    static doublereal h1, k1, k2, cd, dstpointx, dstpointy, pi;
    static integer nl, no;
    static char pn[20*6], sn[20], ch1[1], pn0[20], citysource[50];
    static doublereal hgt, kmx, kmy, hgt2, angt;
    static integer ncol, nlen;
    extern /* Subroutine */ int getz_(integer *, integer *, doublereal *);
    static doublereal kmys;
    static integer nrow;
    static doublereal kmxw;
    static integer nlen2;
    static char fname[24];
    static integer nflag;
    static doublereal apprn, kmxoo, kmyoo, maxan1, maxang, begkmx, begkmy, 
	    endkmx;
    static char totnam[38];
    static doublereal apprnr, sofkmx, sofkmy;
    static logical exists;
    static integer numkmx, numkmy, ndirflg, nbegkmx, ntraflg;
    static doublereal expanse;
    static char citydir[50];
    static doublereal skipkmx, skipkmy;


    /* Fortran I/O blocks */
    static cilist io___12 = { 0, 1, 1, 0, 0 };
    static cilist io___14 = { 0, 1, 1, 0, 0 };
    static cilist io___18 = { 0, 1, 1, 0, 0 };
    static cilist io___35 = { 0, 6, 0, fmt_30, 0 };
    static cilist io___36 = { 0, 6, 0, 0, 0 };
    static cilist io___37 = { 0, 6, 0, 0, 0 };
    static cilist io___58 = { 0, 2, 0, fmt_80, 0 };
    static cilist io___59 = { 0, 2, 0, fmt_80, 0 };
    static cilist io___61 = { 0, 3, 0, "(A20,2(F8.3,1X),F6.1,A1)", 0 };
    static cilist io___62 = { 0, 6, 0, "(I5,3X,A20,2(F8.3,1X),F6.1,A1)", 0 };
    static cilist io___63 = { 0, 1, 0, 0, 0 };


/*       READLST3 is quick fix of READLIST */
/*       It creates long scan regions lists */
/*       that have refraction included in them */
/*       The output still goes SCANLIST.TXT, which is list of */
/*       'place name',kmxo,kmyo,hgt,STARTKMX,SOFKMX,NETZ flag */
/*       LAST REVISED 09/19/2003 */
/* ----------------------------------------------------- */
/* -------------------ARRAYS----------------------------- */
/* ------------------PARAMETERS-------------------------- */
/*       distance between search points (meters) in kmx, kmy */
    dstpointx = 25.;
    dstpointy = 25.;
/*       nearest approach to observation point (kms) for 1,2 search */
    apprnr = 0.f;
/*       start search from this kmx */
    maxang = 30.;
    maxan1 = 45.;
/* >>>>>> remember to change nomentr if change maxang <<<<<<<<<< */
/*       conv deg to rad */
    pi = acos(-1.);
    cd = pi / 180;
    angt = tan((maxang + 2) * cd);
    nl = 0;
    no = 0;

	//fincd jk_c directory
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

/* ----------------------------------------------------------------------- */
/*       READ IN STATUS FILE, THEN INPUT PARAMETERS */
    o__1.oerr = 0;
    o__1.ounit = 1;
    o__1.ofnmlen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/PLACLIS2.TXT" );
    //o__1.ofnm = "C:/jk_c/PLACLIS2.TXT";
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
    ioin__1.inerr = 0;
    ioin__1.infilen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/SCANLIST.TXT" );
    //ioin__1.infile = "C:/jk_c/SCANLIST.TXT";
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
	o__1.ofnmlen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/SCANLIST.TXT" );
	//o__1.ofnm = "C:/jk_c/SCANLIST.TXT";
	o__1.orl = 0;
	o__1.osta = "OLD";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    } else {
	o__1.oerr = 0;
	o__1.ounit = 2;
	o__1.ofnmlen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/SCANLIST.TXT" );
	//o__1.ofnm = "C:/jk_c/SCANLIST.TXT";
	o__1.orl = 0;
	o__1.osta = "NEW";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    }
    ioin__1.inerr = 0;
    ioin__1.infilen = 20;
	strcpy(ioin__1.infile, drivlet[0]);
	strcat(ioin__1.infile, ":/jk_c/PLACLIST.OUT" );
    //ioin__1.infile = "C:/jk_c/PLACLIST.OUT";
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
	o__1.ofnmlen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/PLACLIST.OUT" );
	//o__1.ofnm = "C:/jk_c/PLACLIST.OUT";
	o__1.orl = 0;
	o__1.osta = "OLD";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    } else {
	o__1.oerr = 0;
	o__1.ounit = 3;
	o__1.ofnmlen = 20;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/PLACLIST.OUT" );
	//o__1.ofnm = "C:/jk_c/PLACLIST.OUT";
	o__1.orl = 0;
	o__1.osta = "NEW";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    }
L5:
    i__1 = s_rsle(&io___12);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__9, &c__1, fname, (ftnlen)24);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = e_rsle();
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = s_rsle(&io___14);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__9, &c__1, citydir, (ftnlen)50);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__3, &c__1, (char *)&ntraflg, (ftnlen)sizeof(integer));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__9, &c__1, citysource, (ftnlen)50);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = e_rsle();
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = s_rsle(&io___18);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__9, &c__1, pn0, (ftnlen)20);
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&kmxw, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&kmys, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&hgt, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__3, &c__1, (char *)&nflag, (ftnlen)sizeof(integer));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&maxang, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&apprn, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&begkmx, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__5, &c__1, (char *)&endkmx, (ftnlen)sizeof(doublereal));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = do_lio(&c__3, &c__1, (char *)&ndirflg, (ftnlen)sizeof(integer));
    if (i__1 != 0) {
	goto L900;
    }
    i__1 = e_rsle();
    if (i__1 != 0) {
	goto L900;
    }
/*       find length of city directory */
    *(unsigned char *)totnam = '\'';
    for (i__ = 2; i__ <= 25; ++i__) {
	j = i__ - 1;
	if (*(unsigned char *)&fname[j - 1] == ' ') {
	    goto L7;
	}
	nlen = i__;
	*(unsigned char *)&totnam[i__ - 1] = *(unsigned char *)&fname[j - 1];
/* L6: */
    }
L7:
    for (j = 1; j <= 20; ++j) {
	if (*(unsigned char *)&pn0[j - 1] == ' ') {
	    goto L9;
	}
	nlen2 = nlen + j;
	*(unsigned char *)&totnam[nlen2 - 1] = *(unsigned char *)&pn0[j - 1];
/* L8: */
    }
L9:
    ++nlen2;
    *(unsigned char *)&totnam[nlen2 - 1] = '\'';
/*       If Ntraflg = 1, then this file was already processed */
    if (ntraflg == 1) {
	goto L5;
    }
/*       NFLAG = 0 --> SEARCH FOR LARGEST ALTITUDE IN SQUARE KILOMETER */
/*                     WHERE GIVEN COORDINATES ARE THE SW CORNER */
/*       NFLAG = 1 --> USE GIVEN kmx,kmy,hgt */
/*       MAXANG = 1/2 ANGULAR RANGE OF AZIMUTHS (IF 0, THEN MAXANG=30 DEG) */
/*       APPRN = NEAREST APPROACH TO OBSERVER (IF 0, THEN APPRN=0.5 KM) */
/*               IF APPRN = -1, THEN ALWAYS SEARCH UNTIL EASTERN LIMIT OF DTM = ITMX = 260 */

/*       NDIRFLG= 0 for sunrise */
/*              = 1 for sunset */
    s_copy(sn, "F:/PROM/", (ftnlen)8, (ftnlen)8);
    s_copy(sn + 8, pn0, (ftnlen)8, (ftnlen)8);
    for (i__ = 1; i__ <= 16; ++i__) {
	if (*(unsigned char *)&sn[i__ - 1] == ' ') {
	    *(unsigned char *)&sn[i__ - 1] = '0';
	}
/* L25: */
    }
    s_copy(pn, pn0, (ftnlen)20, (ftnlen)20);
    s_copy(pn + 60, pn0, (ftnlen)20, (ftnlen)20);
    ++nl;
    s_wsfe(&io___35);
    do_fio(&c__1, (char *)&nl, (ftnlen)sizeof(integer));
    do_fio(&c__1, pn0, (ftnlen)20);
    do_fio(&c__1, (char *)&kmxw, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&kmys, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&hgt, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&nflag, (ftnlen)sizeof(integer));
    e_wsfe();
    if (kmxw > 260.f || kmys > 380.f) {
	s_wsle(&io___36);
	do_lio(&c__9, &c__1, "***kmxw or kmys incorrect***", (ftnlen)28);
	e_wsle();
	s_wsle(&io___37);
	do_lio(&c__9, &c__1, "***********ABORT************", (ftnlen)28);
	e_wsle();
	goto L999;
    }
/* -------------------------------------------------------------------------------- */
    if (kmys < 1e3 && kmys >= 860.) {
	kmys += -1000;
    }
    kmyoo = kmys;
    kmxoo = kmxw;
    if (nflag == 1) {
	k1 = kmxoo;
	k2 = kmyoo;
	h1 = hgt;
	goto L70;
    }
    expanse = .5f;
    kmyoo += .5f;
    kmxoo += .5f;
    startkmx = kmxoo - expanse;
    sofkmx = kmxoo + expanse;
    begkmy = kmyoo - expanse;
    sofkmy = kmyoo + expanse;
    endkmx = sofkmx;
    begkmx = startkmx;
    d__1 = begkmx * 40.f;
    nbegkmx = i_dnnt(&d__1);
    begkmx = nbegkmx * .025f;
    skipkmx = dstpointx * .001f;
    skipkmy = dstpointy * .001f;
    d__2 = (d__1 = endkmx - begkmx, abs(d__1)) / skipkmx;
    numkmx = i_dnnt(&d__2) + 1;
    d__2 = (d__1 = sofkmy - begkmy, abs(d__1)) / skipkmy;
    numkmy = i_dnnt(&d__2) + 1;
    i__1 = numkmx;
    for (i__ = 1; i__ <= i__1; ++i__) {
	kmx = begkmx + (i__ - 1) * skipkmx;
	d__1 = (kmx + 20) * 40;
	ncol = i_dnnt(&d__1) + 1;
	i__2 = numkmy;
	for (j = 1; j <= i__2; ++j) {
	    kmy = begkmy + (j - 1) * skipkmy;
	    d__1 = (380 - kmy) * 40;
	    nrow = i_dnnt(&d__1) + 1;
	    getz_(&nrow, &ncol, &hgt2);
	    if (i__ == 1 && j == 1) {
		h1 = hgt2;
		k1 = kmx;
		k2 = kmy;
	    } else {
		if (hgt2 > h1) {
		    h1 = hgt2;
		    k1 = kmx;
		    k2 = kmy;
		}
	    }
/* L55: */
	}
/* L60: */
    }
/*       now write output */
L70:
    if (ndirflg == 0) {
	++no;
	s_wsfe(&io___58);
	do_fio(&c__1, (char *)&no, (ftnlen)sizeof(integer));
	do_fio(&c__1, totnam, (ftnlen)38);
	do_fio(&c__1, (char *)&k1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&k2, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&h1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&begkmx, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&endkmx, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&c__1, (ftnlen)sizeof(integer));
	do_fio(&c__1, (char *)&maxan1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&apprn, (ftnlen)sizeof(doublereal));
	e_wsfe();
    } else if (ndirflg == 1) {
	++no;
	s_wsfe(&io___59);
	do_fio(&c__1, (char *)&no, (ftnlen)sizeof(integer));
	do_fio(&c__1, totnam, (ftnlen)38);
	do_fio(&c__1, (char *)&k1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&k2, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&h1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&begkmx, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&endkmx, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&c__0, (ftnlen)sizeof(integer));
	do_fio(&c__1, (char *)&maxan1, (ftnlen)sizeof(doublereal));
	do_fio(&c__1, (char *)&apprn, (ftnlen)sizeof(doublereal));
	e_wsfe();
    }
    *(unsigned char *)ch1 = ' ';
    if ((d__1 = k1 - startkmx, abs(d__1)) < .001f) {
	*(unsigned char *)ch1 = '*';
    } else if ((d__1 = k1 - sofkmx, abs(d__1)) < .001f) {
	*(unsigned char *)ch1 = '*';
    } else if ((d__1 = k2 - begkmy, abs(d__1)) < .001f) {
	*(unsigned char *)ch1 = '*';
    } else if ((d__1 = k2 - sofkmy, abs(d__1)) < .001f) {
	*(unsigned char *)ch1 = '*';
    }
    s_wsfe(&io___61);
    do_fio(&c__1, pn0, (ftnlen)20);
    do_fio(&c__1, (char *)&k1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&k2, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&h1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, ch1, (ftnlen)1);
    e_wsfe();
    s_wsfe(&io___62);
    do_fio(&c__1, (char *)&nl, (ftnlen)sizeof(integer));
    do_fio(&c__1, pn0, (ftnlen)20);
    do_fio(&c__1, (char *)&k1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&k2, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, (char *)&h1, (ftnlen)sizeof(doublereal));
    do_fio(&c__1, ch1, (ftnlen)1);
    e_wsfe();
    goto L5;
L900:
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
    cl__1.cerr = 0;
    cl__1.cunit = 2;
    cl__1.csta = 0;
    f_clos(&cl__1);
    cl__1.cerr = 0;
    cl__1.cunit = 3;
    cl__1.csta = 0;
    f_clos(&cl__1);
/*       NOW WRITE STATUS FILE */
    ioin__1.inerr = 0;
    ioin__1.infilen = 18;
	strcpy(ioin__1.infile, drivlet[0]);
	strcat(ioin__1.infile, ":/jk_c/STATUS.TXT" );
    //ioin__1.infile = "C:/jk_c/STATUS.TXT";
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
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/STATUS.TXT" );
	//o__1.ofnm = "C:/jk_c/STATUS.TXT";
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
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/jk_c/STATUS.TXT" );
	//o__1.ofnm = "C:/jk_c/STATUS.TXT";
	o__1.orl = 0;
	o__1.osta = "NEW";
	o__1.oacc = 0;
	o__1.ofm = 0;
	o__1.oblnk = 0;
	f_open(&o__1);
    }
    s_wsle(&io___63);
    do_lio(&c__3, &c__1, (char *)&c__1, (ftnlen)sizeof(integer));
    do_lio(&c__3, &c__1, (char *)&no, (ftnlen)sizeof(integer));
    e_wsle();
    cl__1.cerr = 0;
    cl__1.cunit = 1;
    cl__1.csta = 0;
    f_clos(&cl__1);
L999:
    return 0;
} /* MAIN__ */

/* Subroutine */ int getz_(integer *nrow, integer *ncol, doublereal *zm)
{
    /* Initialized data */

    static char chmneo[2] = "XX";
    static integer ib = 0;

    /* Format strings */
    static char fmt_13[] = "(a78)";
    static char fmt_30[] = "(\002 ** ERROR IN GETZ ** \002,a2,8i6)";
    static char fmt_23[] = "(1x,a2,f8.1,2f8.3,6i7)";

    /* System generated locals */
    integer i__1;
    olist o__1;
    cllist cl__1;

    /* Builtin functions */
    integer f_open(olist *), s_rsfe(cilist *), do_fio(integer *, char *, 
	    ftnlen), e_rsfe(void);
    /* Subroutine */ int s_copy(char *, char *, ftnlen, ftnlen);
    integer f_clos(cllist *), s_cmp(char *, char *, ftnlen, ftnlen), s_rdue(
	    cilist *), do_uio(integer *, char *, ftnlen), e_rdue(void), 
	    s_wsfe(cilist *), e_wsfe(void);
    /* Subroutine */ int s_paus(char *, ftnlen);
    integer s_wsle(cilist *), do_lio(integer *, integer *, char *, ftnlen), 
	    e_wsle(void);

    /* Local variables */
    static integer i__, j, n, ic;
    static doublereal ge, gn;
    static integer ni;
    static shortint io[512];
    static char sf[9];
    static integer ir, iz, nn1, nn2, ifn, ios, nrec;
    //extern doublereal hfix_(integer *);
    //extern integer jfix_(integer *);
    static char chmap[2*14*26], chmne[2], doclin[78];
	char *folderDTM;


    /* Fortran I/O blocks */
    static cilist io___67 = { 0, 5, 1, fmt_13, 0 };
    static cilist io___70 = { 0, 5, 1, fmt_13, 0 };
    static cilist io___83 = { 1, 5, 1, 0, 0 };
    static cilist io___86 = { 0, 6, 0, fmt_30, 0 };
    static cilist io___89 = { 0, 6, 0, fmt_23, 0 };
    static cilist io___90 = { 0, 6, 0, 0, 0 };


    if (ib != 0) {
	goto L15;
    }
/*       PULL IN THE DTM TILE SCHEMATIC */
	WhereDTM = WhichDTM();
	drivlet[0] = testdrv[WhereDTM - 1];
	drivlet[1] = 0;
    o__1.oerr = 0;
    o__1.ounit = 5;
    o__1.ofnmlen = 18;
	strcpy(o__1.ofnm, drivlet[0]);
	strcat(o__1.ofnm, ":/DTM/DTM-MAP.LOC" );
    //o__1.ofnm = "C:/DTM/DTM-MAP.LOC";
    o__1.orl = 0;
    o__1.osta = "OLD";
    o__1.oacc = 0;
    o__1.ofm = 0;
    o__1.oblnk = 0;
    f_open(&o__1);
/*      READ THE HEADER LINES */
    for (i__ = 1; i__ <= 3; ++i__) {
	i__1 = s_rsfe(&io___67);
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
	i__1 = s_rsfe(&io___70);
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
/*          PRINT *,'CHMAP)', J,N,' = ', CHMAP(J,N) */
/*          PAUSE 'continue-->' */
/* L7: */
	    }
/* L9: */
	}
    }
L14:
    cl__1.cerr = 0;
    cl__1.cunit = 5;
    cl__1.csta = 0;
    f_clos(&cl__1);
/*        READ(5,5,END=25) ((CHMAP(I,J),I=1,14),J=1,26) */
/* 5       FORMAT(//,(/,2X,14(3X,A2))) */
/*        CLOSE(5) */
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
	strcpy(folderDTM, drivlet[0] );
	strcat(folderDTM, "://DTM//" );
    s_copy(sf, (const char *)folderDTM, (ftnlen)7, (ftnlen)7);
    //s_copy(sf, "C:/DTM/", (ftnlen)7, (ftnlen)7);
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
    ifn = jfix_(i__1) * jfix_(c__801) + jfix_(ic);
    i__1 = (ifn - 1) / 512;
    nrec = (integer) (hfix_(i__1) + 1);
    i__1 = (ifn - 1) % 512;
    ni = (integer) (hfix_(i__1) + 1);
    io___83.cirec = nrec;
    ios = s_rdue(&io___83);
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
    s_wsfe(&io___86);
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
    s_wsfe(&io___89);
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
    gn = 380.f - (*nrow - 1) * .025f;
    ge = (*ncol - 1) * .025f - 20.f;
    s_wsle(&io___90);
    do_lio(&c__9, &c__1, "SOMETHING WRONG AT KMX,KMY:", (ftnlen)27);
    do_lio(&c__5, &c__1, (char *)&ge, (ftnlen)sizeof(doublereal));
    do_lio(&c__5, &c__1, (char *)&gn, (ftnlen)sizeof(doublereal));
    e_wsle();
    s_paus("ZM=-3276.7!!!", (ftnlen)13);
    return 0;
} /* getz_ */

/* Main program alias */ int readlst3_ () { MAIN__ (); return 0; }

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
		{
		   fclose(stream);
		   return i;
		}
	}
	return 0; //return 0 if didn't find anything
}
////////////////////////////////////////////////////////////

//////////////////////////////////////////////////
//function determine on which drive letter jk_c directory resides
int WhichDTM()
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
		strcat(jkdir, ":\\DTM\\DTM-MAP.LOC" );
		if ( (stream = fopen( jkdir, "r" )) != NULL )
		{
		   fclose(stream);
		   return i;
		}
	}
	return 0; //return 0 if didn't find anything
}
////////////////////////////////////////////////////////////

