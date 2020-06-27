
/////////////////////netzski6////////////////////////////////////////
short netzski6(char *fileo, double kmxo, double kmyo, double hgt,
			  ,short nzskflg, float geotd, short nsetflag, short nfil)
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
    double d__1, d__2;


/************* Local variables ******************/
    bool adhocset;
    int nweather;
    double elevintx, elevinty, d__, p, t, z__;
    int nplaceflg;
    bool adhocrise;
    double a1, e2, e3, e4, t3, t6, ac, cd, ec, df, ch, mc, ap, ob, 
	    lg, ra;
    int ne;
    double td, pi, hr, tf, es, mp, lr, dy, lt, ms, et;
    int nn;
    double pt, sr, e2c, al1, ob1, pi2, dy1;
    int ns1, ns2, ns3, ns4;
    double sr1, sr2;
    int nac;
    double aas;
    bool geo;
    double air, ref;
    int ier;
    double fdy, azi, eps;
    int nyd, nyf, nyl, nyr;
    double tss, al1o, alt1, azi1, azi2; //, chd62;

	char CHD46[130] = "";
	float h1,h2,h3,h4,h5,h6,h7,h8;

    //double dobs[1464]	/* was [2][2][366] */;
    bool redo;
    double hran;
    double elev[4803]	/* was [3][1601] */;
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
    double winref[8];
    int nacurr;
    double sumref[8];
    int nloops, ndystp;
    double refrac1;
    int nyltst, nyrstp, nyrtst;

    double stepsec, elevint, winrefo;
    double sumrefo;
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
/*       stryr                    'beginning year to calculate */
/*       endyr                     'last year to calculate */
/* --------------------------<<<<USER INPUT>>>---------------------------------- */
    nplaceflg = 1;
/*                                  '=0 calculates earliest sunrise/latest sunset for selected places in Jerus
alem */
/*                                  '=1 calculation for a specific place--must input place number */
/*                                  '=2 for calculations for LA */
/*      'Nsetflag                   '=0 for calculating sunrises */
/*                                  '=1 for calculating sunsets */
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
    pi = acos(-1.);
    cd = pi / 180.;
    pi2 = pi * 2.;
    ch = 12. / pi;
    hr = 60.;
    geo = false;

/*       summer and winter refraction values for heights 1 m to 999 m */
/*       as calculated by MENAT.FOR */
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

/*       read in elevations from Profile file */
/*       which is a list of the view angle (corrected for refraction) */
/*       as a function of azimuth for the particular observation */
/*       point at kmxo,kmyo,hgt */

//    s_copy(fileo, "e:/cities/acco/netz/1134-aco.pro", (ftnlen)32, (ftnlen)32);
//    strncpy(fileo, "e:/cities/acco/skiy/1134-aco.pro", strlen("e:/cities/acco/skiy/1134-aco.pro"));

    if (nzskflg <= -4) {
	geo = true;
    } else {
	geotd = 2.;
    }

/*       nsetflag = 0 for astr sunrise/visible sunrise/chazos */
/*       nsetflag = 1 for astr sunset/visible sunset/chazos */
/*       nsetflag = 2 for mishor sunrise if hgt = 0/chazos or astro. sunrise/chazos */
/*       nsetflag = 3 for mishor sunset if hgt = 0/chazos or astro. sunset/chazos */
/*       nsetflag = 4 for mishor sunrise/vis. sunrise/chazos */
/*       nsetflag = 5 for mishor sunset/vis. sunset/chazos */

    astro = false;
    if (nsetflag == 2 || nsetflag == 3) {
	astro = true;
    }


    if ( (stream = fopen( fileo, "r" )) != NULL )
		{
			while ( !feof(stream) )
			{
			fscanf( stream, "%s\n", CHD46 );
			fscanf(stream, "%lg,%lg,%lg,%lg,%lg,%lg,%lg,%lg\n", &h1,&h2,&h3,&h4,&h5,&h6,&h7,&h8);
			nazn = 0;
n40:		nazn = nazn++; //increment by 1
//			fscanf(stream,"%lg,%lg,%lg,%lg,%lg,%lg\n", &elev[0][nazn],&elev[1][nazn],&e2,&e3,&elev[2][nazn],&e4);
			fscanf(stream,"%lg,%lg,%lg,%lg,%lg,%lg\n", &elev[nazn * 3 - 3]
				                                     , &elev[nazn * 3 - 2]
													 , &e2,&e3
													 , &elev[nazn * 3 - 1]
													 , &e4);
//			stuff elev array with:  azi,vang,kmx,kmy,dist,hgt
//			if (nazn == 0) maxang = (int)fabs(elev[0][nazn]);
//			if (elev[0][nazn] == maxang) goto n50; //at eof
			if (nazn == 1) maxang = (int)fabs(elev[nazn * 3 - 3]);
			if (elev[nazn * 3 - 3] == maxang) goto n50; //at eof
			goto n40;
			}
		}
	else
	{
		goto L9000;
	}
n50:
        fclose(stream);

/* ----ad hoc summer and winter ranges for different latitudes----- */

    if (geo) {
	lt = kmxo;
	lg = kmyo;
	if (abs(lt) < 20.) {
/*             tropical region, so use tropical atmosphere which */
/*             is approximated well by the summer mid-latitude */
/*             atmosphere for the entire year */
	    nweather = 0;
/*          define days of the year when winter and summer begin */
	} else if (abs(lt) >= 20. && abs(lt) < 30.) {
	    ns1 = 55;
	    ns2 = 320;
/*             no adhoc fixes */
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

/* -------------------initialize astronomical constants--------------------------- */
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

/* ----------------begin astronomical calculations--------------------------- */
    i__1 = endyr;
    for (nyear = stryr; nyear <= i__1; ++nyear) {
/* 	     define year iteration number */
	nyrstp = nyear - stryr + 1;
	nyr = nyear;
/*       year for calculation */
	nyd = nyr - 1996;
	nyf = 0;
	if (nyd < 0) {
	    i__2 = nyr;
	    for (nyri = 1995; nyri >= i__2; --nyri) {
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
	    i__2 = nyr - 1;
	    for (nyri = 1996; nyri <= i__2; ++nyri) {
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
	d__1 = (double) yrend[nyear - stryr];
	for (dy = (double) yrstrt[nyear - stryr]; dy <= d__1; dy += 1.) {
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
        //////////////record astronomical noon in hours/////////////
		if (nfil == 1)
		{
	        tims[nyrstp - 1][6][ndystp - 1] = t6/hr;
		}
		else //average out
		{
			tims[nyrstp - 1][6][ndystp - 1] = (tims[nyrstp - 1][6][ndystp - 1] * (nfil - 1) + t6/hr) / nfil;
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
		    nacurr = 15;
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
		if (nsetflag == 3 || nsetflag == 5) {
 
			//////////////record mishor sunset /////////////
			if (nfil == 1)
			{
	            tims[nyrstp - 1][5][ndystp - 1] = t3;
			}
			else
			{
				if (t3 > tims[nyrstp - 1][5][ndystp - 1] )
					tims[nyrstp - 1][5][ndystp - 1] = t3;
			}
			////////////////////////////////////////////////

		    if (astro) {
/*                     finished calculations for this day */
			goto L925;
		    } else {
/*                     go back to calculate astro sunset, vis sunset */
			redo = true;
		    }
		} else {
 
            //////////////record astro sunset /////////////
			if (nfil == 1)
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
		    nacurr = -15;
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
		if (nsetflag == 2 || nsetflag == 4) {

            //////////////record mishor sunrise /////////////
			if (nfil == 1)
			{
	            tims[nyrstp - 1][4][ndystp - 1] = t3;
			}
			else
			{
				if (t3 < tims[nyrstp - 1][4][ndystp - 1] )
					tims[nyrstp - 1][4][ndystp - 1] = t3;
			}
			////////////////////////////////////////////////

		    if (astro) {
/*                     finished calculations for this day */
			goto L925;
		    } else {
/*                     go back to calculate astro sunrise, vis sunrise */
			goto L688;
		    }
		} else {

            //////////////record astro sunrise /////////////
			if (nfil == 1)
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
	    nn = 0;
	    nfound = 0;
	    nabove = 0;
	    i__2 = nlpend;
	    i__3 = ndirec;
	    for (nloops = nloop; i__3 < 0 ? nloops >= i__2 : nloops <= i__2; 
		    nloops += i__3) {
		++nn;
		hran = sr1 - stepsec * nloops / 3600.;
/*           step in hour angle by stepsec sec */
		dy1 = dy - ndirec * hran / 24.;
		decl_(&dy1, &mp, &mc, &ap, &ac, &ms, &aas, &es, &ob, &d__);
/*           recalculate declination at each new hour angle */
		z__ = sin(lr) * sin(d__) + cos(lr) * cos(d__) * cos(hran / ch);
		z__ = acos(z__);
		alt1 = pi * .5 - z__;
/*           altitude in radians of sun's center */
		azi1 = cos(d__) * sin(hran / ch);
		azi2 = -cos(lr) * sin(d__) + sin(lr) * cos(d__) * cos(hran / 
			ch);
		azi = atan(azi1 / azi2) / cd;
		if (azi > 0.) {
		    azi1 = 90. - azi;
		}
		if (azi < 0.) {
		    azi1 = -(azi + 90.);
		}
		if (abs(azi1) > maxang) {
		    goto L9050;
		}
		azi1 = ndirec * azi1;
		if (nsetflag == 1) {
/*              to first order, the top and bottom of the sun have the */
/*              following altitudes near the sunset */
		    al1 = alt1 / cd + a1;
/*              first guess for apparent top of sun */
		    pt = 1.;
		    refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &
			    ns1, &ns2);
/*              first iteration to find position of upper limb of sun */
		    al1 = alt1 / cd + .2666f + pt * refrac1;
/*              top of sun with refraction */
		    if (niter == 1) {
			for (nac = 1; nac <= 10; ++nac) {
			    al1o = al1;
			    refract_(&al1, &p, &t, &hgt, &nweather, &dy, &
				    refrac1, &ns1, &ns2);
/*                    subsequent iterations */
			    al1 = alt1 / cd + .2666f + pt * refrac1;
/*                    top of sun with refraction */
			    if ((d__2 = al1 - al1o, abs(d__2)) < 4.) {
				goto L695;
			    }
/* L693: */
			}
		    }
L695:
		    ne = (int) ((azi1 + maxang) * 10) + 1;
		    elevinty = elev[(ne + 1) * 3 - 2] - elev[ne * 3 - 2];
		    elevintx = elev[(ne + 1) * 3 - 3] - elev[ne * 3 - 3];
		    elevint = elevinty / elevintx * (azi1 - elev[ne * 3 - 3]) 
			    + elev[ne * 3 - 2];
		    if ((elevint >= al1 || nloops <= 0) && nabove == 1) {
/*                 reached sunset */
			nfound = 1;
/*                 margin of error */
			nacurr = 0;
			t3sub = t3 + ndirec * (stepsec * nloops + nacurr) / 
				3600.;
			if (nloops < 0) {
/*                     visible sunset can't be latter then astronomical sunset! */
			    t3sub = t3;
			}
/*                 <<<<<<<<<distance to obstruction>>>>>> */
			dobsy = elev[(ne + 1) * 3 - 1] - elev[ne * 3 - 1];
			elevintx = elev[(ne + 1) * 3 - 3] - elev[ne * 3 - 3];
			dobss = dobsy / elevintx * (azi1 - elev[ne * 3 - 3]) 
				+ elev[ne * 3 - 1];
/*                 record visual sunset results for time,azimuth,dist. to obstructions */

			/////only record times if they are earlier (sunrise) or latter (sunset)
			if (nfil == 1) //first set of times////////////////////////////////////
			{
     			tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
				azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
				dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
			}
			else
			{
				if (nsetflag == 0 ) //sunrises
				{
					if (t3sub < tims[nyrstp - 1][nsetflag][ndystp - 1] )
					{
						tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
						azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
						dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
					}
				}
				else if (nsetflag == 1) //sunsets
				{
					if (t3sub > tims[nyrstp - 1][nsetflag][ndystp - 1] )
					{
						tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
						azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
						dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
					}
				}
			}
			/////////////////////////////////////////////////////////////

/*                 finished for this day, loop to next day */
			goto L925;
		    } else if (elevint < al1) {
			nabove = 1;
		    }
		} else if (nsetflag == 0) {
		    al1 = alt1 / cd + a1;
/*                 first guess for apparent top of sun */
		    pt = 1.;
		    refract_(&al1, &p, &t, &hgt, &nweather, &dy, &refrac1, &
			    ns1, &ns2);
/*                 first iteration to find top of sun */
		    al1 = alt1 / cd + .2666f + pt * refrac1;
/*                 top of sun with refraction */
		    if (niter == 1) {
			for (nac = 1; nac <= 10; ++nac) {
			    al1o = al1;
			    refract_(&al1, &p, &t, &hgt, &nweather, &dy, &
				    refrac1, &ns1, &ns2);
/*                       subsequent iterations */
			    al1 = alt1 / cd + .2666f + pt * refrac1;
/*                       top of sun with refraction */
			    if ((d__2 = al1 - al1o, abs(d__2)) < .004f) {
				goto L760;
			    }
/* L750: */
			}
		    }
L760:
		    alt1 = al1;
/*              now search digital elevation data */
		    ne = (int) ((azi1 + maxang) * 10) + 1;
		    elevinty = elev[(ne + 1) * 3 - 2] - elev[ne * 3 - 2];
		    elevintx = elev[(ne + 1) * 3 - 3] - elev[ne * 3 - 3];
		    elevint = elevinty / elevintx * (azi1 - elev[ne * 3 - 3]) 
			    + elev[ne * 3 - 2];
		    if (elevint <= alt1) {
			if (nn > 1) {
			    nfound = 1;
			}
			if (nfound == 0) {
			    goto L925;
			}
/*                 <<<<<<<<<distance to obstruction>>>>>> */
			dobsy = elev[(ne + 1) * 3 - 1] - elev[ne * 3 - 1];
			elevintx = elev[(ne + 1) * 3 - 3] - elev[ne * 3 - 3];
			dobss = dobsy / elevintx * (azi1 - elev[ne * 3 - 3]) 
				+ elev[ne * 3 - 1];
/*                 sunrise */
			nacurr = 0;
/*                 margin of error (seconds) to take into */
/*                 account differences between observations and the DTM `` */
			t3sub = t3 + (stepsec * nloops + nacurr) / 3600.;
			if (nloops < 0) {
/*                    sunrise can't be before astronomical sunrise */
/*                    so set it equal to astronomical sunrise time */
/*                    minus the adhoc winter fix if any */
			    t3sub = t3;
			}
/*                 record visual sunrise results for time,azimuth,dist. to obstructions */
//			tims[nyrstp + (nsetflag + 1 + ndystp * 7 << 1) - 17] =
//				 t3sub;
    		tims[nyrstp - 1][nsetflag][ndystp - 1] = t3sub;
//			azils[nyrstp + (nsetflag + 1 + (ndystp << 1) << 1) - 
//				7] = azi1;
            azils[nyrstp - 1][nsetflag][ndystp - 1] = azi1;
//			dobs[nyrstp + (nsetflag + 1 + (ndystp << 1) << 1) - 7]
//				 = dobss;
			dobs[nyrstp - 1][nsetflag][ndystp - 1] = dobss;
/*                 finished for this day, loop to next day */
			goto L925;
		    }
		}
/* L850: */
	    }
/* L900: */
	    if (nfound == 0 && nloop > 0 && nsetflag == 1) {
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
		if (nsetflag == 0) {
		    nloop += -24;
/*               IF (nloop .LT. 0) nloop = 0 */
		} else if (nsetflag == 1) {
		    nloop += 24;
		}
		goto L692;
	    } else if (nfound == 0 && nloop < -100 && nsetflag == 0) {
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
L925:
	    ;
	}
/* L975: */
    }
    goto L9999;
/*     signal read error */
L9000:
    ier = -2;
    goto L9999;
/*     signal nonconverging calculation */
L9050:
    ier = -1;
L9999:
    return 0;
} /* MAIN__ */

/* Subroutine */ short decl_(double *dy1, double *mp, double *mc, 
	double *ap, double *ac, double *ms, double *aas, 
	double *es, double *ob, double *d__)
{
    // Builtin functions 
    //double acos(double), sin(double), asin(double);

    /* Local variables */
    double cd, ec, pi, e2c, pi2;

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
    extern /* Subroutine */ int refwin1_(double *, double *, 
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
	    ;

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
