        PROGRAM NETZSKI6
c       Version 17
c       this program calculates sunrise/sunset tables for lattest format
c       it is called by the WINDOWS 95 program CAL PROGRAM
c       the output filenames reside on the file NETZSKIY.tm3
c       This version allows for arbitrary number of decimal places
c       on coordinates.  It adds an adhoc fix for Eretz Yisroel's
c       sunrise and sunet during two months of the winter (based
c       on observations at Neveh Yaakov. The adhoc fixes are controlled
c       by the logical variables ADHOCRISE (sunrise), ADHOCSET (sunset)
c
c       -------Version Info------------
c       Version 14 only calculates the days that are included in the Hebrew Year
c
c       Version 15 doesn't use COMPARE, rather compares the times in memory
c       and doesn't write itermediate time files for each neighborhood.  Version 15
c       only supports the calculation of one Hebrew Year at a time.
c       Also permits turning off writing any intermediate neighborhood files
c
c		Version 16 adds temperature scaling for Menat atmospheres for a proof of principle test
c		A latter version will use van der Werf's atmospheric model for the entire world simply
c		modified by the appropriate average ground temperature for either sunrise or sunset supplied 
c		via the netzskiy.tm3 file either calculated by the Cal Program and based either on the WorldClim model 
c 		or provided by user input for a certain day whenever that option becomes viable.
c	    For this version, minimum and average temperatures for each month are based on the WorldClim 
c		model for the profile file's coordinates and are calculated by Cal Program and are written to a 
c 		new line in the netzskiy.tm3 file. 
c	    For this version, only Menat's winter atmosphere is used modified by temperature 
c 		scaling laws as determined by Van der Werf's 2003 paper for his own ray tracing calculations. 
c		So even though they don't apply to Menat's modeling, there are used as a proof of principle test. 
c		Both the total atmospheric refraction and refraction at a certain view anglea and height are 
c		according to the scaling laws that Van der Werf published in his paper.
c 
c		From version 16, netzski6 is now given the task of calculating terrestrial refraction, TR, 
c 		using the modified Wikipedia expression for both sunrise and sunset.
c 		I.e., TR (deg.) = 0.0083 * Pathlength (km)**Exponent * K * Ground Pressure (mb)/ (Ground Temperature (K) ** 2).
c		The Ground Temperature is modeled to be different at sunrise and at sunset. 
c		For this version, ground temperature at sunrise is apprximated by the minimum monthly WorldClim
c		temperature for the profile file's coordinates, and the ground temperature at sunset is approximated
c		by the WorldClim average monthly temperature for the profile file's coordinates. This will be tweeked in
c		minor version updates to improve the fit to Rabbi Druk's measurements.
c		Both temepratures are determined for the coordinates using the WorldClim version 2 model and are provided 
c		via the netzskiy.tm3 file from the Cal Program, and if they are zeros, then they are determined by this routine.
c
c		The adhoc sunrise and sunset fixes are NO LONGER USED from version 16, since it is assumed that they were
c		needed due to lack of temperature modeling.
c
c		As of version 17, Menat's refraction is no longer used, and is replaced by van der Werf's 
c		There is also no longer any mention of winter and summer, so the nweather parameter is 
c		currently no longer in use
c
c		Version 18 fixes some wrong application of the vdW temperature scaling on a1, i.e., it needs to only
c		be applied in the visible times section, otherwise it is between added twice
c
C-----------------------------------------------------
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
        PARAMETER(naz = 1601)
c            = 20 * maxang + 1 ;number of azimuth points separated by .1 degrees
        PARAMETER(nplace = 80)
c            = maximum number of places that can be analyzed at once
c--------------------------------------------------------------------------------
c///////////////////commented out for version >= 15 //////////////////////////
c        REAL*4 azin1(366),times1(366), disob(366),DAT1(366)*11 
c        CHARACTER filin1(366)*12, tim1(366)*8, tim(nplace)*8
c        CHARACTER TIMR1*8, timr*36, filin*27, CH80*80
c        REAL*4 azin(nplace)
c        LOGICAL ,SUMMER,WINTER
c//////////////////////////////////////////////////////////////////////////////
        CHARACTER placn(2, nplace)*30,place(2)*8
        CHARACTER name*8, namet*23, filstat*18 
        CHARACTER ext1(2,nplace)*4, ext3(2,nplace)*4
        INTEGER nplac(2)
 
        REAL*4 elev(4, naz), azinss
        REAL*8 lnhgt
        CHARACTER coment*8,fileo*27,CH1*1,CH2*2
        CHARACTER timss*8, comentss*12
        CHARACTER FILTMP*18,f2*27,CHD62*62,CHD46*46,CH4*4
        CHARACTER ts1c*8,ts2c*9,tssc*8,plcfil*27
        CHARACTER t3subc*8,caldayc*11,filout*27,DAT*11,CN4*4
        CHARACTER locname*40,ch3*3
        CHARACTER drivlet*1,testdriv*8,tmpfil*18,CHD48*48

        INTEGER intflag,yrst(2),yrend(2)
        LOGICAL EXISTS,ASTRO,GEO,ADHOCRISE,ADHOCSET
        CHARACTER yesint * 22
        LOGICAL NOINTERNET, NOFILES

c       Version 15-16 parameters
        REAL*8 tsscs(2, 2, 366), azils(2, 2, 366), dobss(2, 2, 366)
        CHARACTER coments(2, 2, 366)*12
c		MT is Minimum Temperature Array, AT is Average Temperature Array		
		INTEGER MT(12), AT(12)
		REAL *8 vdweps(5), vdwref(6),refnorm(5), refexp(2)
        COMMON /t3s/t3subc, caldayc

        pi = dacos(-1.0d0)
        cd=pi/180.0d0
		dcmrad = 1/(cd * 1000.0d0)
        pi2=2.0D0*pi
        ch=12.0D0/pi
        hr = 60.0d0
        GEO = .FALSE.
        NOINTERNET = .TRUE.
		
c		Average radius of Earth in km		
		Rearth = 6356.766

c		vdW dip angle vs height polynomial fit coefficients
		vdweps(1) = 2.77346593151086
		vdweps(2) = 0.497348466526589
		vdweps(3) = 2.53874620975453E-03
		vdweps(4) = 6.75587054940366E-04
		vdweps(5) = 3.94973974451576E-05

c		vdW atmospheric refraction vs height polynomial fit coefficients
		vdwref(1) = 1.16577538442405
		vdwref(2) = 0.468149166683532
		vdwref(3) = -0.019176833246687
		vdwref(4) = -4.8345814464145E-03
		vdwref(5) = -4.90660400743218E-04
		vdwref(6) = -1.60099622077352E-05	

		refrac1 = 9.572702286470884

c		alternative solution with height dependence: exponent for vdw ref part of refraction (height to horizon)		
        refexp(1) = 2.20734384287553
        refexp(2) = -2.86255933358013E-05

c       vdW ref normalization a vs height, where ref=a*(288.15/Tk)**b
        refnorm(1) = 0.767089721048164
        refnorm(2) = 3.89787475596774E-03
        refnorm(3) = -2.02184590999692E-06
        refnorm(4) = 6.5747954161702E-10
        refnorm(5) = -8.27402155995415E-14
		
c       read in scan parameters
c       first find the netzskiy.tm3 file--this determines the directory
        testdriv='cdefghij'
        DO 5 I = 1,8
           drivlet=testdriv(I:I)
           tmpfil(2:18)=':/jk/netzskiy.tm3'
           tmpfil(1:1)=drivlet
           INQUIRE(FILE=tmpfil,EXIST=EXISTS)
           IF (EXISTS) THEN
c             found default drive
              yesint(1:1)=drivlet(1:1)
              GOTO 7
              END IF
5      CONTINUE

c       now read in the internet.yes file
c       if operating internet mode, don't write stat files
7       yesint(2:22)=':/cities/internet.yes'
        INQUIRE(FILE=yesint,EXIST=EXISTS)
        IF (EXISTS) THEN
c          found internet.yes file
           OPEN(1,FILE=yesint,status='old')
           READ(1,'(I1)')intflag
           IF (intflag .EQ. 1) THEN
              NOINTERNET = .FALSE.
              END IF
           CLOSE(1)
        END IF

c'summer and winter refraction values for heights 1 m to 999 m
c'as calculated by MENAT.FOR

c       opening file netzskiy.tm3
        OPEN(10,FILE=tmpfil,STATUS='OLD')
        READ(10,*)nyrheb
        READ(10,*)nzskflg,geotd
        IF (nzskflg .LE. -4) THEN
           GEO = .TRUE.
        ELSE
           geotd = 2.0d0
           END IF
        READ(10,*)numfil
        READ(10,'(A)')CN4
        READ(10,'(A)')locname
        stat = 20.0/numfil
        stat2 = 0.0
        nstat = 1

c       write or zero status file
        IF( NOINTERNET ) THEN
            filstat(2:18) = ':/jk/stat0001.tmp'
            filstat(1:1)=drivlet(1:1)
            INQUIRE(FILE=filstat,EXIST=EXISTS)
            IF (EXISTS) THEN
               OPEN(1,FILE=filstat,STATUS='OLD')
            ELSE
               OPEN(1,FILE=filstat,STATUS='NEW')
               END IF
            WRITE(1,*)0
            CLOSE(1)
            END IF

        
        nplac(1) = 0
        nplac(2) = 0

        DO 1000 nfil = 1,numfil
           READ(10,'(A)')fileo
           READ(10,*) kmxo, kmyo, hgt, Nstrtyr, yrst(1), yrend(1),
     *                Nendyr, yrst(2), yrend(2), Nsetflag
           READ(10,*) MT(1),MT(2),MT(3),MT(4),MT(5),MT(6),MT(7),
     *                MT(8),MT(9),MT(10),MT(11),MT(12)
           READ(10,*) AT(1),AT(2),AT(3),AT(4),AT(5),AT(6),AT(7),
     *                AT(8),AT(9),AT(10),AT(11),AT(12)          

           ASTRO = .FALSE.
           IF ((Nsetflag .EQ. 2) .OR. (Nsetflag .EQ. 3)) THEN
              ASTRO = .TRUE.
              IF (Nsetflag .EQ. 2) Nsetflag = 0
              IF (Nsetflag .EQ. 3) Nsetflag = 1
              END IF
 
          IF (Nsetflag .EQ. 0) THEN

              nplac(1) = nplac(1) + 1
              IF (nplac(1).GT.nplace) THEN
                 IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
                    PRINT *, ' ERROR--exceeded max. num. of places!'
                    PAUSE 'continue'
                    END IF
                 GOTO 1000
                 END IF

              IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
c                prepare names of intermediate neighborhood files
                 PLACN(Nsetflag + 1, nplac(1)) = FILEO(16:23)
c                names used for COMPARE for versions <= 14
                 IF (nplac(1) .LT. 10) THEN
                    WRITE(CH1,'(I1)')nplac(1)
                    EXT1(Nsetflag + 1,nplac(1))(1:3) = '.pl'
                    EXT1(Nsetflag + 1,nplac(1))(4:4)=CH1
                 ELSE IF ((nplac(1).GE.10).AND.(nplac(1).LE.99)) THEN
                    WRITE(CH2,'(I2)')nplac(1)
                    EXT1(Nsetflag + 1,nplac(1))(1:2) = '.p'
                    EXT1(Nsetflag + 1,nplac(1))(3:4)=CH2
                 ELSE IF (nplac(1) .GE. 100) THEN
                    WRITE(CH3,'(I3)')nplac(1)
                    EXT1(Nsetflag + 1,nplac(1))(1:1) = '.'
                    EXT1(Nsetflag + 1,nplac(1))(2:4)=CH3
                    END IF
                 EXT3(Nsetflag + 1,nplac(1))(1:4) = FILEO(24:27)
                 END IF

           ELSEIF (Nsetflag .EQ. 1) THEN
 
             nplac(2) = nplac(2) + 1
             IF (nplac(2).GT.nplace) THEN
                IF ((.NOT. NOFILES) .AND. (NOINTERNET)) THEN
                    PRINT *, ' ERROR--exceeded max. num. of places!'
                    PAUSE 'continue'
                    END IF
                GOTO 1000
                END IF

             IF ((.NOT. NOFILES) .AND. (NOINTERNET)) THEN
c                prepare names of intermediate neighborhood files
                 PLACN(Nsetflag + 1, nplac(2)) = FILEO(16:23)
c                names used for COMPARE
                 IF (nplac(2) .LT. 10) THEN
                    WRITE(CH1,'(I1)')nplac(2)
                    EXT1(Nsetflag + 1,nplac(2))(1:3) = '.pl'
                    EXT1(Nsetflag + 1,nplac(2))(4:4)=CH1
                 ELSE IF ((nplac(2).GE.10).AND.(nplac(2).LE.99)) THEN
                    WRITE(CH2,'(I2)')nplac(2)
                    EXT1(Nsetflag + 1,nplac(2))(1:2) = '.p'
                    EXT1(Nsetflag + 1,nplac(2))(3:4)=CH2
                 ELSE IF (nplac(2) .GE. 100) THEN
                    WRITE(CH3,'(I3)')nplac(1)
                    EXT1(Nsetflag + 1,nplac(2))(1:1) = '.'
                    EXT1(Nsetflag + 1,nplac(2))(2:4)=CH3
                    END IF  
                 EXT3(Nsetflag + 1,nplac(2))(1:4) = FILEO(24:27)
                 END IF
              END IF

c       Nstrtyr                    'beginning year to calculate
c       Nendyr                     'last year to calculate
C--------------------------<<<<USER INPUT>>>----------------------------------

       nplaceflg = 1
c                                  '=0 calculates earliest sunrise/latest sunset for selected places in Jerusalem
c                                  '=1 calculation for a specific place--must input place number
c                                  '=2 for calculations for LA

c      'Nsetflag                   '=0 for calculating sunrises
c                                  '=1 for calculating sunsets
        nweather = 0				
c								   'this parameter is now deprecated
c								   'since there is only one atmosphere		
        ndtm = 5
        nmenat = 1
c                                  '=0 uses tp
c                                  '=1 calculates sunrises over Harei Moav using theo. refrac.
        niter = 1
c                                  '=0 doesn't iterate the sunrise times
c                                  '=1 iterates the sunrise times
        ndruk = 0
c                                  '=0 gives normal output
c                                  '= 1 use to calculate files to compare to druk output/vangeld output
        P = 1013.25
c                                  'mean atmospheric pressure (mb)
        t = 27
c                                  'mean atmospheric temperature (degrees Celsius)
        stepsec = 2.5
c                                  'step size in seconds
        NOFILES = .FALSE.
c                                   =.TRUE. to turn off neighborhood files (automatic for internet calculations)
c                                   =.FALSE. to turn on writing of neighborhood files
        nplaces = 1
c       'N.B. astronomical constants are accurate to 10 seconds only to the year 2050
c       '(Astronomical Almanac, 1996, p. C-24)

c       Addhoc fixes: (15 second fixes based on netz observations at Neveh Yaakov)
c		As of Version 16, ADHOCRISE = .FALSE.
        ADHOCRISE = .FALSE.
        ADHOCSET = .FALSE. 

C--------------------------------------------------------------------------------------
   

        IF (Nsetflag .EQ. 0) THEN
           nloop = 72
c          start search 2 minutes above astronomical horizon
           nlpend = 2000
        ELSE IF (Nsetflag .EQ. 1) THEN
c          sunset
           nloop = 370
           IF (nplaceflg.EQ.0) nloop = 75
           nlpend = -10
           END IF
        nopendruk = 0
        DO 975 nyear = Nstrtyr,Nendyr

c	     define year iteration number	
           nYrStp = nyear - Nstrtyr + 1

        IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
           place(Nsetflag + 1) = fileo(16:23)
           IF ((ndruk.EQ.1).AND.(nopendruk.EQ.0)) THEN
C             open temporary file
              FILTMP(1:1)=drivlet(1:1)
              FILTMP(2:6)=':/JK/'
              FILTMP(7:14)=PLACE(Nsetflag + 1)(1:8)
              FILTMP(15:18)='.TMP'
              OPEN(4,FILE=FILTMP,STATUS='NEW')
              nopendruk = 1
              END IF
           END IF

        IF (GEO) THEN
           lt = kmxo
           lg = kmyo
c		   in this version the above parameters are no longer used, see comment below			  
        ELSE
           CALL CASGEO(kmyo,kmxo,lt,lg)
c		   for this version 16, only use van der Werf's atmosphere which is very
c		   similar to the US Standard Atmosphere, and scale the refraction
c		   according to van der Werf scaling laws, using the minimum temperature 
           END IF

C       read in elevations from Profile file
c       which is a list of the view angle (corrected for refraction)
c       as a function of azimuth for the particular observation
c       point at kmxo,kmyo,hgt
        IF(NOINTERNET) THEN
           PRINT*, '                  PLEASE WAIT--READING DTM FILE'
           END IF
C>>>>>>>>>>>>>>>>>>>>>>>>>>>
        f2 = fileo
        comentss = fileo(16:27)
        OPEN(2,file=f2,status='old')
		
        READ(2,'(A)')CHD46
c		check if this is old format profile file, if so, set flag to remove 
c		old form of terrestrial refraction contribution
		NP1 = INDEX(CHD48,',0')
		NP2 = INDEX(CHD48,',1')
		IF ((NP1 .GT. 0) .AND. (NP2 .EQ. 0)) THEN
c			new format w/o TR calculated, so calculate it		
			NTRCALC = 0
		ELSEIF ((NP1 .EQ. 0) .AND. (NP2 .GT. 0)) THEN
c			average temp TR contrib, remove it and recalculte it using month's average temperature
c			this assumes that the change in the horizon profile will be small when changing the temperature		
			NTRCALC = 1
		ELSE
c			old profile file with old TR calculation, so remove it and recalculate using new expression
c			However, assumption that the horizon profile will not change is more iffy in this case		
			NTRCALC = 2
			ENDIF
			
        READ(2,'(A)')CHD62
        nazn = 0
40      nazn = nazn + 1
c       stuff elev array with:  azi,vang,kmx,kmy,dist,hgt
        READ(2,*,END=50)elev(1,nazn),elev(2,nazn),e2,e3,elev(3,nazn),
     *                   elev(4,nazn)
	 
	    IF (NTRCALC .EQ. 2) THEN
c		   remove old terrestrial refraction values from the view angle	
			deltd = hgt - elev(4,nazn) 
			distd = elev(3,nazn)
			IF (deltd .LE. 0.) THEN
				defm = 7.82D-4 - deltd * 3.11D-7
				defb = deltd * 3.4D-5 - .0141D0
			ELSEIF (deltd .GT. 0.) THEN
				defm = deltd * 3.09D-7 + 7.64D-4
				defb = -.00915D0 - deltd * 2.69D-5
				ENDIF
			avref = defm * distd + defb
			IF (avref .LT. 0.) THEN
				avref = 0.
				ENDIF
			elev(2,nazn) = elev(2,nazn) - avref	
		ELSE IF (NTRCALC .EQ. 1) THEN
c			convert minimum and average temperatures to Kelvin scale
			IF (Nsetflag .EQ. 0) THEN
c				sunrise
c				now find average minimum temperature to use in removing profile's TR	
				MeanTemp = 0.0
				DO 42 ITK = 1,12
					MeanTemp = MeanTemp + MT(ITK)
42				CONTINUE
				MeanTemp = MeanTemp/12.0 + 273.15	
			ELSEIF (Nsetflag .EQ. 1) THEN
c				sunset
c				now find average minimum temperature to use in removing profile's TR	
				MeanTemp = 0.0
				DO 43 ITK = 1,12
					MeanTemp = MeanTemp + AT(ITK)
43				CONTINUE
				MeanTemp = MeanTemp/12.0 + 273.15				
				ENDIF
				
			deltd = hgt - elev(4,nazn) 
			distd = elev(3,nazn)
			
c			calculate conserved portion of Wikipedia's terrestrial refraction expression
c			Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere
			TRRbasis = 8.15 * P * 1000.0 * (0.0342 - 0.0065)/
     *         		(MeanTemp * MeanTemp * 3600)
	 
c		    now calculate pathlength of curved ray using Lehn's parabolic approximation	 
		    PATHLENGTH = DSQRT(dist**2 + 
     *                  + (DABS(deltd)*0.001-0.5*(dist**2)/ Rearth)**2)
c		    now add terrestrial refraction based on modified Wikipedia expression
	        IF (DABS(hgtdif) > 1000.0) THEN
		        Exponent = 0.99
		    ELSE
				Exponent = 0.9965
				ENDIF	 
				
			elev(2,nazn) = elev(2,nazn) 
     *           - TRRbasis * PATHLENGTH**Exponent
			
		    ENDIF
		   
        IF (nazn.EQ.1) maxang = ABS(elev(1, nazn))
        GOTO 40
50      nazn = nazn - 1
        CLOSE(2)
        coment(1:8) = place(Nsetflag + 1)(1:8)

        IF ((.NOT. NOFILES) .AND. (NOINTERNET)) THEN
c          prepare filenames of output files
           nyear1 = nyear
           IF (ABS(nyear1) .LT. 1000) THEN
              nyear1 = ABS(nyear1) + 1000
              END IF
           WRITE(CH4,'(I4)')ABS(nyear1)
           IF ((nplaceflg.EQ.1).OR.(nplaceflg.EQ.2)) THEN
               name(1:4) = place(Nsetflag + 1)(1:4)
               name(5:8) = CH4
               END IF
           IF (Nsetflag .EQ. 0) THEN
              namet(1:1)=drivlet(1:1)
              namet(2:15) = ':/fordtm/netz/'
              namet(16:23) = name(1:8)
           ELSE IF (Nsetflag .EQ. 1) THEN
              namet(1:1)=drivlet(1:1)
              namet(2:15) = ':/fordtm/skiy/'
              namet(16:23) = name(1:8)
              END IF
           plcfil(1:23)=namet(1:23)
           plcfil(24:27)=EXT1(Nsetflag + 1, nplac(Nsetflag + 1))

           INQUIRE(FILE=namet,EXIST=EXISTS)
           IF (EXISTS) THEN
              OPEN(1,FILE=namet,STATUS='OLD')
           ELSE
              OPEN(1,FILE=namet,STATUS='NEW')
              END IF
           INQUIRE(FILE=plcfil,EXIST=EXISTS)
           IF (EXISTS) THEN
              OPEN(2,FILE=plcfil,STATUS='OLD')
           ELSE
              OPEN(2,FILE=plcfil,STATUS='NEW')
              END IF
           END IF

        td = geotd
c       time difference from Greenwich England
        IF (nplaceflg .EQ. 2) td = -8.0d0
        lr = cd * lt
        tf = td * 60.0D0 + 4.0D0 * lg
        ac = cd * .9856003D0
        mc = cd * .9856474D0
        ec = cd * 1.915D0
        e2c = cd * .02D0
        ob1 = cd * (-.0000004D0)
C       change of ecliptic angle per day
   
C       mean anomaly for Jan 0, 1996 00:00 at Greenwich meridian
        ap = cd * 357.528D0  
C       mean longitude of sun Jan 0, 1996 00:00 at Greenwich meridian
        mp = cd * 280.461D0
C       ecliptic angle for Jan 0, 1996 00:00 at Greenwich meridian
        ob = cd * 23.439D0
c       shift the ephemerials for the time zone
        ap = ap - (td / 24.0D0) * ac
        mp = mp - (td / 24.0D0) * mc
        ob = ob - (td / 24.0D0) * ob1
        Nyr = nyear
C       year for calculation
        Nyd = Nyr - 1996
        Nyf = 0
        IF (Nyd .LT. 0) THEN
           DO 100 NYRI=1995,NYR,-1
              NYRTST = NYRI
              NYLTST = 365
              IF (MOD(NYRTST,4) .EQ. 0) NYLTST = 366
              IF ((MOD(NYRTST,100).EQ.0) .AND.
     *            (MOD(NYRTST,400).NE.0)) NYLTST = 365
              Nyf = Nyf - NYLTST
100        CONTINUE
        ELSE IF (Nyd .GE. 0) THEN
           DO 150 NYRI=1996,NYR-1,1
              NYRTST = NYRI
              NYLTST = 365
              IF (MOD(NYRTST,4) .EQ. 0) NYLTST = 366
              IF ((MOD(NYRTST,100).EQ.0) .AND.
     *            (MOD(NYRTST,400).NE.0)) NYLTST = 365
              Nyf = Nyf + NYLTST
150        CONTINUE
           END IF

        Nyl = 365
        IF (MOD(Nyr,4).EQ.0) Nyl = 366
        IF ((MOD(Nyr,100).EQ.0) .AND.
     *      (MOD(Nyr,400).NE.0)) Nyl = 365

        Nyf = Nyf - 1462
        ob = ob + ob1 * Nyf
        mp = mp + Nyf * mc
600     IF (mp .LT. 0) mp = mp + 2.0D0 * pi
610     IF (mp .LT. 0) GOTO 600
620     IF (mp .GT. 2.0D0 * pi) mp = mp - 2.0D0 * pi
630     IF (mp .GT. 2.0D0 * pi) GOTO 620
640     ap = ap + Nyf* ac
650     IF (ap .LT. 0) ap = ap + 2.0D0 * pi
660     IF (ap .LT. 0) GOTO 650
670     IF (ap .GT. 2.0D0 * pi) ap = ap - 2.0D0 * pi
680     IF (ap .GT. 2.0D0 * pi) GOTO 670

        DO 925 dy=yrst(nyear-Nstrtyr+1),yrend(nyear-Nstrtyr+1)
		
c		    determine the minimum and average temperature for this day for current place
c			use Meeus's forumula p. 66 to convert daynumber to month, 
c			no need to interpolate between temepratures -- that is overkill
			K = 2
			IF (Nyl .EQ. 366) K = 1
			M = INT(9*(K + dy)/275 + 0.98)
c			convert minimum and average temperatures to Kelvin scale
			IF (Nsetflag .EQ. 0) THEN
c				sunrise, model temp. as minimum diurnal temperatures
				TK = MT(M) + 273.15
c				now find average minimum temperature to use in removing profile's TR	
			ELSEIF (Nsetflag .EQ. 1) THEN
c				sunset, model temp. as average diurnal temperature			
				TK = AT(M) + 273.15
				ENDIF				
c			calculate van der Werf temperature scaling factor for refraction
			VDWSF = (288.15/TK)**1.687499629598728
c			1.7081
c			calculate van der Werf scaling factor for altitudes
		    VDWALT = (288.15/TK)**0.690
c			calculate van der Werf ref scaling (?? why isn't it the same as VDWSF ??)
			VDWREFR = (288.15/TK)**2.19
		
c			calculate conserved portion of Wikipedia's terrestrial refraction expression
c			Lapse Rate, dt/dh, is set at -0.0065 degress K/m, i.e., the standard atmosphere
			TRbasis = 8.15 * P * 1000.0 * (0.0342 - 0.0065)/
     *           (TK * TK * 3600)

c	     day iteration
           nDyStp = NINT(dy)

           IF (NOINTERNET) THEN
c             update status file
              IF (DMOD(dy,73.0d0) .EQ. 0) THEN
                  stat2 = stat2 + stat
                  filstat(1:1)=drivlet(1:1)
                  filstat(2:10)=':/jk/stat'
                  filstat(15:18)='.tmp'
                  nstat=nstat+1
                  IF (nstat .LE. 9) THEN
                      filstat(11:13)='000'
                      WRITE(CH1,'(I1)')nstat
                      filstat(14:14)=CH1
                  Elseif ((nstat .LE. 99).AND.(nstat.GE.10)) THEN
                      filstat(11:12)='00'
                      WRITE(CH2,'(I2)')nstat
                      filstat(13:14)=CH2
                  Elseif ((nstat .LE. 999).AND.(nstat.GE.100)) THEN
                      WRITE(CH3,'(I3)')nstat
                      filstat(12:14)=CH3
                  Elseif ((nstat .LE. 9999).AND.(nstat.GE.1000)) THEN
                      WRITE(CH4,'(I4)')nstat
                      filstat(11:14)=CH4
                      END IF
                  INQUIRE(FILE=filstat,EXIST=EXISTS)
                  IF (EXISTS) THEN
                     OPEN(3,FILE=filstat,STATUS='OLD')
                  ELSE
                     OPEN(3,FILE=filstat,STATUS='NEW')
                     END IF
                   WRITE(3,'(F8.4)')stat2
                  CLOSE(3)
                  END IF
               END IF
           
           dy1 = dy
c          change ecliptic angle from noon of previous day to noon today
           ob = ob + ob1
           CALL DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
           ra = DATAN(DCOS(ob) * DTAN(es))
           IF (ra .LT. 0) ra = ra + pi
           df = ms - ra
685        IF (DABS(df) .GT. pi * 0.5D0) THEN
              IF (df .GT. 0) THEN
                 df = df - (df/DABS(df)) * pi
              ELSE IF (df .LT. 0) THEN
                 df = df + (df/DABS(df)) * pi
                 END IF
           ELSE
              GOTO 687
              END IF
           GOTO 685
687        et = (df / cd) * 4.0D0
c          equation of time (minutes)
           t6 = 720.0D0 + tf - et
c          astronomical noon in minutes
           fdy = DACOS(-DTAN(d) * DTAN(lr)) * ch / 24.0D0

c          astronomical refraction calculations, hgt in meters
           ref = 0.0d0
           eps = 0.0d0
           IF (hgt .GT. 0) lnhgt = DLOG(hgt*0.001d0)
c		   calculate total atmospheric refraction from the observer's height
c		   to the horizon and then to the end of the atmosphere
c		   All refraction terms have units of mrad
           IF (hgt .LE. 0.0D0) GOTO 690
           ref = EXP(vdwref(1) + vdwref(2) * lnhgt + 
     *           vdwref(3)*lnhgt*lnhgt + vdwref(4)*lnhgt*lnhgt*lnhgt +
     *           vdwref(5)*lnhgt*lnhgt*lnhgt*lnhgt +
     *           vdwref(6)*lnhgt*lnhgt*lnhgt*lnhgt*lnhgt)
           eps = EXP(vdweps(1) + vdweps(2) * lnhgt + 
     *           vdweps(3)*lnhgt*lnhgt + vdweps(4)*lnhgt*lnhgt*lnhgt +
     *           vdweps(5)*lnhgt*lnhgt*lnhgt*lnhgt)
			
		   RefExponent = refexp(1) + refexp(2) * hgt
	       RefNorms = refnorm(1) +  
     *         refnorm(2)*hgt + refnorm(3)*(hgt**2) +
     *         refnorm(4) * (hgt ** 3) + refnorm(5) * (hgt ** 4)
		   refFromExp = RefNorms * (288.15 / TK) ** RefExponent
			
c	       now add the all the contributions together due to the observer's height 
c		   along with the value of atm. ref. from hgt<=0, where view angle, a1 = 0
c		   for this calculation, leave the refraction in units of mrad
690        a1 = 0 
		   air = 90.0d0*cd + (eps+VDWREFR*ref+VDWSF*refrac1)/1.0D3
c		   gives almost the same value when calculated as below
c		   so it is sufficient to use the same temperature dependence as refrac1
c		   air = 90.0d0*cd + (eps+refFromExp+VDWSF*refrac1))/1.0D3
		   
c          a1 = (VDWREFR*ref + VDWSF*refrac1) / 1.0D3 //wrong -> will cause redundant scaling
		   a1 = (ref + refrac1) / 1.0D3
c		   gives almost the same value when calculated as below
c		   so it is sufficient to use the same temperature dependence as refrac1
c		   a1 = (refFromExp + VDWSF*refrac1) / 1.0D3
c		   leave a1 in radians			  
c		   a1 = DATAN(DTAN(a1) * VDWALT) ///Wrong - remove -> redundant temperature scaling		   

           IF (Nsetflag .EQ. 1) THEN
c              astron. sunset
               ndirec = -1
               dy1 = dy - ndirec * fdy
c              calculate sunset/sunrise
               CALL DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
               air = air + 0.2667d0 * (1 - 0.017d0 * DCOS(aas)) * cd
               sr1 = DACOS((-DTAN(lr) * DTAN(d)) +
     *               (DCOS(air) /(DCOS(lr)*DCOS(d)))) * ch
               sr = sr1 * hr
c              hour angle in minutes
               t3 = (t6 - ndirec * sr) / hr
               nacurr = 0
c*************************BEGIN SUNSET WINTER ADHOC FIX************************
              IF ((lt.GE.0).AND.(nweather.EQ.3).AND.((dy.LE.ns3).OR.
     *           (dy.GE.ns4)).AND. ADHOCSET ) THEN
c                 add winter addhoc fix
                  nacurr = 15
                  END IF
c*************************END SUNSET WINTER ADHOC FIX************************
               t3 = t3 + nacurr / 3.6D3
               t3sub = t3
c              astronomical sunrise/sunset in hours.frac hours
               IF (ASTRO) THEN
                  azi1 = DCOS(d) * DSIN(sr1 / ch)
             azi2 = -DCOS(lr)*DSIN(d) + DSIN(lr)*DCOS(d)*DCOS(sr1/ch)
                  azio = DATAN(azi1/azi2)
                  if (azio .GT. 0) azi1 = 90.0D0 - azio / cd
                  if (azio .LT. 0) azi1 = -(90.0D0 + azio / cd)

c                 write astronomical sunset times to file and loop in days

                  IF ((.NOT. NOFILES) .AND. (NOINTERNET)) THEN
c                    write output to intermediate files
                     IF (dy .EQ. 1) THEN
                         PRINT *,'sunset for year ',yr,'place: ',fileo
                         END IF
                     CALL T1500(t3sub,negch)
                     tssc = t3subc
c                    convert dy to calendar date
                     CALL T1750(dy,Nyl,Nyr)
                     WRITE(*,'(1X,A11,3X,A8)')caldayc, tssc
                     WRITE(1,'(A11,3X,A8)')caldayc, tssc
c                    add dist to obst to be consistent w/visible formats                  
                     dobs = 120
                     WRITE(2,710)caldayc, tssc, comentss, azi1, dobs
                     END IF

                  GOTO 920
                  END IF
           ELSE IF (Nsetflag .EQ. 0) THEN
c	         astron. sunrise		
               ndirec = 1
               dy1 = dy + fdy
               CALL DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
               air = air + 0.2667d0 * (1 - 0.017d0 * DCOS(aas)) * cd
c              calculate sunset
               sr2 = DACOS((-DTAN(lr) * DTAN(d)) +
     *               (DCOS(air) /(DCOS(lr)*DCOS(d)))) * ch
               sr2 = sr2 * hr
               tss = (t6 + sr2) / hr
c              sunset in hours.fraction of hours
               dy1 = dy - fdy
c              calculate sunrise
               CALL DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
               sr1 = DACOS((-DTAN(lr) * DTAN(d)) +
     *               (DCOS(air) /(DCOS(lr)*DCOS(d)))) * ch
               sr = sr1 * hr
c              hour angle in minutes
               t3 = (t6 - sr) / hr
               nacurr = 0
c*************************BEGIN SUNRISE WINTER ADHOC FIX************************
              IF ((lt.GE.0).AND.(nweather.EQ.3).AND.((dy.LE.ns3).OR.
     *           (dy.GE.ns4)) .AND. ADHOCRISE ) THEN
c                 add winter addhoc fix
                  nacurr = -15
                  END IF
c*************************END SUNRISE WINTER ADHOC FIX************************
               t3 = t3 + nacurr / 3.6D3
               t3sub = t3
c              astronomical sunrise in hours.frac hours
c              sr1= hour angle in hours at astronomical sunrise
               IF (ASTRO) THEN
                  azi1 = DCOS(d) * DSIN(sr1 / ch)
             azi2 = -DCOS(lr)*DSIN(d) + DSIN(lr)*DCOS(d)*DCOS(sr1/ch)
                  azio = DATAN(azi1/azi2)
                  if (azio .GT. 0) azi1 = 90.0D0 - azio / cd
                  if (azio .LT. 0) azi1 = -(90.0D0 + azio / cd)
c                 write astronomical sunrise time to file and loop in days

                  IF ((.NOT. NOFILES) .AND. (NOINTERNET)) THEN
c                    write output to intermediate files
                      IF (dy .EQ. 1) THEN
                         PRINT *,'sunrise for year ',yr,'place: ',fileo
                         END IF
                     CALL T1500(t3sub,negch)
                     ts1c = t3subc
c                    convert dy to calendar date
                     CALL T1750(dy,Nyl,Nyr)
                     WRITE(*,'(1X,A11,3X,A8)')caldayc, ts1c
                     WRITE(1,'(A11,3X,A8)')caldayc, ts1c
c                    add dist to obst to be consistent w/visible formats                  
                     dobs = 120
                     WRITE(2,785)caldayc, ts1c, comentss, azi1, dobs
                     END IF

                  GOTO 920
                  END IF
               END IF

692      nn = 0
         nfound = 0
         nabove = 0
         DO 850 nloops = nloop,nlpend,ndirec
            nn = nn + 1
            hran = (sr1 - (stepsec * nloops) / 3.6D3)
C           step in hour angle by stepsec sec
            dy1 = dy - ndirec * hran / 24.0D0
            CALL DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
C           recalculate declination at each new hour angle
c			zenith angle at this hour angle is:
            z = DSIN(lr)*DSIN(d) + DCOS(lr)*DCOS(d)*DCOS(hran/ch)
            z = DACOS(z)
c			the view angle, i.e., altitude, alt1, is the complement of z			
            alt1 = (pi * 0.5D0 - z)
c           altitude in radians of sun's center
            azi1 = DCOS(d) * DSIN(hran/ch)
            azi2 = -DCOS(lr)*DSIN(d)+DSIN(lr)*DCOS(d)*DCOS(hran/ch)
            azi = (DATAN(azi1/azi2)) / cd 
            if (azi .GT. 0) azi1 = 90.0D0 - azi
            if (azi .LT. 0) azi1 = -(90.0D0 + azi)
            IF ((DABS(azi1) .GT. maxang).AND.(NOINTERNET)) THEN
               print *,'!!!passed max calculated azi,',maxang,'!!!'
               print *,'for file: ',fileo
               print *,'              ***ABORT***'
               GOTO 9999
               END IF
            azi1 = ndirec * azi1
            IF (Nsetflag .EQ. 1) THEN
C              to first order, the top and bottom of the sun have the
C              following altitudes near the sunset
c			   adjust for temperature by using the vdW scaling of altitude for temp //wrong ->wil produce redunant scaling
c              al1 = DATAN(DTAN(alt1 + a1)*VDWALT)/cd  //wrong -> will produce redundant scaling
			   al1 = (alt1 + a1)/cd
C              first guess for apparent top of sun
               pt = 1
               CALL REFVDW(hgt,al1,refrac1)
C              first iteration to find position of upper limb of sun
c               al1 = alt1 / cd + .2666 + pt*VDWSF*refrac1*dcmrad //wrong -> will produce redundant scaling
			   al1 = alt1 / cd + .2666 + pt * refrac1 * dcmrad
c			   'rescale the altitude for temperature using the vdW scaling exponent
			   al1 = DATAN(DTAN(al1*cd)*VDWALT)/cd
c              top of sun with refraction
               IF (niter .EQ. 1) THEN
                  DO 693 nac = 1,10
                     al1o = al1
                CALL REFVDW(hgt,al1,refrac1)
c                    subsequent iterations
c                     al1 = alt1 / cd + .2666 + pt*VDWSF*refrac1*dcmrad //wrong -> will produce redundant scaling
					 al1 = alt1 / cd + .2666 + pt * refrac1 * dcmrad
c					 apply vdW temperature scaling to this altitude
					 al1 = DATAN(DTAN(al1*cd)*VDWALT)/cd
c                    top of sun with refraction
                     IF (ABS(al1 - al1o) .LT. .004) GOTO 695
693               CONTINUE
                  END IF
695            ne = INT((azi1 +  maxang) * 10) + 1
               elevinty = elev(2, ne + 1) - elev(2, ne)
               elevintx = elev(1, ne + 1) - elev(1, ne)
               elevint = (elevinty/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(2, ne)
	 
c////////////// now modify apparent elevention, elevint, with terrestrial refraction/////
c			   first calculate interpolated difference of heights
			   hgtdifx = elev(4, ne + 1) - elev(4, ne)
			   hgtdif = (hgtdifx/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(4, ne)
			   hgtdif = hgt - hgtdif
c			   convert to kms
			   hgtdif = hgtdif * 0.001
c			   now calculate interpolated distance to obstruction along earth	 
			   distx = elev(3, ne + 1) - elev(3, ne)
			   dist = (distx/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(3, ne)
c			   now calculate pathlength of curved ray using Lehn's parabolic approximation	 
			   PATHLENGTH = DSQRT(dist**2 + 
     *         + (DABS(hgtdif)*0.001-0.5*(dist**2)/ Rearth)**2)
c			   now add terrestrial refraction based on modified Wikipedia expression
		       IF (DABS(hgtdif) > 1000.0) THEN
					Exponent = 0.99
		       ELSE
					Exponent = 0.9965
				 ENDIF	 
			   elevint = elevint + TRbasis * PATHLENGTH**Exponent
c////////////////////////////////////////////////////////////////////////////////////////
	 
               IF (((elevint.GE.al1).OR.(nloops.LE.0)).AND.
     *              (nabove .EQ. 1)) THEN
c                 reached sunset
                  nfound = 1
c                 margin of error
                  nacurr = 0

                  IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
                     t3sub = t3
                     CALL T1500(t3sub,negch)
                     ts1c = t3subc
                     t3sub = ABS((stepsec * nloops + nacurr) /3.6D3)

c                    visible sunset can't be latter then astronomical sunset!
                     IF (nloops .LE. 0) t3sub = 0

                     IF (NOINTERNET .AND. (ndruk .EQ. 1)) THEN
                        WRITE(4,*)dy, t3sub
                        END IF
                     CALL T1500(t3sub,negch)
                     ts2c='         '
                     ts2c(2:9) = t3subc(1:8)
                     IF (nloops .GT. 0) THEN
                        ts2c(1:1)='-'
                        END IF
                     END IF

                  t3sub = t3 + ndirec * (stepsec*nloops+nacurr)/3.6D3
                  IF (nloops .LT. 0) THEN
c                     visible sunset can't be latter then astronomical sunset!
                      t3sub = t3
                      END IF

c                 <<<<<<<<<distance to obstruction>>>>>> 
                  dobsy = elev(3, ne + 1) - elev(3, ne)
                  elevintx = elev(1, ne + 1) - elev(1, ne)
                  dobs = (dobsy/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(3, ne)

                  IF ((NOINTERNET) .AND. (.NOT. NOFILES)) THEN
                     CALL T1500(t3sub,negch)
                     tssc = t3subc
c                    sunset
c                    convert dy to calendar date
                     CALL T1750(dy,Nyl,Nyr)
                     IF (dy .EQ. 1) THEN
                         PRINT *,'sunset for year ',yr,'place: ',fileo
                         END IF

                      WRITE(*,700)caldayc, tssc, ts1c, ts2c, nac
700                   FORMAT(1X,A11,3X,A8,3X,A8,3X,A9,3X,I1)

                      WRITE(1,705)caldayc, tssc, ts1c, ts2c
705                   FORMAT(A11,3X,A8,3X,A8,3X,A9)
                      WRITE(2,710)caldayc, tssc, comentss, azi1, dobs
710                   FORMAT(A11,3X,A8,2X,A12,3X,F7.3,3X,F6.2)
                      END IF

                  GOTO 900

               ELSE IF (elevint .LT. al1) THEN
                  nabove = 1
                  END IF

            ELSE IF (Nsetflag .EQ. 0) THEN
                  al1 = (alt1 + a1)/cd
c                 first guess for apparent top of sun
                  pt = 1 
                  CALL REFVDW(hgt,al1,refrac1)
c                 first iteration to find top of sun
c				  al1 = alt1 / cd + .2666 + pt * VDWSF*refrac1*dcmrad //wrong->redundant temperature scaling
				  al1 = alt1 / cd + .2666 + pt * refrac1 * dcmrad
c				  apply vdW scaling law to altitude
				  al1 = DATAN(DTAN(al1*cd)*VDWALT)/cd				  
c                 top of sun with refraction
                  IF (niter .EQ. 1) THEN
                     DO 750 nac = 1,10
                        al1o = al1
						CALL REFVDW(hgt,al1,refrac1)
c                       subsequent iterations
c						al1 = alt1 / cd + .2666+pt*VDWSF*refrac1*dcmrad //wrong->redundant temperature scaling
						al1 = alt1 / cd + .2666 + pt * refrac1 * dcmrad
c					    apply vdW scaling law to altitude						
						al1 = DATAN(DTAN(al1*cd)*VDWALT)/cd						
c                       top of sun with refraction
                        IF (ABS(al1 - al1o) .LT. .004) GOTO 760
750                  CONTINUE
                     END IF
760               alt1 = al1
c              now search digital elevation data
               ne = INT((azi1 +  maxang) * 10) + 1
               elevinty = elev(2, ne + 1) - elev(2, ne)
               elevintx = elev(1, ne + 1) - elev(1, ne)
               elevint = (elevinty/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(2, ne)
	 
c////////////// now modify apparent elevention, elevint, with terrestrial refraction/////
c			   first calculate interpolated difference of heights
			   hgtdifx = elev(4, ne + 1) - elev(4, ne)
			   hgtdif = (hgtdifx/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(4, ne)
			   hgtdif = hgt - hgtdif
c			   convert to kms
			   hgtdif = hgtdif * 0.001
c			   now calculate interpolated distance to obstruction along earth	 
			   distx = elev(3, ne + 1) - elev(3, ne)
			   dist = (distx/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(3, ne)
c			   now calculate pathlength of curved ray using Lehn's parabolic approximation	 
			   PATHLENGTH = DSQRT(dist**2 + 
     *         + (DABS(hgtdif)*0.001-0.5*(dist**2)/ Rearth)**2)
c			   now add terrestrial refraction based on modified Wikipedia expression
		       IF (DABS(hgtdif) > 1000.0) THEN
					Exponent = 0.99
		       ELSE
					Exponent = 0.9965
				 ENDIF	 
			   elevint = elevint + TRbasis * PATHLENGTH**Exponent
c////////////////////////////////////////////////////////////////////////////////////////
	 
               IF (elevint .LE. alt1) THEN
                  IF (nn .GT. 1) nfound = 1
                  IF (nfound .EQ. 0) GOTO 900
                  IF ((dy .EQ. 1) .AND. (NOINTERNET)) THEN
                     PRINT *,'sunrise for year: ', Nyr, ' place: ',fileo
                     END IF

c                 <<<<<<<<<distance to obstruction>>>>>> 
                  dobsy = elev(3, ne + 1) - elev(3, ne)
                  elevintx = elev(1, ne + 1) - elev(1, ne)
                  dobs = (dobsy/elevintx)*(azi1 - elev(1, ne))
     *                    + elev(3, ne)

c                 sunrise
                  IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
                     t3sub = t3
                     CALL T1500(t3sub,negch)
                     ts1c = t3subc
c                    astronomical time
                     t3sub = (stepsec * nloops + nacurr) / 3.6D3
                     IF (nloops .LT. 0) THEN
                        t3sub = 0
                        END IF
                     IF (ndruk .EQ. 1) THEN
                        WRITE(4,*)dy, t3sub
                        END IF
                     CALL T1500(t3sub,negch)
                     ts2c='         '
                     ts2c(2:9) = t3subc(1:8)
c                    difference between astronom to sunrise
                     END IF 

                  nacurr = 0
c                 margin of error (seconds) to take into
c                 account differences between observations and the DTM ``
                  t3sub = t3 + (stepsec * nloops + nacurr) / 3.6D3

                  IF (nloops .LT. 0) THEN
c                    sunrise can't be before astronomical sunrise
c                    so set it equal to astronomical sunrise time
c                    minus the adhoc winter fix if any
                     t3sub = t3
                     END IF

                  IF (NOINTERNET .AND. (.NOT. NOFILES)) THEN
                     CALL T1500(t3sub,negch)
                     tssc = t3subc
c                    convert dy to calendar date
                     CALL T1750(dy,Nyl,Nyr)
                     WRITE(*,770)caldayc, tssc, ts1c, ts2c, nac
770                  FORMAT(1X,A11,3X,A8,3X,A8,3X,A9,3X,I1)
                     WRITE(1,775)caldayc, tssc, ts1c, ts2c
775                  FORMAT(A11,3X,A8,3X,A8,3X,A9)
                     WRITE(2,785)caldayc, tssc, comentss, azi1, dobs
785                  FORMAT(A11,2X,A8,3X,A12,3X,F7.3,3X,F6.2)
                     END IF

                  GOTO 900

                  END IF
               END IF
850      CONTINUE
900      IF ((nfound .EQ. 0).AND.(nloop .GT. 0)
     *      .AND. (Nsetflag.EQ.1)) THEN
            IF (NOINTERNET) THEN
               PRINT *,'*****search error-changing nloop*****'
               ENDIF
            IF (Nsetflag .EQ. 0) THEN
               nloop = nloop - 24
               IF (nloop .LT. 0) nloop = 0
            ELSE IF (Nsetflag .EQ. 1) THEN
               nloop = nloop + 24
               END IF
            GOTO 692
         ELSEIF ((nfound .EQ. 0).AND.(nloop.GE.-100)
     *       .AND. (Nsetflag.EQ.0)) THEN
            IF (NOINTERNET) THEN
               PRINT *,'*****search error-changing nloop*****'
               END IF
            IF (Nsetflag .EQ. 0) THEN
               nloop = nloop - 24
c               IF (nloop .LT. 0) nloop = 0
            ELSE IF (Nsetflag .EQ. 1) THEN
               nloop = nloop + 24
               END IF
            GOTO 692
         ELSE IF ((nfound.EQ.0).AND.(nloop.LT.-100)
     *                   .AND.(Nsetflag.EQ.0)) THEN
            IF (NOINTERNET) THEN 
               PRINT *,'***search not successful, increase nlpend***'
               END IF
            nlpend = nlpend + 500
            nloop = 72
            GOTO 692
            END IF
         IF (Nsetflag .EQ. 1) THEN
            IF ((nloops+30.LT.nloop).OR.(nloops+30.GT.nloop)) THEN
               nloop = nloops + 30
               END IF
         ELSE IF (Nsetflag .EQ. 0) THEN
            IF (nloops + 30 .LT. nloop) THEN
               nloop = nloops + 30
            ELSE IF (nloops + 30 .GT. nloop) THEN
               nloop = nloops - 30
               END IF
            END IF

c        //////////////calculations complete/////////////////////
c        check for latest sunset times and earliest sunrise times
c        among all the places (this replaces time comparisons of COMPARE)
920      IF (nfil .EQ. 1) THEN
c           first place, so intialize latest sunset arrays
            tsscs(Nsetflag + 1, nYrStp, nDyStp) = t3sub
            coments(Nsetflag + 1, nYrStp, nDyStp) = comentss
            azils(Nsetflag + 1, nYrStp, nDyStp) = azi1
            dobss(Nsetflag + 1, nYrStp, nDyStp) = dobs
         ELSE
c           check if sunrise/sunset for new place is earlier/later
            IF (Nsetflag .EQ. 0) THEN   
c              sunrise
               IF (t3sub .LT. tsscs(Nsetflag + 1, nYrStp, nDyStp)) THEN
                  tsscs(Nsetflag + 1, nYrStp, nDyStp) = t3sub
                  coments(Nsetflag + 1, nYrStp, nDyStp) = comentss
                  azils(Nsetflag + 1, nYrStp, nDyStp) = azi1
                  dobss(Nsetflag + 1, nYrStp, nDyStp) = dobs
                  END IF
            ELSE IF (Nsetflag .EQ. 1) THEN
c              sunset
               IF (t3sub .GT. tsscs(Nsetflag + 1, nYrStp, nDyStp)) THEN
                  tsscs(Nsetflag + 1, nYrStp, nDyStp) = t3sub
                  coments(Nsetflag + 1, nYrStp, nDyStp) = comentss
                  azils(Nsetflag + 1, nYrStp, nDyStp) = azi1
                  dobss(Nsetflag + 1, nYrStp, nDyStp) = dobs
                  END IF
                END IF
            END IF

925     CONTINUE

975     CONTINUE

        CLOSE(1)
        CLOSE(2)
        IF (ndruk.EQ.1) CLOSE(4)

1000    CONTINUE
        CLOSE(10)

c       PROGRAM COMPARE --finds earliest/latest times of nplac .PLC files
c       number of places to search
        filout(16:19)=PLACN(1, 1)(5:8)
        IF (nplac(1).EQ.0) filout(16:19)=PLACN(2, 1)(5:8)

        NSET1 = 0
        NSET2 = 1
        IF (nplac(1).EQ.0) NSET1=1
        IF (nplac(2).EQ.0) NSET2=0
        DO 3000 Nsetflag = NSET1,NSET2
           IF (Nsetflag .EQ. 0) THEN
              filout(1:1)=drivlet(1:1)
              filout(2:15)=':/fordtm/netz/'
              nadd = 0
           ELSE IF (Nsetflag .EQ. 1) THEN
              filout(1:1)=drivlet(1:1)
              filout(2:15)=':/fordtm/skiy/'
              nadd = 1
              END IF

           IF ((nplac(Nsetflag+1) .EQ. 1) .AND. 
     *        (NOFILES .OR. NOINTERNET)) GOTO 2500
c
           DO 2000 nyr = Nstrtyr,Nendyr
              nYrStp = nyr - Nstrtyr + 1

c             determine number of days in this year              
              Nyl = 365
              IF (MOD(nyr,4).EQ.0) Nyl = 366
              IF ((MOD(nyr,100).EQ.0) .AND.
     *            (MOD(nyr,400).NE.0)) Nyl = 365

              nyr1=nyr
              IF (ABS(nyr1) .LT. 1000) THEN
                 nyr1 = ABS(nyr1) + 1000
                 END IF
              WRITE(CH4,'(I4)')ABS(nyr1)
              filout(16:19)=CN4(1:4)
              filout(20:23)=CH4(1:4)

              nleapyr = 0
              IF (nyl .EQ. 366) nleapyr = 1
              filout(24:27) = '.com'
              INQUIRE(file=filout,EXIST=EXISTS)
              IF (EXISTS) THEN
                 OPEN(1,file=filout,STATUS='OLD')
              ELSE
                 OPEN(1,file=filout,STATUS='NEW')
                 END IF

              I1 = yrst(nyr-Nstrtyr+1)
              I2 = yrend(nyr-Nstrtyr+1)
              DO 1500 I = I1, I2

c              convert daynumber to date               
               CALL T1750(DBLE(I),Nyl,nyr)
               DAT = caldayc

c              convert fractional hour to time
               t3sub = tsscs(Nsetflag + 1, nYrStp, I)
               CALL T1500(t3sub,negch)
               timss = t3subc

               comentss = coments(Nsetflag + 1, nYrStp, I)
               azinss = azils(Nsetflag + 1, nYrStp, I)
               dobsf = dobss(Nsetflag + 1, nYrStp, I)

               WRITE(1,1305)DAT,timss(1:8),comentss,azinss,dobsf
1305           FORMAT(A11,3X,A8,6X,A12,3X,F7.3,3X,F6.2)

1500          CONTINUE


2000    CONTINUE
        CLOSE(1)
        GOTO 2500


2500    filout(16:27)='NETZSKIY.TM2'
        INQUIRE (FILE=filout,EXIST=EXISTS)
        IF (EXISTS) THEN
           OPEN(3,FILE=filout,status='OLD')
        ELSE
           OPEN(3,FILE=filout,status='NEW')
           END IF
        IF (nplac(Nsetflag + 1) .LT. 10) THEN
           WRITE(3,'(I1)')nplac(Nsetflag + 1)
        ELSE
           WRITE(3,'(I2)')nplac(Nsetflag + 1)
           END IF
        DO 2700 nyr = Nstrtyr,Nendyr
           filout(16:19)=CN4(1:4)
           WRITE(CH4,'(I4)')nyr

           IF ((nplac(Nsetflag+1) .EQ. 1) .AND. 
     *        (NOFILES .OR. NOINTERNET)) GOTO 2600

           filout(20:23)=CH4(1:4)
           filout(24:27) = '.com'
           WRITE(3,'(A1,A27,A1)')'"',filout,'"'
           GOTO 2700
2600       plcfil(1:15) = filout(1:15)
           plcfil(20:23) = CH4(1:4)
           plcfil(24:27)=EXT1(Nsetflag + 1, nplac(Nsetflag + 1))
           plcfil(16:19) = place(Nsetflag + 1)(1:4)
           WRITE(3,'(A1,A27,A1)')'"',plcfil,'"'
2700    CONTINUE
        CLOSE(3)

3000    CONTINUE

c       now write closing file to tell Cal Program that all
c       calculations have ended
        filstat(1:1)=drivlet(1:1)
        filstat(2:17)=':/jk/netzend.tmp'
        INQUIRE(FILE=filstat,EXIST=EXISTS)
        IF (EXISTS) THEN
           OPEN(1,FILE=filstat,STATUS='OLD')
        ELSE
           OPEN(1,FILE=filstat,STATUS='NEW')
           END IF
        CLOSE(1)

9999  END

        SUBROUTINE DECL(dy1,mp,mc,ap,ac,ms,aas,es,ob,d)
c       calculates new declination
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
        pi = DACOS(-1.0d0)
        pi2 = 2.0D0 * pi
        cd = pi / 180.0d0
        ec = cd * 1.915D0
        e2c = cd * .020D0
        ms = mp + mc * dy1
        IF (ms .GT. pi2)  ms = ms - pi2
        aas = ap + ac * dy1
        IF (aas .GT. pi2) aas = aas - pi2
        es = ms + ec * DSIN(aas) + e2c * DSIN(2.0D0 * aas)
        IF (es .GT. pi2) es = es - pi2
        d = DASIN(DSIN(ob) * DSIN(es))
        RETURN
        END


        SUBROUTINE T1500(t3sub,negch)
c       convert hour.frac hour into hrs:min:sec
c       seconds are rounded UP to nearest second
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
        CHARACTER tshr*2,tsmin*2,tssec*2
        CHARACTER t3subc*8,caldayc*11
        COMMON /t3s/t3subc, caldayc
        tshr = '  '
        tsmin = '  '
        tssec = '  '
c1500
        Nt3hr = INT(t3sub)
        Nt3min = INT((t3sub - Nt3hr) * 60.0)
        t1=REAL(Nt3min/60.0)
        Nt3sec = INT((t3sub - Nt3hr - t1) * 3600 + .5)
        If (Nt3sec .LT. 0) THEN
           negch=-1
           NT3sec=ABS(NT3sec)
        ELSE
           negch=1
           ENDIF
        IF (Nt3sec.EQ.60) THEN
            Nt3min = Nt3min + 1
            Nt3sec = 0
            END IF
        IF (Nt3min .EQ. 60) THEN
            Nt3hr = Nt3hr + 1
            Nt3min = 0
            END IF
        WRITE(tshr,'(I2)')Nt3hr
        WRITE(tsmin,'(I2)')Nt3min
        IF (tsmin(1:1) .EQ. ' ') tsmin(1:1)='0'
        WRITE(tssec,'(I2)')Nt3sec
        IF (tssec(1:1) .EQ. ' ') tssec(1:1)='0'
        t3subc(1:2)=tshr(1:2)
        t3subc(3:3)=':'
        t3subc(4:5)=tsmin(1:2)
        t3subc(6:6)=':'
        t3subc(7:8)=tssec(1:2)
        RETURN
        END

        SUBROUTINE T1750(dy,Nyl,Nyr)
        CHARACTER t3subc*8,caldayc*11,CH2*2,CH4*4
        COMMON /t3s/t3subc,caldayc
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
c       'convert dy at particular yr to calendar date
c1750
        caldayc = '           '
        nleap = 0
        ndy = INT(dy)
        IF (Nyl.EQ.366) nleap = 1
        IF (ndy .LE. 31) THEN
            caldayc(1:4) = 'Jan-'
            WRITE(CH2,'(I2)')ndy
        ELSE IF ((Nyl.EQ.365).AND.(ndy.GT.31).AND.(ndy.LE.59))THEN
            caldayc(1:4) = 'Feb-'
            WRITE(CH2,'(I2)')ndy - 31
            nleap = 0
        ELSE IF ((Nyl.EQ.366).AND.(ndy.GT.31).AND.(ndy.LE.60)) THEN
            caldayc(1:4) = 'Feb-'
            WRITE(CH2,'(I2)')ndy - 31
            nleap = 1
        ELSE IF ((ndy.GT. 59+nleap).AND.(ndy.LE. 90+nleap)) THEN
            caldayc(1:4) = 'Mar-'
            WRITE(CH2,'(I2)')ndy - 59 - nleap
        ELSE IF ((ndy.GT. 90+nleap).AND.(ndy.LE. 120+nleap)) THEN
            caldayc(1:4) = 'Apr-'
            WRITE(CH2,'(I2)')ndy - 90 - nleap
        ELSE IF ((ndy.GT. 120+nleap).AND.(ndy.LE. 151+nleap)) THEN
            caldayc(1:4) = 'May-'
            WRITE(CH2,'(I2)')ndy - 120 - nleap
        ELSE IF ((ndy.GT. 151+nleap).AND.(ndy.LE. 181+nleap)) THEN
            caldayc(1:4) = 'Jun-'
            WRITE(CH2,'(I2)')ndy - 151 - nleap
        ELSE IF ((ndy.GT. 181 + nleap).AND.(ndy.LE. 212+nleap)) THEN
            caldayc(1:4) = 'Jul-'
            WRITE(CH2,'(I2)')ndy - 181 - nleap
        ELSE IF ((ndy.GT. 212+nleap).AND.(ndy .LE. 243+nleap)) THEN
            caldayc(1:4) = 'Aug-'
            WRITE(CH2,'(I2)')ndy - 212 - nleap
        ELSE IF ((ndy.GT. 243+nleap).AND.(ndy .LE. 273+nleap)) THEN
            caldayc(1:4) = 'Sep-'
            WRITE(CH2,'(I2)')ndy - 243 - nleap
        ELSE IF ((ndy.GT. 273+nleap).AND.(ndy .LE. 304+nleap)) THEN
            caldayc(1:4) = 'Oct-'
            WRITE(CH2,'(I2)')ndy - 273 - nleap
        ELSE IF ((ndy.GT. 304+nleap).AND.(ndy .LE. 334+nleap)) THEN
            caldayc(1:4) = 'Nov-'
            WRITE(CH2,'(I2)')ndy - 304 - nleap
        ELSE IF (ndy .GT. 334 + nleap) THEN
            caldayc(1:4) = 'Dec-'
            WRITE(CH2,'(I2)')ndy - 334 - nleap
            END IF
        caldayc(5:6)=CH2(1:2)
        IF (caldayc(5:5).EQ.' ') caldayc(5:5) = '0'
        caldayc(7:7) = '-'
        Nyr1 = Nyr
        IF (ABS(Nyr1) .LT. 1000) THEN
           Nyr1 = ABS(Nyr1) + 1000
           END IF
        WRITE(CH4,'(I4)')ABS(Nyr1)
        caldayc(8:11)=CH4(1:4)
        RETURN
        END

		SUBROUTINE REFVDW(hgt,x,fref)
c		use the polynomial fit coefficients of the vdw refraction 
c		as a function of observer's height, hgt in meters, and the apparent
c`		view angle, x in degrees, to calculate the total atms. refraction in mrad
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
		REAL *8 Coef(5,11), CA(11) 
		
c		polynomical fit coefficients used to calculate
c		vdw total atmospheric refraction vs height
		Coef(1, 1) = 9.56267125268496
		Coef(2, 1) = -8.6718429211079E-04
		Coef(3, 1) = 3.1664677349513E-08
		Coef(4, 1) = -2.04067678864827E-13
		Coef(5, 1) = -6.21413591282229E-17
		Coef(1, 2) = -3.54681762174248
		Coef(2, 2) = 3.05885370538294E-04
		Coef(3, 2) = -3.48413989765623E-09
		Coef(4, 2) = -3.27424677578751E-12
		Coef(5, 2) = 4.85180396156723E-16
		Coef(1, 3) = 1.00487516923555
		Coef(2, 3) = -7.12411305623716E-05
		Coef(3, 3) = -1.30264294792463E-08
		Coef(4, 3) = 6.08386198681256E-12
		Coef(5, 3) = -8.26564806865056E-16
		Coef(1, 4) = -0.234676117102
		Coef(2, 4) = 4.23105602906229E-06
		Coef(3, 4) = 1.25823603467313E-08
		Coef(4, 4) = -4.77064146898649E-12
		Coef(5, 4) = 6.38020633504241E-16
		Coef(1, 5) = 4.55474692911979E-02
		Coef(2, 5) = 4.20546127818185E-06
		Coef(3, 5) = -5.71051596715397E-09
		Coef(4, 5) = 2.05052061222564E-12
		Coef(5, 5) = -2.73486893484326E-16
		Coef(1, 6) = -7.07867490693562E-03
		Coef(2, 6) = -1.77071205323987E-06
		Coef(3, 6) = 1.51606775168871E-09
		Coef(4, 6) = -5.33770875527936E-13
		Coef(5, 6) = 7.12762362198471E-17
		Coef(1, 7) = 8.32295487796478E-04
		Coef(2, 7) = 3.56012191465623E-07
		Coef(3, 7) = -2.51083454939817E-10
		Coef(4, 7) = 8.78112651062853E-14
		Coef(5, 7) = -1.17556897803785E-17
		Coef(1, 8) = -6.96285190742393E-05
		Coef(2, 8) = -4.19488560840601E-08
		Coef(3, 8) = 2.62823518700555E-11
		Coef(4, 8) = -9.18795592984757E-15
		Coef(5, 8) = 1.23368951173372E-18
		Coef(1, 9) = 3.85246830558751E-06
		Coef(2, 9) = 2.93954176835877E-09
		Coef(3, 9) = -1.6910427905383E-12
		Coef(4, 9) = 5.92963984786967E-16
		Coef(5, 9) = -7.98575122972107E-20
		Coef(1, 10) = -1.25306160093963E-07
		Coef(2, 10) = -1.13531297031204E-10
		Coef(3, 10) = 6.10692317760114E-14
		Coef(4, 10) = -2.15222164778055E-17
		Coef(5, 10) = 2.90679288057578E-21
		Coef(1, 11) = 1.80519843190424E-09
		Coef(2, 11) = 1.8626817474919E-12
		Coef(3, 11) = -9.47783437562306E-16
		Coef(4, 11) = 3.36109376083127E-19
		Coef(5, 11) = -4.55153453427451E-23	
		
c		write out all loops inline for speed		
        CA(1) = Coef(1, 1) + Coef(2, 1) * hgt + Coef(3, 1)*hgt*hgt + 
     *          Coef(4, 1)*hgt*hgt*hgt + Coef(5, 1)*hgt*hgt*hgt*hgt
        CA(2) = Coef(1, 2) + Coef(2, 2) * hgt + Coef(3, 2)*hgt*hgt + 
     *          Coef(4, 2)*hgt*hgt*hgt + Coef(5, 2)*hgt*hgt*hgt*hgt
        CA(3) = Coef(1, 3) + Coef(2, 3) * hgt + Coef(3, 3)*hgt*hgt + 
     *          Coef(4, 3)*hgt*hgt*hgt + Coef(5, 3)*hgt*hgt*hgt*hgt
        CA(4) = Coef(1, 4) + Coef(2, 4) * hgt + Coef(3, 4)*hgt*hgt + 
     *          Coef(4, 4)*hgt*hgt*hgt + Coef(5, 4)*hgt*hgt*hgt*hgt
        CA(5) = Coef(1, 5) + Coef(2, 5) * hgt + Coef(3, 5)*hgt*hgt + 
     *          Coef(4, 5)*hgt*hgt*hgt + Coef(5, 5)*hgt*hgt*hgt*hgt
        CA(6) = Coef(1, 6) + Coef(2, 6) * hgt + Coef(3, 6)*hgt*hgt + 
     *          Coef(4, 6)*hgt*hgt*hgt + Coef(5, 6)*hgt*hgt*hgt*hgt
        CA(7) = Coef(1, 7) + Coef(2, 7) * hgt + Coef(3, 7)*hgt*hgt + 
     *          Coef(4, 7)*hgt*hgt*hgt + Coef(5, 7)*hgt*hgt*hgt*hgt
        CA(8) = Coef(1, 8) + Coef(2, 8) * hgt + Coef(3, 8)*hgt*hgt + 
     *          Coef(4, 8)*hgt*hgt*hgt + Coef(5, 8)*hgt*hgt*hgt*hgt
        CA(9) = Coef(1, 9) + Coef(2, 9) * hgt + Coef(3, 9)*hgt*hgt + 
     *          Coef(4, 9)*hgt*hgt*hgt + Coef(5, 9)*hgt*hgt*hgt*hgt
        CA(10)= Coef(1, 10) + Coef(2, 10) * hgt + Coef(3, 10)*hgt*hgt + 
     *          Coef(4, 10)*hgt*hgt*hgt + Coef(5, 10)*hgt*hgt*hgt*hgt
        CA(11)= Coef(1, 11) + Coef(2, 11) * hgt + Coef(3, 11)*hgt*hgt + 
     *          Coef(4, 11)*hgt*hgt*hgt + Coef(5, 11)*hgt*hgt*hgt*hgt


		fref = CA(1) + CA(2)*x + CA(3)*x*x + CA(4)*x*x*x +
     *     	 CA(5)*x*x*x*x + CA(6)*x*x*x*x*x + CA(7)*x*x*x*x*x*x +
     *     	 CA(8)*x*x*x*x*x*x*x + CA(9)*x*x*x*x*x*x*x*x + 
     *     	 CA(10)*x*x*x*x*x*x*x*x*x + CA(11)*x*x*x*x*x*x*x*x*x*x
	 
		RETURN
		END	   


        SUBROUTINE CASGEO(GN,GE,L1,L2)
        IMPLICIT DOUBLE PRECISION (A-H,K-M,O-Z)
        IMPLICIT INTEGER (I,J,N)
        G1=GN*1000.0
        G2=GE*1000.0
        R=57.2957795131
        B2=0.03246816
        F1=206264.806247096
        S1=126763.49
        S2=114242.75
        E4=.006803480836
        C1=3.25600414007E-2
        C2=2.55240717534E-9
        C3=.032338519783
        X1=1170251.56
        Y1=1126867.91
        Y2=G1
C       GN & GE
        X2=G2
        IF(X2.GT.700000.) GOTO 5
        X1=X1-1000000.
5       IF(Y2.GT.550000.) GOTO 10
        Y1=Y1-1000000.
10      X1=X2-X1
        Y1=Y2-Y1
        D1=Y1*B2/2.
        O1=S2+D1
        O2=O1+D1
        A3=O1/F1
        A4=O2/F1
        B3=1.-E4*DSIN(A3)**2
        B4=B3*DSQRT(B3)*C1
        C4=1.-E4*DSIN(A4)**2
        C5=DTAN(A4)*C2*C4**2
        C6=C5*X1**2
        D2=Y1*B4-C6
        C6=C6/3.
C       LAT
        L1=(S2+D2)/F1
        R3=O2-C6
        R4=R3-C6
        R2=R4/F1
        A2=1.-E4*DSIN(L1)**2
        L1=R*(L1)
        A5=DSQRT(A2)*C3
        D3=X1*A5/DCOS(R2)
C       LON
        L2=R*((S1+D3)/F1)
C       THIS IS THE EASTERN HEMISPHERE!
        L2=-L2
        RETURN
        END
