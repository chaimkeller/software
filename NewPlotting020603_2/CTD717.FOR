      PROGRAM CTD717A
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 3 2001
C     CTD717 makes a CALC format file from a CSV file
C     Pressure,Temp,Conductivity,Sound Velocity,Battery,Salinity,Density
      REAL*4 P(2048),T(2048),C(2048),SV(2048),B(2048),S(2048),D(2048)
      REAL*4 SVC(2048),PD,TD,CD,SVD,BD,SD,DD,ALPHA,DSV,TX,TN,SVX,SXN
      CHARACTER*1 KEY,ANS
      CALL HALOIN
      CALL PALETTE
      CALL SETSCREE(1)
      CALL DISPLAY(1)
    1 DMAX=0.0
      TX=0.0
      TN=100.0
      SVX=0.0
      SVN=2000.0
      ABSAV=0.0
      DIFFAV=0.0
      N=0
C     Write header, find out where the CSV data begins
      CALL CALCHDR(KC)
      PRINT 5
    5 FORMAT(' No   Dpres Temø Cond  Vmeas   Vcalc   Diff  Batt Sal',
     *       '    Dens   Absorpt',/,'   ')
C No   Dpres Temø Cond  Vmeas   Vcalc   Diff  Batt Sal    Dens   Absorpt
C+IIII XXX.X XX.X XXX.X XXXX.XX XXXX.XX X.XXX XX.X XX.XXX XXX.XX XX.XX
   10 IF(KC.EQ.19)READ(11,15,END=35) PD,TD,CD,SVD,BD,SD,DD
   15 FORMAT(19X,6(E15.7,1X),E15.7)
      IF(KC.EQ.20)READ(11,20,END=35) PD,TD,CD,SVD,BD,SD,DD
   20 FORMAT(20X,6(E15.7,1X),E15.7)
      IF(KC.EQ.21)READ(11,25,END=35) PD,TD,CD,SVD,BD,SD,DD
   25 FORMAT(21X,6(E15.7,1X),E15.7)
C     Decide if data is valuable
      IF(SVD.LT.1490.0.OR.SVD.GT.1600.0)GO TO 10
      IF(SD.LT.30.0.OR.SD.GT.45.0)GO TO 10
      IF(TD.LT.10.0.OR.TD.GT.45.0)GO TO 10
C     Don't keep measurements less than 20 cm apart
      IF(N.NE.0.AND.ABS(PD-P(N)).LT.0.2)GO TO 10
      IF(TD.GT.TX)TX=TD
      IF(TD.LT.TN)TN=TD
      IF(SVD.GT.SVX)SVX=SVD
      IF(SVD.LT.SVN)SVN=SVD
      N=N+1
      P(N)=PD
      T(N)=TD
      C(N)=CD
      SV(N)=SVD
      B(N)=BD
      S(N)=SD
      D(N)=DD
      IF(PD.GT.DMAX)NMAX=N
      IF(PD.GT.DMAX)DMAX=PD
      IF(N.GE.2048)GO TO 35
      CALL VELCALC(PD,TD,SD,SVC(N))
C     Include these in the Max/min in case the SV Probe is stuck
      IF(SVC(N).GT.SVX)SVX=SVC(N)
      IF(SVC(N).LT.SVN)SVN=SVC(N)
      DSV=SV(N)-SVC(N)
      DIFFAV=DIFFAV+DSV
      CALL ABSCOEF(DD,SVD,TD,SD,ALPHA)
      ABSAV=ABSAV+ALPHA
      PRINT 30, N,P(N),T(N),C(N),SV(N),SVC(N),DSV,B(N),S(N),D(N),ALPHA
   30 FORMAT('+',I4,F6.1,F5.1,F6.1,2F8.2,F6.3,F5.1,F7.3,F7.2,F6.2)
      GO TO 10
   35 CLOSE(UNIT=11)
      ABSAV=ABSAV/FLOAT(N)
      DIFFAV=DIFFAV/FLOAT(N)
      PRINT 40, DMAX,SVN,SVX,DIFFAV,TN,TX,ABSAV
   40 FORMAT(' Dmax=',F9.1,' m',/,
     *       ' SVmin=',F8.1,' m/s',/,
     *       ' SVmax=',F8.1,' m/s',/,
     *       ' Average SVmeas-SVcalc=',F5.2,' m/s',/,
     *       ' Tmin=',F9.1,'ø',/,
     *       ' Tmax=',F9.1,'ø',/,
     *       ' Average Absorption Coefficient=',F5.1,' dB/km')
      PRINT 45
   45 FORMAT(/////////////,
     *       ' Press Enter for Plot of Velocity and Temperature')
      READ 50, KEY
   50 FORMAT(A1)
      CALL PLOTDATA(T,P,SV,SVC,T,N,DMAX,SVX,SVN,TX,TN)
      PRINT 51
   51 FORMAT(' Use the Calculated Velocity Profile Y/N <N>? ')
      ICVP=0
      READ 50, ANS
      IF(ANS.EQ.CHAR(89).OR.ANS.EQ.CHAR(121))ICVP=1
C     Place the downgoing profile into the CALC file
      DO 75 I=1,NMAX
      IF(ICVP.EQ.1)SV(I)=SVC(I)
      IF(P(I).LT.10.0)NDIG=1
      IF(P(I).GE.10.0.AND.P(I).LT.100.0)NDIG=2
      IF(P(I).GE.100.0.AND.P(I).LT.1000.0)NDIG=3
      IF(P(I).GE.1000.0.AND.P(I).LT.10000.0)NDIG=4
      IF(NDIG.EQ.1)WRITE(12,55) P(I),SV(I),T(I)
   55 FORMAT(F4.2,F8.2,F7.3)
      IF(NDIG.EQ.2)WRITE(12,60) P(I),SV(I),T(I)
   60 FORMAT(F5.2,F8.2,F7.3)
      IF(NDIG.EQ.3)WRITE(12,65) P(I),SV(I),T(I)
   65 FORMAT(F6.2,F8.2,F7.3)
      IF(NDIG.EQ.4)WRITE(12,70) P(I),SV(I),T(I)
   70 FORMAT(F7.2,F8.2,F7.3)
   75 CONTINUE
C     Conclude with three zeroes separated by two spaces
      WRITE(12,80)
   80 FORMAT('  0  0  0')
      CLOSE(UNIT=12)
      PRINT 85, NMAX
   85 FORMAT(I6,' Points in the Sound Speed Profile',//,
     *    '       Press Any Number to Do Another, Enter to End: ')
      READ 50, ANS
      IF(ANS.NE.CHAR(32))GO TO 1
      CALL CLOSEGRA
      STOP
      END
      
      SUBROUTINE CALCHDR(KC)
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 4 2001
C     CALCHDR builds the header of the CALC file
      CHARACTER*8 DATE
      CHARACTER*20 DATETIME
      CHARACTER*30 CSVFILE,RELFILE
      CHARACTER*80 LINE
      INTEGER*4 JD
      PRINT 5
    5 FORMAT(' PROGRAM AML CTD-Plus CTD717',/,
     *       ' CSV to REL Conversion Program',/,
     *       ' Dr. John K. Hall - August 4 2001',/,
     *       ' Complete Name of .CSV File? ')
      READ 10, CSVFILE
   10 FORMAT(A30)
      DO 15 IP=1,30
      IC=ICHAR(CSVFILE(IP:IP))
      IF(IC.EQ.46.OR.IC.EQ.32)GO TO 20
   15 CONTINUE
   20 RELFILE=CSVFILE
      OPEN(UNIT=11,FILE=CSVFILE,STATUS='OLD')
      RELFILE(IP:IP+4)='.REL'
      OPEN(UNIT=12,FILE=RELFILE,STATUS='UNKNOWN')
      PRINT 25, CSVFILE,RELFILE
   25 FORMAT(' Input File is ',A30,/,
     *       ' Output File is ',A30)
Cast 1
CtdPlus 717
CTime,Pressure,Temperature,Conductivity,Sound Velocity,Battery,Salinity,Density,
C,dbars,Celsius,ms/cm,m/s,Volts,PSU,,
C7/1/01 11:00:46 AM,7.1195650E-0001,2.8308348E+0001,1.2333735E+0001,0.0000000E+0000,1.3740723E+0001,6.5761700E+0000,1.0662489E+0000
C     Skip the header
      DO 35 I=1,5
      READ(11,30) LINE
   30 FORMAT(A80)
   35 CONTINUE
C     Look for the location of the first space, M, and the first comma;
C      The first comma should be 19, 20, or 21
      KSP=0
      DO 40 IC=1,80
      IF(LINE(IC:IC).EQ.CHAR(32).AND.KSP.EQ.0)KSP=IC
      IF(LINE(IC:IC).EQ.CHAR(77))KM=IC
      IF(LINE(IC:IC).EQ.CHAR(44))GO TO 45
   40 CONTINUE
   45 KC=IC
      PRINT 50, KSP,KM,KC
   50 FORMAT(' Space at',I3,'; M at',I3,'; Comma at',I3)
      PAUSE
      BACKSPACE 11
C     Write the Header, but first extract the date and time from LINE
C     Find the first space, and the M
CALC,3208,21-06-97,-1,meters
CAML SOUND VELOCITY PROFILER S/N:3208
CDATE:85047 TIME:1204
CDEPTH OFFSET (M): 0.0
CDEPTH (M) VELOCITY (M/S) TEMP (C)
C7/1/01 11:00:46 AM,7.1195650E-0001,2.8308348E+0001,1.2333735E+0001,0.0000000E+0000,1.3740723E+0001,6.5761700E+0000,1.0662489E+0000
      DATE(1:KSP-1)=LINE(1:KSP-1)
      NCD=KSP-1
C     Change / to -
      DO 55 IC=1,KSP-1
      IF(DATE(IC:IC).EQ.CHAR(47))DATE(IC:IC)=CHAR(45)
   55 CONTINUE
      IF(NCD.EQ.8)WRITE(12,60) DATE(1:8)
   60 FORMAT('CALC,717,',A8,',-1,meters')
      IF(NCD.EQ.7)WRITE(12,65) DATE(1:7)
   65 FORMAT('CALC,717,',A7,',-1,meters')
      WRITE(12,70)
   70 FORMAT('AML SOUND VELOCITY PROFILER S/N:717')
C     Find the date ID IM IYR in nonliteral form
      OPEN(UNIT=13,STATUS='SCRATCH')
      WRITE(13,30) LINE
      BACKSPACE 13
C     Date as 0/0/00
      IF(LINE(2:2).EQ.CHAR(47).AND.LINE(4:4).EQ.CHAR(47))
     *  READ(13,75) ID,IM,IYR
   75 FORMAT(I1,1X,I1,1X,I2)
C     Date as 0/00/00
      IF(LINE(2:2).EQ.CHAR(47).AND.LINE(5:5).EQ.CHAR(47))
     *  READ(13,80) ID,IM,IYR
   80 FORMAT(I1,1X,I2,1X,I2)
C     Date as 00/00/00
      IF(LINE(3:3).EQ.CHAR(47).AND.LINE(6:6).EQ.CHAR(47))
     *  READ(13,85) ID,IM,IYR
   85 FORMAT(I2,1X,I2,1X,I2)
      IF(IYR.GT.80)IYR=IYR+1900
      IF(IYR.LT.20)IYR=IYR+2000
      CALL JULDAY(ID,IM,IYR,JD)
C     Find the time IHR IMIN ISEC in nonliteral form
      REWIND 13
      WRITE(13,90) (LINE(IC:IC),IC=KSP+1,80)
   90 FORMAT(80A1)
      BACKSPACE 13
C7/1/01 11:00:46 AM,7.1195650E-0001,2.8308348E+0001,1.2333735E+0001,0.0000000E+0000,1.3740723E+0001,6.5761700E+0000,1.0662489E+0000
      IF(KM-KSP.EQ.10)READ(13,95) IHR,IMIN,ISEC
   95 FORMAT(I1,1X,I2,1X,I2)
      IF(KM-KSP.EQ.11)READ(13,100) IHR,IMIN,ISEC
  100 FORMAT(I2,1X,I2,1X,I2)
C     Check if this is PM
      IF(LINE(KM-1:KM-1).EQ.CHAR(80))IHR=IHR+12
      IF(ISEC.GT.30)IMIN=IMIN+1
      IF(IMIN.LT.60)GO TO 105
      IMIN=IMIN-60
      IHR=IHR+1
  105 LHR=IHR*100+IMIN
      WRITE(DATETIME,110) JD,LHR
  110 FORMAT('DATE:',I5,' TIME:',I4)
C     Add a 0 if less than 10 hours
      IF(DATETIME(17:17).EQ.CHAR(32))DATETIME(17:17)=CHAR(48)
      IF(DATETIME(18:18).EQ.CHAR(32))DATETIME(18:18)=CHAR(48)
      WRITE(12,115) DATETIME
  115 FORMAT(A20,/,'DEPTH OFFSET (M): 0.0',/,
     *             'DEPTH (M) VELOCITY (M/S) TEMP (C)')
      RETURN
      END

      SUBROUTINE JULDAY(ID,IM,IYR,JD)
      INTEGER*2 ID,IM,IYR,KLY,MD(13)
      INTEGER*4 JD
      REAL*8 DJ
      DATA MD/0,31,59,90,120,151,181,212,243,273,304,334,365/
      KLY=0 
      DJ=DBLE(IYR-1900)*365.25 
      JD=JFIX(DJ+0.75) 
      IF(DJ-JD)10,5,10
    5 IF(IM.GE.3)KLY=1
   10 JD=JD+JFIX(MD(IM)+ID-1+KLY)
      RETURN 
      END

      SUBROUTINE VELCALC(D,TC,S,SV)
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 3 2001
C     VELCALC uses the ideas in the EM1002 Operation Manual pg. 383, altered to
C      work in the Red Sea, Mediterranean, and Persian Gulf.
      REAL*4 D,SV,TC,S,Z,Z2,Z3,T,T2,T3,SN,C0ST,C2ZST
      Z=0.001*D
      Z2=Z*Z
      T=0.1*TC
      T2=T*T
      T3=T2*T
      SN=S-35.0
C     Coppens, A. B., 1981, J. Acoust. Soc. Am., 69(3), pp. 862-863. Eqn.4.
      C0ST=1449.05+45.7*T-5.21*T2+0.23*T3
     *          +(1.333-0.126*T+0.009*T2)*SN
C     Coppens, A. B., 1981, J. Acoust. Soc. Am., 69(3), pp. 862-863. Eqn. 7.
      C2ZST=C0ST+(16.23+0.253*T)*Z+(0.213-0.1*T)*Z2
     *          +(0.016+0.0002*SN)*SN*T*Z
      SV=C2ZST
      RETURN
      END
     
      SUBROUTINE ABSCOEF(D,SV,T,S,ALPHA)
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - AUGUST 3 2001
C     ABSCOEF calculates the Absorption Coefficient using the equations
C      in the EM1002 Operation Manual, pg. 384.
C     pH can range from 7.6 to 8.2
      REAL*4 D,SV,T,S,ALPHA,PH,Z,SF,SFSQ,A1,A2,P2,P3,SF1,SF2,
     *       TERM1,TERM2,TERM3
      PH=7.9
      Z=D*0.001
      SF=95.0
      SFSQ=SF*SF
      A1=(8.86-10.0**(0.78*PH-5.0))/SV
      A2=(21.44*S*(1.0+0.025*T))/SV
      IF(T.LE.20.0)A3=4.937E-4-T*(2.59E-5-T*(9.11E-7-T*1.5E-8))
      IF(T.GE.20.0)A3=3.964E-4-T*(1.146E-5-T*(1.45E-7-T*6.5E-10))
      P2=1.0-Z*(0.137-0.0062*Z)
      P3=1.0-Z*(0.0383-Z*4.9E-4)
      SF1=2.8*SQRT(S*10.0**(4.0-1245.0/(T+273.0))/35.0)
      SF2=(8.17*10.0**(8.0-1990.0/(T+273.0)))/(1.0+0.0018*(S-35.0))
      TERM1=A1*SF1*SFSQ/(SFSQ+SF1*SF1)
      TERM2=A2*P2*SF2*SFSQ/(SFSQ+SF2*SF2)
      TERM3=A3*P3*SFSQ
      ALPHA=TERM1+TERM2+TERM3
      RETURN
      END
     
      SUBROUTINE PLOTDATA(T,D,SV,SVC,T,NP,DMAX,SVX,SVN,TX,TN)
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - JULY 13 2001
      REAL*4 T(2048),D(2048),SV(2048),SVC(2048)
      REAL*4 DMAX,SVX,SVN,TX,TN,DSV,DT
C     Plot out the profile 
      CALL SETSCREE(1)
      CALL DISPLAY(1)
      CALL SETCOLOR(1)
      CALL CLR
C     Plot out the Sound Speed (1450-1600 m/s) in Red over the profile      
      CALL SETCOLOR(4)
      DSV=(SVX-SVN)
      DO 10 I=1,NP
      IXP=50.0+500.0*(SV(I)-SVN)/DSV
      IYP=30.0+300.0*D(I)/DMAX
      IF(I.NE.1)GO TO 5
      CALL PTABS(IXP,IYP)
      GO TO 10
    5 CALL LNABS(IXP,IYP)
   10 CONTINUE
C     Plot out the Calculated Sound Speed (1450-1600 m/s) in Orange 
      CALL SETCOLOR(5)
      DO 13 I=1,NP
      IXP=50.0+500.0*(SVC(I)-SVN)/DSV
      IYP=30.0+300.0*D(I)/DMAX
      IF(I.NE.1)GO TO 12
      CALL PTABS(IXP,IYP)
      GO TO 13
   12 CALL LNABS(IXP,IYP)
   13 CONTINUE
C     Plot out the Temperature (20-30ø) in Yellow over the profile      
      CALL SETCOLOR(6)
      DT=(TX-TN)
      DO 20 I=1,NP
      IXP=50.0+500.0*(T(I)-TN)/DT
      IYP=30.0+300.0*D(I)/DMAX
      IF(I.NE.1)GO TO 15
      CALL PTABS(IXP,IYP)
      GO TO 20
   15 CALL LNABS(IXP,IYP)
   20 CONTINUE
      PRINT 25, DMAX,SVN,SVX,TX,TN
   25 FORMAT(' Dmax is',F7.1,
     *       ' m; SV (',F6.1,'-',F6.1,' m/s - Red/Orange); ',
     *       'T (',F4.1,'-',F4.1,'ø - Yellow)')
C' Dmax is XXXXX.Xm; SV ( XXX.X-XXX.X m/s - Red/Orange); T (XX.X-XX.Xø - Yellow)
      READ 30, KEY
   30 FORMAT(A1)
      RETURN
      END
     
      SUBROUTINE PALETTE
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - JANUARY 25 1988
      INTEGER*2 PAL(16)
      DATA PAL/0,56,7,4,36,38,54,2,18,51,27,11,9,1,5,63/
C     Change to EGAPAINT Palette
      DO 5 I=1,16
    5 CALL SETXPAL(I-1,PAL(I))
      RETURN
      END

      SUBROUTINE READGTCO(IX,IY,IPB)
C     JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - 9 JANUARY 1995
      INTEGER*2 PORT,BAUD,WORDLEN,PARITY,STOPBITS,RESULT,CHAR
      INTEGER*2 STAT,IX,IY,IPB,IXR,IYR
      DATA IB/0/
      IF(IB.NE.0)GO TO 5
      PORT=2
      BAUD=7
      WORDLEN=8
      PARITY=0
      STOPBITS=1
      RESULT=99
      CALL SERINIT(PORT,BAUD,WORDLEN,PARITY,STOPBITS,RESULT)
      IB=1
    5 CALL SERGTCO(PORT,IPB,IX,IY)
      IY=11000-IY
      RETURN
      END

      SUBROUTINE HALOIN
C     (Copr) JOHN K HALL - GEOLOGICAL SURVEY OF ISRAEL - 25 MARCH 1987
      REAL*4 X,Y
      INTEGER*2 IX,IY,IFUN,IERR,IPB
      CHARACTER*1 BEEP
      CHARACTER*12 GDFIL
      BEEP=CHAR(7)
      GDFIL='HALOIBME.DEV'
      CALL SETDEV(GDFIL)
      CALL INQERR(IFUN,IERR)
      IF(IERR.EQ.0)GO TO 10
      PRINT 5, BEEP,BEEP,BEEP
    5 FORMAT(' Graphics Device Failed to Initialize',3A1)
      stop
   10 CALL INITGR(4)
      CALL INQERR(IFUN,IERR)
      IF(IERR.EQ.0)GO TO 20
      PRINT 15, BEEP,BEEP,BEEP
   15 FORMAT(' HALO Environment failed to Initialize',3A1)
      stop
   20 CALL SETIEE(1)
      CALL INQERR(IFUN,IERR)
      IF(IERR.EQ.0)GO TO 30
      PRINT 25, BEEP,BEEP,BEEP
   25 FORMAT(' IEEE Standard for Floating Point Numbers Failed',3A1)
   30 CALL SETLAT(4)
      CALL INQERR(IFUN,IERR)
      IF(IERR.EQ.0)GO TO 50
      PRINT 45, BEEP,BEEP,BEEP
   45 FORMAT(' Locator not unscaled',3A1)
   50 CALL INITHCUR(4,4,4)
      CALL INQERR(IFUN,IERR)
      IF(IERR.EQ.0)GO TO 60
      PRINT 55, BEEP,BEEP,BEEP
   55 FORMAT(' Crosshair cursor not initialized',3A1)
   60 RETURN
      END

