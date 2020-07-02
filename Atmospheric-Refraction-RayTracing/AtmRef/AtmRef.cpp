// AtmRef.cpp : Defines the entry point for the DLL application.
//
#include "StdAfx.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

//global variables and constants

short CInt( double x );
double Fix( double x );

double fDLNTDH (double *H);
double fDTDH (double *H);
double fTEMP (double *H);
double fFNDPD1 (double *H);
double fFNDPD2 (double *H);
double fSTNDATM (double *H);
double fVAPOR (double *H);
double fGRAVRAT (double *H);
double fPRESSURE (double *H);
double fDVAPDT (double *H);
double fDNDH (double *H, double *fPRESSURE, double *fDTDH);
double fGUESSL (double *BETAM, double *Height);
double fGUESSP (double *BETA, double *Height);

double AD, AW, BD, BW;
double RELH;
double GRAVC, OBSLAT, DEG2RAD, HOBS;
double PDM1;
double PDM10;
double HCROSS;
double TGROUND;
double OLAT, S2, MAXIND;
double HL[50], TL[50], LRL[50]; //layer heights, temperatures, lapse rates //arrays used for other atmopsheres
double PRESSD1[99999], PRESSD2[99999];
double PI;
int NumLayers;

//=================LIST OF CONSTANTS===========================
double RBOLTZ = 8314.472;  //Univ. gas const. = Avogadro//s number x Boltzmann//s const.
double AMASSD = 28.964;  //Molar weight of dry air
double AMASSW = 18.016;  //Molar weight of water
double REARTH = 6356766.0;  //The Earth//s mean radius
double RE;
double RADCON; //Converts arcminutes into radians
double cd;
double HLIMIT = 100000; //Maximum height till which the rays are followed
double HMAXP1 = 30000; // //height where f.FNDPD2 (steps of 10 m) takes over from f.FNDPD1 (steps of 1 m)
double OneSixth = 1/6.0;
double OneHalf = 0.5;
double T7T4 = 1.125;
double OneDiv24 = 1/24.0;
short OPTVAP = 4;
//char FilePath[255] = "c:\\jk"; //root direcoty of rays' path file: to be passed to program

//NAngles is number of possible angle steps = 	(*BETALO - *BETAHI)/*BETAST  + 1


int __declspec (dllexport) __stdcall RayTracing(double *BETALO, double *BETAHI, double *BETAST, double *LastVA,int *NAngles,
											double *DistTo, double *VAwo, double *H21, double *Tolerance, short *FileMode,
											double *HOBSERVER, double *TEMPGROUND, double *HMAXT, LPCSTR pszFile_Path, short *StepSize,
											double *Press0, double *WAVELN,double *RELHUM, double *OBSLATITUDE, int *NSTEPS,
											bool *RecordTLoop, double *TSTART, double *TEND, long cbAddress) {

	//two file modes
	//FileMode = 0  //outputs details of the ray tracing
	//FileMode = 1  //calculates the view angle for a certain observer height = HOBS, distance to obstruction = DistTo, height of obstruction = H21
					//accurate up to a Tolerance amount of meters
	//if FileMode = 1 and RecordTLoop = true, then the terrestrial refrraction is written to a file
	//if FileMode = 0 and RecordTLoop = true, then the total atmospheric refraction is written to a file

	DWORD dwError, dwThreadPri;

	//set high priority

    if(!SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_HIGHEST))
    {
		char buff[255] = ""; 
        dwError = GetLastError();
	    dwThreadPri = GetThreadPriority(GetCurrentThread());
	    sprintf(buff, "%s\n%s%d", "Error in setting priority", "Current priority level: ", dwThreadPri);
		const int result = MessageBox(NULL, (const char *)buff, "Debugging",  MB_OKCANCEL);

		switch (result)
		{
		case IDOK:
			// continue looping
			break;
		case IDCANCEL:
			// exit
			goto LG5;  //skip rest of checks
			break;
		}
    }

LG5:

	  // Declare the function pointer, with one short integer argument.
     typedef void (__stdcall *FUNCPTR)(int parm);
     FUNCPTR vbFunc;

     // Point the function pointer at the passed-in address.
     vbFunc = (FUNCPTR)cbAddress;

	 HOBS = *HOBSERVER; //Height of observer
	 TGROUND = *TEMPGROUND; //temperature at surface of obserer
	 RELH = *RELHUM / 100.0; //relative humidity
	 OBSLAT = *OBSLATITUDE; //Observer's latitude

	//double HMAXT, TLOW, THIGH, Press0;
	//double RELHUM, BETALO, BETAHI, BETAST, WAVELN;
	int i,jstep; //NTLoop;
	double BETAM;
	//double TempStart, TempEnd, TempStep, TLoop;
	double DIST, REFRAC, AIRDRY, AIRVAP, PHI1;
	double BETA1;
	double FKP1, FKR1, FKB1, FKAD1, FKAV1;
	double PHINEW, RNEW, BETANEW, HNEW;
	double FKP2, FKR2, FKB2, FKAD2, FKAV2;
	double FKP3, FKR3, FKB3, FKAD3, FKAV3;
	double FKP4, FKR4, FKB4, FKAD4, FKAV4;
	double PHI2, R2, BETA2, DREFR;
	int ier;
	//double HgtStart, HgtEnd, HgtStep, HLoop;
	//int NHloop;
	//double fPRES; 
	double T, HSTEP;
	double fGRAVRATH,fVAPORH,fDVAPDTH,fDTDHH;
	//short StepSize = 10; //will be inputed -- how many steps to skip
	double H = 0;
	double TempStart;
	double TempEnd;
	short Attempts = 0;
	TempStart = *TSTART;
	TempEnd = *TEND;


	RE = REARTH;
	PI = 4 * atan(1.0);
	RADCON = PI / (60 * 180); //Converts arcminutes into radians
	cd = PI / 180; // //converts degrees into radians

	ier = 0;

	/////////////additions for FileMode = 1//////////////////////
	if (*FileMode == 1) {
		//In this mode, the dll calculates
		*BETALO = *VAwo/RADCON + 60.0; //search within a degree of the view angle without terestrial refraction
		*BETAST = 6.0; //step in 0.1 degrees
		*BETAHI = *VAwo/RADCON; 
	}
	//////////////////////////////////////////////////////////////////

	//===================OPTION-MENU OF FORMULA SATURATED VAPOR PRESSURE: =====================
	//OPTVAP=1: PL2, POWER LAW
	//OPTVAP=2: CC2, CLAUSIUS-CLAPEYRON 2 PAR.
	//OPTVAP=3: CC4, CLAUSIUS-CLAPEYRON 4 PAR
	//OPTVAP=4: ST, SACKUR-TETRODE, 4 PAR. //N.B., not mentioned in van der Werf's paper, but gives identical results to method 3
	//and is set as the default in Van der Werf's original code
	OPTVAP = 4;
	//============INITIALIZATION: THE US 1976 STANDARD ATMOSPHERE=======

	/*
	HOBS = 0;
	TGROUND = 283.15;
	HMAXT = 1000;
	TLOW = 0;
	THIGH = 400;
	Press0 = 1013;
	RELHUM = 0;
	BETALO = 125.0; //0;
	BETAHI = -125.0; //60;
	BETAST = 0.5; //10;
	WAVELN = 0.574;
	OBSLAT = 32; //52;
	NSTEPS = 500; //10000
	*/
	//ROBJ = 15;

	//====CALCULATE HCROSS ========
	//height of troposphere strastosphere boundary
	HCROSS = (TGROUND - 216.65) / 0.0065;

	//============LIST OF DERIVED CONSTANTS==================================
	OLAT = OBSLAT * 60 * RADCON;
	GRAVC = 9.780356 * (1 + 0.0052885 * (pow(sin(OLAT),2.0)) - 0.0000059 * (pow(sin(2 * OLAT), 2.0))); //           //Gravitat. const.
	BD = GRAVC * AMASSD / RBOLTZ; // //Dry air exponent
	BW = GRAVC * AMASSW / RBOLTZ; // //Water exponent
	S2 = 1 / pow(*WAVELN,2.0);
	//CIDDOR//S FORMULAS FOR DRY AIR AND WATER VAPOUR
	AD = 0.00000001 * (5792105.0 / (238.0185 - S2) + 167917.0 / (57.362 - S2)) * 288.15 / 1013.25;
	AW = 0.00000001022 * (295.235 + 2.6422 * S2 - 0.03238 * pow(S2,2.0) + 0.004028 * pow(S2,3.0)) * 293.15 / 13.33;
	MAXIND = *HMAXT + 1;

	////==================================================================
	////==FILL ARRAY PRESSD1 (PARTIAL PRESSURE OF DRY AIR)==
	////== IN STEPS OF 1 METER========
	double P1,P2,DP2DH,FK1,FK2,FK3,FK4,PDM01,I2LOW;
	H = 0.0;
	PRESSD1[1] = *Press0 - RELH * fVAPOR(&H);
	for( i = 1; i <= (31000 + 15); i++) {
		//===Fill PRESSD1. I=1 -> H=0, I=2 -> h=1 m. etc..===
		//===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
		HSTEP = 1.0;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		//STEP 1
		P1 = PRESSD1[i];
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK1 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 1

		//STEP 2
		H = (i - 1) / 1.0 + HSTEP * OneHalf;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD1[i] + FK1 * HSTEP * OneHalf;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK2 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 2

		//STEP 3
		H = (i - 1) / 1.0 + HSTEP * OneHalf;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD1[i] + FK2 * HSTEP * OneHalf;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK3 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 3

		//STEP 4
		H = (i - 1) / 1.0 + HSTEP;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD1[i] + FK3 * HSTEP;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK4 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 4
		PRESSD1[i + 1] = PRESSD1[i] + (HSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
	}
	//===FIND PDM1 AT -1 METER===
	HSTEP = -1.0;
	//STEP 1
	H = 0.0;
	P1 = PRESSD1[1];
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK1 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 1
	//STEP 2
	H = HSTEP * OneHalf;
	P1 = PRESSD1[1] + FK1 * HSTEP * OneHalf;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK2 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 2
	//STEP 3
	H = HSTEP * OneHalf;
	P1 = PRESSD1[1] + FK2 * HSTEP * OneHalf;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK3 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 3
	//STEP 4
	H = HSTEP;
	P1 = PRESSD1[1] + FK3 * HSTEP;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK4 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRATH / fTEMP(&H);
	//END STEP 4
	PDM01 = PRESSD1[1] + (HSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
	//==================================================================
	//=======END OF STORAGE PRESSURE ARRAY PRESSD1 FOR DRY AIR========

	//==FILL ARRAY PRESSD2 (PARTIAL PRESSURE OF DRY AIR)==
	//== IN STEPS OF 10 METER========
	PRESSD2[1] = PRESSD1[1];
	I2LOW = CInt(HMAXP1 / 10);
	for( i = 0; i <= I2LOW; i++) {
		PRESSD2[i + 1] = PRESSD1[10 * i + 1];
	}
	for (i = I2LOW; i <= (HLIMIT / 10 + 5); i++) {
		//===Fill PRESSD2. I=1 -> H=0, I=2 -> h=1 m. etc..===
		//===INTEGRATION BY 4TH ORDER RUNGE-KUTTA===
		HSTEP = 10.0;
		//STEP 1
		H = (i - 1) * 10;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD2[i];
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK1 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 1

		//STEP 2
		H = (i - 1) * 10 + HSTEP * OneHalf;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD2[i] + FK1 * HSTEP * OneHalf;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK2 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 2

		//STEP 3
		H = (i - 1) * 10 + HSTEP * OneHalf;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD2[i] + FK2 * HSTEP * OneHalf;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK3 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 3

		//STEP 4
		H = (i - 1) * 10 + HSTEP;
		T = fTEMP(&H);
		fGRAVRATH = fGRAVRAT(&H);
		fVAPORH = fVAPOR(&H);
		fDVAPDTH = fDVAPDT(&H);
		fDTDHH = fDTDH(&H);

		P1 = PRESSD2[i] + FK3 * HSTEP;
		P2 = RELH * fVAPORH;
		DP2DH = RELH * fDVAPDTH * fDTDHH;
		FK4 = -DP2DH - BD * P1 * fGRAVRATH / T - BW * P2 * fGRAVRATH / T;
		//END STEP 4
		PRESSD2[i + 1] = PRESSD2[i] + (HSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
	}
	//===FIND PDM10 AT -10 METER===
	HSTEP = -10.0;
	//STEP 1
	H = 0.0;
	P1 = PRESSD2[1];
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK1 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 1
	//STEP 2
	H = HSTEP * OneHalf;
	P1 = PRESSD2[1] + FK1 * HSTEP * OneHalf;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK2 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 2
	//STEP 3
	H = HSTEP * OneHalf;
	P1 = PRESSD2[1] + FK2 * HSTEP * OneHalf;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK3 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 3
	//STEP 4
	H = HSTEP;
	P1 = PRESSD2[1] + FK3 * HSTEP;
	P2 = RELH * fVAPOR(&H);
	DP2DH = RELH * fDVAPDT(&H) * fDTDH(&H);
	FK4 = -DP2DH - BD * P1 * fGRAVRAT(&H) / fTEMP(&H) - BW * P2 * fGRAVRAT(&H) / fTEMP(&H);
	//END STEP 4
	PDM10 = PRESSD2[1] + (HSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
	//==================================================================
	//=======END OF STORAGE PRESSURE ARRAY FOR DRY AIR========

	//===================CALCULATION====================================

	//////additions for FileMode = 1///////////////////
	double DIST0, Path0, SLOPE, FitVA, HOBST;
	double STARTANG,ENDANG;
	bool found = false;
	bool AdjustedOnce = false;
	double HeightTolerance;
	HeightTolerance = *Tolerance;
	HOBST = *H21;
		;
	/////////////////////////////////////////////////

	double DPATH,R1,H1,Path,PVAP,PATHLENGTH;
	double PD, PW, fDTDHB, fREFIND, fDNDHB, fRCINVB;
	double fPRESSUREH1, fDENSDRYH1,fDENSVAPH1,PDRY;
	double fPRESSURERNEW,fDENSDRYRNEW,fDENSVAPRNEW;
	double H2,TRUALT,fGUESSLHMAXT,fPRESSUREHNEW, GUESS;
	//double ALFA[2001],ALFT[2001];
	int nloop;
	short jstop;
	char filename[255] = "";  //terminate the filename of the path file
	char buff[1024] = ""; //string buffer for internal writes
	//bool testing = false;

	bool looping;

	FILE *stream;
	FILE *stream2;

	if (*FileMode == 0 && *RecordTLoop) {
		//recording the total refraction as function of temperature for any one height
		sprintf( filename, "%s%s%s%.0lf%s%.0lf%s%.0lf%s%.0lf%s%s", pszFile_Path, "\\", "TR_VDW_", TempStart, "-", TempEnd, "_", HOBS, "_", OBSLAT, ".dat", "\0" );

		if ( !(stream2 = fopen( filename, "a"))) 
		{
			return -1;
		}
	}

	else if (*FileMode == 1 && *RecordTLoop) {
		//recording the Terrestrial refraction as a function of Temperature
		sprintf( filename, "%s%s%s%.0lf%s%.0lf%s%.0lf%s%.0lf%s%.0lf%s%s", pszFile_Path, "\\", "TR_VA_", TempStart, "-", TempEnd, "_", HOBS, "_", HOBST, "_", OBSLAT, ".dat", "\0" );

		if ( !(stream2 = fopen( filename, "a"))) 
		{
			return -1;
		}

	}

	if (*FileMode == 0) //outputing ray tracing details
	{
		sprintf( filename, "%s%s%s%.0lf%s%.0lf%s%.0lf%s%s", pszFile_Path, "\\", "TR_VDW_", TGROUND, "_", HOBS, "_", OBSLAT, ".dat", "\0" );


		if ( !(stream = fopen( filename, "w" ) ) )
		{
			return -1; //can't open the file
		}

	}


STARTLOOP:
	
	GUESS = 0.0;

	fGUESSLHMAXT = fGUESSL(&GUESS, HMAXT);

	jstep = 0;

	if (*FileMode == 0) vbFunc(0);

	for (BETAM = *BETALO; BETAM >= *BETAHI; BETAM -= *BETAST) {
		DPATH = fGUESSP(&BETAM, HMAXT) / *NSTEPS;
		//LOCATE 3, 2: Print "Step size (m) = "; DPATH
		
		jstep++;
		//ALFA[jstep] = BETAM; //view angle without refraction

		//diagnostics
		/*
		if (BETAM == -0.5)
		{
			testing = true;
		}
		*/

		//
		DIST = 0.0;
		DIST0 = 0.0;
		Path0 = 0.0;
		REFRAC = 0.0;
		AIRDRY = 0.0;
		AIRVAP = 0.0;
		PHI1 = 0.0;
		PATHLENGTH = 0.0;
		BETA1 = BETAM * RADCON;
		R1 = REARTH + HOBS;
		H1 = HOBS;
		//PSet (DIST, H1), 10
		Path = -DPATH;
		//===============================
		//DO-LOOP OVER PATH
		looping = true; //set looping flag to only exit if meets criteria within the loop
		nloop = 0;
		jstop = 0;
		do {
			Path = Path + DPATH;
			//
			// FOURTH ORDER RUNGE-KUTTA
			// THE THREE COUPLED FIRST ORDER DIFFERENTIAL EQUATIONS ARE:
			// 1)  dPHI/dPATH=cos(BETA)/R
			// 2)  dR/dPATH = sin(BETA)
			// 3)  dBETA/dPATH = cos(BETA)[1//R+(1/n) dn/dR]
			// WITH (1/n) [dn/dR]=fRCINV(H)
			//
			// STEP 1
			// FIND THE RUNGE-KUTTA K-COEFFICIENTS
			//

			/////////////inline fRCINV////////////////////
 

			fPRESSUREH1 = fPRESSURE(&H1);
			fDTDHB = fDTDH(&H1);
			fDNDHB = fDNDH (&H1, &fPRESSUREH1, &fDTDHB);
			T = fTEMP(&H1);
			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			if ( H1 < HMAXP1) {
				PD = fFNDPD1(&H1);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}else{
				PD = fFNDPD2(&H1);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}

			fRCINVB = fDNDHB / fREFIND;

			PVAP = RELH * fVAPORH;

			if ( H1 < HMAXP1 ) {
				fPRESSUREH1 = fFNDPD1(&H1) + RELH * fVAPORH;
			}else{
				fPRESSUREH1 = fFNDPD2(&H1) + RELH * fVAPORH;
			}

			PDRY = fPRESSUREH1 - PVAP;
			fDENSDRYH1 = (AMASSD / RBOLTZ) * PDRY / T;

			PVAP = RELH * fVAPORH;
			fDENSVAPH1 = (AMASSW / RBOLTZ) * PVAP / T;


			//////////////////////////////////////////////


			FKP1 = cos(BETA1) / R1;
			FKR1 = sin(BETA1);
			FKB1 = cos(BETA1) * (1 / R1 + fRCINVB);
			FKAD1 = fDENSDRYH1;
			FKAV1 = fDENSVAPH1;
			//
			//END OF FIRST STEP
			//
			//STEP 2
			PHINEW = PHI1 + FKP1 * DPATH * 0.5;
			RNEW = R1 + FKR1 * DPATH * 0.5;
			BETANEW = BETA1 + FKB1 * DPATH * 0.5;
			HNEW = RNEW - REARTH; //elevation halfway step

			/////////////inline fRCINV////////////////////

			fPRESSUREHNEW = fPRESSURE(&HNEW);
			fDTDHB = fDTDH(&HNEW);
			fDNDHB = fDNDH (&HNEW, &fPRESSUREHNEW, &fDTDHB);
			T = fTEMP(&HNEW);
			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			if ( HNEW < HMAXP1) {
				PD = fFNDPD1(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}else{
				PD = fFNDPD2(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}

			fRCINVB = fDNDHB / fREFIND;

			PVAP = RELH * fVAPORH;

			if ( HNEW < HMAXP1 ) {
				fPRESSURERNEW = fFNDPD1(&HNEW) + RELH * fVAPORH;
			}else{
				fPRESSURERNEW = fFNDPD2(&HNEW) + RELH * fVAPORH;
			}

			PDRY = fPRESSURERNEW - PVAP;
			fDENSDRYRNEW = (AMASSD / RBOLTZ) * PDRY / T;

			PVAP = RELH * fVAPORH;
			fDENSVAPRNEW = (AMASSW / RBOLTZ) * PVAP / T;


			//////////////////////////////////////////////

			//
			// FIND THE RUNGE-KUTTA K-COEFFICIENTS
			//
			FKP2 = cos(BETANEW) / RNEW;
			FKR2 = sin(BETANEW);
			FKB2 = cos(BETANEW) * (1 / RNEW + fRCINVB);
			FKAD2 = fDENSDRYRNEW; //<<<<<<<<
			FKAV2 = fDENSVAPRNEW;
			//
			//END OF SECOND STEP
			//
			//STEP 3
			PHINEW = PHI1 + FKP2 * DPATH * 0.5;
			RNEW = R1 + FKR2 * DPATH * 0.5;
			BETANEW = BETA1 + FKB2 * DPATH * 0.5;
			HNEW = RNEW - REARTH; //elevation halfway step
			//
			// FIND THE RUNGE-KUTTA K-COEFFICIENTS
			//

			/////////////inline fRCINV////////////////////
			fPRESSUREHNEW = fPRESSURE(&HNEW);
			fDTDHB = fDTDH(&HNEW);
			fDNDHB = fDNDH (&HNEW, &fPRESSUREHNEW, &fDTDHB);
			T = fTEMP(&HNEW);
			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			if ( HNEW < HMAXP1) {
				PD = fFNDPD1(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}else{
				PD = fFNDPD2(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}

			fRCINVB = fDNDHB / fREFIND;

			PVAP = RELH * fVAPORH;

			if ( HNEW < HMAXP1 ) {
				fPRESSURERNEW = fFNDPD1(&HNEW) + RELH * fVAPORH;
			}else{
				fPRESSURERNEW = fFNDPD2(&HNEW) + RELH * fVAPORH;
			}

			PDRY = fPRESSURERNEW - PVAP;
			fDENSDRYRNEW = (AMASSD / RBOLTZ) * PDRY / T;

			PVAP = RELH * fVAPORH;
			fDENSVAPRNEW = (AMASSW / RBOLTZ) * PVAP / T;
			//////////////////////////////////////////////

			FKP3 = cos(BETANEW) / RNEW;
			FKR3 = sin(BETANEW);
			FKB3 = cos(BETANEW) * (1 / RNEW + fRCINVB);
			FKAD3 = fDENSDRYRNEW;
			FKAV3 = fDENSVAPRNEW;

			//
			//END OF THIRD STEP
			//
			//STEP 4
			PHINEW = PHI1 + FKP3 * DPATH;
			RNEW = R1 + FKR3 * DPATH;
			BETANEW = BETA1 + FKB3 * DPATH;
			HNEW = RNEW - REARTH;
			H = HNEW; //elevation at full step
			//
			// FIND THE RUNGE-KUTTA K-COEFFICIENTS
			//
			/////////////inline fRCINV////////////////////
			fPRESSUREHNEW = fPRESSURE(&HNEW);
			fDTDHB = fDTDH(&HNEW);
			fDNDHB = fDNDH (&HNEW, &fPRESSUREHNEW, &fDTDHB);
			T = fTEMP(&HNEW);
			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			if ( HNEW < HMAXP1) {
				PD = fFNDPD1(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}else{
				PD = fFNDPD2(&HNEW);
				PW = RELH * fVAPORH;
				fREFIND = 1 + (AD * PD + AW * PW) / T;
			}

			fRCINVB = fDNDHB / fREFIND;

			PVAP = RELH * fVAPORH;

			if ( HNEW < HMAXP1 ) {
				fPRESSURERNEW = fFNDPD1(&HNEW) + RELH * fVAPORH;
			}else{
				fPRESSURERNEW = fFNDPD2(&HNEW) + RELH * fVAPORH;
			}

			PDRY = fPRESSURERNEW - PVAP;
			fDENSDRYRNEW = (AMASSD / RBOLTZ) * PDRY / T;

			PVAP = RELH * fVAPORH;
			fDENSVAPRNEW = (AMASSW / RBOLTZ) * PVAP / T;
			//////////////////////////////////////////////

			FKP4 = cos(BETANEW) / RNEW;
			FKR4 = sin(BETANEW);
			FKB4 = cos(BETANEW) * (1 / RNEW + fRCINVB);
			//FREF4 = cos(BETANEW) * fRCINV(HNEW);
			FKAD4 = fDENSDRYRNEW;
			FKAV4 = fDENSVAPRNEW;
			//
			//END OF FOURTH AND FINAL STEP
			//
			//FIND R2 AND PHI2
			PHI2 = PHI1 + (FKP1 + 2 * FKP2 + 2 * FKP3 + FKP4) * DPATH * OneSixth;
			R2 = R1 + (FKR1 + 2 * FKR2 + 2 * FKR3 + FKR4) * DPATH * OneSixth;
			BETA2 = BETA1 + (FKB1 + 2 * FKB2 + 2 * FKB3 + FKB4) * DPATH * OneSixth;
			AIRDRY = AIRDRY + (FKAD1 + 2 * FKAD2 + 2 * FKAD3 + FKAD4) * DPATH * OneSixth;
			AIRVAP = AIRVAP + (FKAV1 + 2 * FKAV2 + 2 * FKAV3 + FKAV4) * DPATH * OneSixth;
			H2 = R2 - REARTH;
			DREFR = -BETA2 + BETA1 + PHI2 - PHI1;

			DIST += REARTH * (PHI2 - PHI1);

			//use Lehn's parabolic path approx to ray trajectory and Brutton equation 58
			//if (H2 > 0) {
			//	PATHLENGTH += sqrt(pow(REARTH * (PHI2 - PHI1),2) + pow(H2 - H1 - 0.5 * pow(REARTH * (PHI2 - PHI1),2)/REARTH,2));
			//}
			PATHLENGTH = Path;

		
			//Stop this ray if it hits the ground, or if it seems to never end
			//as may occur for a Novaya-Zemlya atmosphere.

			if (nloop == 0  && *FileMode == 0) {
				buff[0] = 0;
			   TRUALT = BETAM - REFRAC;	
			   sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", 0.0, 0.0, HOBS, BETAM, BETA2 * 1000.0, REFRAC * (1000.0 * RADCON), "\n\0");
			   fprintf (stream, buff);
			}

			if ( H2 < 0 || (DIST > 10 * fGUESSLHMAXT && H2 < *HMAXT) ) {
				looping = false;
				jstop = jstep;
				//ALFT[jstep] = -1000;
				TRUALT = BETAM - REFRAC;

				if (*FileMode == 0) {

					buff[0] = 0;

					if (!(_isnan(PHI2))) {
						sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", DIST, PATHLENGTH, -1000.0, BETAM, BETA2 * 1000.0, REFRAC * (1000.0 * RADCON), "\n\0");
					}else{
						sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", 0.0, PATHLENGTH, -1000.0, BETAM, BETA2 * 1000.0, REFRAC, "\n\0");
					}
					fprintf(stream, buff);
				}

				break; //flag end of ray tracing by jstop = -1
			}else{
				if ( H2 > HLIMIT) {
					// H2 PASSED HLIMIT METER
					TRUALT = BETAM - REFRAC;
					looping = false;
					//ALFT[jstep] = TRUALT; //View angle with refraction

					if (BETAM == 0.0 && *RecordTLoop && *FileMode == 0) {
						buff[0] = 0;
					   sprintf(buff, "%13.5lf%13.5lf%s", TGROUND, REFRAC * (1000.0 * RADCON), "\n\0");
					   fprintf(stream2, buff);
					}

					jstop = -1;
					break; 
				}
			} 

			if (*DistTo >= DIST0 && *DistTo < DIST && *FileMode == 1) {
				//if (HOBST >= H1 && HOBST < H2) {
				if ((HOBST >= H1 && HOBST < H2) || (HOBST <= H1 && HOBST > H2)) {

					//determine if height is within tolerance
					if (fabs(HOBST - H1) < HeightTolerance && fabs(HOBST - H2) < HeightTolerance) {
						SLOPE = -*BETAST/(H2 - H1);
						FitVA = (SLOPE * (HOBST - H1) + BETAM + *BETAST) * RADCON; //interpolate and convert to radians from arc minutes

						if (FitVA <= *LastVA && fabs(*DistTo - DIST) <= 50) { //throw away spurious results

							if (*RecordTLoop) {
								buff[0] = 0;
								sprintf(buff, "%13.1lf,%13.4lf,%13.2lf,%13.9lf,%13.6lf,%13.6lf,%13.6lf%s", TGROUND, DIST-*DistTo, PATHLENGTH, BETA2/60.0, FitVA/cd, *VAwo/cd, (FitVA - *VAwo)/cd, "\n\0" );
								fprintf(stream2, buff);
								//*VAwo = FitVA; //output the fitted view angle in radians
								//printf("%s: %15.4f\n", "Terrestrial refraction (deg.)", (FitVA - VAwo)/cd);
							}

							*LastVA = FitVA;
							found = true;
							break;
						}
					}
				}else{
					if (HOBST > H1 && HOBST > H2 && !AdjustedOnce) {
						//refine the search
						AdjustedOnce = true;
						STARTANG = BETAM + 2.0 * *BETAST;
						ENDANG = BETAM - 2.0 * *BETAST;
					}
				}
			}

			REFRAC = REFRAC + DREFR / RADCON;

			if ( (nloop + 1) % *StepSize == 0  && *FileMode == 0) {
               TRUALT = BETAM - REFRAC; 
			   buff[0] = 0;
			   sprintf(buff, "%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf,%13.5lf%s", DIST, PATHLENGTH, H2, BETAM, BETA2 * 1000.0, REFRAC * (1000.0 * RADCON), "\n\0");
			   fprintf(stream, buff);
			}

			PHI1 = PHI2;
			R1 = R2;
			H1 = H2;
			BETA1 = BETA2;
			DIST0 = DIST;
			Path0 = Path;
			if ( H2 > *HMAXT ) {
				DPATH = fGUESSP(&BETAM, &H2) / *NSTEPS;
			}

			nloop++;

			//END DO-LOOP OVER PATH
		}while(looping);

		if (*FileMode == 0) vbFunc((long)floor(jstep * 100/ *NAngles));

		if (AdjustedOnce) break;

		if (found) break;

		if (jstop != -1) break; //exit the angle loop since ray ray is in outer space

	} 

	if (*FileMode == 0) fclose(stream);


	if (!found && *FileMode == 1) {

		*BETAST = *BETAST * 0.5;

		if (AdjustedOnce) {

			AdjustedOnce = false;
			//refine the search

			*BETALO = STARTANG;
			*BETAHI = ENDANG;

			//if repeats n times, exit with error code
			//signifying no convergence, i.e., TR=0 to the accuracy of this search
			Attempts++;
			if (Attempts > 20) {
				ier = -1;
				return ier;
			}

		}

		goto STARTLOOP;
	}

	if (*RecordTLoop) fclose(stream2);

   return ier; 
}

///////////////////fDLNTDH///////////////////////
double fDLNTDH (double *H00) {
//The deriative of ln(T): (1/T)(dT/dh)
/////////////////////////////////////////////////

	double ftempH,DH,T1,T2,T3,T4,fDTDH;
	double H0, H;

	H0 = *H00;
	H = H0;

	if ( H < HCROSS ) {
		ftempH = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			ftempH = 216.65;
		}else{
			if (H < 32000.0) {
				ftempH = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					ftempH = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						ftempH = 270.65;
					}else{
						if ( H < 71000.0 ){
							ftempH = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								ftempH = 214.65 - 0.002 * (H - 71000.0);
							}else{
								ftempH = 186.65;
							}
						}
					}
				}
			}
		}
	}

	DH = 0.01;
	H = H0 - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;

	return fDTDH / ftempH;

}


///////////////////////fDTDH/////////////////////
double fDTDH (double *H00) {
//////////////////////////////////////////////////

	double DH, T1, T2, T3, T4, H0, H;

	H0 = *H00;
	H = H0;

	DH = 0.01;
	H = H0 - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = H0 + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	return ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;

}

//////////////////fTEMP/////////////////////////
double fTEMP (double *H0) {
///////////////////////////////////////////////

	double H;

	H = *H0;

	if ( H < HCROSS ) {
		return 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			return 216.65;
		}else{
			if (H < 32000.0) {
				return 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					return 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						return 270.65;
					}else{
						if ( H < 71000.0 ){
							return 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								return 214.65 - 0.002 * (H - 71000.0);
							}else{
								return 186.65;
							}
						}
					}
				}
			}
		}
	}

}

/////////////////////fFNDPD1//////////////////////////
double fFNDPD1 (double *H0) {
//////////////////////////////////////////////////////////
	//Interpolation in the array PRESSD1
//DefDbl A-H, O-Z
//SHARED HLIMIT, PDM1, RELH, BD, BW, HMAXP1//

	double T, P1, P2, Y, YSTEP, DH, DP2DY, fVAPORH, HV, fDVAPDT;
	double T1, T2, T3, T4, fGRAVRAT;
	double FK1,FK2,FK3,FK4,fDTDH, H, H00;
	int i;

	H00 = *H0;

	if (H00 < 0 ) H00 = 0;

	if ((H00 < (HMAXP1 + 1) && H00 >= 0)) {
		Y = H00;
		YSTEP = Y - Fix(Y);
		i = Fix(Y);

		////////////////////////////////STEP 1/////////////////////////////////////////////////
		P1 = PRESSD1[i + 1];


		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		default:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}

		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		default:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}

		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK1 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		////////////////////////END STEP 1///////////////////////////////////////

		///////////////////////////////////////STEP 2////////////////////////////////////////
		Y = i + YSTEP * OneHalf;
		P1 = PRESSD1[i + 1] + FK1 * YSTEP * OneHalf;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		default:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		default:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK2 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 2////////////////////////////////////

		///////////////////////////////////////STEP 3////////////////////////////////////////
		Y = i + YSTEP * OneHalf;
		P1 = PRESSD1[i + 1] + FK2 * YSTEP * OneHalf;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		default:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		default:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK3 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 3////////////////////////////////////

		////////////////////////////////////////STEP 4///////////////////////////////////////
		Y = i + YSTEP;
		P1 = PRESSD1[i + 1] + FK3 * YSTEP;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		default:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		default:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK4 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 4//////////////////////////////////

		return PRESSD1[i + 1] + (YSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
	}else{
		return 0.0;
	}
}


////////////////////////fFNDPD2//////////////////////////////
double fFNDPD2 (double *H0) {
////////////////////////////////////////////////////////////////
	//Interpolation in the array PRESSD2
//DefDbl A-H, O-Z
//SHARED HLIMIT, PDM1, RELH, BD, BW, HMAXP1//

	double T, P1, P2, Y, YSTEP, H, DH, DP2DY, fVAPORH, HV, fDVAPDT;
	double T1, T2, T3, T4, fGRAVRAT;
	double FK1,FK2,FK3,FK4,fDTDH;
	int i;


	Y = *H0;
	YSTEP = (Y - Fix(Y)) * 10;
	i = Fix(Y/10);


	////////////////////////////////STEP 1/////////////////////////////////////////////////
	P1 = PRESSD2[i + 1];


	H = Y;
	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}

	P2 = RELH * fVAPORH;

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		fDVAPDT = HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}

	DH = 0.01;
	H = Y - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
	fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

	DP2DY = RELH * fDVAPDT * fDTDH;
	FK1 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
	////////////////////////END STEP 1///////////////////////////////////////

	///////////////////////////////////////STEP 2////////////////////////////////////////
	Y = i + YSTEP * OneHalf;
	P1 = PRESSD2[i + 1] + FK1 * YSTEP * OneHalf;

	H = Y;
	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}
	P2 = RELH * fVAPORH;

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		fDVAPDT = HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}
	DH = 0.01;
	H = Y - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
	fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

	DP2DY = RELH * fDVAPDT * fDTDH;
	FK2 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
	//////////////////////////////////////END STEP 2////////////////////////////////////

	///////////////////////////////////////STEP 3////////////////////////////////////////
	Y = i + YSTEP * OneHalf;
	P1 = PRESSD2[i + 1] + FK2 * YSTEP * OneHalf;

	H = Y;
	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}
	P2 = RELH * fVAPORH;

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		fDVAPDT = HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}
	DH = 0.01;
	H = Y - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
	fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

	DP2DY = RELH * fDVAPDT * fDTDH;
	FK3 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
	//////////////////////////////////////END STEP 3////////////////////////////////////

	////////////////////////////////////////STEP 4///////////////////////////////////////
	Y = i + YSTEP;
	P1 = PRESSD2[i + 1] + FK3 * YSTEP;

	H = Y;
	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}
	P2 = RELH * fVAPORH;

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		fDVAPDT = HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}
	DH = 0.01;
	H = Y - 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T1 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T1 = 216.65;
		}else{
			if (H < 32000.0) {
				T1 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T1 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T1 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T1 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T1 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T1 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y - DH * OneHalf;
	if ( H < HCROSS ) {
		T2 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T2 = 216.65;
		}else{
			if (H < 32000.0) {
				T2 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T2 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T2 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T2 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T2 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T2 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + DH * OneHalf;
	if ( H < HCROSS ) {
		T3 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T3 = 216.65;
		}else{
			if (H < 32000.0) {
				T3 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T3 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T3 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T3 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T3 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T3 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	H = Y + 3 * DH * OneHalf;
	if ( H < HCROSS ) {
		T4 = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T4 = 216.65;
		}else{
			if (H < 32000.0) {
				T4 = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T4 = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T4 = 270.65;
					}else{
						if ( H < 71000.0 ){
							T4 = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T4 = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T4 = 186.65;
							}
						}
					}
				}
			}
		}
	}

	fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
	fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

	DP2DY = RELH * fDVAPDT * fDTDH;
	FK4 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
	//////////////////////////////////////END STEP 4//////////////////////////////////

	return PRESSD2[i + 1] + (YSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
}

////////////////////fSTNDATM/////////////////////////////
double fSTNDATM (double H) {
//STANDARD MUSA76 ATMOSPHERE WITH TROPOSPHERE AT HCROSS = (TGROUND-216.65)/0.0065
////////////////////////////////////////////////////////////////////////////////
	//SHARED HCROSS, TGROUND
	if ( H < HCROSS ) {
		return 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0 ) {
			return 216.65;
		}else{
			if (H < 32000.0 ) {
				return 216.65 + 0.001 * (H - 20000.0);
			}else{
				if (H < 47000.0 ) {
					return 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ) {
						return 270.65;
					}else{
						if ( H < 71000.0 ) {
							return 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ) {
								return 214.65 - 0.002 * (H - 71000.0);
							}else{
								return 186.65;
							}
						}
					}
				}
			}
		}
	}
}

/////////////fVAPOR///////////////////
double fVAPOR (double *H0) {
/////////////////////////////////////////
	double T, H;

	H = *H0;

	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		return pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		return exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		return exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		return pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	default:
		return pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}

}

////////////////////////////fGRAVRAT////////////////////
double fGRAVRAT (double *H) {
/////////////////////////////////////////////////////

   //gravitation at H/ gravitation at H=0
   return pow((REARTH / (REARTH + *H)),2.0);
}

///////////////////fPRESSURE//////////////////////////
double fPRESSURE (double *H00) {
//////////////////////////////////////////////////////////
	//DefDbl A-H, O-Z
	//SHARED HLIMIT, RELH, HMAXP1

	double T, P1, P2, Y, YSTEP, DH, DP2DY, fVAPORH, HV, fDVAPDT;
	double T1, T2, T3, T4, fGRAVRAT,fDTDH;
	double FK1,FK2,FK3,FK4;
	int i;
	double fFNDPD1, fFNDPD2,H, H0;

	H = *H00;

	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}

	if ( H < HMAXP1 ) {
		////////////fFNDPD1//////////////////////

	    H0 = H;

		if (H0 < 0) H0 = 0.0;

		if ((H0 < (HMAXP1 + 1) && H0 >= 0)) {
			Y = H0;
			YSTEP = Y - Fix(Y);
			i = Fix(Y);

			////////////////////////////////STEP 1/////////////////////////////////////////////////
			P1 = PRESSD1[i + 1];


			H = Y;
			if ( H < HCROSS ) {
				T = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T = 216.65;
				}else{
					if (H < 32000.0) {
						T = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T = 270.65;
							}else{
								if ( H < 71000.0 ){
									T = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T = 186.65;
									}
								}
							}
						}
					}
				}
			}

			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}

			P2 = RELH * fVAPORH;

			switch (OPTVAP) {
			case 1:
				HV = pow((T / 247.1), 18.36); //PL2 FORM
				fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
				break;
			case 2:
				HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
				fDVAPDT = HV * 5349 / (T * T);
				break;
			case 3:
				HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
				fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
				break;
			case 4:
				HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
				break;
			}

			DH = 0.01;
			H = Y - 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T1 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T1 = 216.65;
				}else{
					if (H < 32000.0) {
						T1 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T1 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T1 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T1 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T1 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T1 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y - DH * OneHalf;
			if ( H < HCROSS ) {
				T2 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T2 = 216.65;
				}else{
					if (H < 32000.0) {
						T2 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T2 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T2 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T2 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T2 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T2 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + DH * OneHalf;
			if ( H < HCROSS ) {
				T3 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T3 = 216.65;
				}else{
					if (H < 32000.0) {
						T3 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T3 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T3 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T3 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T3 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T3 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T4 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T4 = 216.65;
				}else{
					if (H < 32000.0) {
						T4 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T4 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T4 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T4 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T4 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T4 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
			fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

			DP2DY = RELH * fDVAPDT * fDTDH;
			FK1 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
			////////////////////////END STEP 1///////////////////////////////////////

			///////////////////////////////////////STEP 2////////////////////////////////////////
			Y = i + YSTEP * OneHalf;
			P1 = PRESSD1[i + 1] + FK1 * YSTEP * OneHalf;

			H = Y;
			if ( H < HCROSS ) {
				T = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T = 216.65;
				}else{
					if (H < 32000.0) {
						T = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T = 270.65;
							}else{
								if ( H < 71000.0 ){
									T = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T = 186.65;
									}
								}
							}
						}
					}
				}
			}

			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			P2 = RELH * fVAPORH;

			switch (OPTVAP) {
			case 1:
				HV = pow((T / 247.1), 18.36); //PL2 FORM
				fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
				break;
			case 2:
				HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
				fDVAPDT = HV * 5349 / (T * T);
				break;
			case 3:
				HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
				fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
				break;
			case 4:
				HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
				break;
			}
			DH = 0.01;
			H = Y - 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T1 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T1 = 216.65;
				}else{
					if (H < 32000.0) {
						T1 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T1 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T1 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T1 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T1 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T1 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y - DH * OneHalf;
			if ( H < HCROSS ) {
				T2 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T2 = 216.65;
				}else{
					if (H < 32000.0) {
						T2 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T2 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T2 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T2 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T2 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T2 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + DH * OneHalf;
			if ( H < HCROSS ) {
				T3 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T3 = 216.65;
				}else{
					if (H < 32000.0) {
						T3 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T3 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T3 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T3 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T3 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T3 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T4 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T4 = 216.65;
				}else{
					if (H < 32000.0) {
						T4 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T4 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T4 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T4 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T4 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T4 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
			fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

			DP2DY = RELH * fDVAPDT * fDTDH;
			FK2 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
			//////////////////////////////////////END STEP 2////////////////////////////////////

			///////////////////////////////////////STEP 3////////////////////////////////////////
			Y = i + YSTEP * OneHalf;
			P1 = PRESSD1[i + 1] + FK2 * YSTEP * OneHalf;

			H = Y;
			if ( H < HCROSS ) {
				T = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T = 216.65;
				}else{
					if (H < 32000.0) {
						T = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T = 270.65;
							}else{
								if ( H < 71000.0 ){
									T = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T = 186.65;
									}
								}
							}
						}
					}
				}
			}

			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			P2 = RELH * fVAPORH;

			switch (OPTVAP) {
			case 1:
				HV = pow((T / 247.1), 18.36); //PL2 FORM
				fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
				break;
			case 2:
				HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
				fDVAPDT = HV * 5349 / (T * T);
				break;
			case 3:
				HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
				fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
				break;
			case 4:
				HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
				break;
			}
			DH = 0.01;
			H = Y - 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T1 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T1 = 216.65;
				}else{
					if (H < 32000.0) {
						T1 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T1 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T1 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T1 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T1 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T1 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y - DH * OneHalf;
			if ( H < HCROSS ) {
				T2 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T2 = 216.65;
				}else{
					if (H < 32000.0) {
						T2 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T2 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T2 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T2 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T2 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T2 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + DH * OneHalf;
			if ( H < HCROSS ) {
				T3 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T3 = 216.65;
				}else{
					if (H < 32000.0) {
						T3 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T3 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T3 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T3 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T3 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T3 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T4 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T4 = 216.65;
				}else{
					if (H < 32000.0) {
						T4 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T4 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T4 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T4 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T4 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T4 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
			fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

			DP2DY = RELH * fDVAPDT * fDTDH;
			FK3 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
			//////////////////////////////////////END STEP 3////////////////////////////////////

			////////////////////////////////////////STEP 4///////////////////////////////////////
			Y = i + YSTEP;
			P1 = PRESSD1[i + 1] + FK3 * YSTEP;

			H = Y;
			if ( H < HCROSS ) {
				T = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T = 216.65;
				}else{
					if (H < 32000.0) {
						T = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T = 270.65;
							}else{
								if ( H < 71000.0 ){
									T = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T = 186.65;
									}
								}
							}
						}
					}
				}
			}

			switch (OPTVAP) {
			case 1:
				fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
				break;
			case 2:
				fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
				break;
			case 3:
				fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
				break;
			case 4:
				fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				break;
			}
			P2 = RELH * fVAPORH;

			switch (OPTVAP) {
			case 1:
				HV = pow((T / 247.1), 18.36); //PL2 FORM
				fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
				break;
			case 2:
				HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
				fDVAPDT = HV * 5349 / (T * T);
				break;
			case 3:
				HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
				fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
				break;
			case 4:
				HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
				fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
				break;
			}
			DH = 0.01;
			H = Y - 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T1 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T1 = 216.65;
				}else{
					if (H < 32000.0) {
						T1 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T1 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T1 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T1 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T1 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T1 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y - DH * OneHalf;
			if ( H < HCROSS ) {
				T2 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T2 = 216.65;
				}else{
					if (H < 32000.0) {
						T2 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T2 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T2 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T2 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T2 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T2 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + DH * OneHalf;
			if ( H < HCROSS ) {
				T3 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T3 = 216.65;
				}else{
					if (H < 32000.0) {
						T3 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T3 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T3 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T3 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T3 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T3 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			H = Y + 3 * DH * OneHalf;
			if ( H < HCROSS ) {
				T4 = 216.65 + 0.0065 * (HCROSS - H);
			}else{
				if ( H < 20000.0) {
					T4 = 216.65;
				}else{
					if (H < 32000.0) {
						T4 = 216.65 + 0.001 * (H - 20000.0);
					}else{
						if ( H < 47000.0 ){
							T4 = 228.65 + 0.0028 * (H - 32000.0);
						}else{
							if ( H < 51000.0 ){
								T4 = 270.65;
							}else{
								if ( H < 71000.0 ){
									T4 = 270.65 - 0.0028 * (H - 51000.0);
								}else{
									if ( H < 85000 ){
										T4 = 214.65 - 0.002 * (H - 71000.0);
									}else{
										T4 = 186.65;
									}
								}
							}
						}
					}
				}
			}

			fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
			fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

			DP2DY = RELH * fDVAPDT * fDTDH;
			FK4 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
			//////////////////////////////////////END STEP 4//////////////////////////////////

			fFNDPD1 = PRESSD1[i + 1] + (YSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);
		}


		///////////////////////////////////////////
		return fFNDPD1 + RELH * fVAPORH;
	}else{	
		////////////////fFNDPD2//////////////////
		H0 = H;

		Y = H0;
		YSTEP = (Y - Fix(Y)) * 10;
		i = Fix(Y/10);


		////////////////////////////////STEP 1/////////////////////////////////////////////////
		P1 = PRESSD2[i + 1];


		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}

		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}

		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK1 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		////////////////////////END STEP 1///////////////////////////////////////

		///////////////////////////////////////STEP 2////////////////////////////////////////
		Y = i + YSTEP * OneHalf;
		P1 = PRESSD2[i + 1] + FK1 * YSTEP * OneHalf;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK2 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 2////////////////////////////////////

		///////////////////////////////////////STEP 3////////////////////////////////////////
		Y = i + YSTEP * OneHalf;
		P1 = PRESSD2[i + 1] + FK2 * YSTEP * OneHalf;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK3 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 3////////////////////////////////////

		////////////////////////////////////////STEP 4///////////////////////////////////////
		Y = i + YSTEP;
		P1 = PRESSD2[i + 1] + FK3 * YSTEP;

		H = Y;
		if ( H < HCROSS ) {
			T = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T = 216.65;
			}else{
				if (H < 32000.0) {
					T = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T = 270.65;
						}else{
							if ( H < 71000.0 ){
								T = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T = 186.65;
								}
							}
						}
					}
				}
			}
		}

		switch (OPTVAP) {
		case 1:
			fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
			break;
		case 2:
			fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
			break;
		case 3:
			fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
			break;
		case 4:
			fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			break;
		}
		P2 = RELH * fVAPORH;

		switch (OPTVAP) {
		case 1:
			HV = pow((T / 247.1), 18.36); //PL2 FORM
			fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
			break;
		case 2:
			HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
			fDVAPDT = HV * 5349 / (T * T);
			break;
		case 3:
			HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
			fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
			break;
		case 4:
			HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
			fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
			break;
		}
		DH = 0.01;
		H = Y - 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T1 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T1 = 216.65;
			}else{
				if (H < 32000.0) {
					T1 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T1 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T1 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T1 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T1 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T1 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y - DH * OneHalf;
		if ( H < HCROSS ) {
			T2 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T2 = 216.65;
			}else{
				if (H < 32000.0) {
					T2 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T2 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T2 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T2 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T2 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T2 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + DH * OneHalf;
		if ( H < HCROSS ) {
			T3 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T3 = 216.65;
			}else{
				if (H < 32000.0) {
					T3 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T3 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T3 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T3 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T3 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T3 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		H = Y + 3 * DH * OneHalf;
		if ( H < HCROSS ) {
			T4 = 216.65 + 0.0065 * (HCROSS - H);
		}else{
			if ( H < 20000.0) {
				T4 = 216.65;
			}else{
				if (H < 32000.0) {
					T4 = 216.65 + 0.001 * (H - 20000.0);
				}else{
					if ( H < 47000.0 ){
						T4 = 228.65 + 0.0028 * (H - 32000.0);
					}else{
						if ( H < 51000.0 ){
							T4 = 270.65;
						}else{
							if ( H < 71000.0 ){
								T4 = 270.65 - 0.0028 * (H - 51000.0);
							}else{
								if ( H < 85000 ){
									T4 = 214.65 - 0.002 * (H - 71000.0);
								}else{
									T4 = 186.65;
								}
							}
						}
					}
				}
			}
		}

		fDTDH = ((T3 - T2) * (T7T4) - (T4 - T1) * OneDiv24) / DH;
		fGRAVRAT = pow((REARTH / (REARTH + Y)),2.0);

		DP2DY = RELH * fDVAPDT * fDTDH;
		FK4 = -DP2DY - BD * P1 * fGRAVRAT / T - BW * P2 * fGRAVRAT / T;
		//////////////////////////////////////END STEP 4//////////////////////////////////

		fFNDPD2 = PRESSD2[i + 1] + (YSTEP * OneSixth) * (FK1 + 2 * FK2 + 2 * FK3 + FK4);

		///////////////////////////////////////////
		return fFNDPD2 + RELH * fVAPORH;
	}
}

//////////////////fDVAPDT////////////////////////
double fDVAPDT (double *H0) {
////////////////////////////////////////////////////////
	//DefDbl A-H, O-Z
	//SHARED OPTVAP

	double T, HV, H;

	H = *H0;

	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		return (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		return HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		return HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		return HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	default:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		return HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}


}

////////////////////////////fDENSDRY////////////////////////////
double fDENSDRY (double *H0, double *fFNDPD1, double *fFNDPD2) {
////////////////////////////////////////////////////////////////
	//SHARED HLIMIT, AMASSD, RBOLTZ, RELH
	double PVAP, PDRY, T, fVAPORH, fPRESSURE, H;

	H = *H0;

	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}

	PVAP = RELH * fVAPORH;

	if ( H < HMAXP1 ) {
		fPRESSURE = *fFNDPD1 + RELH * fVAPORH;
	}else{
		fPRESSURE = *fFNDPD2 + RELH * fVAPORH;
	}

	PDRY = fPRESSURE - PVAP;
	return (AMASSD / RBOLTZ) * PDRY / T;
}

/////////////////////////////////////////
double fDNDH (double *H0, double *fPRESSURE, double *fDTDH) {
//////////////////////////////////////////////
	//DefDbl A-H, O-Z
	//SHARED AD, AW, BD, BW, RELH
	double T, HV, fVAPORH, PW, PD, fDVAPDT, fGRAVRAT;
	double 	DPWDH, DPDDH, HV1, HV2, H;

	H = *H0;

	if ( H < HCROSS ) {
		T = 216.65 + 0.0065 * (HCROSS - H);
	}else{
		if ( H < 20000.0) {
			T = 216.65;
		}else{
			if (H < 32000.0) {
				T = 216.65 + 0.001 * (H - 20000.0);
			}else{
				if ( H < 47000.0 ){
					T = 228.65 + 0.0028 * (H - 32000.0);
				}else{
					if ( H < 51000.0 ){
						T = 270.65;
					}else{
						if ( H < 71000.0 ){
							T = 270.65 - 0.0028 * (H - 51000.0);
						}else{
							if ( H < 85000 ){
								T = 214.65 - 0.002 * (H - 71000.0);
							}else{
								T = 186.65;
							}
						}
					}
				}
			}
		}
	}

	switch (OPTVAP) {
	case 1:
		fVAPORH = pow((T / 247.1),18.36); //PL2 VORM
		break;
	case 2:
		fVAPORH = exp(21.39 - 5349.0 / T);  //CC2 FORM
		break;
	case 3:
		fVAPORH = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T); //CC4 FORM
		break;
	case 4:
		fVAPORH = pow((T / 273.15), (2.5)) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		break;
	}

	switch (OPTVAP) {
	case 1:
		HV = pow((T / 247.1), 18.36); //PL2 FORM
		fDVAPDT = (18.36 / 247.1) * pow((T / 247.1), 17.36);
		break;
	case 2:
		HV = exp(21.39 - 5349.0 / T) ; //CC2 FORM
		fDVAPDT = HV * 5349 / (T * T);
		break;
	case 3:
		HV = exp(0.000012378847 * T * T - 0.019121316 * T + 29.33194028 - 6343.1645 / T);//CC4 FORM
		fDVAPDT = HV * (0.000012378847 * 2 * T - 0.019121316 + 6343.1645 / (T * T));
		break;
	case 4:
		HV = pow((T / 273.15) ,2.5) * exp(0.00001782 * T * T - 0.02815 * T + 30.5371935 - 6109.6519 / T); //SACKUR-TETRODE FORM, FIT.
		fDVAPDT = HV * (2.5 / T + 0.00001782 * 2 * T - 0.02815 + 6109.6519 / (T * T));
		break;
	}

	fGRAVRAT = pow((REARTH / (REARTH + H)),2.0);


	PW = RELH * fVAPORH;
	PD = *fPRESSURE - PW;
	DPWDH = RELH * fDVAPDT * *fDTDH;
	DPDDH = -DPWDH - BD * PD * fGRAVRAT / T - BW * PW * fGRAVRAT / T;
	HV1 = (AD * DPDDH + AW * DPWDH) / T;
	HV2 = -(AD * PD + AW * PW) / T / T * *fDTDH;
	return HV1 + HV2;
}

//////////////////////fGUESSL/////////////////////////
double fGUESSL (double *BETAM, double *Height) {
///////////////////////////////////////////////////////
	//Estimate for the distance along the Earth's surface
	//for a ray with tilt angle BETA to cover the height interval
	//from H=0 to H=HEIGHT.
	double A,B,C;
	B = *BETAM * RADCON;
	C = asin((REARTH * cos(B)) / (REARTH + *Height));
	A = (2 * atan(1) - B - C);
	return A * REARTH;
}

//////////////////////fGUESSP//////////////////////////
double fGUESSP (double *BETA, double *Height) {
//////////////////////////////////////////////////////////////
	//'Estimate for the pathlength from H=0 to H=HEIGHT
	//'for a ray with tilt angle BETA
	double A,B,C;
	A = 1;
	B = 2 * REARTH * sin(*BETA * RADCON);
	C = -2 * REARTH * *Height - *Height * *Height;
	return (-B + sqrt(B * B - 4 * A * C)) / 2;
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

/*************** function CInt ***************
This emulates the Visual Basic function Cint() which
converts a decimal number into the nearest integer
****************************************************/
short CInt( double x )
{
	double top, bottom;
	top = ceil(x);
	bottom = floor(x);
	if (abs(top - x) < fabs(bottom - x) )
	{
		return (short)top;
	}
	else
	{
		return (short)bottom;
	}

}



