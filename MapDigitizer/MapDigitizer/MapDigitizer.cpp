// MapDigitizer.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}

//SolveEquations packs the matrices to solve the num_row simultaneous equations to find the C_vectors of the Hardy Quadratic Equations
//It then solves the equations using Gaussian Elimination
//It then determines the heights versus screen coordinates

/////declare global variables//////////////

double	pi = acos( -1 );
double	cd = pi / 180.;
double	pi2 = pi * 2.0;
double	ch = 12. / pi;
double	hr = 60.;

//declare global arrays///////////////////

enum eDatum {
	eWGS84=0,
	eGRS80=1,
	eCLARK80M=2
};

typedef struct _DTM {
	double a;	// a  Equatorial earth radius
	double b;	// b  Polar earth radius
	double f;	// f= (a-b)/a  Flatenning
	double esq;	// esq = 1-(b*b)/(a*a)  Eccentricity Squared
	double e;	// sqrt(esq)  Eccentricity
	// deltas to WGS84
	double dX;
	double dY;
	double dZ;
} DATUM,*PDATUM;

static DATUM Datum[3] = {

	// WGS84 data
	{
		6378137.0,				// a
		6356752.3142,			// b
		0.00335281066474748,  	// f = 1/298.257223563
		0.006694380004260807,	// esq
		0.0818191909289062, 	// e
		// deltas to WGS84
		0,
		0,
		0
	},

	// GRS80 data
	{
		6378137.0,				// a
		6356752.3141,			// b
		0.0033528106811823,		// f = 1/298.257222101
		0.00669438002290272,	// esq
		0.0818191910428276,		// e
		// deltas to WGS84
		-48,
		55,
		52
	},

	// Clark 1880 Modified data
	{
		6378300.789,			// a
		6356566.4116309,		// b
		0.003407549767264,		// f = 1/293.466
		0.006803488139112318,	// esq
		0.08248325975076590,	// e
		// deltas to WGS84
		-235,
		-85,
		264
	}
};

enum gGrid {
	gICS=0,
	gITM=1
};

typedef struct _GRD {
	double lon0;
	double lat0;
	double k0;
	double false_e;
	double false_n;
} GRID,*PGRID;

static GRID Grid[2] = {

	// ICS data
	{
		0.6145667421719,			// lon0 = central meridian in radians of 35.12'43.490"
		0.55386447682762762,		// lat0 = central latitude in radians of 31.44'02.749"
		1.00000,					// k0 = scale factor
		170251.555,					// false_easting
		2385259.0					// false_northing
	},

	// ITM data
	{
		0.61443473225468920,		// lon0 = central meridian in radians 35.12'16.261"
		0.55386965463774187,		// lat0 = central latitude in radians 31.44'03.817"
		1.0000067,					// k0 = scale factor
		219529.584,					// false_easting
		2885516.9488				// false_northing = 3512424.3388-626907.390
									// MAPI says the false northing is 626907.390, and in another place
									// that the meridional arc at the central latitude is 3512424.3388
	}
};

// prototypes of local Grid <=> Lat/Lon conversions
void Grid2LatLon(double N,double E,double& lat,double& lon,gGrid from,eDatum to);
void LatLon2Grid(double lat,double lon,double& N,double& E,eDatum from,gGrid to);
// prototype of Moldensky transformation function
void Molodensky(double ilat,double ilon,double& olat,double& olon,eDatum from,eDatum to);
void itm2wgs84(double N,double E,double& lat,double &lon);
void wgs842itm(double lat,double lon,double& N,double& E);
void ics2wgs84(double N,double E,double *lat, double *lon);
void wgs842ics(double lat,double lon,double& N,double& E);
int Nint(double x);
double Fix( double x );

/////////////////////////////////////////////
//arr must be defined in VB6 as arr(0 to np - 1, 0 to np + 1); 
///////////////////////////////////////////

int __declspec (dllexport) __stdcall SolveEquationsC(double *pnX, double *pnY, double *pnZ,
													 double *pnC_vector,
													 double *pnXcoord, double *pnYcoord,
													 double *pnht, float *pnhtf, short *pnhts, short *HeightMode, 
													 double *xmin, double *StepX, double *ymin, double *StepY,
													 long *np, long cbAddress) {
/*
	long __declspec (dllexport) __stdcall SolveEquationsC(double *pnX, double *pnY, double *pnZ,
													 double *pnarr, double *pnC_vector,
													 double *pnXcoord, double *pnYcoord, double *pnht, 
													 double *xmin, double *StepX, double *ymin, double *StepY,
													 long *np, long cbAddress) {

*/
/*
long __declspec (dllexport) __stdcall SolveEquationsC(double *pnX, double *pnY, double *pnZ,
													 double *pnC_vector,
													 double *pnXcoord, double *pnYcoord, double *pnht, 
													 double *xmin, double *StepX, double *ymin, double *StepY,
													 long *np) {
*/

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


	long i,j, k, l, npm1, rows, cols, ier;
	double dx, dy;

	//dimensions of 2D array to be defined
	rows = *np;
	cols = *np;

     // Declare the function pointer, with one short integer argument.
     typedef void (__stdcall *FUNCPTR)(int parm);
     FUNCPTR vbFunc;

     // Point the function pointer at the passed-in address.
     vbFunc = (FUNCPTR)cbAddress;


	ier = 0;

	//create 2D array from 1D passed array
	//initialize 2-D array in C
	double **arr2D;
	arr2D = new double *[rows + 1];
	for( i = 0; i <= rows + 1; i++ )
		*(arr2D + i) = new double[cols + 2];

	/*
	for (i = 0; i <= rows; i++ ) {
         for (j = 0; j <= cols; j++) {
            arr2D[j][i] = pnarr[j+i* (cols + 2)];

		  
		  //debugging=====================
		  buff[0] = 0; //clear buffer
		  sprintf(buff, "%s%d%s%d%s%f", "arr2D[",j,"][",i,"] =", arr2D[j][i]);

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
        }
		  //===================================
		*/

/*
    //never implemented just an example
	//create 2D array
	double **arr_2D;
	arr_2D = (double **)malloc(rows * sizeof(double *));
	if(arr_2D == NULL)
		{
		MessageBox(NULL, "out of memory\n", "Error",
                   MB_OK | MB_ICONINFORMATION);
		return -1;
		}
	for(i = 0; i < rows; i++)
		{
		arr_2D[i] = (double *)malloc(cols * sizeof(double));
		if(arr_2D[i] == NULL)
			{
			MessageBox(NULL, "out of memory\n", "Error",
					   MB_OK | MB_ICONINFORMATION);
			return -1;
			}
		}
*/

LG5:
	// Call the UpdateStatus  through the function pointer.
    //vbFunc(0);
	
	for (i = 1; i <= *np; i++)
	{
		for (j = 1; j <= i; j++)
		{
			if (i != j) //goto LG10;
			{
				//arr2D[i][j] = 0.0;
			    //goto LG15;
//LG10:
				dx = pnX[j] - pnX[i];
				dy = pnY[j] - pnY[i];

		  	/*
			//debugging=====================
			buff[0] = 0; //clear buffer
			sprintf(buff, "%s%d%s%f%s%d%s%f\n%s%d%s%f%s%d%s%f", "pnX[",j,"] = ", pnX[j], "pnX[",i,"] =", pnX[i], "pnY[",j,"] = ", pnY[j], "pnY[",i,"] =", pnY[i]);
			MessageBox(NULL, (const char *)buff, "Debugging",
				   MB_OK | MB_ICONINFORMATION);
			//===================================
			*/

				if (dx != 0.0 || dy != 0.0) //goto LG25;
				{
					arr2D[i][j] = sqrt(dx * dx + dy * dy);
					arr2D[j][i] = arr2D[i][j];
				}
				else
				{
					goto LG25;
				}
			}
			else
			{
				arr2D[i][j] = 0.0;
			}
//LG15:	;
		}
		arr2D[i][*np + 1] = pnZ[i];

		// Call the UpdateStatus  through the function pointer.
        // vbFunc((long)floor(i * 100/ *np)); //very fast, doesn't need it

		/*
		//debugging=====================
		buff[0] = 0; //clear buffer
		sprintf(buff, "%s%d%s%d%s%f\n%s%d%s%f", "arr2D[",i,"][", *np + 1, "] = ", arr2D[i][*np + 1], "pnZ[",i, "] = ", pnZ[i]);
		MessageBox(NULL, (const char *)buff, "Debugging",
			   MB_OK | MB_ICONINFORMATION);
		//===================================
		*/

	}	
	goto LG40;
LG25:	
	//if got here, two points are the same, remove one and repeat the process
	if (i == *np) goto LG35;
	npm1 = *np - 1;
	for (k = i; k <= npm1; k++)
	{
		l = k + 1;
		pnX[k] = pnX[l];
		pnY[k] = pnY[l];
		pnZ[k] = pnZ[l];
	}
LG35:	
	*np = *np - 1;
	goto LG5;
LG40:	

//////////////////////////original VB6 code/////////////////////////////////////////////
/*
        For i = 1 To np
           For j = 1 To i
               If (i <> j) Then GoTo LG10
               Matrix_A(i, j) = 0#
               GoTo LG15
LG10:
               dX = x(j) - x(i)
               dy = Y(j) - Y(i)
               If (dX = 0# And dy = 0#) Then GoTo LG25
'               Matrix_A(i, j) = Sqr(dX ^ 2 + dy ^ 2)
               Matrix_A(i, j) = Sqr(dX * dX + dy * dy)
               Matrix_A(j, i) = Matrix_A(i, j)
LG15:
           Next j
LG20:
           Matrix_A(i, np + 1) = Z(i)
'           DoEvents
        Next i
        GoTo LG40
    '     two points are the same, remove one and repeat the process
LG25:
        If (i = np) Then GoTo LG35
        npm1 = np - 1
        For k = i To npm1
            l = k + 1
            x(k) = x(l)
            Y(k) = Y(l)
LG30:
            Z(k) = Z(l)
'            DoEvents
        Next k
LG35:
        np = np - 1
        DoEvents
        If npm1 > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * (numDigiHardyPoints - np + 2) / npm1)
        GoTo LG5
LG40:
*/
////////////////////////////////////////////////////////////////////////////////////////////////////////

	const double TINY = 0.00001;
	int num_rows, num_cols;
	int r,c,r2;
	double tmp,factor;

	num_rows = *np;
	num_cols = *np;

	// Call the UpdateStatus  through the function pointer.
    vbFunc(0);

	for (r = 1; r < num_rows; r++)
	{
		//Zero out all entries in column r after this row
		//See if this row has a non-zero entry in column r
		if (fabs(arr2D[r][r]) < TINY)
		{
			//Not a non-zero value.  Try to swap with a later row.
			for (r2 = r + 1; r <= num_rows; r++)
			{
				if (fabs(arr2D[r2][r]) > TINY )
				{
					//This row will work.  Swap them
					for (c = 1; c <= num_cols + 1; c++)
					{
						tmp = arr2D[r][c];
						arr2D[r][c] = arr2D[r2][c];
						arr2D[r2][c] = tmp;
					}
					break; //exit the loop 
				}
			}
		}

		//If this row has a non-zero entry in column r, skip this column
		if (fabs(arr2D[r][r]) > TINY )
		{
			//Zero out this column in later rows.
			for (r2 = r + 1; r2 <= num_rows; r2++)
			{
				factor = -arr2D[r2][r]/arr2D[r][r];
				for (c = r; c <= num_cols + 1; c++)
				{
					arr2D[r2][c] = arr2D[r2][c] + factor * arr2D[r][c];
				}
			}
		}

		// Call the UpdateStatus  through the function pointer.
        vbFunc((long)floor(r * 100 / (num_rows - 1)));

	}

	//See if we have a solution
    if (arr2D[num_rows][num_cols] == 0.0)
	{
        // We have no solution.
        ier = -1;
        return ier;
	}
    else
	{
        // Back solve
        for (r = num_rows; r >= 1; r--)
		{
            tmp = arr2D[r][num_cols + 1];
            for( r2 = r + 1; r2 <= num_rows; r2++)
			{
                tmp = tmp - arr2D[r][r2] * arr2D[r2][num_cols + 2];
			}
            arr2D[r][num_cols + 2] = tmp / arr2D[r][r]; //<-- error must be -7.997 not 0
            pnC_vector[r] = arr2D[r][num_cols + 2];

			/*
			//debugging=====================
			buff[0] = 0; //clear buffer
			sprintf(buff, "%s%d%s%f", "pnC_vector[",r,"] = ", pnC_vector[r]);
			MessageBox(NULL, (const char *)buff, "Debugging",
				   MB_OK | MB_ICONINFORMATION);
			//===================================
			*/

        }
	}


//now SolveEquations code
/*

    For r = 1 To num_rows - 1
        ' Zero out all entries in column r after this row.
        ' See if this row has a non-zero entry in column r.
        If Abs(arr(r, r)) < TINY Then
            ' Not a non-zero value. Try to swap with a later row.
            For r2 = r + 1 To num_rows
                If Abs(arr(r2, r)) > TINY Then
                    ' This row will work. Swap them.
                    For c = 1 To num_cols + 1
                        tmp = arr(r, c)
                        arr(r, c) = arr(r2, c)
                        arr(r2, c) = tmp
                    Next c
                    Exit For
                End If
            Next r2
        End If

        ' If this row has a non-zero entry in column r, skip this column.
        If Abs(arr(r, r)) > TINY Then
            ' Zero out this column in later rows.
            For r2 = r + 1 To num_rows
                factor = -arr(r2, r) / arr(r, r)
                For c = r To num_cols + 1
                    arr(r2, c) = arr(r2, c) + factor * arr(r, c)
                Next c
            Next r2
        End If
        
        DoEvents
        Call UpdateStatus(GDMDIform, 1, 100 * (r / (num_rows - 1)))

    Next r


    ' See if we have a solution.
    If arr(num_rows, num_cols) = 0 Then
        ' We have no solution.
        ier = -1
        SolveEquations = ier
        Exit Function
    Else
        ' Back solve.
        Call UpdateStatus(GDMDIform, 1, 0)
        For r = num_rows To 1 Step -1
            tmp = arr(r, num_cols + 1)
            For r2 = r + 1 To num_rows
                tmp = tmp - arr(r, r2) * arr(r2, num_cols + 2)
            Next r2
            arr(r, num_cols + 2) = tmp / arr(r, r)
            C_vector(r) = arr(r, num_cols + 2)
            
            DoEvents
            Call UpdateStatus(GDMDIform, 1, 100 * ((num_rows - r + 1) / num_rows))

        Next r

    End If

   SolveEquations = ier
*/

	//now calculate the height array at Xcoord, Ycoord that will be used creating coordinate height output as well as DTM
	long ixp, iyp, m;
	double xi, yi, zi;

	/*
	double **ht2D;
	ht2D = new double * [num_rows];
	for( i = 0; i < num_rows; i++ )
		*(ht2D + i) = new double[num_cols];
    */

	/*
	for (i = 0; i < num_rows; i++ ) {
         for (j = 0; j < num_cols; j++) {
            ht2D[j][i] = pnht[j+i*num_cols];
		 }
	}
	*/

	//in case some data has been eliminated as duplicates, recalculate the step sizes
	*StepX = *StepX * (double)(cols - 1)/(double)(num_cols - 1);
	*StepY = *StepY * (double)(rows - 1)/(double)(num_rows - 1);

	// Call the UpdateStatus  through the function pointer.
    vbFunc(0);

	switch (*HeightMode)
	{
	case 0: //short integer elevations
		for (ixp = 1; ixp <= *np; ixp++)
		{
			xi = *xmin + (ixp - 1) * *StepX;
			pnXcoord[ixp - 1] = xi;
			for (iyp = 1; iyp <= *np; iyp++)
			{
				yi = *ymin + (iyp - 1) * *StepY;
				m = *np;
				zi = 0.0;
				for (j = 1; j <= *np; j++)
				{
					dx = xi - pnX[j];
					dy = yi - pnY[j];
					zi += pnC_vector[j] * sqrt(dx * dx + dy * dy);
				}
				pnhts[ixp - 1 + (iyp - 1) * num_rows] = (short)Nint(zi);
				//ht2D[ixp - 1][iyp - 1] = zi;
				pnYcoord[iyp - 1] = yi;
			}

			// Call the UpdateStatus  through the function pointer.
			 vbFunc((long)floor(ixp * 100/ *np));

		}		break;
	case 1: //float precision elevations
		for (ixp = 1; ixp <= *np; ixp++)
		{
			xi = *xmin + (ixp - 1) * *StepX;
			pnXcoord[ixp - 1] = xi;
			for (iyp = 1; iyp <= *np; iyp++)
			{
				yi = *ymin + (iyp - 1) * *StepY;
				m = *np;
				zi = 0.0;
				for (j = 1; j <= *np; j++)
				{
					dx = xi - pnX[j];
					dy = yi - pnY[j];
					zi += pnC_vector[j] * sqrt(dx * dx + dy * dy);
				}
				pnhtf[ixp - 1 + (iyp - 1) * num_rows] = (float)zi;
				//ht2D[ixp - 1][iyp - 1] = zi;
				pnYcoord[iyp - 1] = yi;
			}

			// Call the UpdateStatus  through the function pointer.
			 vbFunc((long)floor(ixp * 100/ *np));

		}		break;
	case 2: //double precision elevations

		for (ixp = 1; ixp <= *np; ixp++)
		{
			xi = *xmin + (ixp - 1) * *StepX;
			pnXcoord[ixp - 1] = xi;
			for (iyp = 1; iyp <= *np; iyp++)
			{
				yi = *ymin + (iyp - 1) * *StepY;
				m = *np;
				zi = 0.0;
				for (j = 1; j <= *np; j++)
				{
					dx = xi - pnX[j];
					dy = yi - pnY[j];
					zi += pnC_vector[j] * sqrt(dx * dx + dy * dy);
				}
				pnht[ixp - 1 + (iyp - 1) * num_rows] = zi;
				//ht2D[ixp - 1][iyp - 1] = zi;
				pnYcoord[iyp - 1] = yi;
			}

			// Call the UpdateStatus  through the function pointer.
			 vbFunc((long)floor(ixp * 100/ *np));

		}

		break;
	}

	//now unpack the height array into ht

	/*
	for (i = 0; i < num_cols; i++)
	{
		for (j = 0; j < num_rows; j++)
		{
			pnht[i + j*(num_rows)] = ht2D[i][j];
		}
	}
	*/


/*
        For ixp = 1 To np
            xi = xmin + (ixp - 1) * StepX
            Xcoord(ixp - 1) = xi
            For iyp = 1 To np
                yi = ymin + (iyp - 1) * StepY
                m = np + 1
                zi = 0#
                For j = 1 To System_DIM 'np
                    dX = xi - X(j)
                    dy = yi - Y(j)
LGL25:
                    zi = zi + C_vector(j) * Sqr(dX * dX + dy * dy)
'                    DoEvents
                Next j
                ht(ixp - 1, iyp - 1) = zi
                Ycoord(iyp - 1) = yi
                If HardyCoordinateOutput Then
                    'convert pixel coordinates to geographic coordinates  'use this to output final values
                    If RSMethod1 Then
                       ier = RS_pixel_to_coord2(CDbl(xi), CDbl(yi), XGeo, YGeo)
                    ElseIf RSMethod2 Then
                       ier = RS_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                    ElseIf RSMethod0 Then
                       ier = Simple_pixel_to_coord(CDbl(xi), CDbl(yi), XGeo, YGeo)
                       End If
                Else
                   'output pixel coordinates
                   XGeo = xi
                   YGeo = yi
                   End If
                   
                If Mid$(LCase(lblX), 1, 3) = "itm" And Mid$(LCase(LblY), 1, 3) = "itm" Then
                   XGeo = CLng(XGeo)
                   YGeo = CLng(YGeo)
                ElseIf Mid$(LCase(lblX), 1, 3) = "lon" And Mid$(LCase(LblY), 1, 3) = "lat" Then
                   XGeo = val(Format(str$(XGeo), "####0.0####"))
                   YGeo = val(Format(str$(YGeo), "####0.0####"))
                   End If
                   
                Write #filnum%, XGeo, YGeo, zi '* 10#
                Write #filtopo%, Xcoord(ixp - 1), Ycoord(iyp - 1), zi '* 10#
                If numDigiHardyPoints > 0 Then Call UpdateStatus(GDMDIform, 1, 100 * ixp / np) 'numDigiHardyPoints)
                DoEvents
                
                '---------------------break on ESC key-------------------------------
                If GetAsyncKeyState(vbKeyEscape) < 0 Then 'escape out of this
                   EscapeKeyPressed = True
                   GoTo LGL40
                   End If
                '---------------------------------------------------------------------
                
            Next iyp
LGL35:
        Next ixp
*/


   return ier; 
}

/*
long __declspec (dllexport) __stdcall SolveEquationsDTM(long *np1, long *np2, long *np3, long *np4, long *np5, long *np6, long *np7) {

	char buff[255] = ""; 
	sprintf(buff, "%s%d, %d, %d, %d, %d, %d, %d", "np1, np2 = ", *np1, *np2, *np3, *np4, *np5, *np6, *np7);
	MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);


	return 0;

}
*/


/*
long __declspec (dllexport) __stdcall SolveEquationsDTM(long *np1, long *np, long *num_rows, long *num_cols,
														double *pnX, double *pnY, double *pnC_vector,
													    double *pnXcoord, double *pnYcoord, double *pnZcoord,
													    long cbAddress, long *np2) {
*/
int __declspec (dllexport) __stdcall SolveEquationsDTM(long *np, long *num_rows, long *num_cols,
														double *pnX, double *pnY, double *pnC_vector,
													    double *pnXcoord, double *pnYcoord,
														double *pnZcoord, float *pnZcoordf, short *pnZcoords, short *HeightMode,
													    long cbAddress) {

	
	double xi, yi, zi, numDigiHardyPoints;
	long i, j, k, l, npm1;
	double dx, dy;

	DWORD dwError, dwThreadPri;

	//set high priority

    if(!SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_HIGHEST))
    {
		char buff[255] = ""; 
        dwError = GetLastError();
	    dwThreadPri = GetThreadPriority(GetCurrentThread());
	    sprintf(buff, "%s\n%s%d", "Error in setting priority", "Current priority level: ", dwThreadPri);
		//const int result = MessageBox(NULL, (const char *)buff, "Debugging",  MB_OKCANCEL);
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

      
	 //diagnostics========================================================
	 /*
	 char buff[255] = "";
	 FILE *stream;
	 bool opened = false;

	 if( (stream  = fopen( "c:\\jk\\testdata.txt", "w" )) == NULL )
	 {
		 opened = false;
	 }
	 else
	 {
		 opened = true;
		 fprintf(stream, "%s\n", "xi, X[k], dx, yi, Y[k], dy, C_vector[k+1], zi");
	 }
	 */
	 //=================================================================
	
	    /*
		buff[0] = 0;
		sprintf(buff, "%s%d,%d,%d,%d,%d,%f,%f,%f,%f,%f,%d", "np1, np2, np, num_rows, num_cols, x(0), y(0), C_vector(0), Xcoord(0), Ycoord(0), np3 = ", *np1, *np2, *np, *num_rows, *num_cols, pnX[0], pnY[0], pnC_vector[0], pnXcoord[0], pnYcoord[0], *np3);
		const int result = MessageBox(NULL, (const char *)buff, "Debugging", MB_OKCANCEL | MB_ICONINFORMATION);
		switch (result)
		{
		case IDOK:
			// continue looping
			break;
		case IDCANCEL:
			// exit
			int ier = -1;  //skip rest of checks
			return ier;
			break;
		}
		*/

		/*
	    for (k = 0; k < *np; k++)
		{
			buff[0] = 0;
			sprintf(buff, "%s%d%s%d%s%f, %f", "x[", k, "], y[", k, "] = ", pnX[k], pnY[k]);
			MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
		}
		*/	
		
	    //remove duplicate x,y's

		numDigiHardyPoints = *np;

	    vbFunc(0);
		for (i = 1; i <= *np; i++)
		{
			for (j = 1; j <= i; j++)
			{
				if (i != j) 
				{
					if ( (pnX[j - 1] == pnX[i - 1]) && (pnY[j - 1] == pnY[i - 1]) )
					{
						goto LG25;
					}
				}
			}
 
		}	
		goto LG40;
LG25:	
		//if got here, two points are the same, remove one and repeat the process
		if (i == *np) goto LG35;
		npm1 = *np - 1;
		for (k = i; k <= npm1; k++)
		{
			l = k + 1;
			pnX[k - 1] = pnX[l - 1];
			pnY[k - 1] = pnY[l - 1];
		}
LG35:	
		*np = *np - 1;
		if (npm1 > 0)
			vbFunc((long)floor(100 * (numDigiHardyPoints - *np + 2) / npm1));

		goto LG5;
LG40:	

		// Call the UpdateStatus  through the function pointer.
		vbFunc(0);

		switch (*HeightMode)
		{
		case 0: //short integer precision elevations
			for (i = 0; i < *num_cols; i++)
			{
				xi = pnXcoord[i];
				
				for(j = 0; j < *num_rows; j++)
				{
					yi = pnYcoord[j];

					zi = 0.0;
					for (k = 0; k < *np; k++)
					{
						dx = xi - pnX[k];
						dy = yi - pnY[k];

						zi += pnC_vector[k + 1] * sqrt(dx * dx + dy * dy);
						
					}
					pnZcoords[i + j * *num_cols ] = (short)Nint(zi);

				}
    
				vbFunc((long)floor(i * 100/ (*num_cols - 1)));
    
			}
			break;
		case 1: //floating precision elevations
			for (i = 0; i < *num_cols; i++)
			{
				xi = pnXcoord[i];
				
				for(j = 0; j < *num_rows; j++)
				{
					yi = pnYcoord[j];

					zi = 0.0;
					for (k = 0; k < *np; k++)
					{
						dx = xi - pnX[k];
						dy = yi - pnY[k];

						zi += pnC_vector[k + 1] * sqrt(dx * dx + dy * dy);
						
					}
					pnZcoordf[i + j * *num_cols ] = (float)zi;

				}
    
				vbFunc((long)floor(i * 100/ (*num_cols - 1)));
    
			}			break;
		case 2: //double precision elevations

			for (i = 0; i < *num_cols; i++)
			{
				xi = pnXcoord[i];
				
				/*
				buff[0] = 0;
				sprintf(buff, "%s%d, %f", "i, xi =", i, xi);
				MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
				*/
				
				for(j = 0; j < *num_rows; j++)
				{
					yi = pnYcoord[j];
					
					/*
					buff[0] = 0;
					sprintf(buff, "%s%d, %f", "j, yi =", j, yi);
					MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
					*/

					zi = 0.0;
					for (k = 0; k < *np; k++)
					{
						dx = xi - pnX[k];
						dy = yi - pnY[k];
						
						/*
						buff[0] = 0;
						sprintf(buff, "%s%d%s%f, %f, %f, %f, %f, %f, %f", "xi, X, dx, yi, Y, dy, C_vector[", k + 1, "] =", xi, pnX[k], dx, yi, pnY[k], dy, pnC_vector[k+1]);
						MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
						*/

						zi += pnC_vector[k + 1] * sqrt(dx * dx + dy * dy);
						
						/*
						buff[0] = 0;
						sprintf(buff, "%s%f", "zi =", zi);
						MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
						*/

						//diagnostics=================================
						/*
						if (opened)
						{
							fprintf(stream, "%f,%f,%f,%f,%f,%f,%f,%f\n", xi, pnX[k], dx, yi, pnY[k], dy, pnC_vector[k+1], zi);
						}
						*/
						//===============================================

					}
					pnZcoord[i + j * *num_cols ] = zi;

					//diagnostics====================================================
					/*
					if (opened)
					{
						fprintf(stream, "%f\n", zi);
					}
					*/
					//===============================================================
					
					/*
					buff[0] = 0;
					sprintf(buff, "%s%d%s%d, %d, %d, %f", "i, j, num_cols, Zcoord[", i + j * *num_cols, "] =", i, j, *num_cols, pnZcoord[i + j * *num_cols ]);
					MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
					*/
		
				}

				/*
				buff[0] = 0;
				sprintf(buff, "%s%d", "iteration = ", (long)floor(i * 100/ (*num_cols - 1)));
				MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
				*/
    
				vbFunc((long)floor(i * 100/ (*num_cols - 1)));
    
			}
			break;
		}

		//diagnostics==========================
		/*
		if (opened)
		{
			fclose(stream);
			opened = false;
		}
		*/
		//=========================================

		/*
		buff[0] = 0;
		sprintf(buff, "%s%d", "num_cols =", *num_cols);
		MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);


		buff[0] = 0;
		sprintf(buff, "%s%d, %d", "i, final callback =", i, (long)floor(i * 100/ (*num_cols - 1)));
		MessageBox(NULL, (const char *)buff, "Debugging", MB_OK | MB_ICONINFORMATION);
		*/


	int ier = 0;



	return ier;
}

////////////////////////version without any file reading and writing/////////////////////////////////////
/* 
int __declspec (dllexport) __stdcall Profiles (double *pnxo, double *pnyo, double *pnzo, 
											   double *pnxx, double *pnyy, double *pnzz,
                                               int *num_array, short *CoordMode, short *HorizMode,
                                               double *pnva, double *pnazi, 
											   double *pnXazi, double *pnYazi, double *pnZazi, double *distazi,
											   int *numazi, double *StepSizeAzi, double *HalfAziRange, 
                                               long cbAddress){



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


	short ier = 0;

	union profile
	{
		float ver[4];
	};

	union profile *prof;

	union dtm
	{
		double ver[4];
	};

	union dtm *dtms;

	//allocate memory for prof array
	free( prof ); //free the memory allocated before for prof, and reallocate
	prof = (union profile *) malloc((*numazi) * sizeof(union profile));

	//allocate memory for dtms array
	free( dtms ); //free the memory allocated before for dtms, and reallocate
	dtms = (union dtm *) malloc((*num_array) * sizeof(union dtm));

	double kmxo, kmyo, hgt, lt, lg;
	kmxo = *pnxo;
	kmyo = *pnyo;
	hgt = *pnzo;

	double N = 0;
	double E = 0;

	double maxang = 0;
	maxang = fabs(*HalfAziRange);

	//convert origin to WGS84

	if (*CoordMode == 1) //ITM
	{
		
		if (kmyo < 1000)
		{
			N = kmyo * 1000 + 1000000;
		}
		else
		{
			N = kmyo;
		}

		if (kmxo < 1000)
		{
			E = kmxo * 1000;
		}
		else
		{
			E = kmxo;
		}

		//use J. Gray's conversion from ITM to WGS84
		//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
		ics2wgs84(N, E, &lt, &lg); 

	}
	else //coordinates already in geographic latitude and longitude
	{
		lt = kmyo;
		lg = kmxo;
	}

	short minview = 10;
	double anav = 0;

	//initialize prof arry and atmospheric refraction constants

	if (hgt >= 0.) {
	    anav = sqrt(hgt) * -.0348f;
	}
	if (hgt < 0.) {
	    anav = -1.f;
	}

	long i__ = 0;

	for ( i__ = 0; i__ < *numazi; ++i__) {
		prof[i__].ver[0] = (float)anav - minview;
		prof[i__].ver[1] = 0;
		prof[i__].ver[2] = 0;
		prof[i__].ver[3] = 0;
	}


	double re, rekm, testazi;
	double hgt2, treehgt = 0, fudge = 0;
	double lg2, lt2, x1, x2, y1, y2, z1, z2, d__1, d__2, d__3;
	double distd, re1, re2, deltd, dist1, dist2, angle, viewang;
	double d__, x1d, y1d, z1d, azicos, x1s, y1s, z1s, x1p, y1p, z1p, azisin, azi;
	double defm, defb, avref, kmxp, kmyp, bne1, bne2,hgt_1, hgt_2;
	int istart1, ndela, ndelazi, ne, nstart2,i__4;
	double az2,vang2,kmy2,az1,vang1,kmy1,bne, vang, proftmp, delazi, dist;
	int invAzimuthStep = 10;

	bool found = false;

	double AzimuthStep = 0;
	AzimuthStep = *StepSizeAzi;
	invAzimuthStep = (int)(1.0 / AzimuthStep);

    re = 6371315.; //radius of earth in meters
    rekm = 6371.315f; //radius of earth in km

	long numdtm = 0; //counter for number of points used in calculation

	double kmx = 0;
	double kmy = 0;

	lt2 = 0;
	lg2 = 0;
	hgt2 = 0;

	//the data is organized as an x,y,z file, increasing x, decreasing y, z where x is the outer loop.
	double kmxcurrent = pnxx[0];
	numdtm = 0;

	long numrctot = 0; //total number of columns processed

	// Call the UpdateStatus  through the function pointer.
    vbFunc(0);

	for (i__ = 0; i__ < *num_array; i__ ++)
	{
		//if Eastern horizon, use only points that longitude are to the East point
		//for Western horizon, use ony points that are to the West of the center point

		kmx = pnxx[i__];
		kmy = pnyy[i__];
		hgt2 = pnzz[i__];

		if (kmx != kmxcurrent) //corresponds to start of new column in y
		{
		    //analyse one column of azimuths

			///////////////////////////////////////////////////////////////////////////////////////////////////////

	//      interpolate to a grid of 0.1 degrees in azimuth 
	//      the interpolated points reside in the array prof() 

			istart1 = 1;
	L150:
			if (*HorizMode == 2) {
				az2 = dtms[istart1].ver[0];
				vang2 = dtms[istart1].ver[1];
				kmy2 = dtms[istart1].ver[2];
				hgt_2 = dtms[istart1].ver[3];
			} else {
				az1 = dtms[istart1].ver[0];
				vang1 = dtms[istart1].ver[1];
				kmy1 = dtms[istart1].ver[2];
				hgt_1 = dtms[istart1].ver[3];
			}
			nstart2 = istart1 + 1;
	L160:
			if (*HorizMode == 2) {
				az1 = dtms[nstart2].ver[0];
				vang1 = dtms[nstart2].ver[1];
				kmy1 = dtms[nstart2].ver[2];
				hgt_1 = dtms[nstart2].ver[3];
			} else {
				az2 = dtms[nstart2].ver[0];
				vang2 = dtms[nstart2].ver[1];
				kmy2 = dtms[nstart2].ver[2];
				hgt_2 = dtms[nstart2].ver[3];
			}


			if (az2 >= -maxang && az1 <= maxang + AzimuthStep) {
				bne1 = (int)(az1 * invAzimuthStep) / (double)invAzimuthStep - AzimuthStep;
				bne2 = (int)(az2 * invAzimuthStep) / (double)invAzimuthStep + AzimuthStep;
				found = false;
				ndela = (int) ((bne2 - bne1) / AzimuthStep + 1);
				i__4 = ndela;
				for (ndelazi = 1; ndelazi <= i__4; ++ndelazi) {
					delazi = bne1 + (ndelazi - 1) * AzimuthStep;
					if (delazi < -maxang) {
						goto L165;}
					if (delazi > maxang + AzimuthStep) {
						goto L170;}


					if (az1 <= delazi && az2 >= delazi) {
						bne = (delazi + maxang) * (double)invAzimuthStep + 1.f;
						ne = Nint(bne);
						if (ne >= 0 && ne <= *numazi) {
							vang = (vang2 - vang1) / (az2 - az1) * (delazi - az1) + vang1;
							proftmp = prof[ne].ver[0];
							if (vang > proftmp) {
								prof[ne].ver[0] = (float)vang;
								prof[ne].ver[1] = (float)(kmxcurrent);
								prof[ne].ver[2] = (float)((kmy2 - kmy1) / (az2 - az1) * (delazi - az1) + kmy1);
								prof[ne].ver[3] = (float)((hgt_2 - hgt_1) / (az2 - az1) * (delazi - az1) + hgt_1);
								}
							found = true; }
					} else if (az1 < delazi && az2 < delazi) {
						goto L170;}
L165:
					;
				}
L170:
				if (found) {
					goto L180; }
				++nstart2;
				if (nstart2 <= numdtm) {
					goto L160;
				} else {
					goto L180; }
			}
L180:
			++istart1;
			if (istart1 < numdtm) {
			goto L150;	}

			numrctot ++;

			vbFunc((long)floor(numrctot * 100/ (sqrt(*num_array))));
			///////////////////////////////////////////////////////////////////////////////////////////////////////

			numdtm = 0;
			kmxcurrent = kmx;
		}

		if (*HorizMode == 1) //Eastern horizon
		{
			if (kmx < kmxo)
				goto loop999; //point is west of center point, so ignore it
		}
		else
		{
			if (kmx > kmxo)
				goto loop999; //point is east of center point, so ignore it
		}

		//now begin further optimizations (it is assumed that all data used for this program are relatively close by, and no need for distant-point optimizations)
		//use crude approximation of azimuth to
		//determine if data will lie within the required azimuth range
		if (kmx != kmxo)
		{
			testazi = (kmy - kmyo)/(kmx - kmxo);
			testazi = atan(testazi)/cd;
			if (testazi > maxang + 10.0 || testazi < -maxang - 10.0 ) goto loop999;
		}

		if (*CoordMode == 1) //ITM
		{

			//convert old ITM grid to WGS84

			if (kmy < 1000)
			{
				N = kmy * 1000 + 1000000;
			}
			else
			{
				N = kmy;
			}

			if (kmx < 1000)
			{
				E = kmx * 1000;
			}
			else
			{
				E = kmx;
			}

			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt2, &lg2);

		}
		else //coordinates already in geographic latitude and longitude
		{
			lt2 = kmy;
			lg2 = kmx;
		}

//      calculate the view angle and azimuth, without 
//      atmospheric refraction for the selected observation point 
//      at longitude kmx,latitude kmy, height hgt2 

		fudge = treehgt;

		//if (hgt2 == -9999. || hgt2 == -32768. || fabs(hgt2) > 8848.) {
		//hgt2 = 0.f;}

		hgt2 += fudge;
		x1 = cos(lt * cd) * cos(-lg * cd);
		x2 = cos(lt2 * cd) * cos(-lg2 * cd);
		y1 = cos(lt * cd) * sin(-lg * cd);
		y2 = cos(lt2 * cd) * sin(-lg2 * cd);
		z1 = sin(lt * cd);
		z2 = sin(lt2 * cd);
		// Computing 2nd power 
		d__1 = x1 - x2;
		// Computing 2nd power 
		d__2 = y1 - y2;
		// Computing 2nd power 
		d__3 = z1 - z2;
		distd = rekm * sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
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
//          view angle in radians 
		viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
		d__ = (dist1 - dist2 * cos(angle)) / dist1;
		x1d = x1 * (1 - d__) - x2;
		y1d = y1 * (1 - d__) - y2;
		z1d = z1 * (1 - d__) - z2;
		x1p = -sin(-lg * cd);
		y1p = cos(-lg * cd);
		z1p = 0.;
		azicos = x1p * x1d + y1p * y1d;
		x1s = -cos(-lg * cd) * sin(lt * cd);
		y1s = -sin(-lg * cd) * sin(lt * cd);
		z1s = cos(lt * cd);
		azisin = x1s * x1d + y1s * y1d + z1s * z1d;
		azi = atan(azisin / azicos);
//      azimuth in degrees 
		azi /= cd;
//      add contribution of atmospheric refraction to view angle 
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

		++numdtm;
		dtms[numdtm].ver[0] = azi;
		dtms[numdtm].ver[1] = viewang / cd + avref;
		dtms[numdtm].ver[2] = kmy;
		dtms[numdtm].ver[3] = hgt2;

loop999: ;
	}

	//finished analysing the data, now output the view angle, etc., vs. azimuth

	vbFunc(0);

	for (i__ = 0; i__ < *numazi; i__++)
	{
		pnva[i__] = prof[i__].ver[0];
		pnazi[i__] = -maxang + i__ * AzimuthStep;
		pnXazi[i__] = prof[i__].ver[1];
		pnYazi[i__] = prof[i__].ver[2];
		pnZazi[i__] = prof[i__].ver[3];

		//now calculate the distance to the obstruction from the calculation center
		kmxp = prof[i__].ver[1];
		kmyp = prof[i__].ver[2];

		if (*CoordMode == 1) //ITM
		{

			//convert old ITM grid to WGS84

			if (kmyp < 1000)
			{
				N = kmyp * 1000 + 1000000;
			}
			else
			{
				N = kmyp;
			}

			if (kmxp < 1000)
			{
				E = kmxp * 1000;
			}
			else
			{
				E = kmxp;
			}

			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt2, &lg2);
			//lg2 = - lg2; //our convention
		}
		else //coordinates already in geographic latitude and longitude
		{
			lt2 = kmyp;
			lg2 = kmxp;
		}

		x1 = cos(lt * cd) * cos(-lg * cd);
		x2 = cos(lt2 * cd) * cos(-lg2 * cd);
		y1 = cos(lt * cd) * sin(-lg * cd);
		y2 = cos(lt2 * cd) * sin(-lg2 * cd);
		z1 = sin(lt * cd);
		z2 = sin(lt2 * cd);
		re = 6371.315;
		re1 = 1.;
		re2 = 1.;
		x1 = re1 * x1;
		y1 = re1 * y1;
		z1 = re1 * z1;
		x2 = re2 * x2;
		y2 = re2 * y2;
		z2 = re2 * z2;
	// Computing 2nd power 
		d__1 = x1 - x2;
	// Computing 2nd power 
		d__2 = y1 - y2;
	// Computing 2nd power 
		d__3 = z1 - z2;
		dist = sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
		dist *= re;
		distazi[i__] = dist; //distances in kms

		vbFunc((long)floor(i__ * 100/ (*numazi - 1)));
	}


	return 0;
}
*/

////////////////profile calculation requiring well formated x,y,z input data, i.e., x (cols) as outside variable, y (rows) as inside variable
int __declspec (dllexport) __stdcall Profiles (LPCSTR pszFile_In, LPCSTR pszFile_Out,
											   double *pnxo, double *pnyo, double *pnzo, 
											   long *nrows, long *ncols, short *CoordMode, short *HorizMode, short *Save_xyz,
											   double *pnva, double *pnazi, 
											   double *pnXazi, double *pnYazi, double *pnZazi, double *distazi,
											   long *numazi, double *StepSizeAzi, double *HalfAziRange, double *aprn,
											   long cbAddress) {

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

	FILE *stream;

	short ier = 0;

	//=========debugging===============
	//char buff[255] = ""; 
	//========================

	union profile
	{
		float ver[4];
	};

	union profile *prof;

	union dtm
	{
		double ver[4];
	};

	union dtm *dtms;

	long rows = *nrows; //number of unique kmy's
	long cols = *ncols; //number of unique kmx's
 
	//debugging=====================
	//buff[0] = 0; //clear buffer
	//sprintf(buff, "%s%s%, %s", "File_In, File_Out = ", pszFile_In, pszFile_Out);
	//MessageBox(NULL, (const char *)buff, "Debugging",
	//	   MB_OK | MB_ICONINFORMATION);
	/*
	buff[0] = 0; //clear buffer
	sprintf(buff, "%s%f, %f, %f, %d, %d, %d, %d, %d", "xo, yo, zo, rows, cols, CoordMode, HorizMode, Save_xyz = ", *pnxo, *pnyo, *pnzo, rows, cols, *CoordMode, *HorizMode, *Save_xyz);
	MessageBox(NULL, (const char *)buff, "Debugging",
		   MB_OK | MB_ICONINFORMATION);

	buff[0] = 0; //clear buffer
	sprintf(buff, "%s%d, %f, %f", "numazi, StepSizeAzi, HalfAziRange = ", *numazi, *StepSizeAzi, *HalfAziRange);
	MessageBox(NULL, (const char *)buff, "Debugging",
		   MB_OK | MB_ICONINFORMATION);
	//===================================
	*/

	if (rows <= 1 || cols <= 1) return -2; //nothing to evaluate

	if (stream = fopen(pszFile_In,"r"))
	{

		//allocate memory for prof array
		//free( prof ); //free the memory allocated before for prof, and reallocate
		prof = (union profile *) malloc((*numazi + 10) * sizeof(union profile));

		//allocate memory for dtms array
		//free( dtms ); //free the memory allocated before for dtms, and reallocate
		dtms = (union dtm *) malloc((*nrows + 10) * sizeof(union dtm));

		double kmxo, kmyo, hgt, lt, lg;
		kmxo = *pnxo;
		kmyo = *pnyo;
		hgt = *pnzo;

		double N = 0;
		double E = 0;
		double apprnr = 0;

		double maxang = 0;
		maxang = fabs(*HalfAziRange);

		//convert origin to WGS84

		if (*CoordMode == 1) //ITM
		{
			
			if (kmyo < 1000)
			{
				N = kmyo * 1000 + 1000000;
			}
			else
			{
				N = kmyo;
			}

			if (kmxo < 1000)
			{
				E = kmxo * 1000;
			}
			else
			{
				E = kmxo;
			}

			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt, &lg); 

			apprnr = *aprn;

		}
		else //coordinates already in geographic latitude and longitude
		{
			lt = kmyo;
			lg = kmxo;

			if (*aprn != 0.) {
				apprnr = *aprn * .0083333333 / cos(lt * cd);}

		}

		short minview = 10;
		double anav = 0;

		//initialize prof arry and atmospheric refraction constants

		if (hgt >= 0.) {
			anav = sqrt(hgt) * -.0348f;
		}
		if (hgt < 0.) {
			anav = -1.f;
		}

		long i__ = 0;

		//initialize the arrays

		for ( i__ = 0; i__ < *numazi; i__++) {
			prof[i__].ver[0] = (float)anav - minview;
			prof[i__].ver[1] = 0;
			prof[i__].ver[2] = 0;
			prof[i__].ver[3] = 0;
		}

		for (i__ = 0; i__ < *nrows + 1; i__++)
		{
			dtms[i__].ver[0] = -9999.;
			dtms[i__].ver[1] = 0;
			dtms[i__].ver[2] = 0;
			dtms[i__].ver[3] = 0;
		}

		double re, rekm, testazi;
		double hgt2, treehgt = 0, fudge = 0;
		double lg2, lt2, x1, x2, y1, y2, z1, z2, d__1, d__2, d__3;
		double distd, re1, re2, deltd, dist1, dist2, angle, viewang;
		double d__, x1d, y1d, z1d, azicos, x1s, y1s, z1s, x1p, y1p, z1p, azisin, azi;
		double defm, defb, avref, kmxp, kmyp, bne1, bne2,hgt_1, hgt_2;
		int istart1, ndela, ndelazi, ne, nstart2,i__4;
		double az2,vang2,kmy2,az1,vang1,kmy1,bne, vang, proftmp, delazi, dist;
		double invAzimuthStep = 10;

		bool found = false;

		double AzimuthStep = 0;
		AzimuthStep = *StepSizeAzi;
		invAzimuthStep = 1.0 / AzimuthStep;

		re = 6371315.; //radius of earth in meters
		rekm = 6371.315f; //radius of earth in km

		long numdtm = 0; //counter for number of points used in calculation

		double kmx = 0;
		double kmy = 0;

		lt2 = 0;
		lg2 = 0;
		hgt2 = 0;

		//the data is organized as an x,y,z file, increasing x, decreasing y, z where x is the outer loop.
		double kmxcurrent = 0;

		long numrctot = 0; //total number of columns processed

		// Call the UpdateStatus  through the function pointer.
		vbFunc(0);

		double startkmx = 0;
		double sofkmx = 0;
		double startkmy = 0;
		double sofkmy = 0;

		long numpoints = rows * cols;

		//debugging=====================
		//buff[0] = 0; //clear buffer
		//sprintf(buff, "%s%d", "numpoints = ", numpoints);
		//MessageBox(NULL, (const char *)buff, "Debugging",
		//	   MB_OK | MB_ICONINFORMATION);
		//===================================
		bool found_start = false;
		long num_skipped = 0;

		for (i__ = 0; i__ < numpoints; i__ ++)
		{
			//if Eastern horizon, use only points that longitude are to the East point
			//for Western horizon, use only points that are to the West of the center point

			fscanf(stream, "%lg,%lg,%lg\n", &kmx, &kmy, &hgt2);

			//debugging=====================
			//buff[0] = 0; //clear buffer
			//sprintf(buff, "%s%lg, %lg, %lg", "kmx, kmy, hgt2 = ", kmx, kmy, hgt2);
			//MessageBox(NULL, (const char *)buff, "Debugging",
			//	   MB_OK | MB_ICONINFORMATION);
			//===================================

			if (*HorizMode == 1) //Eastern horizon
			{
				if (kmx < kmxo)
				{
					++num_skipped;
					goto loop999; //point is west of center point, so ignore it
				}
				else if (kmx > kmxo + apprnr)
				{
					if (!found_start) //first kmx where calculations begin
					{
						found_start = true;
						kmxcurrent = kmx;
						startkmx = kmx;
						startkmy = kmy;
					}
					else //keep on looking for last kmx,kmy
					{
						sofkmx = kmx;
						sofkmy = kmy;
					}

				}
			}
			else //Western horizon
			{
				if (kmx > kmxo)
				{
					++num_skipped;
					goto loop999; //point is east of center point, so ignore it
				}
				else if (kmx < kmxo - apprnr)
				{
					if (i__ == 0) //center point is certainly west of the most easterly point, so this is start
					{
						kmxcurrent = kmx;
						startkmx = kmx;
						startkmy = kmy;
					}
					else //keep on looking for last kmx kmy
					{
						sofkmx = kmx;
						sofkmy = kmy;
					}
				}

			}


			if (kmx != kmxcurrent) //corresponds to start of new column in x
			{
				//analyse one column of azimuths

				//debugging=====================
				//buff[0] = 0; //clear 
				//sprintf(buff, "%s%d, %f", "numrctot, kmx finished, kmx = ", numrctot, kmx);
				//MessageBox(NULL, (const char *)buff, "Debugging",
				//	   MB_OK | MB_ICONINFORMATION);
				//===================================


				///////////////////////////////////////////////////////////////////////////////////////////////////////

		//      interpolate to a grid of 0.1 degrees in azimuth 
		//      the interpolated points reside in the array prof() 

				istart1 = 1;
		L150:
				if (*HorizMode == 1) {
					az2 = dtms[istart1].ver[0];
					vang2 = dtms[istart1].ver[1];
					kmy2 = dtms[istart1].ver[2];
					hgt_2 = dtms[istart1].ver[3];
				} else if (*HorizMode == 2){
					az1 = dtms[istart1].ver[0];
					vang1 = dtms[istart1].ver[1];
					kmy1 = dtms[istart1].ver[2];
					hgt_1 = dtms[istart1].ver[3];
				}
				nstart2 = istart1 + 1;
		L160:
				if (*HorizMode == 1) {
					az1 = dtms[nstart2].ver[0];
					vang1 = dtms[nstart2].ver[1];
					kmy1 = dtms[nstart2].ver[2];
					hgt_1 = dtms[nstart2].ver[3];
				} else if (*HorizMode == 2){
					az2 = dtms[nstart2].ver[0];
					vang2 = dtms[nstart2].ver[1];
					kmy2 = dtms[nstart2].ver[2];
					hgt_2 = dtms[nstart2].ver[3];
				}

				/*
				//debugging=====================
				buff[0] = 0; //clear buffer
				sprintf(buff, "%s%d, %f, %f, %f, %f, %f, %f, %f, %f", "numrctot, az1, az2, vang1, vang2, kmy1, kmy2, hgt_1, hgt_2 = ", numrctot, az1, az2, vang1, vang2, kmy1, kmy2, hgt_1, hgt_2);
				MessageBox(NULL, (const char *)buff, "Debugging",
					   MB_OK | MB_ICONINFORMATION);
				//===================================
				*/


				if (az2 >= -maxang && az1 <= maxang + AzimuthStep) {
					bne1 = Fix(az1 * invAzimuthStep) * AzimuthStep - AzimuthStep;
					bne2 = Fix(az2 * invAzimuthStep) * AzimuthStep + AzimuthStep;
					
					/*

					bne1 = (int)(az1 * invAzimuthStep) / invAzimuthStep - AzimuthStep;
					bne2 = (int)(az2 * invAzimuthStep) / invAzimuthStep + AzimuthStep;

					bne1 = (double)((int)(az1 * invAzimuthStep)); 
					bne2 = (double)((int)(az2 * invAzimuthStep));

					//debugging=====================
					buff[0] = 0; //clear buffer
					sprintf(buff, "%s%d, %f, %f, %f, %f, %f", "1st numrctot, az1, invAzimuthStep, bne1, az2, bne2 = ", numrctot, az1, invAzimuthStep, bne1, az2, bne2);
					MessageBox(NULL, (const char *)buff, "Debugging",
						   MB_OK | MB_ICONINFORMATION);
					//===================================

					bne1 /= (double)invAzimuthStep; 
					bne2 /= (double)invAzimuthStep; 

					//debugging=====================
					buff[0] = 0; //clear buffer
					sprintf(buff, "%s%d, %g, %g", "2nd numrctot, bne1, bne2 = ", numrctot, bne1, bne2);
					MessageBox(NULL, (const char *)buff, "Debugging",
						   MB_OK | MB_ICONINFORMATION);
					//===================================

					bne1 -= AzimuthStep;
					bne2 += AzimuthStep;

					//debugging=====================
					buff[0] = 0; //clear buffer
					sprintf(buff, "%s%d, %g, %g", "3rd numrctot, bne1, bne2 = ", numrctot, bne1, bne2);
					MessageBox(NULL, (const char *)buff, "Debugging",
						   MB_OK | MB_ICONINFORMATION);
					//===================================
					*/

					found = false;
					ndela = (int) ((bne2 - bne1) / AzimuthStep + 1);
					i__4 = ndela;
					for (ndelazi = 1; ndelazi <= i__4; ++ndelazi) {
						delazi = bne1 + (ndelazi - 1) * AzimuthStep;
						/*
						//debugging=====================
						buff[0] = 0; //clear buffer
						sprintf(buff, "%s%d, %d, %f", "numrctot, ndelazi, delazi = ", numrctot, ndelazi, delazi);
						MessageBox(NULL, (const char *)buff, "Debugging",
							   MB_OK | MB_ICONINFORMATION);
						//===================================
						*/
						if (delazi < -maxang) {
							goto L165;}
						if (delazi > maxang + AzimuthStep) {
							goto L170;}


						if (az1 <= delazi && az2 >= delazi) {
							bne = (delazi + maxang) * (double)invAzimuthStep + 1.f;
							ne = Nint(bne);
							if (ne >= 0 && ne <= *numazi) {
								vang = (vang2 - vang1) / (az2 - az1) * (delazi - az1) + vang1;
								proftmp = prof[ne].ver[0];
								/*
								//debugging=====================
								buff[0] = 0; //clear buffer
								sprintf(buff, "%s%d, %f, %f, %f, %f, %f, %f, %f, %f, %f", "numrctot, vang, vang2, vang1, az2, az1, delazi, az1, vang1, proftmp = ", numrctot, vang, vang2, vang1, az2, az1, delazi, az1, vang1, proftmp);
								MessageBox(NULL, (const char *)buff, "Debugging",
									   MB_OK | MB_ICONINFORMATION);
								//===================================
								*/
								if (vang > proftmp) {
									prof[ne].ver[0] = (float)vang;
									prof[ne].ver[1] = (float)(kmxcurrent);
									prof[ne].ver[2] = (float)((kmy2 - kmy1) / (az2 - az1) * (delazi - az1) + kmy1);
									prof[ne].ver[3] = (float)((hgt_2 - hgt_1) / (az2 - az1) * (delazi - az1) + hgt_1);
									/*
									//debugging=====================
									buff[0] = 0; //clear buffer
									sprintf(buff, "%s%d, %f, %f, %f, %f", "numrctot, vang, kmxcurrent, kmy, hgt = ", numrctot, prof[ne].ver[0], prof[ne].ver[1], prof[ne].ver[2], prof[ne].ver[3]);
									MessageBox(NULL, (const char *)buff, "Debugging",
										   MB_OK | MB_ICONINFORMATION);
									//===================================
									*/
									}
								found = true; }
						} else if (az1 < delazi && az2 < delazi) {
							goto L170;}
	L165:
						;
					}
	L170:
					if (found) {
						goto L180; }
					++nstart2;
					if (nstart2 <= numdtm) {
						goto L160;
					} else {
						goto L180; }
				}
	L180:
				++istart1;
				if (istart1 < numdtm) {
				goto L150;	}

				numrctot++;

				//debugging=====================
				//buff[0] = 0; //clear buffer
				//sprintf(buff, "%s%d", "numrctot = ", numrctot);
				//MessageBox(NULL, (const char *)buff, "Debugging",
				//	   MB_OK | MB_ICONINFORMATION);
				//===================================


				if (cols - num_skipped > 0)
				{
					vbFunc((long)floor(numrctot * 100/(cols - num_skipped)));
				}
				///////////////////////////////////////////////////////////////////////////////////////////////////////

				if (numrctot == cols) break; //finished evaluating all the kmx's

				numdtm = 0;
				kmxcurrent = kmx;
			}

			//now begin further optimizations (it is assumed that all data used for this program are relatively close by, and no need for distant-point optimizations)
			//use crude approximation of azimuth to
			//determine if data will lie within the required azimuth range
			if (kmx != kmxo)
			{
				testazi = (kmy - kmyo)/(kmx - kmxo);
				testazi = atan(testazi)/cd;
				if (testazi > maxang + 10.0 || testazi < -maxang - 10.0 ) goto loop999;
			}

			if (*CoordMode == 1) //ITM
			{

				//convert old ITM grid to WGS84

				if (kmy < 1000)
				{
					N = kmy * 1000 + 1000000;
				}
				else
				{
					N = kmy;
				}

				if (kmx < 1000)
				{
					E = kmx * 1000;
				}
				else
				{
					E = kmx;
				}

				//use J. Gray's conversion from ITM to WGS84
				//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
				ics2wgs84(N, E, &lt2, &lg2);

			}
			else //coordinates already in geographic latitude and longitude
			{
				lt2 = kmy;
				lg2 = kmx;
			}

	//      calculate the view angle and azimuth, without 
	//      atmospheric refraction for the selected observation point 
	//      at longitude kmx,latitude kmy, height hgt2 

			fudge = treehgt;

			//if (hgt2 == -9999. || hgt2 == -32768. || fabs(hgt2) > 8848.) {
			//hgt2 = 0.f;}

			hgt2 += fudge;
			x1 = cos(lt * cd) * cos(-lg * cd);
			x2 = cos(lt2 * cd) * cos(-lg2 * cd);
			y1 = cos(lt * cd) * sin(-lg * cd);
			y2 = cos(lt2 * cd) * sin(-lg2 * cd);
			z1 = sin(lt * cd);
			z2 = sin(lt2 * cd);
			// Computing 2nd power 
			d__1 = x1 - x2;
			// Computing 2nd power 
			d__2 = y1 - y2;
			// Computing 2nd power 
			d__3 = z1 - z2;
			distd = rekm * sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
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
	//          view angle in radians 
			viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
			d__ = (dist1 - dist2 * cos(angle)) / dist1;
			x1d = x1 * (1 - d__) - x2;
			y1d = y1 * (1 - d__) - y2;
			z1d = z1 * (1 - d__) - z2;
			x1p = -sin(-lg * cd);
			y1p = cos(-lg * cd);
			z1p = 0.;
			azicos = x1p * x1d + y1p * y1d;
			x1s = -cos(-lg * cd) * sin(lt * cd);
			y1s = -sin(-lg * cd) * sin(lt * cd);
			z1s = cos(lt * cd);
			azisin = x1s * x1d + y1s * y1d + z1s * z1d;
			azi = atan(azisin / azicos);
	//      azimuth in degrees 
			azi /= cd;
	//      add contribution of atmospheric refraction to view angle 
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

			numdtm++;

			if (numdtm > rows + 1) return -1; //the file is corrupted, exit before it corrupts the memory


			//debugging=====================
			//buff[0] = 0; //clear buffer
			//sprintf(buff, "%s%d, %f, %f, %f, %f", "1st numdtm, azi, va, kmy, hgt2 = ", numdtm, azi, viewang / cd + avref, kmy, hgt2);
			//MessageBox(NULL, (const char *)buff, "Debugging",
			//	   MB_OK | MB_ICONINFORMATION);
			//===================================

			dtms[numdtm].ver[0] = azi;
			dtms[numdtm].ver[1] = viewang / cd + avref;
			dtms[numdtm].ver[2] = kmy;
			dtms[numdtm].ver[3] = hgt2;


			//debugging=====================
			//buff[0] = 0; //clear buffer
			//sprintf(buff, "%s%d, %f, %f, %f, %f", "2nd numdtm, azi, va, kmy, hgt2 = ", numdtm, dtms[numdtm].ver[0], dtms[numdtm].ver[1], dtms[numdtm].ver[2], dtms[numdtm].ver[3]);
			//MessageBox(NULL, (const char *)buff, "Debugging",
			//	   MB_OK | MB_ICONINFORMATION);
			//===================================


	loop999: ;
		}

		fclose(stream);

		//finished analysing the data, now output the view angle, etc., vs. azimuth

		bool opened = false;
		bool Save_profile = false;

		if (*Save_xyz) {
			Save_profile = true; }

		if (Save_profile) //record profile file if this option is flagged
		{
			if( stream = fopen(pszFile_Out, "w"))
			{	
				opened = true;
				//print header lines
				fprintf(stream,"%s\n", "kmxo,kmyo,hgt,startkmx,sofkmx,dkmx,dkmy,APPRNR");
				fprintf(stream,"%f,%f,%f,%f,%f,%f,%f,%f\n", kmxo, kmyo, hgt, startkmx, sofkmx, (sofkmx - startkmx)/(cols - 1), (sofkmy - startkmy)/(rows - 1), apprnr );
			}
		}

		vbFunc(0);

		for (i__ = 0; i__ < *numazi; i__++)
		{
			pnva[i__] = prof[i__ + 1].ver[0];
			pnazi[i__] = -maxang + i__ * AzimuthStep;
			pnXazi[i__] = prof[i__ + 1].ver[1];
			pnYazi[i__] = prof[i__ + 1].ver[2];
			pnZazi[i__] = prof[i__ + 1].ver[3];

			//now calculate the distance to the obstruction from the calculation center
			kmxp = prof[i__ + 1].ver[1];
			kmyp = prof[i__ + 1].ver[2];
			
			if (*CoordMode == 1) //ITM
			{

				//convert old ITM grid to WGS84

				if (kmyp < 1000)
				{
					N = kmyp * 1000 + 1000000;
				}
				else
				{
					N = kmyp;
				}

				if (kmxp < 1000)
				{
					E = kmxp * 1000;
				}
				else
				{
					E = kmxp;
				}

				//use J. Gray's conversion from ITM to WGS84
				//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
				ics2wgs84(N, E, &lt2, &lg2);

			}
			else //coordinates already in geographic latitude and longitude
			{
				lt2 = kmyp;
				lg2 = kmxp;
			}


			x1 = cos(lt * cd) * cos(-lg * cd);
			x2 = cos(lt2 * cd) * cos(-lg2 * cd);
			y1 = cos(lt * cd) * sin(-lg * cd);
			y2 = cos(lt2 * cd) * sin(-lg2 * cd);
			z1 = sin(lt * cd);
			z2 = sin(lt2 * cd);
			re = 6371.315;
			re1 = 1.;
			re2 = 1.;
			x1 = re1 * x1;
			y1 = re1 * y1;
			z1 = re1 * z1;
			x2 = re2 * x2;
			y2 = re2 * y2;
			z2 = re2 * z2;
		// Computing 2nd power 
			d__1 = x1 - x2;
		// Computing 2nd power 
			d__2 = y1 - y2;
		// Computing 2nd power 
			d__3 = z1 - z2;
			dist = sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
			dist *= re;
			distazi[i__] = dist; //distances in kms

			if (opened) 
			{
				//write to profile file
				if (*CoordMode == 1) //ITM
				{
					//ITM coordinates
					fprintf(stream, "%5.1f, %8.3f, %10.3f, %11.3f, %8.3f, %8.2f\n", pnazi[i__], pnva[i__], pnXazi[i__], pnYazi[i__], distazi[i__], pnZazi[i__]);
				}
				else
				{
					//geo coordinates
					fprintf(stream, "%5.1f, %8.3f, %11.6f, %10.6f, %8.3f, %8.2f\n", pnazi[i__], pnva[i__], pnXazi[i__], pnYazi[i__], distazi[i__], pnZazi[i__]);
				}
			}

			vbFunc((long)floor(i__ * 100/ (*numazi - 1)));
		}

		if (opened)
		{
			fclose(stream);
			opened = false;
		}

		return 0;

	}
	else //topo_coord.xyz not found
	{
		return -1;
	}

}

////////////////Profiles2 --use for non rectangular grid typical of topo maps/////////////////////////

int __declspec (dllexport) __stdcall Profiles2 (LPCSTR pszFile_In, LPCSTR pszFile_Out,
											   double *pnxo, double *pnyo, double *pnzo, 
											   long *numXYZpoints, short *CoordMode, short *HorizMode, short *Save_xyz,
											   double *pnva, double *pnazi, 
											   double *pnXazi, double *pnYazi, double *pnZazi, double *distazi,
											   long *numazi, double *StepSizeAzi, double *HalfAziRange, double *aprn,
											   long cbAddress) {


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

	FILE *stream;

	short ier = 0;

	union profile
	{
		float ver[4];
	};

	union profile *prof;

	if (stream = fopen(pszFile_In,"r"))
	{

		//allocate memory for prof array
		//free( prof ); //free the memory allocated before for prof, and reallocate
		prof = (union profile *) malloc((*numazi + 10) * sizeof(union profile));

		double kmxo, kmyo, hgt, lt, lg;
		kmxo = *pnxo;
		kmyo = *pnyo;
		hgt = *pnzo;

		double N = 0;
		double E = 0;
		double apprnr = 0;

		double maxang = 0;
		maxang = fabs(*HalfAziRange);

		//int npp = *num_array; 

		//convert origin to WGS84

		if (*CoordMode == 1) //ITM
		{
			
			if (kmyo < 1000)
			{
				N = kmyo * 1000 + 1000000;
			}
			else
			{
				N = kmyo;
			}

			if (kmxo < 1000)
			{
				E = kmxo * 1000;
			}
			else
			{
				E = kmxo;
			}

			//use J. Gray's conversion from ITM to WGS84
			//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
			ics2wgs84(N, E, &lt, &lg); 

			apprnr = *aprn;
		}
		else //coordinates already in geographic latitude and longitude
		{
			lt = kmyo;
			lg = kmxo;

			if (*aprn != 0.) {
				apprnr = *aprn * .0083333333 / cos(lt * cd);
			}

		}

		short minview = 10;
		double anav = 0;

		//initialize prof arry and atmospheric refraction constants

		if (hgt >= 0.) {
			anav = sqrt(hgt) * -.0348f;
		}
		if (hgt < 0.) {
			anav = -1.f;
		}

		long i__ = 0;

		for ( i__ = 0; i__ < *numazi; ++i__) {
			prof[i__].ver[0] = (float)anav - minview;
			prof[i__].ver[1] = 0;
			prof[i__].ver[2] = 0;
			prof[i__].ver[3] = 0;
		}


		double re, rekm, testazi;
		double hgt2, treehgt = 0, fudge = 0;
		double lg2, lt2, x1, x2, y1, y2, z1, z2, d__1, d__2, d__3;
		double distd, re1, re2, deltd, dist1, dist2, angle, viewang;
		double d__, x1d, y1d, z1d, azicos, x1s, y1s, z1s, x1p, y1p, z1p, azisin, azi;
		double defm, defb, avref, kmxp, kmyp;
		int ne;
		double bne, vang, proftmp, dist;

		double invAzimuthStep = 10;

		bool found = false;

		double AzimuthStep = 0;
		AzimuthStep = *StepSizeAzi;
		invAzimuthStep = 1.0 / AzimuthStep;

		re = 6371315.; //radius of earth in meters
		rekm = 6371.315f; //radius of earth in km

		//long numdtm = 0; //counter for number of points used in calculation

		double kmx = 0;
		double kmy = 0;

		lt2 = 0;
		lg2 = 0;
		hgt2 = 0;

		//the data is organized as an x,y,z file, increasing x, decreasing y, z where x is the outer loop.
		double kmxcurrent = 0;

		// Call the UpdateStatus  through the function pointer.
		vbFunc(0);

		double startkmx = 0;
		double sofkmx = 0;
		double startkmy = 0;
		double sofkmy = 0;

		bool found_start = false;

		long numpoint = 0;
		numpoint = 0;

		long numTotpoints = 0;
		numTotpoints = *numXYZpoints;

		while (!feof(stream))
		{
			//if Eastern horizon, use only points that longitude are to the East point
			//for Western horizon, use only points that are to the West of the center point

			fscanf(stream, "%lg,%lg,%lg\n", &kmx, &kmy, &hgt2);

			if (*HorizMode == 1) //Eastern horizon
			{
				if (kmx < kmxo)
				{
					goto loop999; //point is west of center point, so ignore it
				}
				else if (kmx > kmxo + apprnr)
				{
					if (!found_start) //first kmx to start analysis
					{
						found_start = true;
						kmxcurrent = kmx;
						startkmx = kmx;
						startkmy = kmy;
					}
					else //keep on looking for last kmx,kmy
					{
						sofkmx = kmx;
						sofkmy = kmy;
					}

				}
			}
			else //Western horizon
			{
				if (kmx > kmxo)
				{
					goto loop999; //point is east of center point, so ignore it
				}
				else if (kmx < kmxo - apprnr)
				{
					if (numpoint == 0) //center point is certainly west of the most easterly point, so this is start
					{
						kmxcurrent = kmx;
						startkmx = kmx;
						startkmy = kmy;
					}
					else //keep on looking for last kmx kmy
					{
						sofkmx = kmx;
						sofkmy = kmy;
					}
				}

			}

			//now begin further optimizations (it is assumed that all data used for this program are relatively close by, and no need for distant-point optimizations)
			//use crude approximation of azimuth to
			//determine if data will lie within the required azimuth range
			if (kmx != kmxo)
			{
				testazi = (kmy - kmyo)/(kmx - kmxo);
				testazi = atan(testazi)/cd;
				if (testazi > maxang + 10.0 || testazi < -maxang - 10.0 ) goto loop999;
			}

			if (*CoordMode == 1) //ITM
			{

				//convert old ITM grid to WGS84

				if (kmy < 1000)
				{
					N = kmy * 1000 + 1000000;
				}
				else
				{
					N = kmy;
				}

				if (kmx < 1000)
				{
					E = kmx * 1000;
				}
				else
				{
					E = kmx;
				}

				//use J. Gray's conversion from ITM to WGS84
				//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
				ics2wgs84(N, E, &lt2, &lg2);
				//lg2 = - lg2; //our convention
			}
			else //coordinates already in geographic latitude and longitude
			{
				lt2 = kmy;
				lg2 = kmx;
			}

	//      calculate the view angle and azimuth, without 
	//      atmospheric refraction for the selected observation point 
	//      at longitude kmx,latitude kmy, height hgt2 

			fudge = treehgt;

			//if (hgt2 == -9999. || hgt2 == -32768. || fabs(hgt2) > 8848.) {
			//hgt2 = 0.f;}

			hgt2 += fudge;
			x1 = cos(lt * cd) * cos(-lg * cd);
			x2 = cos(lt2 * cd) * cos(-lg2 * cd);
			y1 = cos(lt * cd) * sin(-lg * cd);
			y2 = cos(lt2 * cd) * sin(-lg2 * cd);
			z1 = sin(lt * cd);
			z2 = sin(lt2 * cd);
			// Computing 2nd power 
			d__1 = x1 - x2;
			// Computing 2nd power 
			d__2 = y1 - y2;
			// Computing 2nd power 
			d__3 = z1 - z2;
			distd = rekm * sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
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
	//          view angle in radians 
			viewang = atan((-re1 + re2 * cos(angle)) / (re2 * sin(angle)));
			d__ = (dist1 - dist2 * cos(angle)) / dist1;
			x1d = x1 * (1 - d__) - x2;
			y1d = y1 * (1 - d__) - y2;
			z1d = z1 * (1 - d__) - z2;
			x1p = -sin(-lg * cd);
			y1p = cos(-lg * cd);
			z1p = 0.;
			azicos = x1p * x1d + y1p * y1d;
			x1s = -cos(-lg * cd) * sin(lt * cd);
			y1s = -sin(-lg * cd) * sin(lt * cd);
			z1s = cos(lt * cd);
			azisin = x1s * x1d + y1s * y1d + z1s * z1d;
			azi = atan(azisin / azicos);
	//      azimuth in degrees 
			azi /= cd;
	//      add contribution of atmospheric refraction to view angle 
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

			//now fit the result into the nearest azimuth slot
			bne = (azi + maxang) * (double)invAzimuthStep + 1.f;
			ne = Nint(bne);
			if (ne >= 0 && ne <= *numazi) {
				vang = viewang/cd + avref;
				proftmp = prof[ne].ver[0];
				if (vang > proftmp) {
					prof[ne].ver[0] = (float)vang;
					prof[ne].ver[1] = (float)kmx;
					prof[ne].ver[2] = (float)kmy;
					prof[ne].ver[3] = (float)hgt2;
					}
			}

			numpoint++;
			vbFunc((long)floor(numpoint * 100/ numTotpoints));

	loop999: ;
		}

		fclose(stream);
		bool Save_profile = false;

		if (*Save_xyz) {
			Save_profile = true; }

		//finished analysing the data, now output the view angle, etc., vs. azimuth

		bool opened = false;

		if (Save_profile) //record profile file if this option is flagged
		{
			if( stream = fopen(pszFile_Out, "w"))
			{	
				opened = true;
				//print header lines
				fprintf(stream,"%s\n", "kmxo,kmyo,hgt,startkmx,sofkmx,kmx range,kmy range,APPRNR");
				fprintf(stream,"%f,%f,%f,%f,%f,%f,%f,%f\n", kmxo, kmyo, hgt, startkmx, sofkmx, sofkmx - startkmx, sofkmy - startkmy, apprnr );
			}
		}

		vbFunc(0);

		for (i__ = 0; i__ < *numazi; i__++)
		{
			pnva[i__] = prof[i__].ver[0];
			pnazi[i__] = -maxang + i__ * AzimuthStep;

			if (pnva[i__] - (anav - minview) < 0.001 && fabs(pnazi[i__]) < 10.0)
			{
				pnva[i__] = 0.5 * (prof[i__ - 1].ver[0] + prof[i__ + 1].ver[0]);
				pnXazi[i__] = 0.5 * (prof[i__ - 1].ver[1] + prof[i__ + 1].ver[1]);
				pnYazi[i__] = 0.5 * (prof[i__ - 1].ver[2] + prof[i__ + 1].ver[2]);
				pnZazi[i__] = 0.5 * (prof[i__ - 1].ver[3] + prof[i__ + 1].ver[3]);
			}
			else
			{
				pnXazi[i__] = prof[i__].ver[1];
				pnYazi[i__] = prof[i__].ver[2];
				pnZazi[i__] = prof[i__].ver[3];
			}

			//now calculate the distance to the obstruction from the calculation center
			kmxp = prof[i__].ver[1];
			kmyp = prof[i__].ver[2];
			
			if (*CoordMode == 1) //ITM
			{

				//convert old ITM grid to WGS84

				if (kmyp < 1000)
				{
					N = kmyp * 1000 + 1000000;
				}
				else
				{
					N = kmyp;
				}

				if (kmxp < 1000)
				{
					E = kmxp * 1000;
				}
				else
				{
					E = kmxp;
				}

				//use J. Gray's conversion from ITM to WGS84
				//source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
				ics2wgs84(N, E, &lt2, &lg2);
				//lg2 = - lg2; //our convention
			}
			else //coordinates already in geographic latitude and longitude
			{
				lt2 = kmyp;
				lg2 = kmxp;
			}


			x1 = cos(lt * cd) * cos(-lg * cd);
			x2 = cos(lt2 * cd) * cos(-lg2 * cd);
			y1 = cos(lt * cd) * sin(-lg * cd);
			y2 = cos(lt2 * cd) * sin(-lg2 * cd);
			z1 = sin(lt * cd);
			z2 = sin(lt2 * cd);
			re = 6371.315;
			re1 = 1.;
			re2 = 1.;
			x1 = re1 * x1;
			y1 = re1 * y1;
			z1 = re1 * z1;
			x2 = re2 * x2;
			y2 = re2 * y2;
			z2 = re2 * z2;
		// Computing 2nd power 
			d__1 = x1 - x2;
		// Computing 2nd power 
			d__2 = y1 - y2;
		// Computing 2nd power 
			d__3 = z1 - z2;
			dist = sqrt(d__1 * d__1 + d__2 * d__2 + d__3 * d__3);
			dist *= re;
			distazi[i__] = dist; //distances in kms

			if (opened) 
			{
				//write to profile file
				if (*CoordMode == 1) //ITM
				{
					//ITM coordinates
					fprintf(stream, "%5.1f, %8.3f, %10.3f, %11.3f, %8.3f, %8.2f\n", pnazi[i__], pnva[i__], pnXazi[i__], pnYazi[i__], distazi[i__], pnZazi[i__]);
				}
				else
				{
					//geo coordinates
					fprintf(stream, "%5.1f, %8.3f, %11.6f, %10.6f, %8.3f, %8.2f\n", pnazi[i__], pnva[i__], pnXazi[i__], pnYazi[i__], distazi[i__], pnZazi[i__]);
				}
			}

			vbFunc((long)floor(i__ * 100/ (*numazi - 1)));
		}

		if (opened)
		{
			fclose(stream);
			opened = false;
		}

		return 0;

	}
	else //topo_coord.xyz not found
	{
		return -1;
	}

}

/********************Conversion Functions from ITM to WGS84***********************/
//=================================================
// Israel New Grid (ITM) to WGS84 conversion
//=================================================
void itm2wgs84(double N,double E,double& lat,double &lon)
{
	// 1. Local Grid (ITM) -> GRS80
	double lat80,lon80;
	Grid2LatLon(N,E,lat80,lon80,gITM,eGRS80);

	// 2. Molodensky GRS80->WGS84
	double lat84,lon84;
	Molodensky(lat80,lon80,lat84,lon84,eGRS80,eWGS84);

	// final results
	lat = lat84*180/pi;
	lon = lon84*180/pi;
}

//=================================================
// WGS84 to Israel New Grid (ITM) conversion
//=================================================
void wgs842itm(double lat,double lon,double& N,double& E)
{
	double latr = lat*pi/180;
	double lonr = lon*pi/180;

	// 1. Molodensky WGS84 -> GRS80
	double lat80,lon80;
	Molodensky(latr,lonr,lat80,lon80,eWGS84,eGRS80);

	// 2. Lat/Lon (GRS80) -> Local Grid (ITM)
	LatLon2Grid(lat80,lon80,N,E,eGRS80,gITM);
}

//=================================================
// Israel Old Grid (ICS) to WGS84 conversion
//=================================================
void ics2wgs84(double N,double E,double *lat,double *lon)
{
	// 1. Local Grid (ICS) -> Clark_1880_modified
	double lat80,lon80;
	//printf("inside N, E = %d, %d", N, E);
	//pause();
	Grid2LatLon(N,E,lat80,lon80,gICS,eCLARK80M);
	//printf("grid lat, lon = %f, %f", lat80*180/pi, lon80*180/pi);
	//pause();

	// 2. Molodensky Clark_1880_modified -> WGS84
	double lat84,lon84;
	Molodensky(lat80,lon80,lat84,lon84,eCLARK80M,eWGS84);

	// final results
	*lat = lat84*180/pi;
	*lon = lon84*180/pi;
	//printf("final lat, lon = %f, %f", lat, lon);
	//pause();


	//usage
	//N = int)(*kmyo * 1000 + 1000000);
	//E = (int)(*kmxo * 1000);
	////use J. Gray's conversion from ITM to WGS84
	////source: http://www.technion.ac.il/~zvikabh/software/ITM/isr84lib.cpp
	//ics2wgs84(N, E, lt, lg); 
	//*lg = - *lg; //convert to our sign convention where East longitude is negative
   
}

//=================================================
// WGS84 to Israel Old Grid (ICS) conversion
//=================================================
void wgs842ics(double lat,double lon,double& N,double& E)
{
	double latr = lat*pi/180;
	double lonr = lon*pi/180;

	// 1. Molodensky WGS84 -> Clark_1880_modified
	double lat80,lon80;
	Molodensky(latr,lonr,lat80,lon80,eWGS84,eCLARK80M);

	// 2. Lat/Lon (Clark_1880_modified) -> Local Grid (ICS)
	LatLon2Grid(lat80,lon80,N,E,eCLARK80M,gICS);
}

//====================================
// Local Grid to Lat/Lon conversion
//====================================
void Grid2LatLon(double N,double E,double& lat,double& lon,gGrid from,eDatum to)
{
	//================
	// GRID -> Lat/Lon
	//================
	//printf("gridfrom.falsen = %f", Grid[from].false_n);
	//printf("datumto.a = %f", Datum[to].a);
	//pause();

	double y = N + Grid[from].false_n;
	double x = E - Grid[from].false_e;
	double M = y / Grid[from].k0;

	double a = Datum[to].a;
	double b = Datum[to].b;
	double e = Datum[to].e;
	double esq = Datum[to].esq;

	double mu = M / (a*(1 - e*e/4 - 3*pow(e,4)/64 - 5*pow(e,6)/256));

	double ee = sqrt(1-esq);
	double e1 = (1-ee)/(1+ee);
	double j1 = 3*e1/2 - 27*e1*e1*e1/32;
	double j2 = 21*e1*e1/16 - 55*e1*e1*e1*e1/32;
	double j3 = 151*e1*e1*e1/96;
	double j4 = 1097*e1*e1*e1*e1/512;

	// Footprint Latitude
	double fp =  mu + j1*sin(2*mu) + j2*sin(4*mu) + j3*sin(6*mu) + j4*sin(8*mu);

	double sinfp = sin(fp);
	double cosfp = cos(fp);
	double tanfp = sinfp/cosfp;
	double eg = (e*a/b);
	double eg2 = eg*eg;
	double C1 = eg2*cosfp*cosfp;
	double T1 = tanfp*tanfp;
	double R1 = a*(1-e*e) / pow(1-(e*sinfp)*(e*sinfp),1.5);
	double N1 = a / sqrt(1-(e*sinfp)*(e*sinfp));
	double D = x / (N1*Grid[from].k0);

	double Q1 = N1*tanfp/R1;
	double Q2 = D*D/2;
	double Q3 = (5 + 3*T1 + 10*C1 - 4*C1*C1 - 9*eg2*eg2)*(D*D*D*D)/24;
	double Q4 = (61 + 90*T1 + 298*C1 + 45*T1*T1 - 3*C1*C1 - 252*eg2*eg2)*(D*D*D*D*D*D)/720;
	// result lat
	lat = fp - Q1*(Q2-Q3+Q4);

	double Q5 = D;
	double Q6 = (1 + 2*T1 + C1)*(D*D*D)/6;
	double Q7 = (5 - 2*C1 + 28*T1 - 3*C1*C1 + 8*eg2*eg2 + 24*T1*T1)*(D*D*D*D*D)/120;
	// result lon
	lon = Grid[from].lon0 + (Q5 - Q6 + Q7)/cosfp;
}

//====================================
// Lat/Lon to Local Grid conversion
//====================================
void LatLon2Grid(double lat,double lon,double& N,double& E,eDatum from,gGrid to)
{
	// Datum data for Lat/Lon to TM conversion
	double a = Datum[from].a;
	double e = Datum[from].e; 	// sqrt(esq);
	double b = Datum[from].b;

	//===============
	// Lat/Lon -> TM
	//===============
	double slat1 = sin(lat);
	double clat1 = cos(lat);
	double clat1sq = clat1*clat1;
	double tanlat1sq = slat1*slat1 / clat1sq;
	double e2 = e*e;
	double e4 = e2*e2;
	double e6 = e4*e2;
	double eg = (e*a/b);
	double eg2 = eg*eg;

	double l1 = 1 - e2/4 - 3*e4/64 - 5*e6/256;
	double l2 = 3*e2/8 + 3*e4/32 + 45*e6/1024;
	double l3 = 15*e4/256 + 45*e6/1024;
	double l4 = 35*e6/3072;
	double M = a*(l1*lat - l2*sin(2*lat) + l3*sin(4*lat) - l4*sin(6*lat));
	//double rho = a*(1-e2) / pow((1-(e*slat1)*(e*slat1)),1.5);
	double nu = a / sqrt(1-(e*slat1)*(e*slat1));
	double p = lon - Grid[to].lon0;
	double k0 = Grid[to].k0;
	// y = northing = K1 + K2p2 + K3p4, where
	double K1 = M*k0;
	double K2 = k0*nu*slat1*clat1/2;
	double K3 = (k0*nu*slat1*clat1*clat1sq/24)*(5 - tanlat1sq + 9*eg2*clat1sq + 4*eg2*eg2*clat1sq*clat1sq);
	// ING north
	double Y = K1 + K2*p*p + K3*p*p*p*p - Grid[to].false_n;

	// x = easting = K4p + K5p3, where
	double K4 = k0*nu*clat1;
	double K5 = (k0*nu*clat1*clat1sq/6)*(1 - tanlat1sq + eg2*clat1*clat1);
	// ING east
	double X = K4*p + K5*p*p*p + Grid[to].false_e;

	// final rounded results
	E = (int)(X+0.5);
	N = (int)(Y+0.5);
}

//======================================================
// Abridged Molodensky transformation between 2 datums
//======================================================
void Molodensky(double ilat,double ilon,double& olat,double& olon,eDatum from,eDatum to)
{
	// from->WGS84 - to->WGS84 = from->WGS84 + WGS84->to = from->to
	double dX = Datum[from].dX - Datum[to].dX;
	double dY = Datum[from].dY - Datum[to].dY;
	double dZ = Datum[from].dZ - Datum[to].dZ;

	double slat = sin(ilat);
	double clat = cos(ilat);
	double slon = sin(ilon);
	double clon = cos(ilon);
	double ssqlat = slat*slat;

	//dlat = ((-dx * slat * clon - dy * slat * slon + dz * clat)
	//        + (da * rn * from_esq * slat * clat / from_a)
	//        + (df * (rm * adb + rn / adb )* slat * clat))
	//       / (rm + from.h);

	double from_f = Datum[from].f;
	double df = Datum[to].f - from_f;
	double from_a = Datum[from].a;
	double da = Datum[to].a - from_a;
	double from_esq = Datum[from].esq;
	double adb = 1.0 / (1.0 - from_f);
	double rn = from_a / sqrt(1 - from_esq * ssqlat);
	double rm = from_a * (1 - from_esq) / pow((1 - from_esq * ssqlat),1.5);
	double from_h = 0.0; // we're flat!

	double dlat = (-dX*slat*clon - dY*slat*slon + dZ*clat
				   + da*rn*from_esq*slat*clat/from_a +
				   + df*(rm*adb + rn/adb)*slat*clat) / (rm+from_h);

	// result lat (radians)
	olat = ilat+dlat;

	// dlon = (-dx * slon + dy * clon) / ((rn + from.h) * clat);
	double dlon = (-dX*slon + dY*clon) / ((rn+from_h)*clat);
	// result lon (radians)
	olon = ilon+dlon;
}

/*********************************************************************************/
/////////////////////////////////////////////
//	  emulation of Fortran 95 Nint(double x)
/////////////////////////////////////////////////////////////
//source: http://www.lahey.com/docs/lfprohelp/F95ARNINTFn.htm
////////////////////////////////////////////////////////////
int Nint(double x)
{
	if (x > 0)
	{
		return (int)(x + 0.5);
	}
	else if (x <= 0)
	{
		return (int)(x - 0.5);
	}

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
