Declare in module level:

Public Declare Function RayTracing Lib "AtmRef.dll" (StarAng As Double, EndAng As Double, StepAng As Double, NAngles As Long, _
                                                     HOBS As Double, TGROUND As Double, HMAXT As Double, ByVal File_Path As String, StepSize As Integer, _
                                                     GPress As Double, WAVELN As Double, HUMID As Double, OBSLAT As Double, NSTEPS As Long, _
                                                     ByVal pFunc As Long) As Long

'===============================================================================
Calling the dll

	StartAng = beginning apparent view angle (arcminutes)
	EndAng =  end apparent view angle (arcminutes)
	StepAng = step in apparent view angle (arcminutes)
	NAngles = number of angle steps = (EndAng - StartAng)/StepAng + 1
	HOBS = observer's height
	TLoop = ground temperature (Kelvin)
	HMAXT = same as in your program
	App.Path = foler path (e.g., c:\refracion) to write the rewselting file that has the following information:
		1. Ray's path length along Earth's circumference (meters)
		2. Height of ray above the Earth's surface (meters)
		3. Apparent View Angle at the observer (arcminutes)
		4. Beta (as in  your program), i.e., the completement of the zenith angle (mrads)
		5. Refrac = cumulative atmospheric refraction at the current ray position (mrads)
	StepSize = print out the above information only every StepSize of the loop increment
	Press0 = Ground pressure (mbar)
	HUMID = percent humidity
	OBSLAT = observer's latitude (degrees)
	NSTEPS = as you define it
	MyCallback - a callback function that can be used to run a progress bar (see  example code below)

            ier = RayTracing(StartAng, EndAng, StepAng, NAngles, _
                             HOBS, TLoop, HMAXT, App.Path, StepSize, _
                             Press0, WAVELN, HUMID, OBSLAT, NSTEPS, _
                             AddressOf MyCallback)

===============MyCallback example=================================================
'define in modular level
Public Sub MyCallback(ByVal parm As Long)

   On Error GoTo MyCallback_Error

   Call UpdateStatus(prjAtmRefMainfm, prjAtmRefMainfm.picProgBar, 1, parm)
   
   DoEvents

   On Error GoTo 0
   Exit Sub

MyCallback_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MyCallback of Module modHardy"
End Sub