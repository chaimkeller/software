# software
Software I have written
Folders are projects I have written for zemanim program in either c or VB6

Programs written in VB6 ---------------------------------
1. NewPlotting - plotting program written in VB6 
very user friendly, has also fitting routines for polynomial and spline fits
2. CalProgram - program for calculating zemanim -- needs netzski6.exe module to calculate zemanim as well as the city data
also must have access to the WorldClim 2 temperature model files converted to bil format
3. MapDigitizer - program for digitizing and handling topo maps - full of useful, unique, and nifty routines including
following contours, storing images for realtime editing, contouring maps, determing visible horizon of area on map
Do work properly needs access to a DTM database, typically the JKH 25 m DTM of Israel and the NED 30 m DEM in hgt format
4. Maps & More - IDE used for map interface for visible zemanim calcuilations -- requires uses of c modules = rdhalba4, readDTM
also needs bmp map tiles for EY, and any other map image supported by VB6 for other places, also requires access to DTMS, both
the JKH's 25 m DTM of EY, and the USGEO 30 m DEM of the world in hgt (SRTM hgt file) format.
Also needs access to the WorldClim 2 temperature model files converted to bil format
5. Atmospheric-Refraction-RayTracing - calculates atmospheric raytracing using many methods, i.e., Bruton, Menat, and VDW
allows for using different atmospheres and also custom atmospheres.  Some calculations facilitated using dll writing in c
(AtmRef- todo- extend to any atmosphere that the VB6 code can handle)

Progams written in c ------------------------

1. netzski6_c - porgram for calcuating zemanim, written in c
2. rdhalba4 -program for calculating profiles from JKH's 25 m DTM for Israel
3. readDTM - program for calculating profiles for SRTM and DEM DTMs (SRTM hgt format)
4. StringRoutines - programs used for calculating the ChaiTables

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
<<<<<<<<<<<<<<<<<For any missing files or data, contact the author of this repository at his contact email >>>>>>>>>>>>>>>>>>>>
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////