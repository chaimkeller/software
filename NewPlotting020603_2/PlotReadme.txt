Plot, Version 2.0

----------------------------------------------------
                     MENU ITEMS
-----------------------------------------------------

Open:  Opens common dialog box for selecting file(s) for plotting.
       Allows for multiselection

File:

   Open:  same as Open command above

   Print: opens print preview of plotted files. 
          Use the print button to send the graph to
          the printer for printing.  Requires that
          a printer has been installed.

   Save:  save the current file list for later use

   Restore:  restore the last recorded file list

   Paths:  path to WordPad.exe

   Exit: exit

Format:  Define and store the plot file formats.
         The files must be ascii with any number of
         header lines and data columns and rows.
         The data columns can either be comma delimited or
         space delimited.  The last file types in the buffer
         are reserved for JKH Bathymetric file formats.

Help:  help

------------------------------------------------------------
                       PROGRAM BUTTONS, ETC. 
-------------------------------------------------------------
Plot File Buffer:  The buffer of files that you can plot.  You
              activate any of the files in the list for plotting
              and editing by selecting (highlighting) them.

Plot Wizard:  opens the plot wizard that guides you through
              creating a plot of your files.  The plot wizard
              will not start unless you have highlighted which
              files in the Plot File Buffer you wish to plot.
              You will be requested to enter the plot format number
              of the file (as saved using the Format menu--see above),
              as well as other plotting information.  You can set the
              plot type and color, or leave it on an automatic setting.
 
              For JKH Bathymetric data.  Click on the checkbox.  The data
              is then converted to a SV and Temp data. The conversion
              information is displayed in the list box. To create plot
              files press the "Convert" button.  This 
              generates three files that are added to the plot buffer: 
              (1)depth vs. Temperature = nameTemp.tmp
              (2)depth vs. SV = nameSV.tmp
              (3)depth vs. SVC = nameSVC.tmp
              You can dump 1 or 2 data to the "rel" file by clicking on the
              proper radio button   
 
       
Plot All:  plot all the files in the Plot File Buffer.  You will be
           asked for the format number of all the files.  I.e., to
           use this command, all the selected files must have the
           same format.

Clear:  Clear the plot buffer and erase the last graph (if present)

Show/Edit:  Opens each of the highlighted files in the Plot File Buffer
            using Windows WordPad.    

--------------------------------------------------------
                LAYOUT
--------------------------------------------------------

Show Origin?  To always show the origin no matter where the info is.

Gridlines? To show or hide grid lines     

Index X values:  The plot data has an index from 1 to the number of plot points.
                 Limit the plot to the desired range of points by defining the
                 beginning and ending index values.

X,Y ranges:  Graph axes limits

Labels:  Axes labels

------------------------------------------------
             PLOTTING
-----------------------------------------------

Tips:  Drag WITHIN THE PLOTTING REGION to zoom in.
       Double Click anywhere on the graph to restore the
       original limits.  
       Right click to obtain a readout of the coordinates at
       the click point

--------------------------------------------

This program is freeware.  Enjoy!