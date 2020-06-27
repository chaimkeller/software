Plotting numerical data arrays
==============================


Plot.bas can be used to plot numerical data. Features for plotting are:
- cartesian axes (x and y-axes are not fixed)
- include to show origin yes or no
- include to show gridlines
- add titles for x- and y-axis
- plotting types are 'point', 'connected line' and 'bar'
- different traces can be plot in the same drawing pane
- 10 different colors to plot the traces
- scientific notation is used when differences in X and Y are bigger then 100000 or smaller than 0.00001.
- scientific notation will be shown under an angle when more than 10 lables needs to be shown on the x-axis (to avoid overlapping lables)

Features from the form showing the plot:
- SHIFT + left mouse gives zoom function
- right mouse click gives value of (X,Y)
- double left click draws the starting values

Three demo traces are added, they can be shown with the features as described above.



The syntax to call the plot routine is:

Plot <form>,<data array>, <layout>

The <data array> should be declared in a module.
<layout> is an user defined type GRAPHIC_LAYOUT:

Public Type GRAPHIC_LAYOUT
  XTitle As String 'title X-axis
  Ytitle As String 'title Y-axis
  blnOrigin As Boolean 'origin is included for only pos/neg values when true
  blnGridLine As Boolean 'Gridlines are shown when true
  lStart As Long 'index of start x-Range
  lEnd As Long 'index of end x-Range
  asX As Double 'trace in array to function as "X-value"
  asY() As Variant 'Y-traces to plot
  DrawTrace() As DRAWN_AS
  X0 As Double 'minimum value of domain X-values to draw
  X1 As Double 'maximum value of domain X-values to draw
  Y0 As Double 'minimum value of domain Y-values to draw
  Y1 As Double 'maximum value of domain Y-values to draw
End Type


Peter Wester
wester@kpd.nl
29-01-02
VB6