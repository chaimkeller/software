VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegressionObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MaxO& = 25
Private GlobalO& '"Ordnung" = degree of the polynom expected
Private Finished As Boolean

Private SumX#(0 To 2 * MaxO)
Private SumYX#(0 To MaxO)
Private M#(0 To MaxO, 0 To MaxO + 1)
Private C#(0 To MaxO) 'coefficients in: Y = C(0)*X^0 + C(1)*X^1 + C(2)*X^2 + ...

Private Sub GaussSolve(O&)
'gauss algorithm implementation,
'following R.Sedgewick's "Algorithms in C", Addison-Wesley, with minor modifications
Dim i&, j&, k&, iMax&, T#, O1#
  O1 = O + 1
  'first triangulize the matrix
  For i = 0 To O
    iMax = i: T = Abs(M(iMax, i))
    For j = i + 1 To O 'find the line with the largest absvalue in this row
      If T < Abs(M(j, i)) Then iMax = j: T = Abs(M(iMax, i))
    Next j
    If i < iMax Then 'exchange the two lines
      For k = i To O1
        T = M(i, k)
            M(i, k) = M(iMax, k)
                      M(iMax, k) = T
      Next k
    End If
    For j = i + 1 To O 'scale all following lines to have a leading zero
      T = M(j, i) / M(i, i)
      M(j, i) = 0#
      For k = i + 1 To O1
        M(j, k) = M(j, k) - M(i, k) * T
      Next k
    Next j
  Next i
  'then substitute the coefficients
  For j = O To 0 Step -1
    T = M(j, O1)
    For k = j + 1 To O
      T = T - M(j, k) * C(k)
    Next k
    C(j) = T / M(j, j)
  Next j
  Finished = True
End Sub

Private Sub BuildMatrix(O&)
Dim i&, k&, O1&
  O1 = O + 1
  For i = 0 To O
    For k = 0 To O
      M(i, k) = SumX(i + k)
    Next k
    M(i, O1) = SumYX(i)
  Next i
End Sub

Private Sub FinalizeMatrix(O&)
Dim i&, O1&
  O1 = O + 1
  For i = 0 To O
    M(i, O1) = SumYX(i)
  Next i
End Sub

Private Sub Solve()
Dim O&
  O = GlobalO
  If XYCount <= O Then O = XYCount - 1
  If O < 0 Then Exit Sub
  BuildMatrix O
  On Error Resume Next
    GaussSolve (O)
    While (Err.Number <> 0) And (1 < O)
      Err.Clear
      C(0) = 0#
      O = O - 1
      FinalizeMatrix (O)
    Wend
  On Error GoTo 0
End Sub

Private Sub Class_Initialize()
  Init
  GlobalO = 2
End Sub

Public Sub Init()
Dim i&
  Finished = False
  For i = 0 To MaxO
    SumX(i) = 0#
    SumX(i + MaxO) = 0#
    SumYX(i) = 0#
    C(i) = 0#
  Next i
End Sub

Public Property Get Coeff#(Exponent&)
Dim Ex&, O&
  If Not Finished Then Solve
  Ex = Abs(Exponent)
  O = GlobalO
  If XYCount <= O Then O = XYCount - 1
  If O < Ex Then Coeff = 0# Else Coeff = C(Ex)
End Property

Public Property Get Degree&()
  Degree = GlobalO
End Property
Public Property Let Degree(NewVal&)
  If NewVal < 0 Or MaxO < NewVal Then
    Err.Raise 6000, "RegressionObject", NewVal & " is an invalid property value! Use 0<= Degree <= " & MaxO
    Exit Property
  End If
  Init
  GlobalO = NewVal
End Property

Public Property Get XYCount&()
  XYCount = CLng(SumX(0))
End Property

Public Function XYAdd(ByVal NewX#, ByVal NewY#)
Dim i&, j&, TX#, Max2O&
  Finished = False
  Max2O = 2 * GlobalO
  TX = 1#
  SumX(0) = SumX(0) + 1
  SumYX(0) = SumYX(0) + NewY
  For i = 1 To GlobalO
    TX = TX * NewX
    SumX(i) = SumX(i) + TX
    SumYX(i) = SumYX(i) + NewY * TX
  Next i
  For i = GlobalO + 1 To Max2O
    TX = TX * NewX
    SumX(i) = SumX(i) + TX
  Next i
End Function

Public Function RegVal#(X#)
Dim i&, O&
  If Not Finished Then Solve
  RegVal = 0#
  O = GlobalO
  If XYCount <= O Then O = XYCount - 1
  For i = 0 To O
    RegVal = RegVal + C(i) * X ^ i
  Next i
End Function
