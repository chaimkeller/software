Attribute VB_Name = "modPolyFit"
'Option Explicit
'code based on the most part on: https://www.vb-helper.com/howto_net_polynomial_least_squares.html
Public HasSolution As Boolean
Public PtX As Collection
Public PtY As Collection
Public BestCoeffs As Collection
Public FitPlotType%, FitPlotColor%, NumFitPoints%, MaxPolyDeg As Integer
Public SelectedFileNum() As Integer, FileOffset As Integer, flxFileBuffer() As String
Public notAlreadyFitted As Boolean
Public FitFileName As String
Public CurrentFitFileIndex As Integer
Public OriginalNumPlotFiles As Integer
Public MaxX As Double, MinX As Double, MaxY As Double, MinY As Double
Public SplineType%, FitMethod%, SplineDeg%, FileAddSave() As String, NumAddSave As Integer


' The function.
Public Function f(ByVal coeffs As Collection, ByVal X As Double) As Double
Dim total As Double
Dim x_factor As Double
Dim I As Integer

    total = 0
    x_factor = 1
    For I = 1 To coeffs.Count
        total = total + x_factor * coeffs(I)
        x_factor = x_factor * X
    Next I
    f = total
End Function

' Return the error squared.
Public Function ErrorSquared(ByVal PtX As Collection, ByVal PtY As Collection, ByVal coeffs As Collection) As Double
Dim total As Double
Dim pt As Integer
Dim dy As Double

    total = 0
    For pt = 1 To PtX.Count
        dy = PtY.item(pt) - f(coeffs, PtX.item(pt))
        total = total + dy * dy
    Next pt
    ErrorSquared = total
End Function

' Find the least squares linear fit.
Public Function FindPolynomialLeastSquaresFit(ByVal PtX As Collection, ByVal PtY As Collection, ByVal degree As Integer) As Collection
Dim J As Integer
Dim pt As Integer
Dim a_sub  As Integer
Dim coeff As Variant

' Allocate space for (degree + 1) equations with
' (degree + 2) terms each (including the constant term).
Dim coeffs() As Double
Dim answer() As Double

    ReDim coeffs(degree, degree + 1)

    ' Calculate the coefficients for the equations.
    For J = 0 To degree
        ' Calculate the coefficients for the jth equation.

        ' Calculate the constant term for this equation.
        coeffs(J, degree + 1) = 0
        For pt = 1 To PtX.Count
            coeffs(J, degree + 1) = coeffs(J, degree + 1) - (PtX.item(pt) ^ J) * PtY.item(pt)
        Next pt

        ' Calculate the other coefficients.
        For a_sub = 0 To degree
            ' Calculate the dth coefficient.
            coeffs(J, a_sub) = 0
            For pt = 1 To PtX.Count
                coeffs(J, a_sub) = coeffs(J, a_sub) - PtX.item(pt) ^ (a_sub + J)
            Next pt
        Next a_sub
    Next J

    ' Solve the equations.
    answer = GaussianElimination(coeffs)

    ' Return the result converted into a Collection.
    Set FindPolynomialLeastSquaresFit = New Collection
    For Each coeff In answer
        FindPolynomialLeastSquaresFit.Add CDbl(coeff)
    Next coeff
End Function

' Perform Gaussian elimination on these coefficients.
' Return the array of values that gives the solution.
Private Function GaussianElimination(coeffs() As Double) As Double()
Dim max_equation As Integer
Dim max_coeff As Integer
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim temp As Double
Dim coeff_i_i As Double
Dim coef_j_i As Double
Dim d As Integer
Dim solution() As Double

   On Error GoTo GaussianElimination_Error

    max_equation = UBound(coeffs, 1)
    max_coeff = UBound(coeffs, 2)
    For I = 0 To max_equation
        ' Use equation_coeffs(i, i) to eliminate the ith
        ' coefficient in all of the other equations.

        ' Find a row with non-zero ith coefficient.
        If (coeffs(I, I) = 0) Then
            For J = I + 1 To max_equation
                ' See if this one works.
                If (coeffs(J, I) <> 0) Then
                    ' This one works. Swap equations i and j.
                    ' This starts at k = i because all
                    ' coefficients to the left are 0.
                    For K = I To max_coeff
                        temp = coeffs(I, K)
                        coeffs(I, K) = coeffs(J, K)
                        coeffs(J, K) = temp
                    Next K
                    Exit For
                End If
            Next J
        End If

        ' Make sure we found an equation with
        ' a non-zero ith coefficient.
        coeff_i_i = coeffs(I, I)
        If coeff_i_i = 0 Then Stop

        ' Normalize the ith equation.
        For J = I To max_coeff
            coeffs(I, J) = coeffs(I, J) / coeff_i_i
        Next J

        ' Use this equation value to zero out
        ' the other equations' ith coefficients.
        For J = 0 To max_equation
            ' Skip the ith equation.
            If (J <> I) Then
                ' Zero the jth equation's ith coefficient.
                coef_j_i = coeffs(J, I)
                For d = 0 To max_coeff
                    coeffs(J, d) = coeffs(J, d) - coeffs(I, d) * coef_j_i
                Next d
            End If
        Next J
    Next I

    ' At this point, the ith equation contains
    ' 2 non-zero entries:
    '      The ith entry which is 1
    '      The last entry coeffs(max_coeff)
    ' This means Ai = equation_coef(max_coeff).
    ReDim solution(max_equation)
    For I = 0 To max_equation
        solution(I) = coeffs(I, max_coeff)
    Next I

    ' Return the solution values.
    GaussianElimination = solution

   On Error GoTo 0
   Exit Function

GaussianElimination_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GaussianElimination of Module modPolyFit"
End Function


