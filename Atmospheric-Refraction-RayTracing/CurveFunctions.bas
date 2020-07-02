Attribute VB_Name = "CurveFunctions"
Option Explicit

' The function.
Public Function F(ByVal coeffs As Collection, ByVal x As Double) As Double
Dim total As Double
Dim x_factor As Double
Dim i As Integer

    total = 0
    x_factor = 1
    For i = 1 To coeffs.Count
        total = total + x_factor * coeffs(i)
        x_factor = x_factor * x
    Next i
    F = total
End Function

' Return the error squared.
Public Function ErrorSquared(ByVal PtX As Collection, ByVal PtY As Collection, ByVal coeffs As Collection) As Double
Dim total As Double
Dim pt As Integer
Dim dy As Double

    total = 0
    For pt = 1 To PtX.Count
        dy = PtY.Item(pt) - F(coeffs, PtX.Item(pt))
        total = total + dy * dy
    Next pt
    ErrorSquared = total
End Function

' Find the least squares linear fit.
Public Function FindPolynomialLeastSquaresFit(ByVal PtX As Collection, ByVal PtY As Collection, ByVal degree As Integer) As Collection
Dim j As Integer
Dim pt As Integer
Dim a_sub  As Integer
Dim coeff As Variant

' Allocate space for (degree + 1) equations with
' (degree + 2) terms each (including the constant term).
Dim coeffs() As Double
Dim answer() As Double

    ReDim coeffs(degree, degree + 1)

    ' Calculate the coefficients for the equations.
    For j = 0 To degree
        ' Calculate the coefficients for the jth equation.

        ' Calculate the constant term for this equation.
        coeffs(j, degree + 1) = 0
        For pt = 1 To PtX.Count
            coeffs(j, degree + 1) = coeffs(j, degree + 1) - (PtX.Item(pt) ^ j) * PtY.Item(pt)
        Next pt

        ' Calculate the other coefficients.
        For a_sub = 0 To degree
            ' Calculate the dth coefficient.
            coeffs(j, a_sub) = 0
            For pt = 1 To PtX.Count
                coeffs(j, a_sub) = coeffs(j, a_sub) - PtX.Item(pt) ^ (a_sub + j)
            Next pt
        Next a_sub
    Next j

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
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp As Double
Dim coeff_i_i As Double
Dim coef_j_i As Double
Dim d As Integer
Dim solution() As Double

    max_equation = UBound(coeffs, 1)
    max_coeff = UBound(coeffs, 2)
    For i = 0 To max_equation
        ' Use equation_coeffs(i, i) to eliminate the ith
        ' coefficient in all of the other equations.

        ' Find a row with non-zero ith coefficient.
        If (coeffs(i, i) = 0) Then
            For j = i + 1 To max_equation
                ' See if this one works.
                If (coeffs(j, i) <> 0) Then
                    ' This one works. Swap equations i and j.
                    ' This starts at k = i because all
                    ' coefficients to the left are 0.
                    For k = i To max_coeff
                        temp = coeffs(i, k)
                        coeffs(i, k) = coeffs(j, k)
                        coeffs(j, k) = temp
                    Next k
                    Exit For
                End If
            Next j
        End If

        ' Make sure we found an equation with
        ' a non-zero ith coefficient.
        coeff_i_i = coeffs(i, i)
        If coeff_i_i = 0 Then Stop

        ' Normalize the ith equation.
        For j = i To max_coeff
            coeffs(i, j) = coeffs(i, j) / coeff_i_i
        Next j

        ' Use this equation value to zero out
        ' the other equations' ith coefficients.
        For j = 0 To max_equation
            ' Skip the ith equation.
            If (j <> i) Then
                ' Zero the jth equation's ith coefficient.
                coef_j_i = coeffs(j, i)
                For d = 0 To max_coeff
                    coeffs(j, d) = coeffs(j, d) - coeffs(i, d) * coef_j_i
                Next d
            End If
        Next j
    Next i

    ' At this point, the ith equation contains
    ' 2 non-zero entries:
    '      The ith entry which is 1
    '      The last entry coeffs(max_coeff)
    ' This means Ai = equation_coef(max_coeff).
    ReDim solution(max_equation)
    For i = 0 To max_equation
        solution(i) = coeffs(i, max_coeff)
    Next i

    ' Return the solution values.
    GaussianElimination = solution
End Function

' Find the least squares linear fit.
' Return the total error.
Public Function FindLinearLeastSquaresFit(ByVal PtX As Collection, ByVal PtY As Collection, ByRef m As Double, ByRef b As Double) As Double
Dim S1 As Double
Dim Sx As Double
Dim Sy As Double
Dim Sxx As Double
Dim Sxy As Double
Dim i As Integer

    ' Perform the calculation.
    ' Find the values S1, Sx, Sy, Sxx, and Sxy.
    S1 = PtX.Count
    Sx = 0
    Sy = 0
    Sxx = 0
    Sxy = 0
    For i = 1 To PtX.Count
        Sx = Sx + PtX.Item(i)
        Sy = Sy + PtY.Item(i)
        Sxx = Sxx + PtX.Item(i) * PtX.Item(i)
        Sxy = Sxy + PtX.Item(i) * PtY.Item(i)
    Next i

    ' Solve for m and b.
    m = (Sxy * S1 - Sx * Sy) / (Sxx * S1 - Sx * Sx) 'slope
    b = (Sxy * Sx - Sy * Sxx) / (Sx * Sx - S1 * Sxx) 'y(0) where y(x) = b + m * x

    FindLinearLeastSquaresFit = Sqr(LinErrorSquared(PtX, PtY, m, b))
End Function

' Return the error squared.
Public Function LinErrorSquared(ByVal PtX As Collection, ByVal PtY As Collection, ByVal m As Double, ByVal b As Double) As Double
Dim i As Integer
Dim total As Double
Dim dy As Double

    total = 0
    For i = 1 To PtX.Count
        dy = PtY.Item(i) - (m * PtX.Item(i) + b)
        total = total + dy * dy
    Next i
    LinErrorSquared = total
End Function


