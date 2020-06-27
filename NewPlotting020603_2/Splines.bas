Attribute VB_Name = "modSplines"
'module origin: https://codes-sources.commentcamarche.net/source/18366-interpolation-spline
'==============================================================
' Descrizione.....: Routines di interpolazione con Splines.
' Nome dei Files..: Splines.bas
' Data............: 27/11/1999
' Aggiornamento...: 18/7/2002 (migliorata la Bezier).
' Versione........: 1.0 a 32 bits
' Sistema.........: Visual Basic 6.0 sotto Windows NT.
' Scritto da......: F. Languasco ®
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.virgilio.it/flanguasco/
'                   http://www.flanguasco.org
'==============================================================
'
'   Gli algoritmi usati sono di:
'   P. Bourke     -  Aprile 1989
'   D. Cholaseuk  -  8/Dic./1999
'
'   Le curve Splines vengono calcolate in modo parametrico
'   e quindi, con opportuni adattamenti, possono essere
'   usate per interpolare Point a n dimensioni.
'
'==============================================================
' Modification.....: Modification pour tratement en 3D
' Auteur...........: Cuq
'==============================================================
Option Explicit
'
Public Type P_Type
    X As Double         ' Coordonnee x du point.
    Y As Double         ' Coordonnee y du point.
    Z As Double         ' Coordonnee z du point.
End Type

Public NPI            ' N. de Points dans la courbe.
Public Pi() As P_Type  ' Coordonnees des Points de l'interpolation.
Public NPC            ' N. de point approximant la courbe.
Public Pc() As P_Type  ' Coordonnees des points pour l'approximation.
Public NK             ' Degree pour la B-Spline.
Public VZ             ' Tension de la courbe T-Spline.


Public Sub Bezier(Pi() As P_Type, Pc() As P_Type)
'
'   Ritorna, nel vettore Pc(), i valori della curva di Bezier.
'   La curva e' calcolata in modo parametrico (0 <= u < 1)
'   con il valore 0 di u corrispondente a Pc(0);
'
'   Questa versione elimina alcuni problemi di "underflow"
'   e di "overflow" presentati dalla Bezier_1 e dalla Bezier_C.
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei Point, dati, da
'                           approssimare.
'       Pc(0 to NPC - 1):   Vettore dei Point, calcolati,
'                           della curva approssimante.
'
   Dim I&, K&, NPI_1&, NPC_1&
   Dim u#, u_1#, ue#, BF#
   Static NPI_1_O&, CB_Tav#()
'
    NPI_1 = UBound(Pi) ' N. di Point da approssimare - 1 (deve essere 2 <= NPI_1 <= 1029).
    NPC_1 = UBound(Pc) ' N. di Point sulla curva - 1.
'
    If NPI_1_O <> NPI_1 Then
        ' Prepara la tavola dei coefficienti binomiali:
        ReDim CB_Tav#(0 To NPI_1)
        For K = 0 To NPI_1
            CB_Tav(K) = rncr(NPI_1, K)
            If CB_Tav(K) = -9999 Then Exit Sub 'overflow detected
        Next K
'
        NPI_1_O = NPI_1
    End If
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y
    Pc(0).Z = Pi(0).Z
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ue = 1#
        u_1 = 1# - u
'
        Pc(I).X = 0#
        Pc(I).Y = 0#
        Pc(I).Z = 0#
        For K = 0 To NPI_1
            BF = CB_Tav(K) * ue * u_1 ^ (NPI_1 - K)
'
            Pc(I).X = Pc(I).X + Pi(K).X * BF
            Pc(I).Y = Pc(I).Y + Pi(K).Y * BF
            Pc(I).Z = Pc(I).Z + Pi(K).Z * BF
'
            ue = ue * u
        Next K
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    Pc(NPC_1).Z = Pi(NPI_1).Z
'
'
'
End Sub
Private Function Prodotto(ByVal N2&, Optional ByVal N1& = 2) As Double
'
'   Ritorna il prodotto dei numeri, consecutivi, interi e positivi,
'   da N1 a N2 (0 < N1 <= N2). Se N1 > N2 ritorna 1.
'   Se N1 manca, ritorna il Fattoriale di N2; in questo caso puo'
'   anche essere N2 = 0 perche', per definizione, e' 0! = 1:
'
    Dim Pr#, I&
'
    Pr = 1#
    For I = N1 To N2
        Pr = Pr * CDbl(I)
    Next I
'
    Prodotto = Pr
'
'
'
End Function
Public Sub B_Spline(Pi() As P_Type, ByVal NK, Pc() As P_Type)
'
'   Ritorna, nel vettore Pc(), i valori della curva B-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei Point, dati, da
'                           approssimare.
'       Pc(0 to NPC - 1):   Vettore dei Point, calcolati,
'                           della curva approssimante.
'       NK:                 Numero di nodi della curva
'                           approssimante:
'                           NK = 2    -> segmenti di retta.
'                           NK = 3    -> curve quadratiche.
'                           ..   .       ..................
'                           NK = NPI  -> splines di Bezier.

    Dim NPI_1&, NPC_1&, I&, J&, tmax#, u#, ut#, bn#()
    Const Eps = 0.0000001
'
    NPI_1 = UBound(Pi)  ' N. di Point da approssimare - 1.
    NPC_1 = UBound(Pc)  ' N. di Point sulla curva - 1.
    tmax = NPI_1 - NK + 2
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y
    Pc(0).Z = Pi(0).Z
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ut = u * tmax
        If Abs(ut - CDbl(NPI_1 + NK - 2)) <= Eps Then
            Pc(I).X = Pi(NPI_1).X
            Pc(I).Y = Pi(NPI_1).Y
            Pc(I).Z = Pi(NPI_1).Z
        Else
            Call B_Basis(NPI_1, ut, NK, bn())
            Pc(I).X = 0#
            Pc(I).Y = 0#
            Pc(I).Z = 0#
            For J = 0 To NPI_1
                Pc(I).X = Pc(I).X + bn(J) * Pi(J).X
                Pc(I).Y = Pc(I).Y + bn(J) * Pi(J).Y
                Pc(I).Z = Pc(I).Z + bn(J) * Pi(J).Z
            Next J
        End If
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    Pc(NPC_1).Z = Pi(NPI_1).Z
'
'
End Sub
Private Sub B_Basis(ByVal NPI_1&, ByVal ut#, ByVal K&, bn#())
'
'   Compute the basis function (also called weight)
'   for the B-Spline approximation curve:
'
    Dim NT&, I&, J&
    Dim b0#, b1#, bl0#, bl1#, bu0#, bu1#
    ReDim bn#(0 To NPI_1 + 1), bn0#(0 To NPI_1 + 1), T#(0 To NPI_1 + K + 1)
'
    NT = NPI_1 + K + 1
    For I = 0 To NT
        If (I < K) Then T(I) = 0#
        If ((I >= K) And (I <= NPI_1)) Then T(I) = CDbl(I - K + 1)
        If (I > NPI_1) Then T(I) = CDbl(NPI_1 - K + 2)
    Next I
    For I = 0 To NPI_1
        bn0(I) = 0#
        If ((ut >= T(I)) And (ut < T(I + 1))) Then bn0(I) = 1#
        If ((T(I) = 0#) And (T(I + 1) = 0#)) Then bn0(I) = 0#
    Next I
'
    For J = 2 To K
        For I = 0 To NPI_1
            bu0 = (ut - T(I)) * bn0(I)
            bl0 = T(I + J - 1) - T(I)
            If (bl0 = 0#) Then
                b0 = 0#
            Else
                b0 = bu0 / bl0
            End If
            bu1 = (T(I + J) - ut) * bn0(I + 1)
            bl1 = T(I + J) - T(I + 1)
            If (bl1 = 0#) Then
                b1 = 0#
            Else
                b1 = bu1 / bl1
            End If
            bn(I) = b0 + b1
        Next I
        For I = 0 To NPI_1
            bn0(I) = bn(I)
        Next I
    Next J
'
'
'
End Sub
Public Sub C_Spline(Pi() As P_Type, Pc() As P_Type)
'
'   Ritorna, nel vettore Pc(), i valori della curva C-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei Point, dati, da
'                           interpolare.
'       Pc(0 to NPC - 1):   Vettore dei Point, calcolati,
'                           della curva interpolante.
'
    Dim NPI_1&, NPC_1&, I&, J&
    Dim u#, ui#, uui#
    Dim cof() As P_Type
'
    NPI_1 = UBound(Pi)      ' N. di Point da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di Point sulla curva - 1.
'
    Call Find_CCof(Pi(), NPI_1 + 1, cof())
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y
    Pc(0).Z = Pi(0).Z
    
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        J = Int(u * CDbl(NPI_1)) + 1
        If (J > (NPI_1)) Then J = NPI_1
'
        ui = CDbl(J - 1) / CDbl(NPI_1)
        uui = u - ui
'
        Pc(I).X = cof(4, J).X * uui ^ 3 + cof(3, J).X * uui ^ 2 + cof(2, J).X * uui + cof(1, J).X
        Pc(I).Y = cof(4, J).Y * uui ^ 3 + cof(3, J).Y * uui ^ 2 + cof(2, J).Y * uui + cof(1, J).Y
        Pc(I).Z = cof(4, J).Z * uui ^ 3 + cof(3, J).Z * uui ^ 2 + cof(2, J).Z * uui + cof(1, J).Z

    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    Pc(NPC_1).Z = Pi(NPI_1).Z
'
'
End Sub
Private Function rncr(ByVal N&, ByVal K&) As Double
'
'   Calcola i coefficienti binomiali Cn,k come:
'    rncr = N! / (K! * (N - K)!)
'
'   Nota: La funzione ha senso solo per 0 < N, K <= N
'         e 0 <= K.  Nessun Erreur viene segnalato.
'
    Dim I&, rncr_T#
'
   On Error GoTo rncr_Error

    If ((N < 1) Or (K < 1) Or (N = K)) Then
        rncr = 1#
'
    Else
        rncr_T = 1#
        For I = 1 To N - K
            rncr_T# = rncr_T# * (1# + CDbl(K) / CDbl(I))
        Next I
'
        rncr = rncr_T#
    End If
'
'
'

   On Error GoTo 0
   Exit Function

rncr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rncr of Module modSplines"
    rncr = -9999
End Function
Public Sub T_Spline(Pi() As P_Type, ByVal VZ, Pc() As P_Type)
'
'   Ritorna, nel vettore Pc(), i valori della curva T-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei Point, dati, da
'                           interpolare.
'       Pc(0 to NPC - 1):   Vettore dei Point, calcolati,
'                           della curva interpolante.
'       VZ:                 Valore di tensione della curva
'                           (1 <= VZ <= 100): valori grandi
'                           di VZ appiattiscono la curva.
'
    Dim NPI_1&, NPC_1&, I&, J&
    Dim H#, Z#, z2i#, szh#, u#, u0#, u1#, du1#, du0#
    Dim s() As P_Type
'
    NPI_1 = UBound(Pi)      ' N. di Point da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di Point sulla curva - 1.
    Z = CDbl(VZ)
'
    Call Find_TCof(Pi(), NPI_1 + 1, s(), Z)
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y
    Pc(0).Z = Pi(0).Z
'
    H = 1# / CDbl(NPI_1)
    szh = Sinh(Z * H)
    z2i = 1# / Z / Z
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        J = Int(u * CDbl(NPI_1)) + 1
        If (J > (NPI_1)) Then J = NPI_1
'
        u0 = CDbl(J - 1) / CDbl(NPI_1)
        u1 = CDbl(J) / CDbl(NPI_1)
        du1 = u1 - u
        du0 = u - u0
'
        Pc(I).X = s(J).X * z2i * Sinh(Z * du1) / szh + (Pi(J - 1).X - s(J).X * z2i) * du1 / H
        Pc(I).X = Pc(I).X + s(J + 1).X * z2i * Sinh(Z * du0) / szh + (Pi(J).X - s(J + 1).X * z2i) * du0 / H
    
        Pc(I).Y = s(J).Y * z2i * Sinh(Z * du1) / szh + (Pi(J - 1).Y - s(J).Y * z2i) * du1 / H
        Pc(I).Y = Pc(I).Y + s(J + 1).Y * z2i * Sinh(Z * du0) / szh + (Pi(J).Y - s(J + 1).Y * z2i) * du0 / H
        
        Pc(I).Z = s(J).Z * z2i * Sinh(Z * du1) / szh + (Pi(J - 1).Z - s(J).Z * z2i) * du1 / H
        Pc(I).Z = Pc(I).Z + s(J + 1).Z * z2i * Sinh(Z * du0) / szh + (Pi(J).Z - s(J + 1).Z * z2i) * du0 / H

    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    Pc(NPC_1).Z = Pi(NPI_1).Z
'
'
'
End Sub
Private Sub Find_TCof(Pi() As P_Type, ByVal NPI, s() As P_Type, ByVal Z#)
'
'   Find the coefficients of the T-Spline
'   using constant interval:
'
    Dim I&, H#, a0#, b0#, zh#, Z2#
'
    ReDim s(1 To NPI) As P_Type, f(1 To NPI) As P_Type
    ReDim a(1 To NPI) As Double, B(1 To NPI) As Double, C(1 To NPI) As Double
'
    H = 1# / CDbl(NPI - 1)
    zh = Z * H
    a0 = 1# / H - Z / Sinh(zh)
    b0 = Z * 2# * Cosh(zh) / Sinh(zh) - 2# / H
    For I = 1 To NPI - 2
        a(I) = a0
        B(I) = b0
        C(I) = a0
    Next I
'
    Z2 = Z * Z / H
    For I = 1 To NPI - 2
        f(I).X = (Pi(I + 1).X - 2# * Pi(I).X + Pi(I - 1).X) * Z2
        f(I).Y = (Pi(I + 1).Y - 2# * Pi(I).Y + Pi(I - 1).Y) * Z2
        f(I).Z = (Pi(I + 1).Z - 2# * Pi(I).Z + Pi(I - 1).Z) * Z2
    Next I
'
    Call TRIDAG(a(), B(), C(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).X = s(NPI - 1 - I).X
        s(NPI - I).Y = s(NPI - 1 - I).Y
        s(NPI - I).Z = s(NPI - 1 - I).Z
    Next I
'
    s(1).X = 0#
    s(NPI).X = 0#
    s(1).Y = 0#
    s(NPI).Y = 0#
    s(1).Z = 0#
    s(NPI).Z = 0#
'
'
'
End Sub
Private Sub Find_CCof(Pi() As P_Type, ByVal NPI, cof() As P_Type)
'
'   Find the coefficients of the cubic spline
'   using constant interval parameterization:
'
    Dim I&, H#
'
    ReDim s(1 To NPI) As P_Type, f(1 To NPI) As P_Type, cof(1 To 4, 1 To NPI) As P_Type
    ReDim a(1 To NPI) As Double, B(1 To NPI) As Double, C(1 To NPI) As Double
'
    H = 1# / CDbl(NPI - 1)
    For I = 1 To NPI - 2
        a(I) = 1#
        B(I) = 4#
        C(I) = 1#
    Next I
'
    For I = 1 To NPI - 2
        f(I).X = 6# * (Pi(I + 1).X - 2# * Pi(I).X + Pi(I - 1).X) / H / H
        f(I).Y = 6# * (Pi(I + 1).Y - 2# * Pi(I).Y + Pi(I - 1).Y) / H / H
        f(I).Z = 6# * (Pi(I + 1).Z - 2# * Pi(I).Z + Pi(I - 1).Z) / H / H
    Next I
'
    Call TRIDAG(a(), B(), C(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).X = s(NPI - 1 - I).X
        s(NPI - I).Y = s(NPI - 1 - I).Y
        s(NPI - I).Z = s(NPI - 1 - I).Z
    Next I
'
    s(1).X = 0#
    s(NPI).X = 0#
    s(1).Y = 0#
    s(NPI).Y = 0#
    s(1).Z = 0#
    s(NPI).Z = 0#
    
    For I = 1 To NPI - 1
        cof(4, I).X = (s(I + 1).X - s(I).X) / 6# / H
        cof(4, I).Y = (s(I + 1).Y - s(I).Y) / 6# / H
        cof(4, I).Z = (s(I + 1).Z - s(I).Z) / 6# / H
                
        cof(3, I).X = s(I).X / 2#
        cof(3, I).Y = s(I).Y / 2#
        cof(3, I).Z = s(I).Z / 2#
        
        cof(2, I).X = (Pi(I).X - Pi(I - 1).X) / H - (2# * s(I).X + s(I + 1).X) * H / 6#
        cof(2, I).Y = (Pi(I).Y - Pi(I - 1).Y) / H - (2# * s(I).Y + s(I + 1).Y) * H / 6#
        cof(2, I).Z = (Pi(I).Z - Pi(I - 1).Z) / H - (2# * s(I).Z + s(I + 1).Z) * H / 6#

        cof(1, I).X = Pi(I - 1).X
        cof(1, I).Y = Pi(I - 1).Y
        cof(1, I).Z = Pi(I - 1).Z
    Next I
'
'
'
End Sub
Private Sub TRIDAG(a#(), B#(), C#(), f() As P_Type, s() As P_Type, ByVal NPI_2&)
'
'   Solves the tridiagonal linear system of equations:
'
    Dim J&, bet#
    ReDim gam#(1 To NPI_2)
'
    If B(1) = 0 Then Exit Sub
'
    bet = B(1)
    s(1).X = f(1).X / bet
    s(1).Y = f(1).Y / bet
    s(1).Z = f(1).Z / bet
    
        
    For J = 2 To NPI_2
        gam(J) = C(J - 1) / bet
        bet = B(J) - a(J) * gam(J)
        If (bet = 0) Then Exit Sub
        s(J).X = (f(J).X - a(J) * s(J - 1).X) / bet
        s(J).Y = (f(J).Y - a(J) * s(J - 1).Y) / bet
        s(J).Z = (f(J).Z - a(J) * s(J - 1).Z) / bet
                
    Next J
'
    For J = NPI_2 - 1 To 1 Step -1
        s(J).X = s(J).X - gam(J + 1) * s(J + 1).X
        s(J).Y = s(J).Y - gam(J + 1) * s(J + 1).Y
        s(J).Z = s(J).Z - gam(J + 1) * s(J + 1).Z
    Next J
'
'
'
End Sub
Private Function Cosh(ByVal Z As Double) As Double
'
'   Retourne le Cosinus Hyperbolique de  z#:
'
    Cosh = (Exp(Z) + Exp(-Z)) / 2#
'
'
'
End Function
Private Function Sinh(ByVal Z As Double) As Double
'
'   Retourne le Sinus Hyperbolique de  z#:
'
    Sinh = (Exp(Z) - Exp(-Z)) / 2#
'
'
'
End Function
