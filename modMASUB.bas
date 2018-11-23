Attribute VB_Name = "modMASUB"
'=============================================================
' Description......: Routines for interpolating surfaces from
'                     scattered data points.
' Name of the Files: modMASUB.bas
' Date.............: 21/9/2001
' Version..........: 1.0 at 32 bits.
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.it/flanguasco/
'                    http://www.flanguasco.org
'=============================================================
'
'   Input routine: MASUB (see the parameter description
'                          in the routine)
'
'   Note:   All the vectors and matrices of these routines start from index 1.
'
'   Translated from the FORTRAN program:
'      ALGORITHM 677, COLLECTED ALGORITHMS FROM ACM.
'      THIS WORK PUBLISHED IN TRANSACTIONS ON MATHEMATICAL SOFTWARE,
'      VOL. 15, NO. 4, PP. 365-374.
'
Option Explicit

Public Function MASUB(ByVal IC&, ByVal IEX&, ByVal ND&, XD#(), YD#(), ZD#(), _
     ByVal TP#, ByVal NXI&, ByVal NYI&, XI#(), YI#(), ByRef ZI#(), Optional ByVal ZI_default = 1000000#) As Boolean
'
'   IT CARRIES OUT A SMOOTH BIVARIATE INTERPOLATION WHEN THE DATA
'   POINTS PROJECTION IN THE  X-Y  PLANE IS IRREGULARLY DISTRIBUTED
'   IN THE PLANE: IT ESTIMATES THE INTERPOLANT FUNCTION AT PRESCRIBED
'   RECTANGULAR GRID POINTS BY MEANS OF THE 9-PARAMETERS DISCRETIZED
'   VERSION OF THE NIELSON TRIANGULAR INTERPOLANT. THE REQUIRED DERIVATIVE
'   VALUES ARE CALCULATED BY MINIMIZING SUITABLE TENSION'S FUNCTIONAL
'   AND THE TRIANGULATION OF THE CONVEX HULL OF THE X-Y DATA IS REALIZED
'   BY USING LAWSON'S LOCAL OPTIMIZATION PROCEDURE.
'   IF EXTRAPOLATION IS REQUESTED (IEX PARAMETER) IT ADDS TO THE DATA
'   POINTS SOME ADDITIONAL DATA, EVALUATED BY MEANS OF THE SHEPARD'S
'   METHOD, SO THAT THE NEW CONVEX HULL OF THE DATA POINTS CONTAINS
'   THE RECTANGULAR DOMAIN WHERE THE SURFACE HAS TO BE RECONSTRUCTED.
'
'   THE INPUT PARAMETERS ARE
'    SRELPR =   SINGLE RELATIVE PRECISION,
'    IC =       FLAG OF COMPUTATION (MUST BE 1,2,3 OR 4),
'                = 1 FOR THE FIRST CALL AND FOR NEW IEX,ND,XD-YD,
'                = 2 FOR NEW ZD,TP,
'                = 3 FOR NEW TP,
'                = 4 FOR NEW NXI,NYI,XI-YI,
'    IEX =      FLAG OF COMPUTATION (MUST BE 0 OR 1),
'                = 1 (ONE) ALLOWS THE EXTRAPOLATION, 0 (ZERO) DOES NOT;
'               THIS PARAMETER MUST BE THE SAME IN THE FOLLOWING CALL,
'    ND =       NUMBER OF DATA POINTS (MUST BE 4 OR GREATER);
'               GIVEN IEX AND ND WE DEFINE N AS
'                       N = ND + IEX*(2*INT(ND/25)+4),
'    XD =       ARRAY OF DIMENSION N CONTAINING THE  X  COORDINATES OF
'               THE ND DATA POINTS,
'    YD =       ARRAY OF DIMENSION N CONTAINING THE  Y  COORDINATES OF
'               THE ND DATA POINTS,
'    ZD =       ARRAY OF DIMENSION N CONTAINING THE  Z  COORDINATES OF
'               THE ND DATA POINTS,
'    TP =       TENSION PARAMETER (MUST BE GREATER THAN OR EQUAL TO ZERO).
'    NXI =      NUMBER OF RECTANGULAR GRID POINTS IN THE X COORDINATE (MUST
'               BE 1 OR GREATER),
'    NYI =      NUMBER OF RECTANGULAR GRID POINTS IN THE Y COORDINATE (MUST
'               BE 1 OR GREATER),
'    XI =       ARRAY OF DIMENSION NXI CONTAINING,IN ASCENDING ORDER, THE X
'               COORDINATES OF THE RECTANGULAR GRID POINTS,  WHERE THE
'               SURFACE HAS TO BE RECONSTRUCTED,
'    YI =       ARRAY OF DIMENSION NYI CONTAINING,IN ASCENDING ORDER, THE Y
'               COORDINATES OF THE RECTANGULAR GRID POINTS,  WHERE THE
'               SURFACE HAS TO BE RECONSTRUCTED,
'    NDIM=      DECLARED ROW DIMENSION OF THE ARRAY CONTAINING ZI,
'    NOUT=      LOGICAL UNIT NUMBER FOR THE STANDARD OUTPUT UNIT OF THE
'               SYSTEM.
'
'   THE OUTPUT PARAMETERS ARE
'    XD,YD,ZD = ARRAYS OF DIMENSION N, CONTAINING THE X, Y AND Z
'               COORDINATES OF THE ND DATA POINTS AND,IF EXTRAPOLATION
'               IS REQUESTED (IEX=1), THE X, Y AND Z COORDINATES OF
'               N-ND ADDITIONAL POINTS EVALUATED BY MEANS OF SHEPARD'S
'               METHOD. IF EXTRAPOLATION IS NOT REQUESTED (IEX=0) THEN
'               N=ND AND XD, YD, ZD ARE NOT CHANGED,
'               (Vengono, eventualmente, ridimensionati in questa routine)
'    ZI =       DOUBLY-DIMENSIONED ARRAY OF DIMENSION NXI*NYI, WHERE THE
'               INTERPOLATED  Z  VALUES AT THE RECTANGULAR GRID POINTS ARE
'               TO BE STORED.
'               IF EXTRAPOLATION IS NOT REQUESTED (IEX=0), THE Z VALUES AT
'               THE RECTANGULAR GRID POINTS OUTSIDE THE CONVEX HULL OF THE
'               DATA POINTS, ARE SET TO 1.0E6 OR ZI_default.
'
'   THE OTHER PARAMETERS ARE
'    IWK =      INTEGER ARRAY OF DIMENSION  23*N-31+MAX0(8*N+25,NXI*NYI)
'               USED INTERNALLY AS WORK AREA,
'    WK =       ARRAY OF DIMENSION  19*N-20  USED INTERNALLY AS WORK AREA.
'
'   THE FIRST CALL TO THIS SUBROUTINE AND THE CALL WITH A NEW ND VALUE
'   AND/OR NEW CONTENTS OF THE XD AND YD ARRAYS MUST BE MADE WITH IC=1.
'   THE CALL WITH IC=2 MUST BE PRECEDED BY ANOTHER CALL WITH THE SAME
'   IEX,ND,NXI AND NYI VALUES AND WITH THE SAME CONTENTS OF THE XD,YD,
'   XI AND YI ARRAYS.
'   THE CALL WITH IC=3 MUST BE PRECEDED BY ANOTHER CALL WITH THE SAME
'   IEX,ND,NXI AND NYI VALUES AND WITH THE SAME CONTENTS OF THE XD,YD,
'   ZD,XI AND YI ARRAYS.
'   THE CALL WITH IC=4 MUST BE PRECEDED BY ANOTHER CALL WITH THE SAME
'   IEX AND ND VALUES AND WITH THE SAME CONTENTS OF THE XD,YD AND ZD
'   ARRAYS.
'   IWK AND WK ARRAYS MUST NOT BE DISTURBED BETWEEN THE CALL WITH IC NOT
'   EQUAL TO  1  AND THE PRECEDING ONE.
'   SUBROUTINE MASUB CALLS THE EXTRP,CTANG,ADJAC,PDSTE,PDMIN,ORDGR AND
'   INTRP SUBROUTINES.
'
    Dim I&, IMax&, IMIN&, I1MAX&, I1MIN&, ITI&, ITPV&, IXI&, IYI&, IZI&
    Dim J&, JWNGP&, JWIGP&, KIGP&, KNGP&
    Dim N&, NDP&, NGP0&, NGP1&, NT&
    Dim SRELPR#, TOLL#
'
    N = ND + IEX * (2 * Int(ND / 25) + 4)
    ReDim Preserve XD(1 To N), YD(1 To N), ZD(1 To N)
    ReDim IWK&(1 To 4 * ND + 7), WK#(1 To 7)
'
    Call MACEPS(SRELPR)
    TOLL = 1# - SRELPR
'
    ' ERROR CHECK.
8   If (IC < 1 Or IC > 4) Then GoTo 1000
    If (IEX < 0 Or IEX > 1) Then GoTo 1000
    If (ND < 4) Then GoTo 1000
    If (NXI < 1 Or NYI < 1) Then GoTo 1000
    If (IC > 1) Then GoTo 10
    IWK(1) = ND
    IWK(2) = IEX
    IWK(3) = NXI
    IWK(4) = NYI
    IWK(6) = ND
    If (IEX = 0) Then GoTo 50
    WK(1) = XI(1)
    WK(2) = XI(NXI)
    WK(3) = YI(1)
    WK(4) = YI(NYI)
    GoTo 40
10  If (IWK(1) <> ND) Then GoTo 1000
    If (IWK(2) <> IEX) Then GoTo 1000
    If (IC > 3) Then GoTo 20
    If (IWK(3) <> NXI) Then GoTo 1000
    If (IWK(4) <> NYI) Then GoTo 1000
    If (IC = 2 And IEX = 1) Then GoTo 40
    GoTo 50
20  IWK(3) = NXI
    IWK(4) = NYI
    If (IEX = 0) Then GoTo 50
    If (XI(1) < WK(1) Or XI(NXI) > WK(2)) Then GoTo 1000
    If (YI(1) < WK(3) Or YI(NYI) > WK(4)) Then GoTo 1000
    GoTo 50
'
    ' ADDS SOME EXTERNAL POINTS FOR THE EXTRAPOLATION.
40  Call EXTRP(ND, XD(), YD(), ZD(), IC, IWK(), WK(1), WK(2), WK(3), WK(4))
50  NDP = IWK(6)
    ReDim Preserve XD(1 To NDP), YD(1 To NDP), ZD(1 To NDP)
'
    ' ALLOCATION STORAGE AREAS IN THE IWK AND WK ARRAYS.
    ' In questa versione sono stati usati array dedicati:
    ' si evita percio' di dover passare puntatori all' interno
    ' dei vettori, operazione comune in FORTRAN ma estremamente
    ' inefficiente in Basic (vedi modMASUB_1v.bas ed il sorgente FORTRAN).
'    JIPT = 8
'    JIPL = 6 * NDP - 7
'    JIWP = 12 * NDP - 7 - 1
'    JIND = 13 * NDP - 7
'    JNGP = 19 * NDP - 22
'    JIGP = 23 * NDP - 32
    ReDim IPT&(1 To (6 * NDP - 7) - (8))
    ReDim IPL&(1 To (12 * NDP - 7 - 1) - (6 * NDP - 7))
    ReDim IWP&(1 To (13 * NDP - 7) - (12 * NDP - 7 - 1))
    ReDim IND&(1 To (19 * NDP - 22) - (13 * NDP - 7) + 1)
    ReDim NGP&(1 To (23 * NDP - 32) - (19 * NDP - 22))
    'ReDim IGP&(1 To (23 * N - 31 + MAX0(8 * N + 25, NXI * NYI)) - (23 * NDP - 32))
    ReDim IGP&(1 To NXI * NYI)
'    IPD = 5
'    IAL = 2 * NDP + 5
'    IBE = 5 * NDP - 1
'    IGA = 8 * NDP - 7
'    IEI = 11 * NDP - 13
'    IALS = 14 * NDP - 19
'    IBES = IALS + NDP
'    IGAS = IBES + NDP
'    IZX = IGAS + NDP
'    IZY = IZX + NDP
    ReDim PD#(1 To (2 * NDP + 5) - (5))
    ReDim AL#(1 To (5 * NDP - 1) - (2 * NDP + 5))
    ReDim BE#(1 To (8 * NDP - 7) - (5 * NDP - 1))
    ReDim GA#(1 To (11 * NDP - 13) - (8 * NDP - 7))
    ReDim EI#(1 To (14 * NDP - 19) - (11 * NDP - 13))
    ReDim ALS#(1 To NDP)
    ReDim BES#(1 To NDP)
    ReDim GAS#(1 To NDP)
    ReDim ZX#(1 To NDP)
    ReDim ZY#(1 To NDP)
    If (IC > 1) Then GoTo 60
'
    ' TRIANGULATES THE X-Y PLANE.
    Call CTANG(NDP, XD(), YD(), NT, IPT(), IPL(), IND(), IWP(), PD())
    If NT = 0 Then
        MASUB = False
        Exit Function
    End If
    IWK(5) = NT
'
    ' CONSTRUCTS THE ADJACENCIES MONODIMENSIONAL ARRAY.
    Call ADJAC(NT, IPT(), NDP, IPL(), IWP())
60  NT = IWK(5)
    If (IC > 3) Then GoTo 70
    If (IC = 3) Then GoTo 65
'
    ' ESTIMATES PARTIAL DERIVATIVES AT ALL DATA POINTS.
    If Not PDSTE(NDP, XD(), YD(), ZD(), NT, IPT(), PD(), IND()) Then GoTo 1000
65  If Not PDMIN(SRELPR, NDP, XD(), YD(), ZD(), IPL(), IWP(), IND(), TP, PD(), _
       AL(), BE(), GA(), EI(), ALS(), BES(), GAS(), ZX(), ZY()) Then GoTo 1000
    If (IC > 1) Then GoTo 80
'
    ' SORTS RECTANGULAR GRID POINTS ACCORDING TO THEIR BELONGING TO THE
    ' TRIANGLES.
70  Call ORDGR(XD(), YD(), NT, IPT(), NXI, NYI, XI(), YI(), NGP(), IGP())
80  For I = 1 To NXI
        For J = 1 To NYI
            ZI(I, J) = ZI_default
        Next J
    Next I
    ITPV = 0
    IMax = 0
    I1MIN = NXI * NYI + 1
    For KNGP = 1 To NT
        ITI = KNGP
        JWNGP = 1 - 1 + KNGP
        NGP0 = NGP(JWNGP)
        If (NGP0 = 0) Then GoTo 100
        IMIN = IMax + 1
        IMax = IMax + NGP0
        For KIGP = IMIN To IMax
            JWIGP = 1 + KIGP - 1
            IZI = IGP(JWIGP)
            IYI = ((IZI - 1) \ NXI) + 1
            IXI = IZI - NXI * (IYI - 1)
            Call INTRP(SRELPR, TOLL, XD(), YD(), ZD(), IPT(), PD(), ITI, ITPV _
               , XI(IXI), YI(IYI), ZI(IXI, IYI))
        Next KIGP
100     JWNGP = 1 + 2 * NT - KNGP
        NGP1 = NGP(JWNGP)
        If (NGP1 = 0) Then GoTo 120
        I1MAX = I1MIN - 1
        I1MIN = I1MIN - NGP1
        For KIGP = I1MIN To I1MAX
            JWIGP = 1 + KIGP - 1
            IZI = IGP(JWIGP)
            IYI = ((IZI - 1) \ NXI) + 1
            IXI = IZI - NXI * (IYI - 1)
            Call INTRP(SRELPR, TOLL, XD(), YD(), ZD(), IPT(), PD(), ITI, ITPV _
               , XI(IXI), YI(IYI), ZI(IXI, IYI))
        Next KIGP
120     ' CONTINUE
    Next KNGP
    MASUB = True
    Exit Function
'
'
    ' ERROR EXIT.
1000
    MsgBox "IMPROPER INPUT PARAMETER VALUE(S)." & vbNewLine _
          & "IEX = " & IEX & " IC = " & IC & " ND = " & ND, vbCritical, " MASUB"
    MASUB = False
'
'
'
End Function

Private Sub MACEPS(ByRef SRELPR#)
'
'   IT CARRIES OUT AN APPROXIMATION OF MACHINE PRECISION
'
    Dim I&
'
    SRELPR = 1#
10  SRELPR = 0.5 * SRELPR
    ' THE DO LOOP STATEMENT IS NECESSARY WHEN THE ARITHMETIC UNIT HAS MORE
    ' BITS THAN IN STORAGE (INTEL 8087 FAMILY OF ARITHMETIC UNITS IS OF
    ' THIS TYPE), INFACT A DO LOOP INVOLVES A STORE FROM REGISTER TO MEMORY
    ' OF THE VALUE FOR SRELPR.
    For I = 1 To 2
    Next I
    If (SRELPR + 1# > 1#) Then GoTo 10
    SRELPR = 2# * SRELPR
'
'
'
End Sub

Private Function IF_ARI(ByVal vX As Variant) As Long
'
'   To be used as:      On IF_ARI(vX) GoTo Ln1, Ln2, Ln3
'   To implement:       If (vX) Ln1, Ln2, Ln3
'                       which is the arithmetic IF of FORTRAN.
'
'   CAUTION: in the calling program DO NOT USE the :
'   after the line numbers, otherwise you get a fatal error
'    "Unreferenced Memory Call at ...".
'
    If vX < 0 Then
        IF_ARI = 1
    ElseIf vX = 0 Then
        IF_ARI = 2
    Else
        IF_ARI = 3
    End If
'
'
'
End Function

Private Sub EXTRP(ByVal ND&, X#(), Y#(), Z#(), ByVal KC&, _
    ByRef IWK&(), ByRef A#, ByRef B#, ByRef C#, ByRef D#)
'
'   IT ADDS SOME EXTERNAL POINTS TO THE DATA POINT SET AND ESTIMATES
'   THE Z COORDINATE AT THESE POINTS FOLLOWING THE SHEPARD METHOD.
'
'   THE INPUT PARAMETERS ARE
'     ND  =     NUMBER OF DATA POINTS,
'     X,Y,Z =   ARRAYS OF DIMENSION  ND  CONTAINING THE  X,Y  AND  Z
'               COORDINATES OF THE DATA POINTS,
'     KC =      FLAG OF COMPUTATION,
'     A,B,C,D = EXTREME OF THE RECTANGULAR GRID POINTS.
'
'   THE OUTPUT PARAMETERS ARE
'     X,Y,Z =   ARRAYS OF DIMENSION  NDP  CONTAINING THE  X,Y  AND  Z
'               COORDINATES OF THE NEW SET OF DATA POINTS WHERE
'                NDP = ND+2*INT(ND/25)+4,
'     A,B,C,D = EXTREME OF THE RECTANGULAR REGION CONTAINING THE DATA
'               POINT SET AND THE RECTANGULAR GRID POINTS.
'
'   THE OTHER PARAMETER IS
'     IWK =     INTEGER ARRAY OF DIMENSION 4*ND+7 USED AS WORK AREA.
'
    Dim I&, IA&, IC&, ID&, IP1&, IP2&, IP2P1&
    Dim J&, JA&, JMX&, JB&, IB&, J1&, J2&, JC&, JD&
    Dim N&, N1&, NCP&, NDP&, NX&, NY&
    Dim ANUM#, ADEN#, BMA#, DMC#, DM#, DMX#, HX#, HY#, R4#, X0#, XP#, Y0#, YP#
    Dim DIST#(1 To 5), IPC0&(1 To 5)
'
    NDP = IWK(6)
    NCP = IWK(7)
    If (KC = 2) Then GoTo 200
'
    ' ESTIMATES THE SMALLEST RECTANGLE CONTAINING THE DATA POINT SET
    ' AND THE RECTANGULAR GRID POINTS.
    JA = 0
    IA = 7
    JB = 0
    IB = IA + ND
    JC = 0
    IC = IB + ND
    JD = 0
    ID = IC + ND
    For I = 1 To ND
        On IF_ARI(X(I) - A) GoTo 16, 18, 20
16      JA = 0
18      JA = JA + 1
        IWK(IA + JA) = I
        A = X(I)
        GoTo 30
20      On IF_ARI(B - X(I)) GoTo 26, 28, 30 'IF(B-X(I))26,28,30
26      JB = 0
28      JB = JB + 1
        IWK(IB + JB) = I
        B = X(I)
30      On IF_ARI(Y(I) - C) GoTo 36, 38, 40 'IF(Y(I)-C)36,38,40
36      JC = 0
38      JC = JC + 1
        IWK(IC + JC) = I
        C = Y(I)
        GoTo 50
40      On IF_ARI(D - Y(I)) GoTo 46, 48, 50 'IF(D-Y(I))46,48,50
46      JD = 0
48      JD = JD + 1
        IWK(ID + JD) = I
        D = Y(I)
50  ' CONTINUE
    Next I
'
    ' ESTIMATES THE NUMBER OF POINTS AND WHERE THEY HAVE TO BE ADJOINTED.
    N = ND \ 25
    BMA = B - A
    DMC = D - C
    HX = BMA
    HY = DMC
    NX = 1
    NY = 1
    If (N = 0) Then GoTo 75
    For I = 1 To N
        If (HX > HY) Then GoTo 60
        NY = NY + 1
        HY = DMC / CDbl(NY)
        GoTo 70
60      NX = NX + 1
        HX = BMA / CDbl(NX)
70      ' CONTINUE
    Next I
    HX = BMA / CDbl(NX)
    HY = DMC / CDbl(NY)
'
    ' ADDS THE NEW EXTERNAL POINTS AND CHECKS THAT THEY ARE NOT
    ' COINCIDENT WITH THE OLD ONES.
75  NDP = ND + 1
    YP = C
    For I = 1 To NY
        If (JA = 0) Then GoTo 80
        For J = 1 To JA
            If (YP = Y(IWK(IA + J))) Then GoTo 90
        Next J
80      X(NDP) = A
        Y(NDP) = YP
        NDP = NDP + 1
90      YP = YP + HY
    Next I
    XP = A
    For I = 1 To NX
        If (JD = 0) Then GoTo 110
        For J = 1 To JD
            If (XP = X(IWK(ID + J))) Then GoTo 120
        Next J
110     X(NDP) = XP
        Y(NDP) = D
        NDP = NDP + 1
120     XP = XP + HX
    Next I
    YP = D
    For I = 1 To NY
        If (JB = 0) Then GoTo 140
        For J = 1 To JB
            If (YP = Y(IWK(IB + J))) Then GoTo 150
        Next J
140     X(NDP) = B
        Y(NDP) = YP
        NDP = NDP + 1
150     YP = YP - HY
    Next I
    XP = B
    For I = 1 To NX
        If (JC = 0) Then GoTo 170
        For J = 1 To JC
            If (XP = X(IWK(IC + J))) Then GoTo 180
        Next J
170     X(NDP) = XP
        Y(NDP) = C
        NDP = NDP + 1
180     XP = XP - HX
    Next I
    NDP = NDP - 1
    IWK(6) = NDP
    NCP = 5
    If (ND <= 5) Then NCP = 3
    IWK(7) = NCP
'
    ' ESTIMATES THE FUNCTION VALUE AT THE NEW EXTERNAL POINTS.
200 N1 = ND + 1
    For IP1 = N1 To NDP
        X0 = X(IP1)
        Y0 = Y(IP1)
        DMX = 0#
        For IP2 = 1 To ND
            DM = DINF(X0, Y0, X(IP2), Y(IP2))
            DIST(IP2) = DM
            IPC0(IP2) = IP2
            If (DM <= DMX) Then GoTo 210
            DMX = DM
            JMX = IP2
210         If (IP2 >= NCP) Then GoTo 230
        Next IP2
230     IP2P1 = IP2 + 1
        For IP2 = IP2P1 To ND
            DM = DINF(X0, Y0, X(IP2), Y(IP2))
            If (DM >= DMX) Then GoTo 250
            DIST(JMX) = DM
            IPC0(JMX) = IP2
            DMX = 0#
            For J1 = 1 To NCP
                If (DIST(J1) <= DMX) Then GoTo 240
                DMX = DIST(J1)
                JMX = J1
240             ' CONTINUE
            Next J1
250         ' CONTINUE
        Next IP2
        ANUM = 0#
        ADEN = 0#
        For J2 = 1 To NCP
            IP2 = IPC0(J2)
            R4 = DSQF2(X0, Y0, X(IP2), Y(IP2))
            If (R4 = 0) Then GoTo 260
            ANUM = ANUM + Z(IP2) / R4
            ADEN = ADEN + 1# / R4
260       ' CONTINUE
        Next J2
        Z(IP1) = ANUM / ADEN
    Next IP1
'
'
'
End Sub

Private Function DINF(ByVal U1#, ByVal V1#, ByVal U2#, ByVal V2#) As Double
'
'   STATEMENT FUNCTIONS.
'   DINF(U1, V1, U2, V2) = AMAX1(Abs(U1 - U2), Abs(V1 - V2))
'
    DINF = DMAX1(Abs(U1 - U2), Abs(V1 - V2))
'
'
'
End Function
Private Function DSQF(ByVal U1#, ByVal V1#, ByVal U2#, ByVal V2#) As Double
'
'   STATEMENT FUNCTIONS.
'   DSQF(U1,V1,U2,V2)=(U2-U1)**2+(V2-V1)**2
'
    DSQF = (U2 - U1) ^ 2 + (V2 - V1) ^ 2
'
'
'
End Function
Private Function DSQF2(ByVal U1#, ByVal V1#, ByVal U2#, ByVal V2#) As Double
'
'   STATEMENT FUNCTIONS.
'   DSQF(U1, V1, U2, V2) = ((U2 - U1) ^ 2 + (V2 - V1) ^ 2) ^ 2
'
    DSQF2 = ((U2 - U1) ^ 2 + (V2 - V1) ^ 2) ^ 2
'
'
'
End Function
Private Sub CTANG(ByVal NDP&, XD#(), YD#(), ByRef NT&, _
    ByRef IPT&(), ByRef IPL&(), ByRef IWL&(), _
    ByRef IWP&(), ByRef WK#())
'
'   IT CARRIES OUT TRIANGULATION BY DIVIDING THE X-Y PLANE INTO A
'   NUMBER OF TRIANGLES ACCORDING TO THE GIVEN DATA POINTS IN THE PLANE.
'   AT THE END OF THE OPERATIONS,THE INDICES OF THE VERTEXES OF THE
'   TRIANGLES ARE LISTED COUNTER-CLOCKWISE.
'   SUBROUTINE CTANG CALLS THE MAXMN FUNCTION.
'
'   THE INPUT PARAMETERS ARE
'     NDP = NUMBER OF DATA POINTS,
'     XD  = ARRAY OF DIMENSION NDP CONTAINING THE X COORDINATES OF THE
'           DATA POINTS,
'     YD  = ARRAY OF DIMENSION NDP CONTAINING THE Y COORDINATES OF THE
'           DATA POINTS,
'     NOUT= LOGICAL UNIT NUMBER FOR THE STANDARD OUTPUT UNIT OF THE
'           SYSTEM.
'
'   THE OUTPUT PARAMETERS ARE
'     NT  = NUMBER OF TRIANGLES,
'     IPT = INTEGER ARRAY OF DIMENSION 6*NDP-15, WHERE THE INDICES OF
'           THE VERTEXES OF THE (IT)TH TRIANGLE ARE TO BE STORED AS THE
'           (3*IT-2)ND, (3*IT-1)ST AND (3*IT)TH ELEMENTS, IT=1,2,..,NT.
'
'   THE OTHER PARAMETERS ARE
'     IPL = INTEGER ARRAY OF DIMENSION 6*NDP USED INTERNALLY AS WORK
'           AREA,
'     IWL = INTEGER ARRAY OF DIMENSION 18*NDP USED INTERNALLY AS WORK
'           AREA,
'     IWP = INTEGER ARRAY OF DIMENSION NDP USED INTERNALLY AS WORK
'           AREA,
'     WK  = ARRAY OF DIMENSION NDP USED INTERNALLY AS WORK AREA.
'
    Dim I&, IP1&, IP2&, IP3&, JP1&, JP2&, NDP0&, NDPM1&, IPMN1&, IPMN2&, IP1P1&
    Dim JPMN&, ITS&, JP&, IP&, JPMX&, JPC&, NT0&, NTT3&, NL0&, NLT3&, NSH&, JWL1&
    Dim NSHT3&, JP2T3&, JP3T3&, JWL&, IPL1&, IPL2&, It&, NLN&, NLNT3&, ITT3&
    Dim IPTI&, NLF&, NTT3P3&, IREP&, ILF&, ILFT2&, NTF&, ITT3R&, IPT1&, IPT2&, IPT3&
    Dim IT1T3&, IPTI1&, IPTI2&, IT2T3&, JLT3&, IPLJ1&, IPLJ2&, NLFC&, JWL1MN&, NLFT2&
    Dim DSQMN#, x1#, y1#, DSQI#, DSQ12#, YDMP#, XDMP#, AR#, DX21#, DY21#
    Dim DXMN#, DYMN#, ARMN#, DXMX#, DYMX#, DSQMX#, ARMX#, DX#, DY#
    Dim ITF&(1 To 2)
    Const RATIO# = 0.000001, NREP& = 100
'
    ' PRELIMINARY PROCESSING.
10  NDP0 = NDP
    NDPM1 = NDP0 - 1
    If (NDP0 < 4) Then GoTo 90
'
    ' DETERMINES THE CLOSEST PAIR OF DATA POINTS AND THEIR MIDPOINTS.
20  DSQMN = DSQF(XD(1), YD(1), XD(2), YD(2))
    IPMN1 = 1
    IPMN2 = 2
    For IP1 = 1 To NDPM1
        x1 = XD(IP1)
        y1 = YD(IP1)
        IP1P1 = IP1 + 1
        For IP2 = IP1P1 To NDP0
            DSQI = DSQF(x1, y1, XD(IP2), YD(IP2))
            If (DSQI = 0#) Then GoTo 91
            If (DSQI >= DSQMN) Then GoTo 21
            DSQMN = DSQI
            IPMN1 = IP1
            IPMN2 = IP2
21      ' CONTINUE
        Next IP2
    Next IP1
    DSQ12 = DSQMN
    XDMP = (XD(IPMN1) + XD(IPMN2)) / 2#
    YDMP = (YD(IPMN1) + YD(IPMN2)) / 2#
'
    ' SORTS THE OTHER (NDP-2) DATA POINTS IN ASCENDING ORDER OF DISTANCE
    ' FROM THE MIDPOINTS AND STORES THE STORED DATA POINTS NUMBERS IN THE
    ' IWP ARRAY.
30  JP1 = 2
    For IP1 = 1 To NDP0
        If ((IP1 = IPMN1) Or (IP1 = IPMN2)) Then GoTo 31
        JP1 = JP1 + 1
        IWP(JP1) = IP1
        WK(JP1) = DSQF(XDMP, YDMP, XD(IP1), YD(IP1))
31  ' CONTINUE
    Next IP1
    For JP1 = 3 To NDPM1
        DSQMN = WK(JP1)
        JPMN = JP1
        For JP2 = JP1 To NDP0
            If (WK(JP2) >= DSQMN) Then GoTo 32
            DSQMN = WK(JP2)
            JPMN = JP2
32      ' CONTINUE
        Next JP2
        ITS = IWP(JP1)
        IWP(JP1) = IWP(JPMN)
        IWP(JPMN) = ITS
        WK(JPMN) = WK(JP1)
    Next JP1
'
    ' IF NECESSARY, MODIFIES THE ORDERING SO THAT THE
    ' FIRST THREE DATA POINTS ARE NOT COLLINEAR.
35  AR = DSQ12 * RATIO
    x1 = XD(IPMN1)
    y1 = YD(IPMN1)
    DX21 = XD(IPMN2) - x1
    DY21 = YD(IPMN2) - y1
    For JP = 3 To NDP0
        IP = IWP(JP)
        If (Abs((YD(IP) - y1) * DX21 - (XD(IP) - x1) * DY21) > AR) Then GoTo 37
    Next JP
    GoTo 92
37  If (JP = 3) Then GoTo 40
    JPMX = JP
    JP = JPMX + 1
    For JPC = 4 To JPMX
        JP = JP - 1
        IWP(JP) = IWP(JP - 1)
    Next JPC
    IWP(3) = IP
'
    ' FORMS THE FIRST TRIANGLE, STORES POINT NUMBERS OF THE VERTEXES OF
    ' THE TRIANGLES IN THE IPT ARRAY, AND STORES POINT NUMBERS OF THE
    ' BORDER LINE SEGMENTS AND THE TRIANGLE NUMBER IN THE IPL ARRAY.
40  IP1 = IPMN1
    IP2 = IPMN2
    IP3 = IWP(3)
    If (SIDE(XD(IP1), YD(IP1), XD(IP2), YD(IP2), XD(IP3), YD(IP3)) >= 0#) Then GoTo 41
    IP1 = IPMN2
    IP2 = IPMN1
41  NT0 = 1
    NTT3 = 3
    IPT(1) = IP1
    IPT(2) = IP2
    IPT(3) = IP3
    NL0 = 3
    NLT3 = 9
    IPL(1) = IP1
    IPL(2) = IP2
    IPL(3) = 1
    IPL(4) = IP2
    IPL(5) = IP3
    IPL(6) = 1
    IPL(7) = IP3
    IPL(8) = IP1
    IPL(9) = 1
'
    ' ADDS THE REMAINING (NDP-3) DATA POINTS, ONE BY ONE.
50  For JP1 = 4 To NDP0
        IP1 = IWP(JP1)
        x1 = XD(IP1)
        y1 = YD(IP1)
'
        ' DETERMINES THE VISIBLE LINE SEGMENTS.
        IP2 = IPL(1)
        JPMN = 1
        DXMN = XD(IP2) - x1
        DYMN = YD(IP2) - y1
        DSQMN = DXMN ^ 2 + DYMN ^ 2
        ARMN = DSQMN * RATIO
        JPMX = 1
        DXMX = DXMN
        DYMX = DYMN
        DSQMX = DSQMN
        ARMX = ARMN
        For JP2 = 2 To NL0
            IP2 = IPL(3 * JP2 - 2)
            DX = XD(IP2) - x1
            DY = YD(IP2) - y1
            AR = DY * DXMN - DX * DYMN
            If (AR > ARMN) Then GoTo 51
            DSQI = DX ^ 2 + DY ^ 2
            If ((AR >= (-ARMN)) And (DSQI >= DSQMN)) Then GoTo 51
            JPMN = JP2
            DXMN = DX
            DYMN = DY
            DSQMN = DSQI
            ARMN = DSQMN * RATIO
51          AR = DY * DXMX - DX * DYMX
            If (AR < (-ARMX)) Then GoTo 52
            DSQI = DX ^ 2 + DY ^ 2
            If ((AR <= ARMX) And (DSQI >= DSQMX)) Then GoTo 52
            JPMX = JP2
            DXMX = DX
            DYMX = DY
            DSQMX = DSQI
            ARMX = DSQMX * RATIO
52      ' CONTINUE
        Next JP2
        If (JPMX < JPMN) Then JPMX = JPMX + NL0
        NSH = JPMN - 1
        If (NSH <= 0) Then GoTo 60
'
        ' SHIFTS (ROTATES) THE IPL ARRAY SO THAT THE INVISIBLE BORDER LINE
        ' SEGMENTS ARE CONTAINED IN THE FIRST PART OF THE IPL ARRAY.
        NSHT3 = NSH * 3
        For JP2T3 = 3 To NSHT3 Step 3
            JP3T3 = JP2T3 + NLT3
            IPL(JP3T3 - 2) = IPL(JP2T3 - 2)
            IPL(JP3T3 - 1) = IPL(JP2T3 - 1)
            IPL(JP3T3) = IPL(JP2T3)
        Next JP2T3
        For JP2T3 = 3 To NLT3 Step 3
            JP3T3 = JP2T3 + NSHT3
            IPL(JP2T3 - 2) = IPL(JP3T3 - 2)
            IPL(JP2T3 - 1) = IPL(JP3T3 - 1)
            IPL(JP2T3) = IPL(JP3T3)
        Next JP2T3
        JPMX = JPMX - NSH
'
        ' ADDS TRIANGLES TO THE IPT ARRAY, UPDATES BORDER LINE SEGMENTS IN
        ' THE IPL ARRAY, AND SETS FLAGS FOR THE BORDER LINE SEGMENTS TO BE
        ' REEXAMINED IN THE IWL ARRAY.
60      JWL = 0
        For JP2 = JPMX To NL0
            JP2T3 = JP2 * 3
            IPL1 = IPL(JP2T3 - 2)
            IPL2 = IPL(JP2T3 - 1)
            It = IPL(JP2T3)
'
            ' ADDS A TRIANGLE TO THE IPT ARRAY.
            NT0 = NT0 + 1
            NTT3 = NTT3 + 3
            IPT(NTT3 - 2) = IPL2
            IPT(NTT3 - 1) = IPL1
            IPT(NTT3) = IP1
'
            ' UPDATES THE BORDER LINE SEGMENTS IN THE IPL ARRAY.
            If (JP2 <> JPMX) Then GoTo 61
            IPL(JP2T3 - 1) = IP1
            IPL(JP2T3) = NT0
61          If (JP2 <> NL0) Then GoTo 62
            NLN = JPMX + 1
            NLNT3 = NLN * 3
            IPL(NLNT3 - 2) = IP1
            IPL(NLNT3 - 1) = IPL(1)
            IPL(NLNT3) = NT0
'
            ' DETERMINES THE VERTEX THAT DOES NOT LIE ON THE BORDER LINE SEGMENTS.
62          ITT3 = It * 3
            IPTI = IPT(ITT3 - 2)
            If ((IPTI <> IPL1) And (IPTI <> IPL2)) Then GoTo 63
            IPTI = IPT(ITT3 - 1)
            If ((IPTI <> IPL1) And (IPTI <> IPL2)) Then GoTo 63
            IPTI = IPT(ITT3)
'
            ' CHECKS WHETHER THE EXCHANGE IS NECESSARY.
63          If (MAXMN(XD, YD, IP1, IPTI, IPL1, IPL2) = 0) Then GoTo 64
'
            ' MODIFIES THE IPT ARRAY WHEN NECESSARY.
            IPT(ITT3 - 2) = IPTI
            IPT(ITT3 - 1) = IPL1
            IPT(ITT3) = IP1
            IPT(NTT3 - 1) = IPTI
            If (JP2 = JPMX) Then IPL(JP2T3) = It
            If ((JP2 = NL0) And (IPL(3) = It)) Then IPL(3) = NT0

            ' SETS FLAGS IN THE IWL ARRAY.
            JWL = JWL + 4
            IWL(JWL - 3) = IPL1
            IWL(JWL - 2) = IPTI
            IWL(JWL - 1) = IPTI
            IWL(JWL) = IPL2
64      ' CONTINUE
        Next JP2
        NL0 = NLN
        NLT3 = NLNT3
        NLF = JWL \ 2
        If (NLF = 0) Then GoTo 79
'
        ' IMPROVES THE TRIANGULATION.
70      NTT3P3 = NTT3 + 3
        For IREP = 1 To NREP
            For ILF = 1 To NLF
                ILFT2 = ILF * 2
                IPL1 = IWL(ILFT2 - 1)
                IPL2 = IWL(ILFT2)
'
                ' LOCATES IN THE IPT ARRAY TWO TRIANGLES ON BOTH SIDES OF THE
                ' FLAGGED LINE SEGMENT.
                NTF = 0
                For ITT3R = 3 To NTT3 Step 3
                    ITT3 = NTT3P3 - ITT3R
                    IPT1 = IPT(ITT3 - 2)
                    IPT2 = IPT(ITT3 - 1)
                    IPT3 = IPT(ITT3)
                    If ((IPL1 <> IPT1) And (IPL1 <> IPT2) And (IPL1 <> IPT3)) Then GoTo 71
                    If ((IPL2 <> IPT1) And (IPL2 <> IPT2) And (IPL2 <> IPT3)) Then GoTo 71
                    NTF = NTF + 1
                    ITF(NTF) = ITT3 \ 3
                    If (NTF = 2) Then GoTo 72
71              ' CONTINUE
                Next ITT3R
                If (NTF < 2) Then GoTo 76
'
                ' DETERMINES THE VERTEXES OF THE TRIANGLES THAT DO NOT LIE ON THE
                ' LINE SEGMENT.
72              IT1T3 = ITF(1) * 3
                IPTI1 = IPT(IT1T3 - 2)
                If ((IPTI1 <> IPL1) And (IPTI1 <> IPL2)) Then GoTo 73
                IPTI1 = IPT(IT1T3 - 1)
                If ((IPTI1 <> IPL1) And (IPTI1 <> IPL2)) Then GoTo 73
                IPTI1 = IPT(IT1T3)
73              IT2T3 = ITF(2) * 3
                IPTI2 = IPT(IT2T3 - 2)
                If ((IPTI2 <> IPL1) And (IPTI2 <> IPL2)) Then GoTo 74
                IPTI2 = IPT(IT2T3 - 1)
                If ((IPTI2 <> IPL1) And (IPTI2 <> IPL2)) Then GoTo 74
                IPTI2 = IPT(IT2T3)
'
                ' CHECKS WHETHER THE EXCHANGE IS NECESSARY.
74              If (MAXMN(XD, YD, IPTI1, IPTI2, IPL1, IPL2) = 0) Then GoTo 76
'
                ' MODIFIES THE IPT ARRAY WHEN NECESSARY.
                IPT(IT1T3 - 2) = IPTI1
                IPT(IT1T3 - 1) = IPTI2
                IPT(IT1T3) = IPL1
                IPT(IT2T3 - 2) = IPTI2
                IPT(IT2T3 - 1) = IPTI1
                IPT(IT2T3) = IPL2
'
                ' SETS NEW FLAGS.
                JWL = JWL + 8
                IWL(JWL - 7) = IPL1
                IWL(JWL - 6) = IPTI1
                IWL(JWL - 5) = IPTI1
                IWL(JWL - 4) = IPL2
                IWL(JWL - 3) = IPL2
                IWL(JWL - 2) = IPTI2
                IWL(JWL - 1) = IPTI2
                IWL(JWL) = IPL1
                For JLT3 = 3 To NLT3 Step 3
                    IPLJ1 = IPL(JLT3 - 2)
                    IPLJ2 = IPL(JLT3 - 1)
                    If (((IPLJ1 = IPL1) And (IPLJ2 = IPTI2)) _
                    Or ((IPLJ2 = IPL1) And (IPLJ1 = IPTI2))) _
                    Then IPL(JLT3) = ITF(1)
                    If (((IPLJ1 = IPL2) And (IPLJ2 = IPTI1)) _
                    Or ((IPLJ2 = IPL2) And (IPLJ1 = IPTI1))) _
                    Then IPL(JLT3) = ITF(2)
                Next JLT3
76          ' CONTINUE
            Next ILF
            NLFC = NLF
            NLF = JWL \ 2
            If (NLF = NLFC) Then GoTo 79
'
            ' RESETS THE IWL ARRAY FOR THE NEXT ROUND.
            JWL = 0
            JWL1MN = (NLFC + 1) * 2
            NLFT2 = NLF * 2
            For JWL1 = JWL1MN To NLFT2 Step 2
                JWL = JWL + 2
                IWL(JWL - 1) = IWL(JWL1 - 1)
                IWL(JWL) = IWL(JWL1)
            Next JWL1
            NLF = JWL \ 2
        Next IREP
79  ' CONTINUE
    Next JP1
'
    ' REARRANGES THE IPT ARRAY SO THAT THE VERTEXES OF EACH TRIANGLE ARE
    ' LISTED COUNTER-CLOCKWISE.
80  For ITT3 = 3 To NTT3 Step 3
        IP1 = IPT(ITT3 - 2)
        IP2 = IPT(ITT3 - 1)
        IP3 = IPT(ITT3)
        If (SIDE(XD(IP1), YD(IP1), XD(IP2), YD(IP2), XD(IP3), YD(IP3)) >= 0#) Then GoTo 81
        IPT(ITT3 - 2) = IP2
        IPT(ITT3 - 1) = IP1
81    ' CONTINUE
    Next ITT3
    NT = NT0
'
    Dim N&, I1&, I2&, I3&, L&
    For N = 1 To NT
        L = 3 * N
        I1 = IPT(L - 2)
        I2 = IPT(L - 1)
        I3 = IPT(L)
    Next N
    Exit Sub
'
'
    ' ERROR EXIT.
90
    MsgBox " ***   NDP LESS THAN 4." & vbNewLine _
             & NDP0, vbCritical, " CTANG"
    GoTo 93
91
    MsgBox " ***   IDENTICAL DATA POINTS." & vbNewLine _
          & NDP0 & " " & IP1 & " " & IP2 & " " & x1 & " " & y1 _
          , vbCritical, " CTANG"
    GoTo 93
92
    MsgBox " ***   ALL COLLINEAR DATA POINTS." & vbNewLine _
          & NDP0, vbCritical, " CTANG"
93
    'MsgBox " ERROR DETECTED IN ROUTINE CTANG"
    NT = 0
'
'
'
End Sub

Private Function SIDE(ByVal U1#, ByVal V1#, ByVal U2#, ByVal V2#, _
    ByVal U3#, ByVal V3#) As Double
'
'   STATEMENT FUNCTIONS.
'   SIDE(U1, V1, U2, V2, U3, V3) = (V3 - V1) * (U2 - U1) - (U3 - U1) * (V2 - V1)
'
    SIDE = (V3 - V1) * (U2 - U1) - (U3 - U1) * (V2 - V1)
'
'
'
End Function

Private Function MAXMN(X#(), Y#(), ByVal I1&, ByVal I2&, ByVal I3&, ByVal I4&) As Long
'
'   IT DETERMINES WHETHER THE EXCHANGE OF TWO TRIANGLES IS NECESSARY
'   OR NOT ON THE BASIS OF THE MAX-MIN-ANGLE CRITERION BY C.LAWSON.
'
'   THE INPUT PARAMETERS ARE
'     X,Y =         ARRAYS CONTAINING THE COORDINATES OF THE DATA POINTS,
'     I1,I2,I3,I4 = POINT NUMBERS OF FOUR POINTS P1,P2,P3 AND P4
'                   FORMING A QUADRILATERAL WITH P3 AND P4
'                   DIAGONALLY CONNECTED.
'
'   FUNCTION MAXMN RETURNS AN INTEGER VALUE 1 (ONE) WHEN AN EXCHANGE IS
'   NECESSARY, OTHERWISE 0 (ZERO).
'
    Dim IDX&
    Dim U1#, U2#, U3#, U4#
    Dim x1#, y1#, x2#, y2#, X3#, Y3#, X4#, Y4#
    Dim A1SQ#, A2SQ#, A3SQ#, A4SQ#
    Dim B1SQ#, B2SQ#, B3SQ#, B4SQ#
    Dim C1SQ#, C2SQ#, C3SQ#, C4SQ#
    Dim S1SQ#, S2SQ#, S3SQ#, S4SQ#
'
    ' PRELIMINARY PROCESSING.
    x1 = X(I1)
    y1 = Y(I1)
    x2 = X(I2)
    y2 = Y(I2)
    X3 = X(I3)
    Y3 = Y(I3)
    X4 = X(I4)
    Y4 = Y(I4)
'
    ' CALCULATION.
    IDX = 0
    U3 = (y2 - Y3) * (x1 - X3) - (x2 - X3) * (y1 - Y3)
    U4 = (y1 - Y4) * (x2 - X4) - (x1 - X4) * (y2 - Y4)
    If (U3 * U4 <= 0#) Then GoTo 10
    U1 = (Y3 - y1) * (X4 - x1) - (X3 - x1) * (Y4 - y1)
    U2 = (Y4 - y2) * (X3 - x2) - (X4 - x2) * (Y3 - y2)
'
      'EQUIVALENCE (C2SQ,C1SQ),(A3SQ,B2SQ),(B3SQ,A1SQ),
     '1            (A4SQ,B1SQ),(B4SQ,A2SQ),(C4SQ,C3SQ)
'
    A1SQ = (x1 - X3) ^ 2 + (y1 - Y3) ^ 2: B3SQ = A1SQ
    B1SQ = (X4 - x1) ^ 2 + (Y4 - y1) ^ 2: A4SQ = B1SQ
    C1SQ = (X3 - X4) ^ 2 + (Y3 - Y4) ^ 2: C2SQ = C1SQ
    A2SQ = (x2 - X4) ^ 2 + (y2 - Y4) ^ 2: B4SQ = A2SQ
    B2SQ = (X3 - x2) ^ 2 + (Y3 - y2) ^ 2: A3SQ = B2SQ
    C3SQ = (x2 - x1) ^ 2 + (y2 - y1) ^ 2: C4SQ = C3SQ
'
    S1SQ = U1 * U1 / (C1SQ * DMAX1(A1SQ, B1SQ))
    S2SQ = U2 * U2 / (C2SQ * DMAX1(A2SQ, B2SQ))
    S3SQ = U3 * U3 / (C3SQ * DMAX1(A3SQ, B3SQ))
    S4SQ = U4 * U4 / (C4SQ * DMAX1(A4SQ, B4SQ))
    If (DMIN1(S1SQ, S2SQ) < DMIN1(S3SQ, S4SQ)) Then IDX = 1
'
10  MAXMN = IDX
'
'
'
End Function

Private Sub ADJAC(ByVal NT&, IPT&(), ByVal N&, _
    ByRef IADVE&(), ByRef NADVE&())
'
'   IT ESTIMATES THE ADJACENCIES MONODIMENSIONAL ARRAY CONTAINING
'   FOR EACH VERTEX THE INDICES OF THE VERTEXES ADJACENT IN THE
'   TRIANGULATION.
'
'   THE INPUT PARAMETERS ARE
'     NT  =   NUMBER OF TRIANGLES,
'     IPT =   INTEGER ARRAY OF DIMENSION 3*NT CONTAINING THE INDICES
'             OF THE VERTEXES OF THE TRIANGLES,
'     N =     NUMBER OF DATA POINTS.
'
'   THE OUTPUT PARAMETERS ARE
'     IADVE = INTEGER ARRAY OF DIMENSION  6*N-12  CONTAINING FOR EACH
'             VERTEX THE INDICES OF THE VERTEXES ADJACENT IN THE
'             TRIANGULATION,
'     NADVE = INTEGER ARRAY OF DIMENSION  N+1  CONTAINING FOR EACH
'             VERTEX THE NUMBER OF THE VERTEXES ADJACENT IN THE
'             TRIANGULATION.
'
    Dim I&, I1&, I2&, J&, J1&, J2&, JIN&, JFIN&, L&, KIN&, NT3&
    'Dim ITEM&(1 To 30)
    ReDim ITEM&(1 To MAX0(30, NT)) ' To treat large numbers of triangulations.
'
    NADVE(1) = 0
    KIN = 0
    NT3 = 3 * NT
    For I = 1 To N
        I2 = 0
        ' STORES THE INDICES OF THE ADJACENT VERTEXES.
        For J1 = 1 To NT3 Step 3
            J2 = J1 + 2
            For J = J1 To J2
                If (I = IPT(J)) Then GoTo 20
            Next J
            GoTo 30
20          I1 = I2 + 1
            I2 = I2 + 2
            ITEM(I1) = IPT(J1 + (J Mod 3))
            ITEM(I2) = IPT(J1 + ((J + 1) Mod 3))
30      ' CONTINUE
        Next J1
'
        ' DISCARDS THE INDICES THAT HAVE BEEN STORED TWICE.
        JIN = KIN + 1
        KIN = KIN + 2
        JFIN = KIN
        IADVE(JIN) = ITEM(1)
        IADVE(JFIN) = ITEM(2)
        If (I2 = 2) Then GoTo 60
        For J = 3 To I2
            For L = JIN To JFIN
                If (ITEM(J) = IADVE(L)) Then GoTo 50
            Next L
            KIN = KIN + 1
            IADVE(KIN) = ITEM(J)
            JFIN = KIN
50      ' CONTINUE
        Next J
60      NADVE(I + 1) = KIN
    Next I
'
'
'
End Sub

Private Function PDSTE(ByVal N&, X#(), Y#(), Z#(), ByVal NT&, _
    IPT&(), ByRef PD#(), ByRef IPD&()) As Boolean
'
'   IT ESTIMATES THE FIRST ORDER PARTIAL DERIVATIVE VALUES
'   AT THE DATA POINTS FOLLOWING THE KLUCEWICZ METHOD.
'
'   THE INPUT PARAMETERS ARE
'     N   = NUMBER OF DATA POINTS,
'     X,Y,Z = ARRAYS OF DIMENSION N CONTAINING THE X,Y AND Z
'             COORDINATES OF THE DATA POINTS,
'     NT  = NUMBER OF TRIANGLES,
'     IPT = INTEGER ARRAY OF DIMENSION 3*NT CONTAINING THE INDICES
'           OF THE VERTEXES OF THE TRIANGLES.
'
'   THE OUTPUT PARAMETER IS
'     PD  = ARRAY OF DIMENSION 2*N CONTAINING THE PARTIAL DERIVATIVE
'           VALUES AT THE DATA POINTS.
'
'   THE OTHER PARAMETER IS
'     IPD = INTEGER ARRAY OF DIMENSION N USED INTERNALLY AS WORK AREA.
'
    Dim I&, I1&, I2&, I3&, J&, L&, N2&
    Dim C#, DEN#, DX#, DY#, X21#, X31#, Y21#, Y31#, Z21#, Z31#
'
    'On Error GoTo PDSTE_ERR
    On Error GoTo 0
'
    ' PRELIMINARY PROCESSING.
    N2 = N + N
    For I = 1 To N2
        PD(I) = 0#
        IPD(I) = 0
    Next I
'
    ' ESTIMATES FOR EACH TRIANGLE THE SLOPES OF THE PLANE THROUGH THE
    ' FUNCTION'S VALUE AT THE VERTEXES.
    For I = 1 To NT
        L = 3 * I
        I1 = IPT(L - 2)
        I2 = IPT(L - 1)
        I3 = IPT(L)
        X21 = X(I2) - X(I1)
        X31 = X(I3) - X(I1)
        Y21 = Y(I2) - Y(I1)
        Y31 = Y(I3) - Y(I1)
        Z21 = Z(I2) - Z(I1)
        Z31 = Z(I3) - Z(I1)
        C = Y21 * X31 - X21 * Y31
        DX = (Y21 * Z31 - Z21 * Y31) / C
        DY = (Z21 * X31 - X21 * Z31) / C
'
    ' UPDATES THE IPD AND PD ARRAYS.
        IPD(I1) = IPD(I1) + 1
        I1 = I1 + I1
        PD(I1 - 1) = PD(I1 - 1) + DX
        PD(I1) = PD(I1) + DY
        IPD(I2) = IPD(I2) + 1
        I2 = I2 + I2
        PD(I2 - 1) = PD(I2 - 1) + DX
        PD(I2) = PD(I2) + DY
        IPD(I3) = IPD(I3) + 1
        I3 = I3 + I3
        PD(I3 - 1) = PD(I3 - 1) + DX
        PD(I3) = PD(I3) + DY
    Next I
'
    ' AVERAGES THE DERIVATIVE VALUES STORED IN THE PD ARRAY.
    J = 0
    For I = 2 To N2 Step 2
        J = J + 1
        DEN = IPD(J)
        PD(I - 1) = PD(I - 1) / DEN
        PD(I) = PD(I) / DEN
    Next I
'
'
PDSTE_ERR:
    PDSTE = (Err = 0)
'
'
'
End Function

Private Function PDMIN(ByVal SRELPR#, ByVal N&, X#(), Y#(), Z#(), _
    IADVE&(), NADVE&(), ByRef INDEX&(), _
    ByVal TP#, ByRef PD#(), _
    ByRef ALFA#(), ByRef BETA#(), ByRef GAMMA#(), _
    ByRef EIJQ#(), _
    ByRef ALFAS#(), ByRef BETAS#(), ByRef GAMMAS#(), _
    ByRef ZX#(), ByRef ZY#()) As Boolean
'
'   IT ESTIMATES THE FIRST ORDER PARTIAL DERIVATIVE VALUES AT
'   THE DATA POINTS BY MEANS OF A GLOBAL METHOD BASED ON A MINIMUM
'   NORM UNDER TENSION NETWORK .
'
'   THE INPUT PARAMETERS ARE
'     SRELPR = SINGLE RELATIVE PRECISION
'     N   = NUMBER OF DATA POINTS,
'     X,Y,Z = ARRAY OF DIMENSION N CONTAINING THE X,Y AND Z
'             COORDINATES OF THE DATA POINTS,
'     NOUT = LOGICAL UNIT NUMBER FOR THE STANDARD OUTPUT UNIT OF THE
'            SYSTEM,
'     IADVE = INTEGER ARRAY OF DIMENSION 6*N-12 CONTAINING THE INDICES
'             OF THE VERTEXES ADJACENT TO EACH VERTEX IN THE
'             TRIANGULATION,
'     NADVE = INTEGER ARRAY OF DIMENSION N+1 CONTAINING THE NUMBER OF
'             THE VERTEXES ADJACENT TO EACH VERTEX IN THE TRIANGULATION
'     TP  = TENSION PARAMETER,
'     PD  = ARRAY OF DIMENSION 2*N CONTAINING AN INITIAL EVALUATION
'           OF THE PARTIAL DERIVATIVE VALUES AT THE DATA POINTS.
'
'   THE OUTPUT PARAMETER IS
'     PD  = ARRAY OF DIMENSION 2*N CONTAINING THE PARTIAL DERIVATIVE
'           VALUES AT THE DATA POINTS.
'
'   THE OTHER PARAMETERS ARE
'     INDEX = INTEGER ARRAY OF DIMENSION 6*N-15 USED INTERNALLY AS
'             WORK AREA,
'     ALFA,BETA,GAMMA,EIJQ = ARRAYS OF DIMENSION 3*N-6 USED INTERNALLY
'                            AS WORK AREAS,
'     ALFAS,BETAS,GAMMAS,ZX,ZY = ARRAYS OF DIMENSION N USED INTERNALLY
'                                AS WORK AREAS.
'
'   THE  RELER  CONSTANT IN THE DATA  INITIALIZATION STATEMENT  IS A
'   RELATIVE ERROR TOLERANCE TO STOP THE ITERATIVE METHOD.
'   THEREFORE IT IS MACHINE DEPENDENT;
'   THE ABSOLUTE ERROR TOLERANCE  TAU  IS THEN OBTAINED BY
'         TAU=RELER*AMAX1(ABS(PD(I)),I=1,2*N)+2*N*SRELPR.
'
    Dim I&, IND&, J&, JIN&, JFIN&, K&, L&, LIN&, LFIN&, M&, ITER&, IPI&
    Dim DX#, DXQ#, DY#, DYQ#, DZ#, PDM#, ALUNQ#, AL3#, AL3P2#
    Dim TPQ#, TPQ15#, TPQ40#, TPQ60#, AL#, BE#, GA#, Z1#, Z2#
    Dim SQEIJQ#, tau#, ERQ#, ADX#, BDX#, BDY#, GDY#, det#, S1#, S2#
    Dim PDXN#, PDYN#, ERX#, ERY#
    Const RELER# = 0.00001
'
    ' CALCULATES THE PART OF MATRIX COEFFICIENTS INDEPENDENT
    ' FROM THE TENSION PARAMETER TP.
    K = 0
    PDM = 0#
    For I = 1 To N
        J = I + I
        PDM = DMAX1(PDM, Abs(PD(J)), Abs(PD(J - 1)))
        ZX(I) = 0#
        ZY(I) = 0#
        ALFAS(I) = 0#
        BETAS(I) = 0#
        GAMMAS(I) = 0#
        JIN = NADVE(I) + 1
        JFIN = NADVE(I + 1)
        For J = JIN To JFIN
            IND = IADVE(J)
            DX = X(I) - X(IND)
            DY = Y(I) - Y(IND)
            DZ = Z(I) - Z(IND)
            DXQ = DX * DX
            DYQ = DY * DY
            ALUNQ = DXQ + DYQ
            AL3 = ALUNQ * Sqr(ALUNQ)
            ZX(I) = ZX(I) + DZ * DX / AL3
            ZY(I) = ZY(I) + DZ * DY / AL3
            If (IND > I) Then GoTo 30
            LIN = NADVE(IND) + 1
            LFIN = NADVE(IND + 1)
            For L = LIN To LFIN
                If (I = IADVE(L)) Then GoTo 20
            Next L
20          INDEX(J) = INDEX(L)
            GoTo 40
30          K = K + 1
            INDEX(J) = K
            AL3P2 = AL3 + AL3
            EIJQ(K) = ALUNQ
            ALFA(K) = DXQ / AL3P2
            BETA(K) = DX * DY / AL3P2
            GAMMA(K) = DYQ / AL3P2
40          ALFAS(I) = ALFAS(I) + ALFA(INDEX(J))
            BETAS(I) = BETAS(I) + BETA(INDEX(J))
            GAMMAS(I) = GAMMAS(I) + GAMMA(INDEX(J))
        Next J
        ZX(I) = 3# * ZX(I) / 2#
        ZY(I) = 3# * ZY(I) / 2#
        ALFAS(I) = 2# * ALFAS(I)
        BETAS(I) = 2# * BETAS(I)
        GAMMAS(I) = 2# * GAMMAS(I)
    Next I
    If (TP = 0) Then GoTo 100
'
    ' CALCULATES THE PART OF MATRIX COEFFICIENTS DEPENDING FROM
    ' THE TENSION PARAMETER TP.
    TPQ = TP * TP
    TPQ60 = TPQ / 60#
    TPQ40 = TPQ / 40#
    TPQ15 = TPQ / 15#
    For I = 1 To N
        AL = 0#
        BE = 0#
        GA = 0#
        Z1 = 0#
        Z2 = 0#
        JIN = NADVE(I) + 1
        JFIN = NADVE(I + 1)
        For J = JIN To JFIN
            K = INDEX(J)
            IND = IADVE(J)
            DZ = Z(I) - Z(IND)
            AL = AL + ALFA(K) * EIJQ(K)
            BE = BE + BETA(K) * EIJQ(K)
            GA = GA + GAMMA(K) * EIJQ(K)
            SQEIJQ = Sqr(EIJQ(K))
            Z1 = Z1 + (X(I) - X(IND)) * DZ / SQEIJQ
            Z2 = Z2 + (Y(I) - Y(IND)) * DZ / SQEIJQ
        Next J
        ALFAS(I) = ALFAS(I) + TPQ15 * AL
        BETAS(I) = BETAS(I) + TPQ15 * BE
        GAMMAS(I) = GAMMAS(I) + TPQ15 * GA
        ZX(I) = ZX(I) + TPQ40 * Z1
        ZY(I) = ZY(I) + TPQ40 * Z2
    Next I
    M = NADVE(N + 1) \ 2
    For I = 1 To M
        ALFA(I) = ALFA(I) - TPQ60 * ALFA(I) * EIJQ(I)
        BETA(I) = BETA(I) - TPQ60 * BETA(I) * EIJQ(I)
        GAMMA(I) = GAMMA(I) - TPQ60 * GAMMA(I) * EIJQ(I)
    Next I
'
    ' CALCULATES THE SOLUTIONS OF THE SYSTEM FOLLOWING THE GAUSS-SIEDEL
    ' METHOD.
100 ITER = 1
    tau = RELER * PDM + 2 * N * SRELPR
110 ERQ = 0#
    For I = 1 To N
        ADX = 0#
        BDX = 0#
        BDY = 0#
        GDY = 0#
        JIN = NADVE(I) + 1
        JFIN = NADVE(I + 1)
        For J = JIN To JFIN
            IND = 2 * IADVE(J)
            K = INDEX(J)
            ADX = ADX + ALFA(K) * PD(IND - 1)
            BDY = BDY + BETA(K) * PD(IND)
            BDX = BDX + BETA(K) * PD(IND - 1)
            GDY = GDY + GAMMA(K) * PD(IND)
        Next J
        det = ALFAS(I) * GAMMAS(I) - BETAS(I) * BETAS(I)
        S1 = ADX + BDY - ZX(I)
        S2 = BDX + GDY - ZY(I)
        PDXN = (-GAMMAS(I) * S1 + BETAS(I) * S2) / det
        PDYN = (BETAS(I) * S1 - ALFAS(I) * S2) / det
        IPI = I + I
        ERX = PDXN - PD(IPI - 1)
        ERY = PDYN - PD(IPI)
        ERQ = ERQ + ERX * ERX + ERY * ERY
        PD(IPI - 1) = PDXN
        PD(IPI) = PDYN
    Next I
'
    ' CHECKS WHETHER CONVERGENCE IS REACHED WITH THE PRESCRIBED TOLERANCE.
    If (ERQ < tau) Then
        PDMIN = True
        Exit Function
    End If
    If (ITER = 20) Then GoTo 150
    ITER = ITER + 1
    GoTo 110
'
'
    ' ERROR EXIT.
150
    MsgBox "MINIMIZATION NOT COMPLETED", vbCritical, " PDMIN"
    PDMIN = False
'
'
'
End Function

Private Sub ORDGR(XD#(), YD#(), ByVal NT&, IPT&(), _
    ByVal NXI&, ByVal NYI&, XI#(), YI#(), ByRef NGP&(), ByRef IGP&())
'
'   IT ORGANIZES GRID POINTS FOR SURFACE RECONSTRUCTION BY
'   SORTING THEM ACCORDING TO THEIR BELONGING TO THE TRIANGLES.
'
'   THE INPUT PARAMETERS ARE
'     XD,YD = ARRAY OF DIMENSION N CONTAINING THE X AND Y COORDINATES
'             OF THE DATA POINTS, WHERE N IS THE NUMBER OF THE DATA
'             POINTS,
'     NT  =   NUMBER OF TRIANGLES,
'     IPT =   INTEGER ARRAY OF DIMENSION 3*NT CONTAINING THE INDICES OF
'             THE VERTEXES OF THE TRIANGLES,
'     NXI =   NUMBER OF GRID POINTS IN THE X COORDINATES,
'     NYI =   NUMBER OF GRID POINTS IN THE Y COORDINATES,
'     XI,YI = ARRAY OF DIMENSION NXI AND NYI CONTAINING THE X AND Y
'             COORDINATES OF THE GRID POINTS,RESPECTIVELY.
'
'   THE OUTPUT PARAMETERS ARE
'     NGP =   INTEGER ARRAY OF DIMENSION 2*NT WHERE THE NUMBER OF GRID
'             POINTS BELONGING TO EACH TRIANGLE IS TO BE STORED,
'     IGP =   INTEGER ARRAY OF DIMENSION NXI*NYI WHERE THE INDICES OF THE
'             GRID POINTS ARE TO BE STORED ACCORDING TO THEIR BELONGING
'             TO THE TRIANGLES CONSIDERED IN ASCENDING ORDER NUMBERS.
'
    Dim NT0&, NXI0&, NYI0&, NXINYI&, JNGP0&, JNGP1&, JIGP0&, JIGP1&, IT0&
    Dim NGP0&, NGP1&, IT0T3&, IP1&, IP2&, IP3&, INSD&, IXI&, IYI&, L&, IZI&
    Dim IXIMX&, IXIMN&, JIGP1I&
    Dim x1#, y1#, x2#, y2#, X3#, Y3#, XII#, XMN#, XMX#, YMN#, YMX#, YII#
'
    Dim N&, I1&, I2&, I3&
    For N = 1 To NT
        L = 3 * N
        I1 = IPT(L - 2)
        I2 = IPT(L - 1)
        I3 = IPT(L)
    Next N
'
    NT0 = NT
'
    ' PRELIMINARY PROCESSING.
    NXI0 = NXI
    NYI0 = NYI
    NXINYI = NXI0 * NYI0
'
    ' DETERMINES GRID POINTS INSIDE THE DATA AREA.
    JNGP0 = 0
    JNGP1 = 2 * NT0 + 1
    JIGP0 = 0
    JIGP1 = NXINYI + 1
    For IT0 = 1 To NT0
        NGP0 = 0
        NGP1 = 0
        IT0T3 = IT0 * 3
        IP1 = IPT(IT0T3 - 2)
        IP2 = IPT(IT0T3 - 1)
        IP3 = IPT(IT0T3)
        x1 = XD(IP1)
        y1 = YD(IP1)
        x2 = XD(IP2)
        y2 = YD(IP2)
        X3 = XD(IP3)
        Y3 = YD(IP3)
        XMN = DMIN1(x1, x2, X3)
        XMX = DMAX1(x1, x2, X3)
        YMN = DMIN1(y1, y2, Y3)
        YMX = DMAX1(y1, y2, Y3)
        INSD = 0
        For IXI = 1 To NXI0
            If (XI(IXI) >= XMN And XI(IXI) <= XMX) Then GoTo 10
            If (INSD = 0) Then GoTo 20
            IXIMX = IXI - 1
            GoTo 30
10          If (INSD = 1) Then GoTo 20
            INSD = 1
            IXIMN = IXI
20      ' CONTINUE
        Next IXI
        If (INSD = 0) Then GoTo 150
        IXIMX = NXI0
30      For IYI = 1 To NYI0 ' DO 140
            YII = YI(IYI)
            If (YII < YMN Or YII > YMX) Then GoTo 140
            For IXI = IXIMN To IXIMX    ' DO 130
                XII = XI(IXI)
                L = 0
                On IF_ARI(SIDE(x1, y1, x2, y2, XII, YII)) GoTo 130, 40, 50 'IF (SIDE(X1,Y1,X2,Y2,XII,YII)) 130, 40, 50
40              L = 1
50              On IF_ARI(SIDE(x2, y2, X3, Y3, XII, YII)) GoTo 130, 60, 70 'IF (SIDE(X2,Y2,X3,Y3,XII,YII)) 130, 60, 70
60              L = 1
70              On IF_ARI(SIDE(X3, Y3, x1, y1, XII, YII)) GoTo 130, 80, 90 'IF (SIDE(X3,Y3,X1,Y1,XII,YII)) 130, 80, 90
80              L = 1
90              IZI = NXI0 * (IYI - 1) + IXI
                If (L = 1) Then GoTo 100
                NGP0 = NGP0 + 1
                JIGP0 = JIGP0 + 1
                IGP(JIGP0) = IZI
                GoTo 130
100             If (JIGP1 > NXINYI) Then GoTo 120
                For JIGP1I = JIGP1 To NXINYI
                    If (IZI = IGP(JIGP1I)) Then GoTo 130
                Next JIGP1I
120             NGP1 = NGP1 + 1
                JIGP1 = JIGP1 - 1
                IGP(JIGP1) = IZI
130         ' CONTINUE
            Next IXI
140     ' CONTINUE
        Next IYI
150     JNGP0 = JNGP0 + 1
        NGP(JNGP0) = NGP0
        JNGP1 = JNGP1 - 1
        NGP(JNGP1) = NGP1
    Next IT0
'
'
'
End Sub

Private Sub INTRP(ByVal SRELPR#, ByVal Tol#, X#(), Y#(), Z#(), IPT&(), _
    PD#(), ByVal ITI&, ByRef ITPV&, ByVal XII#, ByVal YII#, ByRef ZII#)
'
'   IT CARRIES OUT PUNCTUAL INTERPOLATION, I.E., IT DETERMINES
'   THE Z VALUE AT A GIVEN POINT IN A TRIANGLE BY MEANS OF THE
'   9-PARAMETER DISCRETIZED VERSION OF NIELSON'S SCHEME.
'
'   THE INPUT PARAMETERS ARE
'    SRELPR,TOL = SINGLE RELATIVE PRECISION AND TOLERANCE,
'    X,Y,Z=       ARRAYS OF DIMENSION N CONTAINING THE X,Y AND Z COORDINATES
'                 OF THE DATA POINTS, WHERE N IS THE NUMBER OF THE DATA POINTS,
'    IPT =        INTEGER ARRAY OF DIMENSION 3*NT, WHERE NT IS THE NUMBER OF
'                 TRIANGLES, CONTAINING THE INDICES OF THE VERTEXES OF THE
'                 TRIANGLES-THEMSELVES,
'    PD =         ARRAY OF DIMENSION 2*N CONTAINING THE PARTIAL DERIVATIVE
'                 VALUES AT THE DATA POINTS,
'    ITI =        INDEX OF THE TRIANGLE WHERE THE POINT FOR WHICH
'                 INTERPOLATION HAS TO BE PERFORMED, LIES,
'    ITPV =       INDEX OF THE TRIANGLE CONSIDERED IN THE PREVIOUS CALL,
'    XII,YII =    X AND Y COORDINATES OF THE POINT FOR WHICH INTER-
'                 POLATION HAS TO BE PERFORMED.
'
'   THE OUTPUT PARAMETER IS
'    ZII =        INTERPOLATED Z VALUE.
'
    Static ITO&, ITO3&, IND1&, IND2&, IND3&, INDEX&
    Static X12#, X13#, X23#, Y12#, Y13#, Y23#, D#, D3V1#, D2V1#, D1V2#, D3V2#
    Static D2V3#, D1V3#, E#, E1#, E2#, E3#, ALFA21#, ALFA31#, ALFA12#, ALFA32#, ALFA13#, ALFA23#
    Static B1#, B2#, B3#, w#, QB#, WB#
'
    ITO = ITI
    If (ITO = ITPV) Then GoTo 10
'
    ' SELECTS THE TRIANGLE CONTAINING THE POINT (XII,YII).
    ITPV = ITO
    ITO3 = 3 * ITO
    IND1 = IPT(ITO3 - 2)
    IND2 = IPT(ITO3 - 1)
    IND3 = IPT(ITO3)
'
    ' CALCULATES THE BASIC QUANTITIES RELATIVES TO THE SELECTED TRIANGLE.
    X12 = X(IND1) - X(IND2)
    X13 = X(IND1) - X(IND3)
    X23 = X(IND2) - X(IND3)
    Y12 = Y(IND1) - Y(IND2)
    Y13 = Y(IND1) - Y(IND3)
    Y23 = Y(IND2) - Y(IND3)
    D = X13 * Y23 - X23 * Y13
    INDEX = 2 * IND1
    D3V1 = -X13 * PD(INDEX - 1) - Y13 * PD(INDEX)
    D2V1 = -X12 * PD(INDEX - 1) - Y12 * PD(INDEX)
    INDEX = 2 * IND2
    D1V2 = X12 * PD(INDEX - 1) + Y12 * PD(INDEX)
    D3V2 = -X23 * PD(INDEX - 1) - Y23 * PD(INDEX)
    INDEX = 2 * IND3
    D2V3 = X23 * PD(INDEX - 1) + Y23 * PD(INDEX)
    D1V3 = X13 * PD(INDEX - 1) + Y13 * PD(INDEX)
    E1 = X23 * X23 + Y23 * Y23
    E2 = X13 * X13 + Y13 * Y13
    E3 = X12 * X12 + Y12 * Y12
    E = 2# * E1
    ALFA21 = (E2 + E1 - E3) / E
    ALFA31 = (E3 + E1 - E2) / E
    E = 2# * E2
    ALFA12 = (E1 + E2 - E3) / E
    ALFA32 = (E3 + E2 - E1) / E
    E = 2# * E3
    ALFA13 = (E1 + E3 - E2) / E
    ALFA23 = (E2 + E3 - E1) / E
'
    ' CALCULATES THE REMAINING QUANTITIES NECESSARY FOR THE ZI EVALUATION
    ' DEPENDING FROM THE INTERPOLATION POINT.
10  B1 = ((XII - X(IND3)) * Y23 - (YII - Y(IND3)) * X23) / D
    B2 = ((YII - Y(IND1)) * X13 - (XII - X(IND1)) * Y13) / D
    B3 = 1# - B1 - B2
    If (B1 >= Tol) Then GoTo 30
    If (B2 >= Tol) Then GoTo 40
    If (B1 + B2 <= SRELPR) Then GoTo 50
    w = B1 * B2 * B3 / (B1 * B2 + B1 * B3 + B2 * B3)
    QB = B1 * B1
    WB = w * B1
    ZII = Z(IND1) * (QB * (3# - 2# * B1) + 6# * WB * (B3 * ALFA12 + B2 * ALFA13)) _
        + D3V1 * (QB * B3 + WB * (3# * B3 * ALFA12 + B2 - B3)) _
        + D2V1 * (QB * B2 + WB * (3# * B2 * ALFA13 + B3 - B2))
    QB = B2 * B2
    WB = w * B2
    ZII = ZII + Z(IND2) * (QB * (3# - 2# * B2) + 6# * WB * (B1 * ALFA23 + B3 * ALFA21)) _
        + D1V2 * (QB * B1 + WB * (3# * B1 * ALFA23 + B3 - B1)) _
        + D3V2 * (QB * B3 + WB * (3# * B3 * ALFA21 + B1 - B3))
    QB = B3 * B3
    WB = w * B3
    ZII = ZII + Z(IND3) * (QB * (3# - 2# * B3) + 6# * WB * (B2 * ALFA31 + B1 * ALFA32)) _
        + D2V3 * (QB * B2 + WB * (3# * B2 * ALFA31 + B1 - B2)) _
        + D1V3 * (QB * B1 + WB * (3# * B1 * ALFA32 + B2 - B1))
    Exit Sub
'
30  ZII = Z(IND1)
    Exit Sub
'
40  ZII = Z(IND2)
    Exit Sub
'
50  ZII = Z(IND3)
'
'
'
End Sub
