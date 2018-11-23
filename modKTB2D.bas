Attribute VB_Name = "modKTB2D"
'==============================================================
' Description......: Routines for interpolating surfaces from
'                     scattered data points.
' Name of the Files: modKTB2D.bas
' Date.............: 12/10/2001
' Version..........: 1.0 at 32 bits.
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.it/flanguasco/
'                    http://www.flanguasco.org
'==============================================================
'
'   Routine di ingresso: KTB2D (vedere nella routine e nella
'                               definizione del tipo Par la
'                               descrizione dei parametri)
'
'   Note:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1.
'
'   Tradotto dal programma FORTRAN:
'
'                 Kriging of a 2-D Rectangular Grid
'                 *********************************
'------------------------------------------------------------------------
'                                                                        |
'   Copyright (C) 1996, The Board of Trustees of the Leland Stanford     |
'   Junior University.  All rights reserved.                             |
'                                                                        |
'   The programs in GSLIB are distributed in the hope that they will be  |
'   useful, but WITHOUT ANY WARRANTY.  No author or distributor accepts  |
'   responsibility to anyone for the consequences of using them or for   |
'   whether they serve any particular purpose or work at all, unless he  |
'   says so in writing.  Everyone is granted permission to copy, modify  |
'   and redistribute the programs in GSLIB, but only under the condition |
'   that this notice and the above copyright notice remain intact.       |
'                                                                        |
'------------------------------------------------------------------------
'
Option Explicit
'
Public Type ParType
    tmin As Double          ' Trimming limits.
    tmax As Double          '     "       "
    nxdis As Long           ' Number of discretization points/block in X.
    nydis As Long           ' Number of discretization points/block in Y.
    ndmin As Long           ' Minimum number of data required for kriging.
    ndmax As Long           ' Maximum number of samples to use in kriging.
    radius As Double        ' Maximum search radius closest ndmax samples will be retained.
    ktype As Long           ' Indicator for simple kriging (0=No, 1=Yes).
    skmean As Double        ' Mean for simple kriging (used if ktype=1).
    c0 As Double            ' Nugget constant (isotropic).
    Nst As Long             ' Number of nested structures (max. 4).
    It(1 To 4) As Long      ' Type of each nested structure:
                            '  it(nst) = 1. spherical model of range a;
                            '  it(nst) = 2. exponential model of parameter a;
                            '               i.e. practical range is 3a
                            '  it(nst) = 3. gaussian model of parameter a;
                            '               i.e. practical range is a*sqrt(3)
                            '  it(nst) = 4. power model of power a (a must be > 0  and
                            '               < 2).  if linear model, a=1,c=slope.
    cc(1 To 4) As Double    ' Multiplicative factor of each nested structure.
    ang(1 To 4) As Double   ' Azimuth angle for the principal direction of
                            ' continuity (measured clockwise in degrees from Y)
    AA(1 To 4) As Double    ' Parameter "a" of each nested structure.
    a2(1 To 4) As Double    ' a2(i)/aa(i) = Anisotropy (radius in minor direction
                            '               at 90 degrees from "ang" divided by the
                            '               principal radius in direction "ang")
End Type
Private Sub KSOL(ByVal Nright&, ByVal NEQ&, ByVal Nsb&, A#(), R#(), _
    ByRef S#(), ByRef Ising&)
'-----------------------------------------------------------------------
'
'                Solution of a System of Linear Equations
'                ****************************************
'
'
'
'   INPUT VARIABLES:
'
'    nright,nsb     number of columns in right hand side matrix.
'                   for KB2D: nright=1, nsb=1
'    neq            number of equations
'    a()            upper triangular left hand side matrix (stored
'                   columnwise)
'    r()            right hand side matrix (stored columnwise)
'                   for kb2d, one column per variable
'
'
'
'   OUTPUT VARIABLES:
'
'    s()            solution array, same dimension as r above.
'    ising          singularity indicator
'                     0,  no singularity problem
'                    -1,  neq .le. 1
'                     k,  a null pivot appeared at the kth iteration
'
'
'
'   PROGRAM NOTES:
'
'    1. Requires the upper triangular left hand side matrix.
'    2. Pivots are on the diagonal.
'    3. Does not search for max. element for pivot.
'    4. Several right hand side matrices possible.
'    5. USE for ok and sk only, NOT for UK.
'
'
'-----------------------------------------------------------------------
'
    Dim I&, II&, IJ&, IJm&, IN1&, IV&, J&, K&, KK&, LL&, LLb&, LL1&, Lp&
    Dim KM1&, M1&, Nm&, NM1&, Nn&
    Dim Ak#, Ap#, Piv#, Tol#
'
' If there is only one equation then set ising and return:
'
    If (NEQ <= 1) Then
        Ising = -1
        Exit Sub
    End If
'
' Initialize:
'
    Tol = 0.0000001
    Ising = 0
    Nn = NEQ * (NEQ + 1) / 2
    Nm = Nsb * NEQ
    M1 = NEQ - 1
    KK = 0
'
' Start triangulation:
'
    For K = 1 To M1
        KK = KK + K
        Ak = A(KK)
        If (Abs(Ak) < Tol) Then
            Ising = K
            Exit Sub
        End If
        KM1 = K - 1
        For IV = 1 To Nright
            NM1 = Nm * (IV - 1)
            II = KK + Nn * (IV - 1)
            Piv = 1# / A(II)
            Lp = 0
            For I = K To M1
                LL = II
                II = II + I
                Ap = A(II) * Piv
                Lp = Lp + 1
                IJ = II - KM1
                For J = I To M1
                    IJ = IJ + J
                    LL = LL + J
                    A(IJ) = A(IJ) - Ap * A(LL)
                Next J
                For LLb = K To Nm Step NEQ
                    IN1 = LLb + Lp + NM1
                    LL1 = LLb + NM1
                    R(IN1) = R(IN1) - Ap * R(LL1)
                Next LLb
            Next I
        Next IV
    Next K
'
' Error checking - singular matrix:
'
    IJm = IJ - Nn * (Nright - 1)
    If (Abs(A(IJm)) < Tol) Then
        Ising = NEQ
        Exit Sub
    End If
'
' Finished triangulation, start solving back:
'
    For IV = 1 To Nright
        NM1 = Nm * (IV - 1)
        IJ = IJm + Nn * (IV - 1)
        Piv = 1# / A(IJ)
        For LLb = NEQ To Nm Step NEQ
            LL1 = LLb + NM1
            S(LL1) = R(LL1) * Piv
        Next LLb
        I = NEQ
        KK = IJ
        For II = 1 To M1
            KK = KK - I
            Piv = 1# / A(KK)
            I = I - 1
            For LLb = I To Nm Step NEQ
                LL1 = LLb + NM1
                IN1 = LL1
                Ap = R(IN1)
                IJ = KK
                For J = I To M1
                    IJ = IJ + J
                    IN1 = IN1 + 1
                    Ap = Ap - A(IJ) * S(IN1)
                Next J
                S(LL1) = Ap * Piv
            Next LLb
        Next II
    Next IV
'
' Finished solving back, return:
'
'
'
End Sub
Public Sub KTB2D(ByVal ND&, XD#(), YD#(), ZD#(), Par As ParType, _
    ByVal NX&, ByVal XMN#, ByVal xsiz#, ByVal NY&, ByVal YMN#, ByVal ysiz#, _
    ByRef Z#(), ByRef IER&)
'
'   Parametri in ingresso:
'    ND:            N° di punti dati.
'    XD(1 To ND):   vettore delle ascisse dei punti dati.
'    YD(1 To ND):   vettore delle ordinate dei punti dati.
'    ZD(1 To ND):   vettore dei valori della superficie ai punti dati.
'    Par:           struttura dei parametri di controllo.
'    NX:            N° di colonne della griglia dei punti interpolati.
'    XMN:           ascissa minima della griglia dei punti interpolati.
'    xsiz:          distanza fra le ascisse della griglia dei punti interpolati.
'    NY:            N° di righe della griglia dei punti interpolati.
'    YMN:           ordinata minima della griglia dei punti interpolati.
'    ysiz:          distanza fra le ordinate della griglia dei punti interpolati.
'    IER:           se il parametro IER viene passato con valore > 0 viene
'                   generato un file "Debug.txt" con informazioni sul
'                   funzionamento della routine.
'
'   Parametri in uscita:
'    Z(1 To NX, 1 To NY):   matrice dei valori interpolati.
'    IER:           codice di errore ritornato dalla routine
'                    0 = No errors.
'                    1 = Exceeded available memory for data.
'                    2 = ndmax is too big - modify PARAMETERS.
'                    3 = nst is too big - modify PARAMETERS.
'                    4 = INVALID power variogram.
'                    5 = Too many discretization points:
'                        Increase MAXDIS or lower n[xy]dis.
'
'-----------------------------------------------------------------------
'
'           Ordinary/Simple Kriging of a 2-D Rectangular Grid
'           *************************************************
'
'   This subroutine estimates point or block values of one variable by
'   ordinary kriging.  All of the samples are rescanned for each block
'   estimate; this makes the program simple but inefficient.  The data
'   should NOT contain any missing values.  Unestimated points are
'   returned as -1.0e21
'
'
'
'   Original:  A.G.Journel 1978
'   Revisions: B.E. Buxton                                     Apr. 1983
'-----------------------------------------------------------------------
'
'                       *******************
'
'   The following Parameters control static dimensioning within KTB2D:
'
'    MAXX      maximum nodes in X
'    MAXY      maximum nodes in Y
'    MAXDAT    maximum number of data points
'    MAXSAM    maximum number of data points to use in one kriging system
'    MAXDIS    maximum number of discretization points per block
'    MAXNST    maximum number of nested structures
'
'   User Adjustable:
'
    Const MAXDAT& = 10000
    Const MAXSAM& = 120
    Const MAXDIS& = 64
    Const MAXNST& = 4
'
'   Fixed:
'
    Const MAXKD& = MAXSAM + 1, MAXKRG& = MAXKD * MAXKD
    Const UNEST# = -999, EPSLON# = 0.0000001
'
'-----------------------------------------------------------------------
'
'   Variable Declaration:
'
    Dim NUMS&(1 To MAXSAM)
    Dim NDr&, N1&, Na&, Ndb&, NEQ&, Nn&, Nk&, I&, IA&, Isam&, II&, ID&, IN1&, Ix&, Iy&
    Dim Ldbg%, K&, JK&, J&, J1&, JJ&, Idbg&, Ising&
'
    Dim X#(1 To MAXDAT), Y#(1 To MAXDAT), VR#(1 To MAXDAT)
    Dim xdb#(1 To MAXDIS), ydb#(1 To MAXDIS), xa#(1 To MAXSAM), ya#(1 To MAXSAM)
    Dim vra#(1 To MAXSAM), DIST#(1 To MAXSAM), PMX#, anis#(1 To MAXNST)
    Dim R#(1 To MAXSAM + 1), rr#(1 To MAXSAM + 1), S#(1 To MAXSAM + 1), A#(1 To MAXKRG)
    Dim v#, ss#, xloc#, yloc#, xdis#, ydis#, cbb#, cb1#, rad2#, cov#, unbias#
    Dim Ak#, AV#, vk#, DX#, DY#, h2#, est#, estv#, XX#, YY#, cb#, sumw#, vrt#
    Dim First As Boolean
'
    First = True
    PMX = 9999#
'
'   Read Input Parameters:
'
    Idbg = IER  ' Debug flag
    If (Idbg > 0) Then
'       Unit numbers:
        Ldbg = FreeFile
        Open App.Path & "\Debug.txt" For Output As Ldbg
    End If
'
    If (ND > MAXDAT) Then
        If Idbg > 0 Then
            Print #Ldbg, " ERROR: Exceeded available memory for data"
            IER = 1
            Exit Sub
        End If
    End If
'
    IER = 0
'
    If (Par.ndmin < 0) Then Par.ndmin = 0
    If (Par.ndmax > MAXSAM) Then
        If Idbg > 0 Then
            Print #Ldbg, "ndmax is too big - modify PARAMETERS"
            IER = 2
            Exit Sub
        End If
    End If
'
    If (Par.Nst > MAXNST) Then
        If Idbg > 0 Then
            Print #Ldbg, "nst is too big - modify PARAMETERS"
            IER = 3
            Exit Sub
        End If
    End If
'
    If (Par.Nst < 0) Then
        Par.Nst = 1
        Par.It(1) = 1
        Par.cc(1) = 0#
        Par.ang(1) = 0#
        Par.AA(1) = 0#
        anis(1) = 0#
    Else
        For I = 1 To Par.Nst
            anis(I) = Par.a2(I) / Par.AA(I)
'
            If (Par.It(I) = 4) Then
                If ((Par.AA(I) < 0#) Or (Par.AA(I) > 2#)) Then
                    If Idbg > 0 Then
                        Print #Ldbg, "INVALID power variogram"
                        IER = 4
                        Exit Sub
                    End If
                End If
            End If
        Next I
    End If
'
'   Read the data:
'
    AV = 0#
    ss = 0#
'
    NDr = 0
    For I = 1 To ND
        vrt = ZD(I)
        If (vrt < Par.tmin Or vrt > Par.tmax) Then GoTo 7
'
        NDr = NDr + 1
        X(NDr) = XD(I)
        Y(NDr) = YD(I)
        VR(NDr) = vrt
        AV = AV + vrt
        ss = ss + vrt * vrt
7   ' CONTINUE
    Next I
'
'   Echo the input data if debugging flag >0:
'
    If Idbg > 0 Then
        Print #Ldbg, "tmin,tmax", Par.tmin, Par.tmax
        Print #Ldbg, "xmn,ymn", XMN, YMN
        Print #Ldbg, "xsiz,ysiz", xsiz, ysiz
        Print #Ldbg, "nxdis,nydis", Par.nxdis, Par.nydis
        Print #Ldbg, "ndmin", Par.ndmin
        Print #Ldbg, "ndmax", Par.ndmax
        Print #Ldbg, "radius", Par.radius
        Print #Ldbg, "ktype", Par.ktype
        Print #Ldbg, "skmean", Par.skmean
        Print #Ldbg, "nst", Par.Nst
        Print #Ldbg, "c0", Par.c0
        For I = 1 To Par.Nst
            Print #Ldbg, "it ", I, Par.It(I)
            Print #Ldbg, "cc ", I, Par.cc(I)
            Print #Ldbg, "ang", I, Par.ang(I)
            Print #Ldbg, "aa ", I, Par.AA(I)
            Print #Ldbg, "a2 ", I, Par.a2(I)
        Next I
        Print #Ldbg, "NDr,nx,ny", NDr, NX, NY
'
        For ID = 1 To NDr
            Print #Ldbg, ID, X(ID), Y(ID), VR(ID)
        Next ID
'
'   Compute the averages and variances as an error check for the user:
'
        AV = AV / DMAX1(CDbl(NDr), 1#)
        ss = (ss / DMAX1(CDbl(NDr), 1#)) - AV * AV
        Print #Ldbg, "av,ss", AV, ss
    End If
'
' Set up the discretization points per block.  Figure out how many
' are needed, the spacing, and fill the xdb and ydb arrays with the
' offsets relative to the block center (this only gets done once):
'
    Ndb = Par.nxdis * Par.nydis
    If (Ndb > MAXDIS) Then
        If Idbg > 0 Then
                Print #Ldbg, "ERROR KB2D: Too many discretization points."
                Print #Ldbg, "            Increase MAXDIS"
                Print #Ldbg, "            or lower n[xy]dis."
        End If
        IER = 5
        Exit Sub
    End If
    xdis = xsiz / DMAX1(CDbl(Par.nxdis), 1#)
    ydis = ysiz / DMAX1(CDbl(Par.nydis), 1#)
    xloc = -0.5 * (xsiz + xdis)
    I = 0
    For Ix = 1 To Par.nxdis
        xloc = xloc + xdis
        yloc = -0.5 * (ysiz + ydis)
        For Iy = 1 To Par.nydis
            yloc = yloc + ydis
            I = I + 1
            xdb(I) = xloc
            ydb(I) = yloc
        Next Iy
    Next Ix
'
'   Initialize accumulators:
'
    cbb = 0#
    rad2 = Par.radius * Par.radius
'
'   Calculate Block Covariance. Check for point kriging.
'
    cov = COVA2(xdb(1), ydb(1), xdb(1), ydb(1), Par.Nst, Par.c0 _
              , PMX, Par.cc(), Par.AA(), Par.It(), Par.ang(), anis(), First)
'
'   Keep this value to use for the unbiasedness constraint:
'
    unbias = cov
    First = False
    If (Ndb <= 1) Then
        cbb = cov
    Else
        For I = 1 To Ndb
            For J = 1 To Ndb
                cov = COVA2(xdb(I), ydb(I), xdb(J), ydb(J), _
                            Par.Nst, Par.c0, PMX, _
                            Par.cc(), Par.AA(), Par.It(), _
                            Par.ang(), anis(), First)
                If (I = J) Then cov = cov - Par.c0
                cbb = cbb + cov
            Next J
        Next I
        cbb = cbb / CDbl(Ndb * Ndb)
    End If
'
'   MAIN LOOP OVER ALL THE BLOCKS IN THE GRID:
'
    Nk = 0
    Ak = 0#
    vk = 0#
    For Iy = 1 To NY
        yloc = YMN + (Iy - 1) * ysiz
        For Ix = 1 To NX
            xloc = XMN + (Ix - 1) * xsiz
'
'   Find the nearest samples within each octant: First initialize
'   the counter arrays:
'
            Na = 0
            For Isam = 1 To Par.ndmax
                DIST(Isam) = 1E+20
                NUMS(Isam) = 0
            Next Isam
'
'   Scan all the samples (this is inefficient and the user with lots of
'   data should move to ktb3d):
'
            For ID = 1 To NDr
                DX = X(ID) - xloc
                DY = Y(ID) - yloc
                h2 = DX * DX + DY * DY
                If (h2 > rad2) Then GoTo 6
'
'   Do not consider this sample if there are enough close ones:
'
                If (Na = Par.ndmax) Then
                    If (h2 > DIST(Na)) Then GoTo 6
                End If
'
'   Consider this sample (it will be added in the correct location):
'
                If (Na < Par.ndmax) Then Na = Na + 1
                NUMS(Na) = ID
                DIST(Na) = h2
                If (Na = 1) Then GoTo 6
'
'   Sort samples found thus far in increasing order of distance:
'
                N1 = Na - 1
                For II = 1 To N1
                    K = II
                    If (h2 < DIST(II)) Then
                        JK = 0
                        For JJ = K To N1
                            J = N1 - JK
                            JK = JK + 1
                            J1 = J + 1
                            DIST(J1) = DIST(J)
                            NUMS(J1) = NUMS(J)
                        Next JJ
                        DIST(K) = h2
                        NUMS(K) = ID
                        GoTo 6
                    End If
                Next II
6           ' CONTINUE
            Next ID
'
'   Is there enough samples?
'
            If (Na < Par.ndmin) Then
                If Idbg > 0 Then Print #Ldbg, "Block ", Ix, Iy, " not estimated"
                est = UNEST
                estv = UNEST
                GoTo 1
            End If
'
'   Put coordinates and values of neighborhood samples into xa,ya,vra:
'
            For IA = 1 To Na
                JJ = NUMS(IA)
                xa(IA) = X(JJ)
                ya(IA) = Y(JJ)
                vra(IA) = VR(JJ)
            Next IA
'
'   Handle the situation of only one sample:
'
            If (Na = 1) Then
                cb1 = COVA2(xa(1), ya(1), xa(1), ya(1), Par.Nst, Par.c0, _
                            PMX, Par.cc(), Par.AA(), Par.It(), Par.ang(), _
                            anis(), First)
                XX = xa(1) - xloc
                YY = ya(1) - yloc
'
'   Establish Right Hand Side Covariance:
'
                If (Ndb <= 1) Then
                    cb = COVA2(XX, YY, xdb(1), ydb(1), _
                               Par.Nst, Par.c0, PMX, _
                               Par.cc(), Par.AA(), Par.It(), _
                               Par.ang(), anis(), First)
                Else
                    cb = 0#
                    For I = 1 To Ndb
                        cb = cb + COVA2(XX, YY, xdb(I), ydb(I), _
                                        Par.Nst, Par.c0, PMX, _
                                        Par.cc(), Par.AA(), Par.It(), _
                                        Par.ang(), anis(), First)
                        DX = XX - xdb(I)
                        DY = YY - ydb(I)
                        If ((DX * DX + DY * DY) < EPSLON) Then cb = cb - Par.c0
                    Next I
                    cb = cb / CDbl(Ndb)
                End If
                If (Par.ktype = 0) Then
                    S(1) = cb / cbb
                    est = S(1) * vra(1) + (1# - S(1)) * Par.skmean
                    estv = cbb - S(1) * cb
                Else
                    est = vra(1)
                    estv = cbb - 2# * cb + cb1
                End If
            Else
'
'   Solve the Kriging System with more than one sample:
'
                NEQ = Na + Par.ktype
                Nn = (NEQ + 1) * NEQ / 2
'
'   Set up kriging matrices:
'
                IN1 = 0
                For J = 1 To Na
'
'   Establish Left Hand Side Covariance Matrix:
'
                    For I = 1 To J
                        IN1 = IN1 + 1
                        A(IN1) = COVA2(xa(I), ya(I), xa(J), ya(J), _
                                       Par.Nst, Par.c0, PMX, _
                                       Par.cc(), Par.AA(), Par.It(), _
                                       Par.ang(), anis(), First)
                    Next I
                    XX = xa(J) - xloc
                    YY = ya(J) - yloc
'
'   Establish Right Hand Side Covariance:
'
                    If (Ndb <= 1) Then
                        cb = COVA2(XX, YY, xdb(1), ydb(1), _
                                   Par.Nst, Par.c0, PMX, _
                                   Par.cc(), Par.AA(), Par.It(), _
                                   Par.ang(), anis(), First)
                    Else
                        cb = 0#
                        For J1 = 1 To Ndb
                            cb = cb + COVA2(XX, YY, _
                                            xdb(J1), ydb(J1), _
                                            Par.Nst, Par.c0, PMX, _
                                            Par.cc(), Par.AA(), _
                                            Par.It(), Par.ang(), _
                                            anis(), First)
                            DX = XX - xdb(J1)
                            DY = YY - ydb(J1)
                            If ((DX * DX + DY * DY) < EPSLON) Then cb = cb - Par.c0
                        Next J1
                        cb = cb / CDbl(Ndb)
                    End If
                    R(J) = cb
                    rr(J) = R(J)
                Next J
'
'   Set the unbiasedness constraint:
'
                If (Par.ktype = 1) Then
                    For I = 1 To Na
                        IN1 = IN1 + 1
                        A(IN1) = unbias
                    Next I
                    IN1 = IN1 + 1
                    A(IN1) = 0#
                    R(NEQ) = unbias
                    rr(NEQ) = R(NEQ)
                End If
'
'   Solve the Kriging System:
'
                Call KSOL(1, NEQ, 1, A, R, S, Ising)
'
'   Write a warning if the matrix is singular:
'
                If (Ising <> 0) Then
                    If Idbg > 0 Then
                        Print #Ldbg, "WARNING KB2D: singular matrix"
                        Print #Ldbg, "              for block", Ix, Iy
                    End If
                    est = UNEST
                    estv = UNEST
                    GoTo 1
                End If
'
'   Compute the estimate and the kriging variance:
'
                est = 0#
                estv = cbb
                sumw = 0#
                If (Par.ktype = 1) Then estv = estv - (S(Na + 1))
                For I = 1 To Na
                    sumw = sumw + (S(I))
                    est = est + (S(I)) * vra(I)
                    estv = estv - (S(I) * rr(I))
                Next I
                If (Par.ktype = 0) Then est = est + (1# - sumw) * Par.skmean
            End If
'
'   Write the result to the output matrix:

1           Z(Ix, Iy) = est

            If (est > UNEST) Then
                  Nk = Nk + 1
                  Ak = Ak + est
                  vk = vk + est * est
            End If
'
'   END OF MAIN LOOP OVER ALL THE BLOCKS:
'
4           ' CONTINUE
        Next Ix
    Next Iy
'
' Finished:
'
100 ' CONTINUE
    If Idbg > 0 Then Close (Ldbg)
'
'
'
End Sub
Private Function COVA2(ByVal x1#, ByVal y1#, ByVal x2#, ByVal y2#, _
    ByVal Nst&, ByVal c0#, ByVal PMX#, cc#(), AA#(), It&(), ang#(), _
    anis#(), ByVal First As Boolean) As Double
'-----------------------------------------------------------------------
'
'              Covariance Between Two Points (2-D Version)
'              *******************************************
'
'   This function returns the covariance associated with a variogram model
'   that is specified by a nugget effect and possibly four different
'   nested varigoram structures.  The anisotropy definition can be
'   different for each of the nested structures (spherical, exponential,
'   gaussian, or power).
'
'   INPUT VARIABLES:
'
'    x1,y1              Coordinates of first point
'    x2,y2              Coordinates of second point
'    nst                Number of nested structures (max. 4).
'    c0                 Nugget constant (isotropic).
'    PMX                Maximum variogram value needed for kriging when
'                       using power model.  A unique value of PMX is
'                       used for all nested structures which use the
'                       power model.  therefore, PMX should be chosen
'                       large enough to account for the largest single
'                       structure which uses the power model.
'    cc(nst)            Multiplicative factor of each nested structure.
'    aa(nst)            Parameter "a" of each nested structure.
'    it(nst)            Type of each nested structure:
'                        1. spherical model of range a;
'                        2. exponential model of parameter a;
'                           i.e. practical range is 3a
'                        3. gaussian model of parameter a;
'                           i.e. practical range is a*sqrt(3)
'                        4. power model of power a (a must be gt. 0  and
'                           lt. 2).  if linear model, a=1,c=slope.
'    ang(nst)           Azimuth angle for the principal direction of
'                       continuity (measured clockwise in degrees from Y)
'    anis(nst)          Anisotropy (radius in minor direction at 90 degrees
'                       from "ang" divided by the principal radius in
'                       direction "ang")
'    first              A logical variable which is set to true if the
'                       direction specifications have changed - causes
'                       the rotation matrices to be recomputed.
'
'
'
'   OUTPUT VARIABLES:   returns "COVA2" the covariance obtained from the
'                       variogram model.
'
'
'
'-----------------------------------------------------------------------
'
    Dim azmuth#, DX#, DY#, Dx1#, Dy1#, H#, hh#, hr#, cov1#, COVA2T#
    Static rotmat#(1 To 4, 1 To 4), maxcov#
    Dim IS1&
    Const DTOR# = 3.14159265358979 / 180#, EPSLON# = 0.0000001
'
'   The first time around, re-initialize the cosine matrix for the
'   variogram structures:
'
    If (First) Then
        maxcov = c0
        For IS1 = 1 To Nst
            azmuth = (90# - ang(IS1)) * DTOR
            rotmat(1, IS1) = Cos(azmuth)
            rotmat(2, IS1) = Sin(azmuth)
            rotmat(3, IS1) = -Sin(azmuth)
            rotmat(4, IS1) = Cos(azmuth)
            If (It(IS1) = 4) Then
                maxcov = maxcov + PMX
            Else
                maxcov = maxcov + cc(IS1)
            End If
        Next IS1
    End If
'
'   Check for very small distance:
'
    DX = x2 - x1
    DY = y2 - y1
    If ((DX * DX + DY * DY) < EPSLON) Then
        COVA2 = maxcov
        Exit Function
    End If
'
'   Non-zero distance, loop over all the structures:
'
    COVA2T = 0#
    For IS1 = 1 To Nst
'
'       Compute the appropriate structural distance:
'
        Dx1 = (DX * rotmat(1, IS1) + DY * rotmat(2, IS1))
        Dy1 = (DX * rotmat(3, IS1) + DY * rotmat(4, IS1)) / anis(IS1)
        H = Sqr(DMAX1((Dx1 * Dx1 + Dy1 * Dy1), 0#))
        If (It(IS1) = 1) Then
'
'           Spherical model:
'
            hr = H / AA(IS1)
            If (hr < 1#) Then
                COVA2T = COVA2T + cc(IS1) * (1# - hr * (1.5 - 0.5 * hr * hr))
            End If
        ElseIf (It(IS1) = 2) Then
'
'           Exponential model:
'
            COVA2T = COVA2T + cc(IS1) * Exp(-H / AA(IS1))
        ElseIf (It(IS1) = 3) Then
'
'           Gaussian model:
'
            hh = -(H * H) / (AA(IS1) * AA(IS1))
            COVA2T = COVA2T + cc(IS1) * Exp(hh)
        Else
'
'           Power model:
'
            cov1 = PMX - cc(IS1) * (H ^ AA(IS1))
            COVA2T = COVA2T + cov1
        End If
    Next IS1
'
    COVA2 = COVA2T
'
'
'
End Function
