Attribute VB_Name = "modQSHEP2D"
'==============================================================
' Description......: Routines for interpolating surfaces from
'                     scattered data points.
' Name of the Files: modQSHEP2D.bas
' Date.............: 12/10/2001
' Version..........: 1.0 at 32 bits.
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.it/flanguasco/
'                    http://www.flanguasco.org
'==============================================================
'
'   Routine di impostazione: QSHEP2
'   Routines di interpolazione: QS2GRD e QS2VAL
'   (vedere nelle routines la descrizione dei parametri ed
'   in Main un esempio d' uso)
'
'   Note:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1.
'
'   Tradotto dal programma FORTRAN:
'    ALGORITHM 660, COLLECTED ALGORITHMS FROM ACM.
'    THIS WORK PUBLISHED IN TRANSACTIONS ON MATHEMATICAL SOFTWARE,
'    VOL. 14, NO. 2, P.149.
'
Option Explicit

Public Sub QSHEP2(ByVal N&, X#(), Y#(), F#(), ByVal NQ&, ByVal NW&, ByVal NR&, _
    ByRef LCELL&(), ByRef LNEXT&(), ByRef XMin#, ByRef YMin#, ByRef DX#, ByRef DY#, _
    ByRef RMAX#, ByRef RSQ#(), ByRef A#(), ByRef IER&)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'   1 / 8 / 90
'
'   THIS SUBROUTINE COMPUTES A SET OF PARAMETERS A AND RSQ
'   DEFINING A SMOOTH (ONCE CONTINUOUSLY DIFFERENTIABLE) BI-
'   VARIATE FUNCTION Q(X,Y) WHICH INTERPOLATES DATA VALUES F
'   AT SCATTERED NODES (X,Y).  THE INTERPOLANT Q MAY BE EVAL-
'   UATED AT AN ARBITRARY POINT BY FUNCTION QS2VAL, AND ITS
'   FIRST DERIVATIVES ARE COMPUTED BY SUBROUTINE QS2GRD.
'   THE INTERPOLATION SCHEME IS A MODIFIED QUADRATIC SHEPARD
'   METHOD --
'
'    Q = (W(1)*Q(1)+W(2)*Q(2)+..+W(N)*Q(N))/(W(1)+W(2)+..+W(N))
'
'   FOR BIVARIATE FUNCTIONS W(K) AND Q(K).  THE NODAL FUNC-
'   TIONS ARE GIVEN BY
'
'    Q(K)(X,Y) = A(1,K)*(X-X(K))**2 + A(2,K)*(X-X(K))*(Y-Y(K))
'              + A(3,K)*(Y-Y(K))**2 + A(4,K)*(X-X(K))
'              + A(5,K)*(Y-Y(K))    + F(K) .
'
'   THUS, Q(K) IS A QUADRATIC FUNCTION WHICH INTERPOLATES THE
'   DATA VALUE AT NODE K.  ITS COEFFICIENTS A(,K) ARE OBTAINED
'   BY A WEIGHTED LEAST SQUARES FIT TO THE CLOSEST NQ DATA
'   POINTS WITH WEIGHTS SIMILAR TO W(K).  NOTE THAT THE RADIUS
'   OF INFLUENCE FOR THE LEAST SQUARES FIT IS FIXED FOR EACH
'   K, BUT VARIES WITH K.
'   THE WEIGHTS ARE TAKEN TO BE
'
'    W(K)(X,Y) = ( (R(K)-D(K))+ / R(K)*D(K) )**2
'
'   WHERE (R(K)-D(K))+ = 0 IF R(K) .LE. D(K) AND D(K)(X,Y) IS
'   THE EUCLIDEAN DISTANCE BETWEEN (X,Y) AND (X(K),Y(K)).  THE
'   RADIUS OF INFLUENCE R(K) VARIES WITH K AND IS CHOSEN SO
'   THAT NW NODES ARE WITHIN THE RADIUS.  NOTE THAT W(K) IS
'   NOT DEFINED AT NODE (X(K),Y(K)), BUT Q(X,Y) HAS LIMIT F(K)
'   AS (X,Y) APPROACHES (X(K),Y(K)).
'
'   ON INPUT --
'
'    N =     NUMBER OF NODES AND ASSOCIATED DATA VALUES.
'            N .GE. 6.
'
'    X,Y =   ARRAYS OF LENGTH N CONTAINING THE CARTESIAN
'            COORDINATES OF THE NODES.
'
'    F =     ARRAY OF LENGTH N CONTAINING THE DATA VALUES
'            IN ONE-TO-ONE CORRESPONDENCE WITH THE NODES.
'
'    NQ =    NUMBER OF DATA POINTS TO BE USED IN THE LEAST
'            SQUARES FIT FOR COEFFICIENTS DEFINING THE NODAL
'            FUNCTIONS Q(K).  A HIGHLY RECOMMENDED VALUE IS
'            NQ = 13.  5 .LE. NQ .LE. MIN(40,N-1).
'
'    NW =    NUMBER OF NODES WITHIN (AND DEFINING) THE RADII
'            OF INFLUENCE R(K) WHICH ENTER INTO THE WEIGHTS
'            W(K).  FOR N SUFFICIENTLY LARGE, A RECOMMENDED
'            VALUE IS NW = 19.  1 .LE. NW .LE. MIN(40,N-1).
'
'    NR =    NUMBER OF ROWS AND COLUMNS IN THE CELL GRID DE-
'            FINED IN SUBROUTINE STORE2.  A RECTANGLE CON-
'            TAINING THE NODES IS PARTITIONED INTO CELLS IN
'            ORDER TO INCREASE SEARCH EFFICIENCY.  NR =
'            SQRT(N/3) IS RECOMMENDED.  NR .GE. 1.
'
'   THE ABOVE PARAMETERS ARE NOT ALTERED BY THIS ROUTINE.
'
'    LCELL = ARRAY OF LENGTH .GE. NR**2.
'
'    LNEXT = ARRAY OF LENGTH .GE. N.
'
'    RSQ =   ARRAY OF LENGTH .GE. N.
'
'    A =     ARRAY OF LENGTH .GE. 5N.
'
'   ON OUTPUT --
'
'    LCELL = NR BY NR ARRAY OF NODAL INDICES ASSOCIATED
'            WITH CELLS.  REFER TO STORE2.
'
'    LNEXT = ARRAY OF LENGTH N CONTAINING NEXT-NODE INDI-
'            CES.  REFER TO STORE2.
'
'    XMIN,YMIN,DX,DY = MINIMUM NODAL COORDINATES AND CELL
'            DIMENSIONS.  REFER TO STORE2.
'
'    RMAX =  SQUARE ROOT OF THE LARGEST ELEMENT IN RSQ --
'            MAXIMUM RADIUS R(K).
'
'    RSQ =   ARRAY CONTAINING THE SQUARES OF THE RADII R(K)
'            WHICH ENTER INTO THE WEIGHTS W(K).
'
'    A =     5 BY N ARRAY CONTAINING THE COEFFICIENTS FOR
'            QUADRATIC NODAL FUNCTION Q(K) IN COLUMN K.
'
'   NOTE THAT THE ABOVE OUTPUT PARAMETERS ARE NOT DEFINED
'   UNLESS IER = 0.
'
'    IER =   ERROR INDICATOR --
'             IER = 0 IF NO ERRORS WERE ENCOUNTERED.
'             IER = 1 IF N, NQ, NW, OR NR IS OUT OF RANGE.
'             IER = 2 IF DUPLICATE NODES WERE ENCOUNTERED.
'             IER = 3 IF ALL NODES ARE COLLINEAR.
'
'   MODULES REQUIRED BY QSHEP2 -- GETNP2, GIVENS, ROTATE,
'   SETUP2, STORE2
'
'   INTRINSIC FUNCTIONS CALLED BY QSHEP2 -- ABS, AMIN1, FLOAT,
'   MAX0, MIN0, SQRT
'
'***********************************************************
'
    Dim I&, IB&, IERR&, IP1&, IRM1&, IRow&, J&, JP1&, K&, LMAX&
    Dim LNP&, NEQ&, Nn&, NNQ&, NNR&, NNW&, NP&, NPTS&(1 To 40), NQWMAX&
    Dim AV#, AVSQ#, B#(1 To 6, 1 To 6), C#, DDX#, DDY#, DMin#, FK#
    Dim RQ#, RS#, RSMX#, RSOLD#, RWS#, S#, sum#, t#
    Dim XK#, XMN#, YK#, YMN#
'
    Const RTOL# = 0.00001, DTOL# = 0.01, SF# = 1#
'
    ' LOCAL PARAMETERS --
'
'    AV =         ROOT-MEAN-SQUARE DISTANCE BETWEEN K AND THE
'                 NODES IN THE LEAST SQUARES SYSTEM (UNLESS
'                 ADDITIONAL NODES ARE INTRODUCED FOR STABIL-
'                 ITY).  THE FIRST 3 COLUMNS OF THE MATRIX
'                 ARE SCALED BY 1/AVSQ, THE LAST 2 BY 1/AV
'    AVSQ =       AV * AV
'    B =          TRANSPOSE OF THE AUGMENTED REGRESSION MATRIX
'    C =          FIRST COMPONENT OF THE PLANE ROTATION USED TO
'                 ZERO THE LOWER TRIANGLE OF B**T -- COMPUTED
'                 BY SUBROUTINE GIVENS
'    DDX,DDY =    LOCAL VARIABLES FOR DX AND DY
'    DMIN =       MINIMUM OF THE MAGNITUDES OF THE DIAGONAL
'                 ELEMENTS OF THE REGRESSION MATRIX AFTER
'                 ZEROS ARE INTRODUCED BELOW THE DIAGONAL
'    DTOL =       TOLERANCE FOR DETECTING AN ILL-CONDITIONED
'                 SYSTEM.  THE SYSTEM IS ACCEPTED WHEN DMIN
'                 .GE.DTOL
'    FK =         DATA VALUE AT NODE K -- F(K)
'    I =          INDEX FOR A, B, AND NPTS
'    IB =         DO-LOOP INDEX FOR BACK SOLVE
'    IERR =       ERROR FLAG FOR THE CALL TO STORE2
'    IP1 =        I + 1
'    IRM1 =       IROW - 1
'    IROW =       ROW INDEX FOR B
'    J =          INDEX FOR A AND B
'    JP1 =        J + 1
'    K =          NODAL FUNCTION INDEX AND COLUMN INDEX FOR A
'    LMAX =       MAXIMUM NUMBER OF NPTS ELEMENTS (MUST BE CON-
'                 SISTENT WITH THE DIMENSION STATEMENT ABOVE)
'    LNP =        CURRENT LENGTH OF NPTS
'    NEQ =        NUMBER OF EQUATIONS IN THE LEAST SQUARES FIT
'    NN,NNQ,NNR = LOCAL COPIES OF N, NQ, AND NR
'    NNW =        LOCAL COPY OF NW
'    NP =         NPTS ELEMENT
'    NPTS =       ARRAY CONTAINING THE INDICES OF A SEQUENCE OF
'                 NODES TO BE USED IN THE LEAST SQUARES FIT
'                 OR TO COMPUTE RSQ.  THE NODES ARE ORDERED
'                 BY DISTANCE FROM K AND THE LAST ELEMENT
'                 (USUALLY INDEXED BY LNP) IS USED ONLY TO
'                 DETERMINE RQ, OR RSQ(K) IF NW .GT. NQ
'    NQWMAX =     Max(NQ, NW)
'    RQ =         RADIUS OF INFLUENCE WHICH ENTERS INTO THE
'                 WEIGHTS FOR Q(K) (SEE SUBROUTINE SETUP2)
'    RS =         SQUARED DISTANCE BETWEEN K AND NPTS(LNP) --
'                 USED TO COMPUTE RQ AND RSQ(K)
'    RSMX =       MAXIMUM RSQ ELEMENT ENCOUNTERED
'    RSOLD =      SQUARED DISTANCE BETWEEN K AND NPTS(LNP-1) --
'                 USED TO COMPUTE A RELATIVE CHANGE IN RS
'                 BETWEEN SUCCEEDING NPTS ELEMENTS
'    RTOL =       TOLERANCE FOR DETECTING A SUFFICIENTLY LARGE
'                 RELATIVE CHANGE IN RS.  IF THE CHANGE IS
'                 NOT GREATER THAN RTOL, THE NODES ARE
'                 TREATED AS BEING THE SAME DISTANCE FROM K
'    RWS =        CURRENT VALUE OF RSQ(K)
'    S =          SECOND COMPONENT OF THE PLANE GIVENS ROTATION
'    SF =         MARQUARDT STABILIZATION FACTOR USED TO DAMP
'                 OUT THE FIRST 3 SOLUTION COMPONENTS (SECOND
'                 PARTIALS OF THE QUADRATIC) WHEN THE SYSTEM
'                 IS ILL-CONDITIONED.  AS SF INCREASES, THE
'                 FITTING FUNCTION APPROACHES A LINEAR
'    SUM =        SUM OF SQUARED EUCLIDEAN DISTANCES BETWEEN
'                 NODE K AND THE NODES USED IN THE LEAST
'                 SQUARES FIT (UNLESS ADDITIONAL NODES ARE
'                 ADDED FOR STABILITY)
'    T =          TEMPORARY VARIABLE FOR ACCUMULATING A SCALAR
'                 PRODUCT IN THE BACK SOLVE
'    XK,YK =      COORDINATES OF NODE K -- X(K), Y(K)
'    XMN,YMN =    LOCAL VARIABLES FOR XMIN AND YMIN
'
    Nn = N
    NNQ = NQ
    NNW = NW
    NNR = NR
    NQWMAX = MAX0(NNQ, NNW)
    LMAX = MIN0(40, Nn - 1)
    If (5 > NNQ Or 1 > NNW Or NQWMAX > LMAX Or NNR < 1) Then GoTo 20
'
'   CREATE THE CELL DATA STRUCTURE, AND INITIALIZE RSMX.
'
    Call STORE2(Nn, X(), Y(), NNR, LCELL(), LNEXT(), XMN, YMN, DDX, DDY, IERR)
    If (IERR <> 0) Then GoTo 22
    RSMX = 0#
'
    ' OUTER LOOP ON NODE K
'
    For K = 1 To Nn
        XK = X(K)
        YK = Y(K)
        FK = F(K)
'
        ' MARK NODE K TO EXCLUDE IT FROM THE SEARCH FOR NEAREST
        ' NEIGHBORS.
'
        LNEXT(K) = -LNEXT(K)
'
        ' INITIALIZE FOR LOOP ON NPTS.
'
        RS = 0#
        sum = 0#
        RWS = 0#
        RQ = 0#
        LNP = 0
'
        ' COMPUTE NPTS, LNP, RWS, NEQ, RQ, AND AVSQ.
'
1       sum = sum + RS
        If (LNP = LMAX) Then GoTo 3
        LNP = LNP + 1
        RSOLD = RS
        Call GETNP2(XK, YK, X(), Y(), NNR, LCELL(), LNEXT(), XMN, YMN, DDX, DDY, NP, RS)
        If (RS = 0#) Then GoTo 21
        NPTS(LNP) = NP
        If ((RS - RSOLD) / RS < RTOL) Then GoTo 1
        If (RWS = 0# And LNP > NNW) Then RWS = RS
        If (RQ <> 0# Or LNP <= NNQ) Then GoTo 2
'
        ' RQ = 0 (NOT YET COMPUTED) AND LNP .GT. NQ.  RQ =
        ' SQRT(RS) IS SUFFICIENTLY LARGE TO (STRICTLY) INCLUDE
        ' NQ NODES.  THE LEAST SQUARES FIT WILL INCLUDE NEQ =
        ' LNP - 1 EQUATIONS FOR 5 .LE. NQ .LE. NEQ .LT. LMAX
        ' .LE.N - 1#
'
        NEQ = LNP - 1
        RQ = Sqr(RS)
        AVSQ = sum / CDbl(NEQ)
'
        ' BOTTOM OF LOOP -- TEST FOR TERMINATION.
'
2       If (LNP > NQWMAX) Then GoTo 4
        GoTo 1
'
        ' ALL LMAX NODES ARE INCLUDED IN NPTS.  RWS AND/OR RQ**2 IS
        ' (ARBITRARILY) TAKEN TO BE 10 PERCENT LARGER THAN THE
        ' DISTANCE RS TO THE LAST NODE INCLUDED.
'
3       If (RWS = 0#) Then RWS = 1.1 * RS
        If (RQ <> 0#) Then GoTo 4
        NEQ = LMAX
        RQ = Sqr(1.1 * RS)
        AVSQ = sum / CDbl(NEQ)
'
        ' STORE RSQ(K), UPDATE RSMX IF NECESSARY, AND COMPUTE AV.
'
4       RSQ(K) = RWS
        If (RWS > RSMX) Then RSMX = RWS
        AV = Sqr(AVSQ)
'
        ' SET UP THE AUGMENTED REGRESSION MATRIX (TRANSPOSED) AS THE
        ' COLUMNS OF B, AND ZERO OUT THE LOWER TRIANGLE (UPPER
        ' TRIANGLE OF B) WITH GIVENS ROTATIONS -- QR DECOMPOSITION
        ' WITH ORTHOGONAL MATRIX Q NOT STORED.
'
        I = 0
5       I = I + 1
        NP = NPTS(I)
        IRow = MIN0(I, 6)
        'Call SETUP2(XK, YK, FK, X(NP), Y(NP), F(NP), AV, AVSQ, RQ, B(1, IROW))
        Call SETUP2(XK, YK, FK, X(NP), Y(NP), F(NP), AV, AVSQ, RQ, B(), 0, IRow)
        If (I = 1) Then GoTo 5
        IRM1 = IRow - 1
        For J = 1 To IRM1
            JP1 = J + 1
            Call GIVENS(B(J, J), B(J, IRow), C, S)
            'Call ROTATE(6 - J, C, S, B(JP1, J), B(JP1, IROW))
            Call ROTATE(6 - J, C, S, B(), JP1 - 1, J, B(), JP1 - 1, IRow)
6       Next J
        If (I < NEQ) Then GoTo 5
'
        ' TEST THE SYSTEM FOR ILL-CONDITIONING.
'
        DMin = DMIN1(Abs(B(1, 1)), Abs(B(2, 2)), Abs(B(3, 3)), _
                     Abs(B(4, 4)), Abs(B(5, 5)))
        If (DMin * RQ >= DTOL) Then GoTo 13
        If (NEQ = LMAX) Then GoTo 10
'
        ' INCREASE RQ AND ADD ANOTHER EQUATION TO THE SYSTEM TO
        ' IMPROVE THE CONDITIONING.  THE NUMBER OF NPTS ELEMENTS
        ' IS ALSO INCREASED IF NECESSARY.
'
7       RSOLD = RS
        NEQ = NEQ + 1
        If (NEQ = LMAX) Then GoTo 9
        If (NEQ = LNP) Then GoTo 8
'
        ' NEQ .LT. LNP
'
        NP = NPTS(NEQ + 1)
        RS = (X(NP) - XK) ^ 2 + (Y(NP) - YK) ^ 2
        If ((RS - RSOLD) / RS < RTOL) Then GoTo 7
        RQ = Sqr(RS)
        GoTo 5
'
        ' ADD AN ELEMENT TO NPTS.
'
8       LNP = LNP + 1
        Call GETNP2(XK, YK, X(), Y(), NNR, LCELL(), LNEXT(), XMN, YMN, DDX, DDY, NP, RS)
        If (NP = 0) Then GoTo 21
        NPTS(LNP) = NP
        If ((RS - RSOLD) / RS < RTOL) Then GoTo 7
        RQ = Sqr(RS)
        GoTo 5
'
9       RQ = Sqr(1.1 * RS)
        GoTo 5
'
        ' STABILIZE THE SYSTEM BY DAMPING SECOND PARTIALS -- ADD
        ' MULTIPLES OF THE FIRST THREE UNIT VECTORS TO THE FIRST
        ' THREE EQUATIONS.
'
10      For I = 1 To 3
            B(I, 6) = SF
            IP1 = I + 1
            For J = IP1 To 6
                B(J, 6) = 0#
11          Next J
            For J = I To 5
                JP1 = J + 1
                Call GIVENS(B(J, J), B(J, 6), C, S)
                'Call ROTATE(6 - J, C, S, B(JP1, J), B(JP1, 6))
12              Call ROTATE(6 - J, C, S, B(), JP1 - 1, J, B(), JP1 - 1, 6)
            Next J
        Next I
'
        ' TEST THE STABILIZED SYSTEM FOR ILL-CONDITIONING.
'
        DMin = DMIN1(Abs(B(1, 1)), Abs(B(2, 2)), Abs(B(3, 3)), _
                     Abs(B(4, 4)), Abs(B(5, 5)))
        If (DMin * RQ < DTOL) Then GoTo 22
'
        ' SOLVE THE 5 BY 5 TRIANGULAR SYSTEM FOR THE COEFFICIENTS
'
13      For IB = 1 To 5
            I = 6 - IB
            t = 0#
            If (I = 5) Then GoTo 15
            IP1 = I + 1
            For J = IP1 To 5
14              t = t + B(J, I) * A(J, K)
            Next J
15          A(I, K) = (B(6, I) - t) / B(I, I)
        Next IB

'
        ' SCALE THE COEFFICIENTS TO ADJUST FOR THE COLUMN SCALING.
'
        For I = 1 To 3
16          A(I, K) = A(I, K) / AVSQ
        Next I
        A(4, K) = A(4, K) / AV
        A(5, K) = A(5, K) / AV
'
        ' UNMARK K AND THE ELEMENTS OF NPTS.
'
        LNEXT(K) = -LNEXT(K)
        For I = 1 To LNP
            NP = NPTS(I)
17          LNEXT(NP) = -LNEXT(NP)
        Next I
18      ' CONTINUE
    Next K
'
    ' NO ERRORS ENCOUNTERED.
'
    XMin = XMN
    YMin = YMN
    DX = DDX
    DY = DDY
    RMAX = Sqr(RSMX)
    IER = 0
    Exit Sub
'
    ' N, NQ, NW, OR NR IS OUT OF RANGE.
'
20  IER = 1
    Exit Sub
'
    ' DUPLICATE NODES WERE ENCOUNTERED BY GETNP2.
'
21  IER = 2
    Exit Sub
'
    ' NO UNIQUE SOLUTION DUE TO COLLINEAR NODES.
'
22  XMin = XMN
    YMin = YMN
    DX = DDX
    DY = DDY
    IER = 3
'
'
'
End Sub
Private Sub GETNP2(ByVal PX#, ByVal PY#, X#(), Y#(), ByVal NR&, LCELL&(), LNEXT&(), _
    ByVal XMin#, ByVal YMin#, ByVal DX#, ByVal DY#, ByRef NP&, ByRef DSQ#)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'
'   GIVEN A SET OF N NODES AND THE DATA STRUCTURE DEFINED IN
'   SUBROUTINE STORE2, THIS SUBROUTINE USES THE CELL METHOD TO
'   FIND THE CLOSEST UNMARKED NODE NP TO A SPECIFIED POINT P.
'   NP IS THEN MARKED BY SETTING LNEXT(NP) TO -LNEXT(NP).  (A
'   NODE IS MARKED IF AND ONLY IF THE CORRESPONDING LNEXT ELE-
'   MENT IS NEGATIVE.  THE ABSOLUTE VALUES OF LNEXT ELEMENTS,
'   HOWEVER, MUST BE PRESERVED.)  THUS, THE CLOSEST M NODES TO
'   P MAY BE DETERMINED BY A SEQUENCE OF M CALLS TO THIS ROU-
'   TINE.  NOTE THAT IF THE NEAREST NEIGHBOR TO NODE K IS TO
'   BE DETERMINED (PX = X(K) AND PY = Y(K)), THEN K SHOULD BE
'   MARKED BEFORE THE CALL TO THIS ROUTINE.
'   THE SEARCH IS BEGUN IN THE CELL CONTAINING (OR CLOSEST
'   TO) P AND PROCEEDS OUTWARD IN RECTANGULAR LAYERS UNTIL ALL
'   CELLS WHICH CONTAIN POINTS WITHIN DISTANCE R OF P HAVE
'   BEEN SEARCHED, WHERE R IS THE DISTANCE FROM P TO THE FIRST
'   UNMARKED NODE ENCOUNTERED (INFINITE IF NO UNMARKED NODES
'   ARE PRESENT).
'
'   ON INPUT --
'
'   PX,PY = CARTESIAN COORDINATES OF THE POINT P WHOSE
'           NEAREST UNMARKED NEIGHBOR IS TO BE FOUND.
'
'   X,Y =   ARRAYS OF LENGTH N, FOR N .GE. 2, CONTAINING
'           THE CARTESIAN COORDINATES OF THE NODES.
'
'   NR =    NUMBER OF ROWS AND COLUMNS IN THE CELL GRID.
'           NR .GE. 1.
'
'   LCELL = NR BY NR ARRAY OF NODAL INDICES ASSOCIATED
'           WITH CELLS.
'
'   LNEXT = ARRAY OF LENGTH N CONTAINING NEXT-NODE INDI-
'           CES (OR THEIR NEGATIVES).
'
'   XMIN,YMIN,DX,DY = MINIMUM NODAL COORDINATES AND CELL
'           DIMENSIONS.  DX AND DY MUST BE
'           POSITIVE.
'
'   INPUT PARAMETERS OTHER THAN LNEXT ARE NOT ALTERED BY
'   THIS ROUTINE.  WITH THE EXCEPTION OF (PX,PY) AND THE SIGNS
'   OF LNEXT ELEMENTS, THESE PARAMETERS SHOULD BE UNALTERED
'   FROM THEIR VALUES ON OUTPUT FROM SUBROUTINE STORE2.
'
'   ON OUTPUT --
'
'   NP =    INDEX (FOR X AND Y) OF THE NEAREST UNMARKED
'           NODE TO P, OR 0 IF ALL NODES ARE MARKED OR NR
'           .LT. 1 OR DX .LE. 0 OR DY .LE. 0.  LNEXT(NP)
'           .LT. 0 IF NP .NE. 0.
'
'   DSQ =   SQUARED EUCLIDEAN DISTANCE BETWEEN P AND NODE
'           NP, OR 0 IF NP = 0.
'
'   MODULES REQUIRED BY GETNP2 -- NONE
'
'   INTRINSIC FUNCTIONS CALLED BY GETNP2 -- IABS, IFIX, SQRT
'
'***********************************************************
'
    Dim I&, I0&, I1&, I2&, IMIN&, IMax&, J&, J0&, J1&, J2&, JMIN&, JMAX&
    Dim L&, LMIN&, LN&
    Dim DELX#, DELY#, R#, RSMIN#, RSQ#, XP#, YP#
    Dim First As Boolean
'
    XP = PX
    YP = PY
'
    ' TEST FOR INVALID INPUT PARAMETERS.
'
    If (NR < 1 Or DX <= 0# Or DY <= 0#) Then GoTo 9
'
    ' INITIALIZE PARAMETERS --
'
    '   FIRST = TRUE IFF THE FIRST UNMARKED NODE HAS YET TO BE
    '           ENCOUNTERED,
    '   IMIN,IMAX,JMIN,JMAX = CELL INDICES DEFINING THE RANGE OF
    '           THE SEARCH,
    '   DELX,DELY = PX-XMIN AND PY-YMIN,
    '   I0,J0 = CELL CONTAINING OR CLOSEST TO P,
    '   I1,I2,J1,J2 = CELL INDICES OF THE LAYER WHOSE INTERSEC-
    '           TION WITH THE RANGE DEFINED BY IMIN,...,
    '           JMAX IS CURRENTLY BEING SEARCHED.
'
    First = True
    IMIN = 1
    IMax = NR
    JMIN = 1
    JMAX = NR
    DELX = XP - XMin
    DELY = YP - YMin
    I0 = Int(DELX / DX) + 1
    If (I0 < 1) Then I0 = 1
    If (I0 > NR) Then I0 = NR
    J0 = Int(DELY / DY) + 1
    If (J0 < 1) Then J0 = 1
    If (J0 > NR) Then J0 = NR
    I1 = I0
    I2 = I0
    J1 = J0
    J2 = J0
'
    ' OUTER LOOP ON LAYERS, INNER LOOP ON LAYER CELLS, EXCLUDING
    ' THOSE OUTSIDE THE RANGE (IMIN,IMAX) X (JMIN,JMAX).
'
1   For J = J1 To J2
        If (J > JMAX) Then GoTo 7
        If (J < JMIN) Then GoTo 6
        For I = I1 To I2
            If (I > IMax) Then GoTo 6
            If (I < IMIN) Then GoTo 5
            If (J <> J1 And J <> J2 And I <> I1 And I <> I2) Then GoTo 5
'
            ' SEARCH CELL (I,J) FOR UNMARKED NODES L.
'
            L = LCELL(I, J)
            If (L = 0) Then GoTo 5
'
            ' LOOP ON NODES IN CELL (I,J).
'
2           LN = LNEXT(L)
            If (LN < 0) Then GoTo 4
'
            ' NODE L IS NOT MARKED.
'
            RSQ = (X(L) - XP) ^ 2 + (Y(L) - YP) ^ 2
            If (Not First) Then GoTo 3
'
            ' NODE L IS THE FIRST UNMARKED NEIGHBOR OF P ENCOUNTERED.
            ' INITIALIZE LMIN TO THE CURRENT CANDIDATE FOR NP, AND
            ' RSMIN TO THE SQUARED DISTANCE FROM P TO LMIN.  IMIN,
            ' IMAX, JMIN, AND JMAX ARE UPDATED TO DEFINE THE SMAL-
            ' LEST RECTANGLE CONTAINING A CIRCLE OF RADIUS R =
            ' SQRT(RSMIN) CENTERED AT P, AND CONTAINED IN (1,NR) X
            ' (1,NR) (EXCEPT THAT, IF P IS OUTSIDE THE RECTANGLE
            ' DEFINED BY THE NODES, IT IS POSSIBLE THAT IMIN .GT.
            ' NR, IMAX .LT. 1, JMIN .GT. NR, OR JMAX .LT. 1).  FIRST
            ' IS RESET TO FALSE.
'
            LMIN = L
            RSMIN = RSQ
            R = Sqr(RSMIN)
            IMIN = Int((DELX - R) / DX) + 1
            If (IMIN < 1) Then IMIN = 1
            IMax = Int((DELX + R) / DX) + 1
            If (IMax > NR) Then IMax = NR
            JMIN = Int((DELY - R) / DY) + 1
            If (JMIN < 1) Then JMIN = 1
            JMAX = Int((DELY + R) / DY) + 1
            If (JMAX > NR) Then JMAX = NR
            First = False
            GoTo 4
'
            ' TEST FOR NODE L CLOSER THAN LMIN TO P.
'
3           If (RSQ >= RSMIN) Then GoTo 4
'
            ' UPDATE LMIN AND RSMIN.
'
            LMIN = L
            RSMIN = RSQ
'
            ' TEST FOR TERMINATION OF LOOP ON NODES IN CELL (I,J).
'
4           If (Abs(LN) = L) Then GoTo 5
            L = Abs(LN)
            GoTo 2
5           ' CONTINUE
        Next I
6       ' CONTINUE
    Next J
'
    ' TEST FOR TERMINATION OF LOOP ON CELL LAYERS.
'
7   If (I1 <= IMIN And I2 >= IMax And J1 <= JMIN And J2 >= JMAX) Then GoTo 8
    I1 = I1 - 1
    I2 = I2 + 1
    J1 = J1 - 1
    J2 = J2 + 1
    GoTo 1
'
    ' UNLESS NO UNMARKED NODES WERE ENCOUNTERED, LMIN IS THE
    ' CLOSEST UNMARKED NODE TO P.
'
8   If (First) Then GoTo 9
    NP = LMIN
    DSQ = RSMIN
    LNEXT(LMIN) = -LNEXT(LMIN)
    Exit Sub
'
    ' ERROR -- NR, DX, OR DY IS INVALID OR ALL NODES ARE MARKED.
'
9   NP = 0
    DSQ = 0#
'
'
'
End Sub

Private Sub SETUP2(ByVal XK#, ByVal YK#, ByVal FK#, _
    ByVal XI#, ByVal YI#, ByVal FI#, ByVal S1#, ByVal S2#, _
    ByVal R#, ByRef Row#(), ByVal IRow&, ByVal JRow&)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'
'   THIS ROUTINE SETS UP THE I-TH ROW OF AN AUGMENTED RE-
'   GRESSION MATRIX FOR A WEIGHTED LEAST-SQUARES FIT OF A
'   QUADRATIC FUNCTION Q(X,Y) TO A SET OF DATA VALUES F, WHERE
'   Q(XK,YK) = FK.  THE FIRST 3 COLUMNS (QUADRATIC TERMS) ARE
'   SCALED BY 1/S2 AND THE FOURTH AND FIFTH COLUMNS (LINEAR
'   TERMS) ARE SCALED BY 1/S1.  THE WEIGHT IS (R-D)/(R*D) IF
'   R .GT. D AND 0 IF R .LE. D, WHERE D IS THE DISTANCE
'   BETWEEN NODES I AND K.
'
'   ON INPUT --
'
'    XK,YK,FK = COORDINATES AND DATA VALUE AT NODE K --
'               INTERPOLATED BY Q.
'
'    XI,YI,FI = COORDINATES AND DATA VALUE AT NODE I.
'
'    S1,S2 =    RECIPROCALS OF THE SCALE FACTORS.
'
'    R =        RADIUS OF INFLUENCE ABOUT NODE K DEFINING THE
'               WEIGHT.
'
'    ROW =      ARRAY OF LENGTH 6.
'
'   INPUT PARAMETERS ARE NOT ALTERED BY THIS ROUTINE.
'
'   ON OUTPUT --
'
'    ROW =      VECTOR CONTAINING A ROW OF THE AUGMENTED
'               REGRESSION MATRIX.
'
'   MODULES REQUIRED BY SETUP2 -- NONE
'
'   INTRINSIC FUNCTION CALLED BY SETUP2 -- SQRT
'
'***********************************************************
'
    Dim I&, DX#, DY#, DXSQ#, DYSQ#, D#, w#, W1#, W2#
'
    ' LOCAL PARAMETERS -
'
    ' I =    DO-LOOP INDEX
    ' DX =   XI - XK
    ' DY =   YI - YK
    ' DXSQ = DX * DX
    ' DYSQ = DY * DY
    ' D =    DISTANCE BETWEEN NODES K AND I
    ' W =    WEIGHT ASSOCIATED WITH THE ROW
    ' W1 =   W / S1
    ' W2 =   W / S2
'
    DX = XI - XK
    DY = YI - YK
    DXSQ = DX * DX
    DYSQ = DY * DY
    D = Sqr(DXSQ + DYSQ)
    If (D <= 0# Or D >= R) Then GoTo 1
    w = (R - D) / R / D
    W1 = w / S1
    W2 = w / S2
    Row(1 + IRow, JRow) = DXSQ * W2
    Row(2 + IRow, JRow) = DX * DY * W2
    Row(3 + IRow, JRow) = DYSQ * W2
    Row(4 + IRow, JRow) = DX * W1
    Row(5 + IRow, JRow) = DY * W1
    Row(6 + IRow, JRow) = (FI - FK) * w
    Exit Sub
'
    ' NODES K AND I COINCIDE OR NODE I IS OUTSIDE OF THE RADIUS
    ' OF INFLUENCE.  SET ROW TO THE ZERO VECTOR.
'
1   For I = 1 To 6
2       Row(I + IRow, JRow) = 0#
    Next I
'
'
'
End Sub
Public Sub STORE2(ByVal N&, X#(), Y#(), ByVal NR&, ByRef LCELL&(), ByRef LNEXT&(), _
    ByRef XMin#, ByRef YMin#, ByRef DX#, ByRef DY#, ByRef IER&)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'
'   GIVEN A SET OF N ARBITRARILY DISTRIBUTED NODES IN THE
'   PLANE, THIS SUBROUTINE CREATES A DATA STRUCTURE FOR A
'   CELL-BASED METHOD OF SOLVING CLOSEST-POINT PROBLEMS.  THE
'   SMALLEST RECTANGLE CONTAINING THE NODES IS PARTITIONED
'   INTO AN NR BY NR UNIFORM GRID OF CELLS, AND NODES ARE AS-
'   SOCIATED WITH CELLS.  IN PARTICULAR, THE DATA STRUCTURE
'   STORES THE INDICES OF THE NODES CONTAINED IN EACH CELL.
'   FOR A UNIFORM RANDOM DISTRIBUTION OF NODES, THE NEAREST
'   NODE TO AN ARBITRARY POINT CAN BE DETERMINED IN CONSTANT
'   EXPECTED TIME.
'
'   ON INPUT --
'
'    N =   NUMBER OF NODES.  N .GE. 2.
'
'    X,Y = ARRAYS OF LENGTH N CONTAINING THE CARTESIAN
'          COORDINATES OF THE NODES.
'
'    NR =  NUMBER OF ROWS AND COLUMNS IN THE GRID.  THE
'          CELL DENSITY (AVERAGE NUMBER OF NODES PER CELL)
'          IS D = N/(NR**2).  A RECOMMENDED VALUE, BASED
'          ON EMPIRICAL EVIDENCE, IS D = 3 -- NR =
'          SQRT(N/3).  NR .GE. 1.
'
'   THE ABOVE PARAMETERS ARE NOT ALTERED BY THIS ROUTINE.
'
'    LCELL = ARRAY OF LENGTH .GE. NR**2.
'
'    LNEXT = ARRAY OF LENGTH .GE. N.
'
'   ON OUTPUT --
'
'    LCELL = NR BY NR CELL ARRAY SUCH THAT LCELL(I,J)
'            CONTAINS THE INDEX (FOR X AND Y) OF THE
'            FIRST NODE (NODE WITH SMALLEST INDEX) IN
'            CELL (I,J), OR LCELL(I,J) = 0 IF NO NODES
'            ARE CONTAINED IN THE CELL.  THE UPPER RIGHT
'            CORNER OF CELL (I,J) HAS COORDINATES (XMIN+
'            I*DX,YMIN+J*DY).  LCELL IS NOT DEFINED IF
'            IER .NE. 0.
'
'    LNEXT = ARRAY OF NEXT-NODE INDICES SUCH THAT
'            LNEXT(K) CONTAINS THE INDEX OF THE NEXT NODE
'            IN THE CELL WHICH CONTAINS NODE K, OR
'            LNEXT(K) = K IF K IS THE LAST NODE IN THE
'            CELL FOR K = 1,...,N.  (THE NODES CONTAINED
'            IN A CELL ARE ORDERED BY THEIR INDICES.)
'            IF, FOR EXAMPLE, CELL (I,J) CONTAINS NODES
'            2, 3, AND 5 (AND NO OTHERS), THEN LCELL(I,J)
'            = 2, LNEXT(2) = 3, LNEXT(3) = 5, AND
'            LNEXT(5) = 5.  LNEXT IS NOT DEFINED IF
'            IER .NE. 0.
'
'    XMIN,YMIN = CARTESIAN COORDINATES OF THE LOWER LEFT
'            CORNER OF THE RECTANGLE DEFINED BY THE
'            NODES (SMALLEST NODAL COORDINATES) UN-
'            LESS IER = 1.  THE UPPER RIGHT CORNER IS
'            (XMAX,YMAX) FOR XMAX = XMIN + NR*DX AND
'            YMAX = YMIN + NR*DY.
'
'    DX,DY = DIMENSIONS OF THE CELLS UNLESS IER = 1.
'            DX = (XMAX - XMIN) / NR And DY = (YMAX - YMIN) / NR
'            WHERE XMIN, XMAX, YMIN, AND YMAX ARE THE
'            EXTREMA OF X AND Y.
'
'    IER =  ERROR INDICATOR --
'            IER = 0 IF NO ERRORS WERE ENCOUNTERED.
'            IER = 1 IF N .LT. 2 OR NR .LT. 1.
'            IER = 2 IF DX = 0 OR DY = 0.
'
'   MODULES REQUIRED BY STORE2 -- NONE
'
'   INTRINSIC FUNCTIONS CALLED BY STORE2 -- FLOAT, IFIX
'
'***********************************************************
'
    Dim I&, J&, K&, L&, KB&, Nn&, NNR&, NP1&
    Dim XMN#, XMX#, YMN#, YMX#, DELX#, DELY#
'
    Nn = N
    NNR = NR
    If (Nn < 2 Or NNR < 1) Then GoTo 4
'
    ' COMPUTE THE DIMENSIONS OF THE RECTANGLE CONTAINING THE
    ' NODES.
'
    XMN = X(1)
    XMX = XMN
    YMN = Y(1)
    YMX = YMN
    For K = 2 To Nn
        If (X(K) < XMN) Then XMN = X(K)
        If (X(K) > XMX) Then XMX = X(K)
        If (Y(K) < YMN) Then YMN = Y(K)
1       If (Y(K) > YMX) Then YMX = Y(K)
    Next K
    XMin = XMN
    YMin = YMN
'
    ' COMPUTE CELL DIMENSIONS AND TEST FOR ZERO AREA.
'
    DELX = (XMX - XMN) / CDbl(NNR)
    DELY = (YMX - YMN) / CDbl(NNR)
    DX = DELX
    DY = DELY
    If (DELX = 0# Or DELY = 0#) Then GoTo 5
'
    ' INITIALIZE LCELL.
'
    For J = 1 To NNR
        For I = 1 To NNR
2           LCELL(I, J) = 0
        Next I
    Next J
'
    ' LOOP ON NODES, STORING INDICES IN LCELL AND LNEXT.
'
    NP1 = Nn + 1
    For K = 1 To Nn
        KB = NP1 - K
        I = Int((X(KB) - XMN) / DELX) + 1
        If (I > NNR) Then I = NNR
        J = Int((Y(KB) - YMN) / DELY) + 1
        If (J > NNR) Then J = NNR
        L = LCELL(I, J)
        LNEXT(KB) = L
        If (L = 0) Then LNEXT(KB) = KB
3       LCELL(I, J) = KB
    Next K
'
    ' NO ERRORS ENCOUNTERED
'
    IER = 0
    Exit Sub
'
    ' INVALID INPUT PARAMETER
'
4   IER = 1
    Exit Sub
'
    ' DX = 0 Or DY = 0
'
5   IER = 2
'
'
'
End Sub
Private Sub GIVENS(ByRef A#, ByRef B#, ByRef C#, ByRef S#)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'
'   THIS ROUTINE CONSTRUCTS THE GIVENS PLANE ROTATION --
'        ( C  S)
'    G = (     ) WHERE C*C + S*S = 1 -- WHICH ZEROS THE SECOND
'        (-S  C)
'    ENTRY OF THE 2-VECTOR (A B)-TRANSPOSE.  A CALL TO GIVENS
'    IS NORMALLY FOLLOWED BY A CALL TO ROTATE WHICH APPLIES
'    THE TRANSFORMATION TO A 2 BY N MATRIX.  THIS ROUTINE WAS
'    TAKEN FROM LINPACK.
'
'   ON INPUT --
'
'    A,B =  COMPONENTS OF THE 2-VECTOR TO BE ROTATED.
'
'   ON OUTPUT --
'
'    A =    VALUE OVERWRITTEN BY R = +/-SQRT(A*A + B*B)
'
'    B =    VALUE OVERWRITTEN BY A VALUE Z WHICH ALLOWS C
'           AND S TO BE RECOVERED AS FOLLOWS --
'           C = SQRT(1-Z*Z), S=Z     IF ABS(Z) .LE. 1.
'           C = 1/Z, S = SQRT(1-C*C) IF ABS(Z) .GT. 1.
'
'    C =    +/-(A/R)
'
'    S =    +/-(B/R)
'
'   MODULES REQUIRED BY GIVENS -- NONE
'
'   INTRINSIC FUNCTIONS CALLED BY GIVENS - ABS, SQRT
'
'***********************************************************
'
    Dim AA#, BB#, R#, u#, v#
'
    ' LOCAL PARAMETERS --
'
    ' AA,BB = LOCAL COPIES OF A AND B
    ' R =     C*A + S*B = +/-SQRT(A*A+B*B)
    ' U,V =   VARIABLES USED TO SCALE A AND B FOR COMPUTING R
'
    AA = A
    BB = B
    If (Abs(AA) <= Abs(BB)) Then GoTo 1
'
    ' ABS(A) .GT. ABS(B)
'
    u = AA + AA
    v = BB / u
    R = Sqr(0.25 + v * v) * u
    C = AA / R
    S = v * (C + C)
'
    ' NOTE THAT R HAS THE SIGN OF A, C .GT. 0, AND S HAS
    ' SIGN(A)*SIGN(B).
'
    B = S
    A = R
    Exit Sub
'
    ' ABS(A) .LE. ABS(B)
'
1   If (BB = 0#) Then GoTo 2
    u = BB + BB
    v = AA / u
'
    ' STORE R IN A.
'
    A = Sqr(0.25 + v * v) * u
    S = BB / A
    C = v * (S + S)
'
    ' NOTE THAT R HAS THE SIGN OF B, S .GT. 0, AND C HAS
    ' SIGN(A)*SIGN(B).
'
    B = 1#
    If (C <> 0#) Then B = 1# / C
    Exit Sub
'
    ' A = B = 0#
'
2   C = 1#
    S = 0#
'
'
'
End Sub
Private Sub ROTATE(ByVal N&, ByVal C#, ByVal S#, _
    ByRef X#(), ByVal Ix&, ByVal Jx&, ByRef Y#(), ByVal Iy&, ByVal Jy&)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'
'                                            ( C  S)
'   THIS ROUTINE APPLIES THE GIVENS ROTATION (     ) TO THE
'                                            (-S  C)
'                 (X(1) ... X(N))
'   2 BY N MATRIX (             ).
'                 (Y(1) ... Y(N))
'
'   ON INPUT --
'
'    N =    NUMBER OF COLUMNS TO BE ROTATED.
'
'    C,S =  ELEMENTS OF THE GIVENS ROTATION.  THESE MAY BE
'           DETERMINED BY SUBROUTINE GIVENS.
'
'    X,Y =  ARRAYS OF LENGTH .GE. N CONTAINING THE VECTORS
'           TO BE ROTATED.
'
'   PARAMETERS N, C, AND S ARE NOT ALTERED BY THIS ROUTINE.
'
'   ON OUTPUT --
'
'    X,Y =  ROTATED VECTORS.
'
'   MODULES REQUIRED BY ROTATE -- NONE
'
'***********************************************************
'
    Dim I&, XI#, YI#
'
    ' LOCAL PARAMETERS --
'
    ' I =     DO-LOOP INDEX
    ' XI, YI = X(I), Y(I)
'
    If (N <= 0 Or (C = 1# And S = 0#)) Then Exit Sub
    For I = 1 To N
        XI = X(I + Ix, Jx)
        YI = Y(I + Iy, Jy)
        X(I + Ix, Jx) = C * XI + S * YI
        Y(I + Iy, Jy) = -S * XI + C * YI
1       ' CONTINUE
    Next I
'
'
'
End Sub
Public Sub Main()
'
'   ALGORITHM 660, COLLECTED ALGORITHMS FROM ACM.
'   THIS WORK PUBLISHED IN TRANSACTIONS ON MATHEMATICAL SOFTWARE,
'   VOL. 14, NO. 2, P.149.
'
'   QS2TEST
'
'   THIS PROGRAM TESTS THE SCATTERED DATA INTERPOLATION
'   PACKAGE QSHEP2D BY PRINTING THE MAXIMUM ERRORS ASSOCIATED
'   WITH INTERPOLATED VALUES AND GRADIENTS ON A 10 BY 10
'   UNIFORM GRID IN THE UNIT SQUARE.  THE DATA SET CONSISTS
'   OF 36 NODES WITH DATA VALUES TAKEN FROM A QUADRATIC FUNC-
'   TION FOR WHICH THE METHOD IS EXACT.  THE RATIO OF MAXIMUM
'   INTERPOLATION ERROR RELATIVE TO THE MACHINE PRECISION IS
'   ALSO PRINTED.  THIS SHOULD BE O(1).  THE INTERPOLATED
'   VALUES FROM QS2VAL AND QS2GRD ARE COMPARED FOR AGREEMENT.
'
    Dim I&, IER&, J&, K&
    Dim EPS#, EP1#, EQ#, EQX#, EQY#, PX#, PY#, Q#, QX#, QY#, Q1#, RQ#
    Dim XX#, YK#, YY#, XMin#, YMin#, DX#, DY#, RMAX#
    Dim LCELL&(1 To 3, 1 To 3), LNEXT&(1 To 36)
    Dim X#(1 To 36), Y#(1 To 36), F#(1 To 36)
    Dim RSQ#(1 To 36), A#(1 To 5, 1 To 36), P#(1 To 10)
'
    ' QSHEP2 PARAMETERS AND LOGICAL UNIT FOR OUTPUT
'
    Const N& = 36, NQ& = 13, NW& = 19, NR& = 3 ', LOUT& = 6
'
    ' QUADRATIC TEST FUNCTION AND PARTIAL DERIVATIVES
'
    'FQ(XX, YY) = ((XX + 2# * YY) / 3#) ^ 2
    'FX(XX, YY) = 2# * (XX + 2# * YY) / 9#
    'FY(XX, YY) = 4# * (XX + 2# * YY) / 9#
'
    ' GENERATE A 6 BY 6 GRID OF NODES IN THE UNIT SQUARE WITH
    ' THE NATURAL ORDERING.
'
    K = 0
    For J = 1 To 6
        YK = CDbl(6 - J) / 5#
        For I = 1 To 6
            K = K + 1
            X(K) = CDbl(I - 1) / 5#
1           Y(K) = YK
        Next I
    Next J
'
    ' COMPUTE THE DATA VALUES.
'
    For K = 1 To N
2       'F(K) = FQ(X(K), Y(K))
        F(K) = ((X(K) + 2# * Y(K)) / 3#) ^ 2
    Next K
'
    ' COMPUTE PARAMETERS DEFINING THE INTERPOLANT Q.
'
    Call QSHEP2(N, X(), Y(), F(), NQ, NW, NR, LCELL(), LNEXT(), _
             XMin, YMin, DX, DY, RMAX, RSQ(), A(), IER)
    If (IER <> 0) Then GoTo 6
'
    ' GENERATE A 10 BY 10 UNIFORM GRID OF INTERPOLATION POINTS
    ' (P(I),P(J)) IN THE UNIT SQUARE.  THE FOUR CORNERS COIN-
    ' CIDE WITH NODES.
'
    For I = 1 To 10
3       P(I) = CDbl(I - 1) / 9#
    Next I
'
    ' COMPUTE THE MACHINE PRECISION EPS.
'
    EPS = 1#
4   EPS = EPS / 2#
    EP1 = EPS + 1#
    'IF (STORE(EP1) .GT. 1.) GO TO 4
    If EP1 > 1# Then GoTo 4
    EPS = EPS * 2#
'
    ' COMPUTE INTERPOLATION ERRORS AND TEST FOR AGREEMENT IN THE
    ' Q VALUES RETURNED BY QS2VAL AND QS2GRD.
'
    EQ = 0#
    EQX = 0#
    EQY = 0#
    For J = 1 To 10
        PY = P(J)
        For I = 1 To 10
            PX = P(I)
            Q1 = QS2VAL(PX, PY, N, X(), Y(), F(), NR, LCELL(), LNEXT(), _
                        XMin, YMin, DX, DY, RMAX, RSQ(), A())
            Call QS2GRD(PX, PY, N, X(), Y(), F(), NR, LCELL(), LNEXT(), _
                        XMin, YMin, DX, DY, RMAX, RSQ(), A(), Q, QX, QY, IER)
'
            If (IER <> 0) Then GoTo 7
            If (Abs(Q1 - Q) > 3# * Abs(Q) * EPS) Then GoTo 8
'
            'EQ = AMAX1(EQ, Abs(FQ(PX, PY) - Q))
            EQ = DMAX1(EQ, Abs(((PX + 2# * PY) / 3#) ^ 2 - Q))
'            EQX = AMAX1(EQX, Abs(FX(PX, PY) - QX))
'5           EQY = AMAX1(EQY, Abs(FY(PX, PY) - QY))
            EQX = DMAX1(EQX, Abs(2# * (PX + 2# * PY) / 9# - QX))
5           EQY = DMAX1(EQY, Abs(4# * (PX + 2# * PY) / 9# - QY))
        Next I
    Next J
'
    ' PRINT ERRORS AND THE RATIO EQ/EPS.
'
    Dim M$
    RQ = EQ / EPS
    M$ = "MAXIMUM ABSOLUTE ERRORS IN THE INTERPOLANT Q AND PARTIAL" & vbNewLine _
       & "DERIVATIVES QX AND QY RELATIVE TO MACHINE PRECISION EPS" & vbNewLine _
       & vbNewLine _
       & "FUNCTION  MAX ERROR  MAX ERROR/EPS/" & vbNewLine _
       & "Q                   " & Format$(EQ, "0.0000E-00") & "     " & RQ & vbNewLine _
       & "QX                  " & Format$(EQX, "0.0000E-00") & vbNewLine _
       & "QY                  " & Format$(EQY, "0.0000E-00") & vbNewLine
    MsgBox M$, vbInformation, " QSHEP2D"
'    WRITE (LOUT,100)
'    WRITE (LOUT,110) EQ, RQ
'    WRITE (LOUT,120) EQX
'    WRITE (LOUT,130) EQY
'    Stop
'100 FORMAT (///1H ,31HMAXIMUM ABSOLUTE ERRORS IN THE ,
'   .        25HINTERPOLANT Q AND PARTIAL/
'   .        1H ,31HDERIVATIVES QX AND QY RELATIVE ,
'   .        24HTO MACHINE PRECISION EPS//
'   .        1H ,10X,8HFUNCTION,3X,9HMAX ERROR,3X,
'   .        13HMAX ERROR/EPS/)
'110 FORMAT (1H ,13X,1HQ,7X,E9.3,7X,F4.2)
'120 FORMAT (1H ,13X,2HQX,6X,E9.3)
'130 FORMAT (1H ,13X,2HQY,6X,E9.3)
'
    ' ERROR IN QSHEP2
'
6 'WRITE (LOUT,200) IER
    If (IER <> 0) Then MsgBox "ERROR  IN QSHEP2 -- IER = " & IER
    Exit Sub 'Stop
'200 FORMAT (///1H ,28H*** ERROR IN QSHEP2 -- IER =,I2, 4H ***)
'
    ' ERROR IN QS2GRD
'
7   'WRITE (LOUT,210) IER
    MsgBox "ERROR  IN QS2GRD -- IER = " & IER
    Exit Sub 'Stop
'210 FORMAT (///1H ,28H*** ERROR IN QS2GRD -- IER =,I2, 4H ***)
'
    ' VALUES RETURNED BY QS2VAL AND QS2GRD DIFFER BY A RELATIVE
    ' AMOUNT GREATER THAN 3*EPS.
'
8   'WRITE (LOUT,220) Q1, Q
    MsgBox "*** ERROR -- INTERPOLATED VALUES ," & vbNewLine _
         & "(QS2VAL) AND Q2 (QS2GRD) DIFFER -- " & Q1 & "  " & Q, vbCritical, " QSHEP2D"
    'Stop
'220 FORMAT (///1H ,33H*** ERROR -- INTERPOLATED VALUES ,
'.        37HQ1 (QS2VAL) AND Q2 (QS2GRD) DIFFER --//
'.        1H ,5X,5HQ1 = ,E21.14,5X,5HQ2 = ,E21.14)
'
'
'
End Sub
Public Function QS2VAL(ByVal PX#, ByVal PY#, ByVal N&, X#(), Y#(), F#(), ByVal NR&, _
    LCELL&(), LNEXT&(), ByVal XMin#, ByVal YMin#, ByVal DX#, ByVal DY#, _
    ByVal RMAX#, RSQ#(), A#()) As Double
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'   10 / 28 / 87
'
'   THIS FUNCTION RETURNS THE VALUE Q(PX,PY) WHERE Q IS THE
'   WEIGHTED SUM OF QUADRATIC NODAL FUNCTIONS DEFINED IN SUB-
'   ROUTINE QSHEP2.  QS2GRD MAY BE CALLED TO COMPUTE A GRADI-
'   ENT OF Q ALONG WITH THE VALUE, AND/OR TO TEST FOR ERRORS.
'
'   ON INPUT --
'
'    PX,PY = CARTESIAN COORDINATES OF THE POINT P AT
'            WHICH Q IS TO BE EVALUATED.
'
'    N =     NUMBER OF NODES AND DATA VALUES DEFINING Q.
'            N .GE. 6.
'
'    X,Y,F = ARRAYS OF LENGTH N CONTAINING THE NODES AND
'            DATA VALUES INTERPOLATED BY Q.
'
'    NR =    NUMBER OF ROWS AND COLUMNS IN THE CELL GRID.
'            REFER TO STORE2.  NR .GE. 1.
'
'    LCELL = NR BY NR ARRAY OF NODAL INDICES ASSOCIATED
'            WITH CELLS.  REFER TO STORE2.
'
'    LNEXT = ARRAY OF LENGTH N CONTAINING NEXT-NODE INDI-
'            CES.  REFER TO STORE2.
'
'    XMIN,YMIN,DX,DY = MINIMUM NODAL COORDINATES AND CELL
'            DIMENSIONS.  DX AND DY MUST BE
'            POSITIVE.  REFER TO STORE2.
'
'    RMAX =  SQUARE ROOT OF THE LARGEST ELEMENT IN RSQ --
'            MAXIMUM RADIUS.
'
'    RSQ =   ARRAY OF LENGTH N CONTAINING THE SQUARED RADII
'            WHICH ENTER INTO THE WEIGHTS DEFINING Q.
'
'    A =     5 BY N ARRAY CONTAINING THE COEFFICIENTS FOR THE
'            NODAL FUNCTIONS DEFINING Q.
'
'   INPUT PARAMETERS ARE NOT ALTERED BY THIS FUNCTION.  THE
'   PARAMETERS OTHER THAN PX AND PY SHOULD BE INPUT UNALTERED
'   FROM THEIR VALUES ON OUTPUT FROM QSHEP2.  THIS FUNCTION
'   SHOULD NOT BE CALLED IF A NONZERO ERROR FLAG WAS RETURNED
'   BY QSHEP2.
'
'   ON OUTPUT --
'
'    QS2VAL = FUNCTION VALUE Q(PX,PY) UNLESS N, NR, DX,
'             DY, OR RMAX IS INVALID, IN WHICH CASE NO
'             VALUE IS RETURNED.
'
'   MODULES REQUIRED BY QS2VAL -- NONE
'
'   INTRINSIC FUNCTIONS CALLED BY QS2VAL -- IFIX, SQRT
'
'***********************************************************
'
    Dim I&, IMIN&, IMax&, J&, JMIN&, JMAX&, K&, KP&
    Dim DELX#, DELY#, DS#, DXSQ#, DYSQ#, RD#, RS#, RDS#, SW#, SWQ#, w#, XP#, YP#
'
    XP = PX
    YP = PY
    If (N < 6 Or NR < 1 Or DX <= 0# _
     Or DY <= 0# Or RMAX < 0#) Then Exit Function
'
    ' SET IMIN, IMAX, JMIN, AND JMAX TO CELL INDICES DEFINING
    ' THE RANGE OF THE SEARCH FOR NODES WHOSE RADII INCLUDE
    ' P.  THE CELLS WHICH MUST BE SEARCHED ARE THOSE INTER-
    ' SECTED BY (OR CONTAINED IN) A CIRCLE OF RADIUS RMAX
    ' CENTERED AT P.
'
    IMIN = Int((XP - XMin - RMAX) / DX) + 1
    IMax = Int((XP - XMin + RMAX) / DX) + 1
    If (IMIN < 1) Then IMIN = 1
    If (IMax > NR) Then IMax = NR
    JMIN = Int((YP - YMin - RMAX) / DY) + 1
    JMAX = Int((YP - YMin + RMAX) / DY) + 1
    If (JMIN < 1) Then JMIN = 1
    If (JMAX > NR) Then JMAX = NR
'
    ' THE FOLLOWING IS A TEST FOR NO CELLS WITHIN THE CIRCLE
    ' OF RADIUS RMAX.
'
    If (IMIN > IMax Or JMIN > JMAX) Then GoTo 5
'
    ' ACCUMULATE WEIGHT VALUES IN SW AND WEIGHTED NODAL FUNCTION
    ' VALUES IN SWQ.  THE WEIGHTS ARE W(K) = ((R-D)+/(R*D))**2
    ' FOR R**2 = RSQ(K) AND D = DISTANCE BETWEEN P AND NODE K.
'
    SW = 0#
    SWQ = 0#
'
    ' OUTER LOOP ON CELLS (I,J).
'
    For J = JMIN To JMAX
        For I = IMIN To IMax
            K = LCELL(I, J)
            If (K = 0) Then GoTo 3
'
            ' INNER LOOP ON NODES K.
'
1           DELX = XP - X(K)
            DELY = YP - Y(K)
            DXSQ = DELX * DELX
            DYSQ = DELY * DELY
            DS = DXSQ + DYSQ
            RS = RSQ(K)
            If (DS >= RS) Then GoTo 2
            If (DS = 0#) Then GoTo 4
            RDS = RS * DS
            RD = Sqr(RDS)
            w = (RS + DS - RD - RD) / RDS
            SW = SW + w
            SWQ = SWQ + w * (A(1, K) * DXSQ + A(2, K) * DELX * DELY _
                + A(3, K) * DYSQ + A(4, K) * DELX _
                + A(5, K) * DELY + F(K))
'
            ' BOTTOM OF LOOP ON NODES IN CELL (I,J).
'
2           KP = K
            K = LNEXT(KP)
            If (K <> KP) Then GoTo 1
3           ' CONTINUE
        Next I
    Next J
'
    ' SW = 0 IFF P IS NOT WITHIN THE RADIUS R(K) FOR ANY NODE K.
'
    If (SW = 0#) Then GoTo 5
    QS2VAL = SWQ / SW
    Exit Function
'
    ' (PX,PY) = (X(K),Y(K))
'
4   QS2VAL = F(K)
    Exit Function
'
    ' ALL WEIGHTS ARE 0 AT P.
'
5   QS2VAL = 0#
'
'
'
End Function
Public Sub QS2GRD(ByVal PX#, ByVal PY#, ByVal N&, X#(), Y#(), F#(), ByVal NR&, _
    LCELL&(), LNEXT&(), ByRef XMin#, ByVal YMin#, ByVal DX#, ByVal DY#, _
    ByVal RMAX#, RSQ#(), A#(), ByRef Q#, ByRef QX#, ByRef QY#, ByRef IER&)
'
'***********************************************************
'
'                       ROBERT RENKA
'                   UNIV. OF NORTH TEXAS
'                         (817) 565-2767
'   10 / 28 / 87
'
'   THIS SUBROUTINE COMPUTES THE VALUE AND GRADIENT AT
'   (PX,PY) OF THE INTERPOLATORY FUNCTION Q DEFINED IN SUB-
'   ROUTINE QSHEP2.  Q(X,Y) IS A WEIGHTED SUM OF QUADRATIC
'   NODAL FUNCTIONS.
'
'   ON INPUT --
'
'    PX,PY = CARTESIAN COORDINATES OF THE POINT AT WHICH
'            Q AND ITS PARTIALS ARE TO BE EVALUATED.
'
'    N =     NUMBER OF NODES AND DATA VALUES DEFINING Q.
'            N .GE. 6.
'
'    X,Y,F = ARRAYS OF LENGTH N CONTAINING THE NODES AND
'            DATA VALUES INTERPOLATED BY Q.
'
'    NR =    NUMBER OF ROWS AND COLUMNS IN THE CELL GRID.
'            REFER TO STORE2.  NR .GE. 1.
'
'    LCELL = NR BY NR ARRAY OF NODAL INDICES ASSOCIATED
'            WITH CELLS.  REFER TO STORE2.
'
'    LNEXT = ARRAY OF LENGTH N CONTAINING NEXT-NODE INDI-
'            CES.  REFER TO STORE2.
'
'    XMIN,YMIN,DX,DY = MINIMUM NODAL COORDINATES AND CELL
'            DIMENSIONS.  DX AND DY MUST BE
'            POSITIVE.  REFER TO STORE2.
'
'    RMAX =  SQUARE ROOT OF THE LARGEST ELEMENT IN RSQ --
'            MAXIMUM RADIUS.
'
'    RSQ =   ARRAY OF LENGTH N CONTAINING THE SQUARED RADII
'            WHICH ENTER INTO THE WEIGHTS DEFINING Q.
'
'    A =     5 BY N ARRAY CONTAINING THE COEFFICIENTS FOR THE
'            NODAL FUNCTIONS DEFINING Q.
'
'   INPUT PARAMETERS ARE NOT ALTERED BY THIS SUBROUTINE.
'   THE PARAMETERS OTHER THAN PX AND PY SHOULD BE INPUT UNAL-
'   TERED FROM THEIR VALUES ON OUTPUT FROM QSHEP2.  THIS SUB-
'   ROUTINE SHOULD NOT BE CALLED IF A NONZERO ERROR FLAG WAS
'   RETURNED BY QSHEP2.
'
'   ON OUTPUT --
'
'    Q =     VALUE OF Q AT (PX,PY) UNLESS IER .EQ. 1, IN
'            WHICH CASE NO VALUES ARE RETURNED.
'
'    QX,QY = FIRST PARTIAL DERIVATIVES OF Q AT (PX,PY)
'            UNLESS IER .EQ. 1.
'
'    IER =   ERROR INDICATOR
'             IER = 0 IF NO ERRORS WERE ENCOUNTERED.
'             IER = 1 IF N, NR, DX, DY OR RMAX IS INVALID.
'             IER = 2 IF NO ERRORS WERE ENCOUNTERED BUT
'                     (PX,PY) IS NOT WITHIN THE RADIUS R(K)
'                     FOR ANY NODE K (AND THUS Q=QX=QY=0).
'
'   MODULES REQUIRED BY QS2GRD -- NONE
'
'   INTRINSIC FUNCTIONS CALLED BY QS2GRD -- IFIX, SQRT
'
'***********************************************************
'
    Dim I&, IMIN&, IMax&, J&, JMIN&, JMAX&, K&, KP&
    Dim DELX#, DELY#, DS#, DXSQ#, DYSQ#, RD#, QK#, QKX#, QKY#, RS#, RDS#
    Dim SW#, SWS#, SWX#, SWY#, SWQ#, SWQX#, SWQY#, t#, w#, WX#, WY#, XP#, YP#
'
    XP = PX
    YP = PY
    If (N < 6 Or NR < 1 Or DX <= 0# _
     Or DY <= 0# Or RMAX < 0#) Then GoTo 5
'
    ' SET IMIN, IMAX, JMIN, AND JMAX TO CELL INDICES DEFINING
    ' THE RANGE OF THE SEARCH FOR NODES WHOSE RADII INCLUDE
    ' P.  THE CELLS WHICH MUST BE SEARCHED ARE THOSE INTER-
    ' SECTED BY (OR CONTAINED IN) A CIRCLE OF RADIUS RMAX
    ' CENTERED AT P.
'
    IMIN = Int((XP - XMin - RMAX) / DX) + 1
    IMax = Int((XP - XMin + RMAX) / DX) + 1
    If (IMIN < 1) Then IMIN = 1
    If (IMax > NR) Then IMax = NR
    JMIN = Int((YP - YMin - RMAX) / DY) + 1
    JMAX = Int((YP - YMin + RMAX) / DY) + 1
    If (JMIN < 1) Then JMIN = 1
    If (JMAX > NR) Then JMAX = NR
'
    ' THE FOLLOWING IS A TEST FOR NO CELLS WITHIN THE CIRCLE
    ' OF RADIUS RMAX.
'
    If (IMIN > IMax Or JMIN > JMAX) Then GoTo 6
'
    ' Q = SWQ/SW = SUM(W(K)*Q(K))/SUM(W(K)) WHERE THE SUM IS
    ' FROM K = 1 TO N, Q(K) IS THE QUADRATIC NODAL FUNCTION,
    ' AND W(K) = ((R-D)+/(R*D))**2 FOR RADIUS R(K) AND DIST-
    ' ANCE D(K).  THUS
    '
    '  QX = (SWQX*SW - SWQ*SWX)/SW**2  AND
    '  QY = (SWQY*SW - SWQ*SWY)/SW**2
    '
    ' WHERE SWQX AND SWX ARE PARTIAL DERIVATIVES WITH RESPECT
    ' TO X OF SWQ AND SW, RESPECTIVELY.  SWQY AND SWY ARE DE-
    ' FINED SIMILARLY.
'
    SW = 0#
    SWX = 0#
    SWY = 0#
    SWQ = 0#
    SWQX = 0#
    SWQY = 0#
'
    ' OUTER LOOP ON CELLS (I,J).
'
    For J = JMIN To JMAX
        For I = IMIN To IMax
            K = LCELL(I, J)
            If (K = 0) Then GoTo 3
'
            ' INNER LOOP ON NODES K.
'
1           DELX = XP - X(K)
            DELY = YP - Y(K)
            DXSQ = DELX * DELX
            DYSQ = DELY * DELY
            DS = DXSQ + DYSQ
            RS = RSQ(K)
            If (DS >= RS) Then GoTo 2
            If (DS = 0#) Then GoTo 4
            RDS = RS * DS
            RD = Sqr(RDS)
            w = (RS + DS - RD - RD) / RDS
            t = 2# * (RD - RS) / (DS * RDS)
            WX = DELX * t
            WY = DELY * t
            QKX = 2# * A(1, K) * DELX + A(2, K) * DELY
            QKY = A(2, K) * DELX + 2# * A(3, K) * DELY
            QK = (QKX * DELX + QKY * DELY) / 2#
            QKX = QKX + A(4, K)
            QKY = QKY + A(5, K)
            QK = QK + A(4, K) * DELX + A(5, K) * DELY + F(K)
            SW = SW + w
            SWX = SWX + WX
            SWY = SWY + WY
            SWQ = SWQ + w * QK
            SWQX = SWQX + WX * QK + w * QKX
            SWQY = SWQY + WY * QK + w * QKY
'
            ' BOTTOM OF LOOP ON NODES IN CELL (I,J).
'
2           KP = K
            K = LNEXT(KP)
            If (K <> KP) Then GoTo 1
3           ' CONTINUE
        Next I
    Next J
'
    ' SW = 0 IFF P IS NOT WITHIN THE RADIUS R(K) FOR ANY NODE K.
'
    If (SW = 0#) Then GoTo 6
    Q = SWQ / SW
    SWS = SW * SW
    QX = (SWQX * SW - SWQ * SWX) / SWS
    QY = (SWQY * SW - SWQ * SWY) / SWS
    IER = 0
    Exit Sub
'
    ' (PX,PY) = (X(K),Y(K))
'
4   Q = F(K)
    QX = A(4, K)
    QY = A(5, K)
    IER = 0
    Exit Sub
'
    ' INVALID INPUT PARAMETER.
'
5   IER = 1
    Exit Sub
'
    ' NO CELLS CONTAIN A POINT WITHIN RMAX OF P, OR
    ' SW = 0 AND THUS DS .GE. RSQ(K) FOR ALL K.
'
6   Q = 0#
    QX = 0#
    QY = 0#
    IER = 2
'
'
'
End Sub
