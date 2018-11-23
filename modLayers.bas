Attribute VB_Name = "modLayers"
'==============================================================
' Description......: Subroutine CONREC
' Name of the Files: modLayers.bas
' Date.............: Nov. 1990 - 1?Versione in QuickBasic,
'                    F. Languasco - URL Colworth
' Update...........: 26/7/1999
' Update...........: 24/7/2001 (addition of labeling of the
'                     level lines).
' Update...........: 20/10/2001 (added the generation of points
'                     of the level lines: CONREC_pK).
' Version..........: 2.0 at 32 bits.
' System...........: Visual Basic 6.0 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.virgilio.it/flanguasco/
'                    http://www.flanguasco.org
'==============================================================
'
'   CONREC subroutine for tracking level curves:
'
'   The values of the surface to be traced must be passed in the matrix
'    D(1 to IUB, 1 to JUB): this represents a grid with columns from 1 to IUB
'    and rows from 1 to JUB.
'
'   The actual coordinates of columns and rows, to which the level values of
'    the matrix D(.., ..) correspond, are passed in the vectors X(1 to IUB),
'    Y(1 to JUB): these do not necessarily have to be equispaced and normally,
'    they correspond to the scaling of the sheet (Page) on which you want to
'    draw.
'
'   The NL values to which the level lines are to be drawn are passed into the
'    vector ZLType(1 to NL).dLinVal:
'   these values do not necessarily have to be equispaced but must be in
'    increasing order. In the ZLTipo(1 to NL).lLinCol and ZLTipo(1 to NL)
'    .lLinSps vectors, the colors (in RGB format) and the thicknesses with
'    which to draw the corresponding level curves are specified.
'
'   In the vectors ZLTipo(1 to NL).sLblTes, ZLTipo(1 to NL).lLblCol and
'    ZLTipo(1 to NL).lLBLFSz are passed the texts, colors and sizes of the
'    labels that you want to write in correspondence of each line of level.
'
'   Page is the Picture Box on which you want to draw: it can be, instead of a
'    PictureBox, any object that supports the Line method.
'
'   Msg$ is returned empty if CONREC did not find errors, otherwise it contains
'    the description of the error.
'
'   The variable fStop can be set to True, in the calling program, to stop the
'    tracing of the level curves.
'
'   Note:   The present version of CONREC (CONREC_pK) has been modified, with
'            respect to the original, with precedence to the exploration by
'            level with respect to the exploration by box.
'           This allowed me to construct, when necessary, an ordered vector
'            containing the succession of the coordinates of the points
'            representing one (or more) curves of a certain level.
'           For reference purposes only, the Sort routine that performs the
'            above is reported.
'
'   Note:   All the vectors and matrices of these routines start from index 1.
'
Option Explicit
'
' Editable settings:
Private Const lblFNm$ = "Courier New"   ' Fonts of the labels.
Private Const lblFTr As Boolean = True  ' Transparent labels background.
Private Const lblFBl As Boolean = True  ' Text of the "Bold" labels.
'
' Structure of the information related to the level lines to be drawn
'  (CONREC_pK):
Public Type LineaLivello_Type
    dLinVal As Double      ' Value     of the level line.
    lLinCol As Long        ' Color     "   "    "    " (default = vbBlack).
    lLinSps As Long        ' Thickness "   "    "    " (default = 1 [Pixel]).
    sLblTes As String      ' Text      of   the label.
    lLblCol As Long        ' Color     "       "  (default = vbBlack).
    lLblFSz As Long        ' Font size "       "  (default = 8 [Points]).
End Type
'
' Support structure for writing layer labels (CONREC_pK):
Private Type zLbl_Type
    LblW As Double  ' Width of the label.
    LblH As Double  ' Height of the label.
    x1 As Single    ' Positions of the corners
    x2 As Single    '  of the area engaged
    y1 As Single    '  by the label.
    y2 As Single    '
End Type
'
'----------------------------------------------------------------------------------
'   Declarations for Order:
Private Type SegOrg_Type
    NS As Boolean   ' Segment still to be selected.
    x1 As Long      ' Segment coordinates: they must be
    y1 As Long      '  transformed by UserScale
    x2 As Long      '  (used by CONREC_pK) to Pixels.
    y2 As Long      '
End Type
Dim PO() As SegOrg_Type ' Vector of the disordered segments
                        '  as designed, for a certain level,
                        '  by CONREC_pK.
'
Private Type SegOrd_Type
    NC As Long      ' Number of the curve of belonging.
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type
Dim PT() As SegOrd_Type ' Vector of segments ordered in
                        '  succession and by curve.

Private Sub Ordina()
'
'
    Dim I&, I1&, NPi&, NPu&, NCu&, NPO&
    Dim SegDisp As Boolean  ' There are still segments available.
    Dim SegForL As Boolean  ' Segments were used in the For loop.
'
    ' Choose consecutive segments:
    I1 = 0
    NCu = 0
    NPu = 0
    NPO = UBound(PO)
    Do
        ' Search for the first segment still available in PO():
        I1 = PrimoDisponibile(PO(), I1 + 1, NPO)
'
        NCu = NCu + 1   ' Number of the current curve.
        NPu = NPu + 1   ' Pointer in PT () to the last segment in the current curve.
        NPi = NPu       ' Pointer in PT () to the first segment in the current curve.
        ReDim Preserve PT(1 To NPu)
        PT(NPi).x1 = PO(I1).x1
        PT(NPi).y1 = PO(I1).y1
        PT(NPi).x2 = PO(I1).x2
        PT(NPi).y2 = PO(I1).y2
        PT(NPi).NC = NCu
        PO(I1).NS = False
'
        Do
            SegDisp = False
            SegForL = False
            For I = I1 + 1 To NPO
                If PO(I).NS Then
                    ' Search for segments to add at the end of the current curve:
                    If (PT(NPu).x2 = PO(I).x1) And (PT(NPu).y2 = PO(I).y1) Then
                        GoSub Coda
                        PT(NPu).x1 = PO(I).x1
                        PT(NPu).y1 = PO(I).y1
                        PT(NPu).x2 = PO(I).x2
                        PT(NPu).y2 = PO(I).y2
                    ElseIf (PT(NPu).x2 = PO(I).x2) And (PT(NPu).y2 = PO(I).y2) Then
                        GoSub Coda
                        PT(NPu).x1 = PO(I).x2
                        PT(NPu).y1 = PO(I).y2
                        PT(NPu).x2 = PO(I).x1
                        PT(NPu).y2 = PO(I).y1
'
                    ' Search for segments to add at the top of the current curve:
                    ElseIf (PT(NPi).x1 = PO(I).x2) And (PT(NPi).y1 = PO(I).y2) Then
                        GoSub Testa
                        PT(NPi).x1 = PO(I).x1
                        PT(NPi).y1 = PO(I).y1
                        PT(NPi).x2 = PO(I).x2
                        PT(NPi).y2 = PO(I).y2
                    ElseIf (PT(NPi).x1 = PO(I).x1) And (PT(NPi).y1 = PO(I).y1) Then
                        GoSub Testa
                        PT(NPi).x1 = PO(I).x2
                        PT(NPi).y1 = PO(I).y2
                        PT(NPi).x2 = PO(I).x1
                        PT(NPi).y2 = PO(I).y1
'
                    Else
                        SegDisp = True   ' There is still at least one segment available.
                    End If
                End If
            Next I
        Loop While (SegDisp And SegForL) ' Are there still any segments available and
                                         '  the For loop has not failed?
'
    Loop While SegDisp  ' The For loop has failed but there are
                        '  still segments available?
'
    Exit Sub
'
'
Testa:
    NPu = NPu + 1
    ReDim Preserve PT(1 To NPu)
    PO(I).NS = False
    PT(NPi).NC = NCu
    SegForL = True
    FaiSpazio PT(), NPi, NPu
    Return
'
Coda:
    NPu = NPu + 1
    ReDim Preserve PT(1 To NPu)
    PO(I).NS = False
    PT(NPu).NC = NCu
    SegForL = True
    Return
'
'
'
End Sub

Private Sub FaiSpazio(Seg() As SegOrd_Type, ByVal II&, ByVal Iu&)
'
'   Create a free space in the vector Seg(1 To Iu) at the Ii location:
'
    Dim I&
'
    For I = Iu To II + 1 Step -1
        Seg(I) = Seg(I - 1)
    Next I
'
'
'
End Sub

Private Function PrimoDisponibile(Seg() As SegOrg_Type _
    , ByVal I1&, ByVal I2&) As Long
'
'   Returns the position of the first segment of Seg (),
'    with index> = I1, having the flag NS = True:
'
    Dim I&
'
    For I = I1 To I2
        If Seg(I).NS Then
            PrimoDisponibile = I
            Exit Function
        End If
    Next I
'
    PrimoDisponibile = 0
'
'
'
End Function

Public Sub CONREC_pK(ByVal Page As PictureBox, D#(), X#(), Y#(), _
    ZLTipo() As LineaLivello_Type, _
    Optional ByRef Msg$, _
    Optional ByRef fStop As Boolean = False)
'
'   Subroutine CONREC
'   Da: A countouring subroutine.
'   Di: P. Bourke
'   Byte Jun/1987 - Pg. 143, 150.
'
'   Input Variables to CONREC:
'
'   IUB, JUB                Index bounds of the Data array.
'   D(1 to IUB, 1 to JUB)   Matrix for the Data Surface.
'   X(1 to IUB)             Data array for Column coordinates.
'   Y(1 to JUB)             Data array for Row    coordinates.
'   NL                      Number of Contour Levels.
'   ZLTipo(1 to NL)         A structure array containing
'                           the Contour Levels information.
'
    Dim IUB&, JUB&, NL&
    Dim K&, K1&, J&, I&, M&, M1&, M2&, M3&, CAS&
    Dim x1#, y1#, x2#, y2#, DMin#, DMax#
'
    ' Definitions for layer labels:
    Dim PageFN$, PageFS&, PageDW&, PageFC&
    Dim fzLbl As Boolean, PageFB As Boolean, PageFT As Boolean
'
    Dim H#(0 To 4), XH#(0 To 4), YH#(0 To 4), ISH&(0 To 4)
'
    Dim IM&(0 To 3)
    IM(0) = 0: IM(1) = 1: IM(2) = 1:  IM(3) = 0
'
    Dim JM&(0 To 3)
    JM(0) = 0: JM(1) = 0: JM(2) = 1: JM(3) = 1
'
    Dim CT&(0 To 2, 0 To 2, 0 To 2)
    CT(0, 0, 0) = 0: CT(0, 0, 1) = 0: CT(0, 0, 2) = 8
    CT(0, 1, 0) = 0: CT(0, 1, 1) = 2: CT(0, 1, 2) = 5
    CT(0, 2, 0) = 7: CT(0, 2, 1) = 6: CT(0, 2, 2) = 9
    CT(1, 0, 0) = 0: CT(1, 0, 1) = 3: CT(1, 0, 2) = 4
    CT(1, 1, 0) = 1: CT(1, 1, 1) = 3: CT(1, 1, 2) = 1
    CT(1, 2, 0) = 4: CT(1, 2, 1) = 3: CT(1, 2, 2) = 0
    CT(2, 0, 0) = 9: CT(2, 0, 1) = 6: CT(2, 0, 2) = 7
    CT(2, 1, 0) = 5: CT(2, 1, 1) = 2: CT(2, 1, 2) = 0
    CT(2, 2, 0) = 8: CT(2, 2, 1) = 0: CT(2, 2, 2) = 0
'
   On Error GoTo CONREC_ERR
'
    IUB = UBound(X)
    JUB = UBound(Y)
    NL = UBound(ZLTipo)
'
    If IUB <= 1 Or JUB <= 1 Or NL < 1 Then
        Err.Raise 1001, "CONREC", "Error in Input Parameters"
    End If
    For K = 2 To NL
        If ZLTipo(K).dLinVal <= ZLTipo(K - 1).dLinVal Then
            Err.Raise 1001, "CONREC", "Error in Contour Levels"
            Exit For
        End If
    Next K
'
    ' Calculate the length in [Characters] of the labels and,
    '  if necessary, display them:
    ReDim zLbl(1 To NL) As zLbl_Type
    For K = 1 To NL
        zLbl(K).LblW = Len(Trim(ZLTipo(K).sLblTes))
        If zLbl(K).LblW > 0 Then
            zLbl(K).LblW = zLbl(K).LblW + 2
            fzLbl = True
        End If
    Next K
'
    PageDW = Page.DrawWidth
    If fzLbl Then   ' We want to write the values of the levels.
        ' Save the current settings, assign the new ones and calculate
        '  the position parameters:
        PageFN$ = Page.FontName
        PageFS = Page.FontSize
        PageFC = Page.ForeColor
        PageFB = Page.FontBold
        PageFT = Page.FontTransparent
        Page.FontName = lblFNm$
        Page.FontBold = lblFBl
        Page.FontTransparent = lblFTr
        For K = 1 To NL
            If ZLTipo(K).lLblFSz < 8 Then ZLTipo(K).lLblFSz = 8
            Page.FontSize = ZLTipo(K).lLblFSz
            ' Calculate the width and height of the label in [ScaleUnits] and
            '  use it's FontSize:
            zLbl(K).LblW = Page.TextWidth("O") * zLbl(K).LblW
            zLbl(K).LblH = Abs(Page.TextHeight("O"))
        Next K
    End If
'
    For K = 1 To NL
        If ZLTipo(K).lLinSps = 0 Then ZLTipo(K).lLinSps = 1
        Page.DrawWidth = ZLTipo(K).lLinSps
        For J = 1 To JUB - 1
            For I = 1 To IUB - 1
'
                If D(I, J) < D(I, J + 1) Then
                    DMin = D(I, J)
                Else
                    DMin = D(I, J + 1)
                End If
                If D(I + 1, J) < DMin Then DMin = D(I + 1, J)
                If D(I + 1, J + 1) < DMin Then DMin = D(I + 1, J + 1)
'
                If D(I, J) > D(I, J + 1) Then
                    DMax = D(I, J)
                Else
                    DMax = D(I, J + 1)
                End If
                If D(I + 1, J) > DMax Then DMax = D(I + 1, J)
                If D(I + 1, J + 1) > DMax Then DMax = D(I + 1, J + 1)
'
                If ZLTipo(K).dLinVal < DMin _
                Or DMax < ZLTipo(K).dLinVal _
                Or DMax = DMin Then GoTo NOITRI
'
                For M = 4 To 0 Step -1
                    If M > 0 Then
                        H(M) = D(I + IM(M - 1), J + JM(M - 1)) - ZLTipo(K).dLinVal
                        XH(M) = X(I + IM(M - 1))
                        YH(M) = Y(J + JM(M - 1))
                    End If
                    If M = 0 Then
                        H(0) = (H(1) + H(2) + H(3) + H(4)) / 4#
                        XH(0) = (X(I) + X(I + 1)) / 2#
                        YH(0) = (Y(J) + Y(J + 1)) / 2#
                    End If
                    If H(M) > 0 Then ISH(M) = 2
                    If H(M) < 0 Then ISH(M) = 0
                    If H(M) = 0 Then ISH(M) = 1
                Next M
'
                For M = 1 To 4
                    M1 = M
                    M2 = 0
                    M3 = M + 1
                    If M3 = 5 Then M3 = 1
                    CAS = CT(ISH(M1), ISH(M2), ISH(M3))
                    If CAS = 0 Then GoTo CASE0
                    Select Case CAS
                        Case 1
                        x1 = XH(M1)
                        y1 = YH(M1)
                        x2 = XH(M2)
                        y2 = YH(M2)
'
                        Case 2
                        x1 = XH(M2)
                        y1 = YH(M2)
                        x2 = XH(M3)
                        y2 = YH(M3)
'
                        Case 3
                        x1 = XH(M3)
                        y1 = YH(M3)
                        x2 = XH(M1)
                        y2 = YH(M1)
'
                        Case 4
                        x1 = XH(M1)
                        y1 = YH(M1)
                        x2 = (H(M3) * XH(M2) - H(M2) * XH(M3)) / (H(M3) - H(M2))
                        y2 = (H(M3) * YH(M2) - H(M2) * YH(M3)) / (H(M3) - H(M2))
'
                        Case 5
                        x1 = XH(M2)
                        y1 = YH(M2)
                        x2 = (H(M1) * XH(M3) - H(M3) * XH(M1)) / (H(M1) - H(M3))
                        y2 = (H(M1) * YH(M3) - H(M3) * YH(M1)) / (H(M1) - H(M3))
'
                        Case 6
                        x1 = XH(M3)
                        y1 = YH(M3)
                        x2 = (H(M2) * XH(M1) - H(M1) * XH(M2)) / (H(M2) - H(M1))
                        y2 = (H(M2) * YH(M1) - H(M1) * YH(M2)) / (H(M2) - H(M1))
'
                        Case 7
                        x1 = (H(M2) * XH(M1) - H(M1) * XH(M2)) / (H(M2) - H(M1))
                        y1 = (H(M2) * YH(M1) - H(M1) * YH(M2)) / (H(M2) - H(M1))
                        x2 = (H(M3) * XH(M2) - H(M2) * XH(M3)) / (H(M3) - H(M2))
                        y2 = (H(M3) * YH(M2) - H(M2) * YH(M3)) / (H(M3) - H(M2))
'
                        Case 8
                        x1 = (H(M3) * XH(M2) - H(M2) * XH(M3)) / (H(M3) - H(M2))
                        y1 = (H(M3) * YH(M2) - H(M2) * YH(M3)) / (H(M3) - H(M2))
                        x2 = (H(M1) * XH(M3) - H(M3) * XH(M1)) / (H(M1) - H(M3))
                        y2 = (H(M1) * YH(M3) - H(M3) * YH(M1)) / (H(M1) - H(M3))
'
                        Case 9
                        x1 = (H(M1) * XH(M3) - H(M3) * XH(M1)) / (H(M1) - H(M3))
                        y1 = (H(M1) * YH(M3) - H(M3) * YH(M1)) / (H(M1) - H(M3))
                        x2 = (H(M2) * XH(M1) - H(M1) * XH(M2)) / (H(M2) - H(M1))
                        y2 = (H(M2) * YH(M1) - H(M1) * YH(M2)) / (H(M2) - H(M1))
                    End Select
'
                    ' Management of a possible end-of-cycle flag:
                    DoEvents
                    If fStop Then GoTo CONREC_END
'
DRAWIT:
                    Page.Line (x1, y1)-(x2, y2), ZLTipo(K).lLinCol
                    If ZLTipo(K).lLinSps = 1 Then Page.PSet (x2, y2), ZLTipo(K).lLinCol
'
                    ' Level line labels:
                    If zLbl(K).LblW > 0 Then    ' Try writing the value of the
                                                '  level line K to position
                                                '  x2, y2.
                        ' Discard the out-of-square positions:
                        If (X(IUB) < x2 + zLbl(K).LblW) _
                        Or (y2 - zLbl(K).LblH < Y(1)) Then GoTo CASE0
                        ' and those already engaged:
                        For K1 = 1 To NL
                            If (K1 <> K) _
                            And (zLbl(K1).x1 < x2 + zLbl(K).LblW) _
                            And (x2 < zLbl(K1).x2) _
                            And (zLbl(K1).y1 < y2) _
                            And (y2 - zLbl(K).LblH < zLbl(K1).y2) Then GoTo CASE0
                        Next K1
'
                        ' The position x2, y2 is OK:
                        zLbl(K).x1 = x2                 ' Area
                        zLbl(K).x2 = x2 + zLbl(K).LblW  '  occupied
                        zLbl(K).y1 = y2 - zLbl(K).LblH  '  by the
                        zLbl(K).y2 = y2                 '  K label.
                        zLbl(K).LblW = 0                '  K-level labeled.
'
                        Page.ForeColor = ZLTipo(K).lLblCol
                        Page.FontSize = ZLTipo(K).lLblFSz
'
                        Page.Print ZLTipo(K).sLblTes
                    End If
'
CASE0:          Next M
'
NOITRI:     Next I
'
NOIBOX: Next J
'
    Next K
'
CONREC_END:
    Page.DrawWidth = PageDW
    If fzLbl Then
        ' Reset settings:
        Page.FontName = PageFN$
        Page.FontSize = PageFS
        Page.ForeColor = PageFC
        Page.FontBold = PageFB
        Page.FontTransparent = PageFT
    End If
'
CONREC_ERR:
    Msg$ = IIf(Err = 0, "", Err.Description)
'
'
'
End Sub
