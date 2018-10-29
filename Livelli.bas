Attribute VB_Name = "modLivelli"
'==============================================================
' Descrizione.....: Subroutine CONREC
' Nome dei Files..: Livelli.bas
' Data............: Nov. 1990 - 1° Versione in QuickBasic,
'                   F. Languasco - URL Colworth
' Aggiornamento...: 26/7/1999
' Aggiornamento...: 24/7/2001 (aggiunta etichettatura delle
'                   linee di livello).
' Aggiornamento...: 20/10/2001 (aggiunta la generazione dei
'                   punti delle linee di livello: CONREC_pK).
' Versione........: 2.0 a 32 bits.
' Sistema.........: Visual Basic 6.0 sotto Windows NT.
' Scritto da......: F. Languasco ®
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.virgilio.it/flanguasco/
'                   http://www.flanguasco.org
'==============================================================
'
'   Subroutine CONREC per il tracciamento delle curve di livello:
'
'   I valori della superficie da tracciare devono essere passati
'   nella matrice D(1 to IUB, 1 to JUB): questa rappresenta una
'   griglia con colonne da 1 a IUB e righe da 1 a JUB.
'
'   Le coordinate reali di colonne e righe, alle quali corrispondono
'   i valori di livello della matrice D(.., ..), vengono passate nei
'   vettori X(1 to IUB), Y(1 to JUB): queste non devono necessariamente
'   essere equispaziate e normalmente, corrispondono alla scalatura
'   del foglio (Page) su cui si vuole disegnare.
'
'   Gli NL valori a cui si intendono tracciare le linee di livello
'   vengono passati nel vettore ZLTipo(1 to NL).dLinVal: questi valori
'   non devono necessariamente essere equispaziati ma devono essere
'   in ordine crescente.  Nei vettori ZLTipo(1 to NL).lLinCol e
'   ZLTipo(1 to NL).lLinSps vengono specificati i colori (in formato RGB)
'   e gli spessori con cui tracciare le corrispondenti curve di livello.
'
'   Nei vettori ZLTipo(1 to NL).sLblTes, ZLTipo(1 to NL).lLblCol e
'   ZLTipo(1 to NL).lLblFSz vengono passati i testi, i colori e le
'   dimensioni delle etichette che si vogliono scrivere in corrispondenza
'   di ciascuna line di livello.
'
'   Page e' il Picture Box sul quale si vuole disegnare: puo' essere,
'   anziche' un PictureBox, qualsiasi oggetto che supporti il metodo
'   Line.
'
'   Msg$ viene ritornata vuota se CONREC non ha trovato errori,
'   altrimenti contiene la descrizione dell' errore.
'
'   La variabile fStop puo' essere impostata a True, nel programma
'   chiamante, per interrompere il tracciamento delle curve di livello.
'
'   Nota:   La presente versione di CONREC (CONREC_pK) e' stata
'           modificata, rispetto all' originale, con precedenza
'           all' esplorazione per livello rispetto all' esplorazione
'           per casella.  Questo mi ha permesso di costruire, quando
'           necessario, un vettore ordinato contenente la successione
'           delle coordinate dei punti rappresentanti una (o piu')
'           curve di un certo livello.  A solo titolo di riferimento
'           e' riportata la routine Ordina che esegue quanto sopra.
'
'   Nota:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1.
'
Option Explicit
'
' Impostazioni modificabili:
Private Const lblFNm$ = "Courier New"   ' Font delle etichette.
Private Const lblFTr As Boolean = True  ' Fondo delle etichette trasparente.
Private Const lblFBl As Boolean = True  ' Testo delle etichette "Bold".
'
' Struttura delle informazioni relative
' alle linee di livello da disegnare (CONREC_pK):
Public Type LineaLivello_Type
    dLinVal As Double      ' Valore   della linea di livello.
    lLinCol As Long        ' Colore     "     "    "    " (default = vbBlack).
    lLinSps As Long        ' Spessore   "     "    "    " (default = 1 [Pixel]).
    sLblTes As String      ' Testo    dell' etichetta.
    lLblCol As Long        ' Colore     "      "  (default = vbBlack).
    lLblFSz As Long        ' Font size  "      "  (default = 8 [Points]).
End Type
'
' Struttura di appoggio per scrivere
' le etichette dei livelli (CONREC_pK):
Private Type zLbl_Type
    LblW As Double  ' Larghezza dell' etichetta.
    LblH As Double  ' Altezza   dell' etichetta.
    x1 As Single    ' Posizioni degli angoli
    x2 As Single    ' dell' area impegnata
    y1 As Single    ' dall' etichetta.
    y2 As Single    '
End Type
'
'----------------------------------------------------------------------------------
'   Dichiarazioni per Ordina:
Private Type SegOrg_Type
    NS As Boolean   ' Segmento ancora da selezionare.
    x1 As Long      ' Coordinate dei segmenti: devono
    y1 As Long      ' essere trasformate da UserScale
    x2 As Long      ' (usata da CONREC_pK) a Pixels.
    y2 As Long      '
End Type
Dim PO() As SegOrg_Type ' Vettore dei segmenti disordinati
                        ' come disegnati, per un certo
                        ' livello, da CONREC_pK.
'
Private Type SegOrd_Type
    NC As Long      ' N° della curva di appartenenza.
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type
Dim PT() As SegOrd_Type ' Vettore dei segmenti ordinati in
                        ' successione e per curva.

Private Sub Ordina()
'
'
    Dim I&, I1&, NPi&, NPu&, NCu&, NPO&
    Dim SegDisp As Boolean  ' Ci sono ancora segmenti disponibili.
    Dim SegForL As Boolean  ' Sono stati utilizzati dei segmenti nel ciclo For.
'
    ' Sceglie i segmenti consecutivi:
    I1 = 0
    NCu = 0
    NPu = 0
    NPO = UBound(PO)
    Do
        ' Cerca il primo segmento ancora disponibile in PO():
        I1 = PrimoDisponibile(PO(), I1 + 1, NPO)
'
        NCu = NCu + 1   ' N° della curva corrente.
        NPu = NPu + 1   ' Puntatore in PT() all' ultimo segmento nella curva corrente.
        NPi = NPu       ' Puntatore in PT() al primo segmento nella curva corrente.
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
                    ' Cerca i segmenti da aggiungere in coda alla curva corrente:
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
                    ' Cerca i segmenti da aggiungere in testa alla curva corrente:
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
                        SegDisp = True   ' C' e' ancora almeno un segmento disponibile.
                    End If
                End If
            Next I
        Loop While (SegDisp And SegForL) ' Ci sono ancora segmenti disponibili
                                         ' ed il ciclo For non e' andato a vuoto?
'
    Loop While SegDisp  ' Il ciclo For e' andato a vuoto ma
                        ' ci sono ancora segmenti disponibili?
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
'   Crea uno spazio libero nel vettore Seg(1 To Iu)
'   alla posizione Ii:
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
'   Ritorna la posizione del primo segmento di Seg(),
'   con indice >= I1, avente il flag NS = True:
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
    ' Definizioni per le etichette dei livelli:
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
    ' Calcola la lunghezza in [Caratteri] delle etichette
    ' e ne richiede, eventualmente, la visualizzazione:
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
    If fzLbl Then   ' Si vogliono scrivere i valori dei livelli.
        ' Salva le impostazioni correnti,
        ' assegna quelle nuove e calcola
        ' i parametri di posizione:
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
            ' Calcola larghezza ed altezza dell' etichetta
            ' in [ScaleUnits] ed usando il suo FontSize:
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
                    ' Gestione di un eventuale Flag di fine ciclo:
                    DoEvents
                    If fStop Then GoTo CONREC_END
'
DRAWIT:
                    Page.Line (x1, y1)-(x2, y2), ZLTipo(K).lLinCol
                    If ZLTipo(K).lLinSps = 1 Then Page.PSet (x2, y2), ZLTipo(K).lLinCol
'
                    ' Etichette delle linee di livello:
                    If zLbl(K).LblW > 0 Then    ' Prova a scrivere il
                                                ' valore della linea
                                                ' di livello K alla
                                                ' posizione x2, y2.
                        ' Scarta le posizioni fuori quadro:
                        If (X(IUB) < x2 + zLbl(K).LblW) _
                        Or (y2 - zLbl(K).LblH < Y(1)) Then GoTo CASE0
                        ' e quelle gia' impegnate:
                        For K1 = 1 To NL
                            If (K1 <> K) _
                            And (zLbl(K1).x1 < x2 + zLbl(K).LblW) _
                            And (x2 < zLbl(K1).x2) _
                            And (zLbl(K1).y1 < y2) _
                            And (y2 - zLbl(K).LblH < zLbl(K1).y2) Then GoTo CASE0
                        Next K1
'
                        ' La posizione x2, y2 e' OK:
                        zLbl(K).x1 = x2                 ' Area
                        zLbl(K).x2 = x2 + zLbl(K).LblW  ' impegnata
                        zLbl(K).y1 = y2 - zLbl(K).LblH  ' dalla
                        zLbl(K).y2 = y2                 ' etichetta K.
                        zLbl(K).LblW = 0                ' Livello K etichettato.
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
        ' Ripristina le impostazioni:
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
