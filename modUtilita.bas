Attribute VB_Name = "modUtilita"
'================================================================
' Descrizione.....: Collezione di routines e costanti di utilita'
'                   per il "Surface Fitting".
' Nome dei Files..: modUtilita.bas
' Data............: 21/9/2001
' Versione........: 1.0 a 32 bits.
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.it/flanguasco/
'                   http://www.flanguasco.org
'================================================================
'
Option Explicit
'
Public Const PI# = 3.14159265358979    ' 4# * Atn(1#)
Public Const PI2# = 2# * PI
Public Const PI_2# = PI / 2#           ' 90� in [Rad].
'
'--- GetLocale: ----------------------------------------------------------------
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, _
    ByVal cchData As Long) As Long
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
'
Private Const LOCALE_SDECIMAL& = &HE
Private Const LOCALE_STHOUSAND& = &HF
Private Const LOCALE_SDATE& = &H1D
Private Const LOCALE_STIME& = &H1E

Public Sub UpdateRecentFiles(ByVal mnuRecent As Object, ByVal FileRec$ _
    , Optional ByVal MAXFIL& = 5)
'
'   Add the FileRec $ file to the mnuRecent () menu.
'   mnuRecent is a menu item, with an index:
'   mnuRecent (0) is the title, the files to be memorized start from
'   the mnuRecent position (1).
'   If FileRec $ is already present it moves it to the first place.
'   MAXFIL is the maximum number of recent files to remember:
'   it can be different depending on the application using this routine (0 <MAXFIL <10).
'
    Dim I&, F&, FILN&
    Dim A$, NewFile As Boolean
'
    If FileRec$ = "" Then Exit Sub
'
    FILN = mnuRecent.UBound
'
    NewFile = True
    For F = 1 To FILN
        A$ = mnuRecent(F).Caption
        If UCase$(Right$(A$, Len(A$) - 3)) = UCase$(FileRec$) Then
            NewFile = False
            Exit For
        End If
    Next F
    If F > FILN Then F = FILN
'
    If (FILN < MAXFIL) And NewFile Then
        FILN = FILN + 1
        F = FILN
        Load mnuRecent(FILN)
        mnuRecent(FILN).Visible = True
    End If
'
    mnuRecent(0).Visible = (FILN > 0)
'
    For I = F To 2 Step -1
        A$ = mnuRecent(I - 1).Caption
        A$ = "&" & Trim$(Str$(I)) & " " & Right$(A$, Len(A$) - 3)
        mnuRecent(I).Caption = A$
    Next I
'
    mnuRecent(1).Caption = "&1 " & FileRec$
'
'
'
End Sub

Public Sub LoadRecentFiles(ByVal mnuRecent As Object, ByVal Section$ _
    , Optional ByVal MAXFIL& = 5, Optional ByVal VerificationFE As Boolean = True)
'
'   It reads from the Windows Registry and if the "Product Name" of the project
'   is defined, the recent files are added to the menu mnuRecent ().
'   If VerificationFE = True also verifies the existence of the files to be added:
'
    Dim I&, FileRec$
'
    If App.ProductName = "" Then Exit Sub
'
    For I = MAXFIL To 1 Step -1
        FileRec$ = GetSetting(App.ProductName, Section$, Str$(I), "")
        If FileRec$ <> "" Then
            If VerificationFE Then
                If FileExists(FileRec$) Then _
                    UpdateRecentFiles mnuRecent, FileRec$, MAXFIL
            Else
                UpdateRecentFiles mnuRecent, FileRec$, MAXFIL
            End If
        End If
    Next I
'
'
'
End Sub

Public Sub SalvaFilesRecenti(ByVal mnuRecent As Object, ByVal Section$)
'
'   Salva nel Registry di Windows e se e' definito il "Product Name"
'   del progetto, i files recenti contenuti nel menu mnuRecent():
'
    Dim I&, FileRec$
'
    If App.ProductName = "" Then Exit Sub
'
    For I = 1 To mnuRecent.UBound
        FileRec$ = mnuRecent(I).Caption
        FileRec$ = Right$(FileRec$, Len(FileRec$) - 3)
        SaveSetting App.ProductName, Section$, Str$(I), FileRec$
    Next I
'
'
'
End Sub
Public Sub SalvaPosizioneForm(ByVal frmF As Form _
    , Optional ByVal Dimensioni As Boolean = False)
'
'   Salva, sul Registry di Window e se e' definito il
'   "Product Name" del progetto, la posizione finale
'   del Form frmF.  Da usare nell' evento frmF_Unload.
'   Se Dimensioni = True salva anche le dimensioni:
'
    If (frmF.WindowState <> vbMinimized) And (App.ProductName <> "") Then
        SaveSetting App.ProductName, "FormsPositions", frmF.Name & "_Left", frmF.Left
        SaveSetting App.ProductName, "FormsPositions", frmF.Name & "_Top", frmF.Top
'
        If Dimensioni Then
            SaveSetting App.ProductName, "FormsPositions", frmF.Name & "_Width", frmF.Width
            SaveSetting App.ProductName, "FormsPositions", frmF.Name & "_Height", frmF.Height
        End If
    End If
'
'
'
End Sub

Public Sub LoadPositionForm(ByVal frmF As Form, ByRef frmF_Left&, ByRef frmF_Top& _
    , Optional ByRef frmF_Width&, Optional ByRef frmF_Height&)
'
'   It reads, if the "Product Name" of the project is defined, the initial
'   position and the dimensions of the Form frmF stored on the Window Registry.
'   To be used in the frmF_Load event and with frmF.StartUpPosition = vbStartUpManual:
'
    If App.ProductName <> "" Then
        frmF_Left = Val(GetSetting(App.ProductName, "FormsPositions", frmF.Name & "_Left", 0))
        frmF_Top = Val(GetSetting(App.ProductName, "FormsPositions", frmF.Name & "_Top", 0))
'
        frmF_Width = Val(GetSetting(App.ProductName, "FormsPositions", frmF.Name & "_Width", Screen.Width))
        frmF_Height = Val(GetSetting(App.ProductName, "FormsPositions", frmF.Name & "_Height", Screen.Height))
    Else
        frmF_Left = 0
        frmF_Top = 0
'
        frmF_Width = frmF.Width
        frmF_Height = frmF.Height
    End If
'
'
'
End Sub
Public Function DATAN2(ByVal Y#, ByVal X#) As Double
'
'   Ritorna il Valore dell' ArcoTangente di y/x
'   come implementata dal FORTRAN.
'   E':    -PI < DATAN2 <= PI.
'
    Select Case X
        Case Is > 0#
        DATAN2 = Atn(Y / X)
'
        Case Is < 0#
        If Y >= 0# Then
            DATAN2 = Atn(Y / X) + PI
        ElseIf Y < 0# Then
            DATAN2 = Atn(Y / X) - PI
        End If
'
        Case Is = 0#
        DATAN2 = Sgn(Y) * PI / 2#
    End Select
'
'
'
End Function
Public Function KAscNumInteri(ByVal KA%, _
    Optional ByVal NEG As Boolean = False) As Integer
'
'   Ritorna solo i caratteri validi per un campo
'   contenente un numero intero (anche negativo se
'   viene passato il valore Neg = True).
'   Tutti gli altri caratteri vengono annullati.
'   Va' usato nella procedura KeyPress di TextB come:
'   KeyAscii = KAscNumInteri(KeyAscii [, Neg])
'
    Dim KeyMinus%
    Dim TextB As TextBox    ' Solo per TextBoxes.
    'Dim TextB As Control    ' Anche per Combo, etc...
'
    Const myKeyMinus% = 45  ' E' il valore effettivamente ritornato
                            ' alla pressione del tasto - ;
                            ' vbKeySubtract (= 109) non funziona.
'
    Set TextB = Screen.ActiveControl
'
    ' Filtro per il segno "-":
    If (Left$(TextB.Text, 1) <> "-" Or TextB.SelText = TextB.Text) _
    And NEG And TextB.SelStart = 0 Then KeyMinus = myKeyMinus
'
    Select Case KA
        Case vbKey0 To vbKey9, KeyMinus, vbKeyBack
        KAscNumInteri = KA
'
        Case Else
        KAscNumInteri = 0
    End Select
'
'
'
End Function
Public Function KAscNumReali(ByVal KA As Integer _
    , Optional ByVal NEG As Boolean = False) As Integer
'
'   Ritorna solo i caratteri validi per un campo
'   contenente un numero reale (anche negativo se
'   viene passato il valore Neg = True).
'   Tutti gli altri caratteri vengono annullati.
'   Usa le impostazioni locali del separatore decimale.
'   Va' usato nella procedura KeyPress di TextB come:
'   KeyAscii = KAscNumReali(KeyAscii [, Neg])
'   Questa versione accetta anche numeri in notazione
'   scientifica.
'
    Dim KeyDecimal%, KeyMinus%, KeyE%
    Dim TextB As TextBox    ' Solo per TextBoxes.
    'Dim TextB As Control    ' Anche per Combo, etc...
    Dim SD$, SM$, myKeyDecimal%
'
    Const myKeyMinus% = 45      ' E' il valore effettivamente ritornato
                                ' alla pressione del -;
                                ' vbKeySubtract (= 109)  non funziona.
    GetLocale SD$, SM$          ' Trova le impostazioni locali dei
    myKeyDecimal% = Asc(SD$)    ' separatori decimale e delle migliaia.
'
    Set TextB = Screen.ActiveControl
'
    ' Filtro per il separatore decimale:
    If (InStr(TextB.Text, SD$) = 0 _
    And Not (TextB.SelStart = 0 And Left$(TextB.Text, 1) = "-")) _
    Or TextB.SelText = TextB.Text Then KeyDecimal = myKeyDecimal
'
    ' Filtro per il segno "-":
    If (Left$(TextB.Text, 1) <> "-" Or TextB.SelText = TextB.Text) _
    And NEG And TextB.SelStart = 0 Then KeyMinus = myKeyMinus
'
    ' Filtro per la notazione scientifica:
    If TextB.SelStart > 0 Then
        KA = Asc(UCase$(Chr$(KA)))
        If (InStr(TextB.Text, "E") = 0 _
        And Not (TextB.SelStart = 0 Or TextB.SelText = TextB.Text)) _
        And Mid$(TextB.Text, TextB.SelStart, 1) <> "-" Then KeyE = vbKeyE
'
        If Mid$(TextB.Text, TextB.SelStart, 1) = "E" Then KeyMinus = myKeyMinus
'
        If (InStr(TextB.Text, "E") > 0) _
        And (TextB.SelStart - InStr(TextB.Text, "E") >= 0) Then KeyDecimal = 0
    End If
'
    Select Case KA
        Case vbKey0 To vbKey9, KeyDecimal, KeyMinus, vbKeyBack, KeyE
        KAscNumReali = KA
'
        Case Else
        KAscNumReali = 0
    End Select
'
'
'
End Function
Private Sub GetLocale(Optional ByRef DS$, Optional ByRef MS$, _
    Optional ByRef GS$, Optional ByRef TS$)
'
'   Trova i separatori del sistema:
'    DS$:   separatore decimale.
'    MS$:   separatore delle migliaia.
'    GS$:   separatore dei giorni.
'    TS$:   separatore delle ore.
'
    DS$ = "  "
    MS$ = "  "
    GS$ = "  "
    TS$ = "  "
'
    GetLocaleInfo GetThreadLocale(), LOCALE_SDECIMAL, DS$, Len(DS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_STHOUSAND, MS$, Len(MS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_SDATE, GS$, Len(GS$)
    GetLocaleInfo GetThreadLocale(), LOCALE_STIME, TS$, Len(TS$)
'
    DS$ = Left$(DS$, 1)
    MS$ = Left$(MS$, 1)
    GS$ = Left$(GS$, 1)
    TS$ = Left$(TS$, 1)
'
'
'
End Sub

Public Function IsLoaded(ByVal frmF As Form) As Boolean
'
'   Ritorna True se il Form frmF e' gia' caricato in memoria:
'
    Dim I&
'
    For I = 0 To Forms.Count - 1
        If Forms(I) Is frmF Then
            IsLoaded = True
            Exit Function
        End If
    Next I
'
    IsLoaded = False
'
'
'
End Function

Public Function ColorTable(ByVal NCol&) As Long()
'
'   Returns a vector containing NCol (2, 16, 256, 1280 or 1792) colors in RGB
'   format. The color scale, for NCOL = 1280, goes from Purple to Red; for
'   NCOL = 1792, go from White to Black:
'
    Dim C&, C1&, C2&
    Dim R&, G&, B&
    ReDim TCol(0 To NCol - 1) As Long ' Table of color.
'
    Select Case NCol
        Case 2
        ' Prepare the 2-color table:
        TCol(0) = vbWhite
        TCol(1) = &H808080 ' Grey.
'
        Case 16
        ' Prepare the table in 16 colors:
        For C = 0 To 15
            R = 255 * (Sqr(C) / Sqr(15))
            If C < 8 Then
                G = CLng(255 * (Sqr(C) / Sqr(7)))
            Else
                G = CLng(255 * (Sqr(15 - C) / Sqr(7)))
            End If
            B = 255 * (1! - Sqr(C) / Sqr(15))
'
            TCol(C) = RGB(R, G, B)
        Next C
'
        Case 256
        ' Prepare the 256 color table:
        For C = 0 To 255
            R = 255 * (Sqr(C) / Sqr(255))
            If C < 128 Then
                G = CLng(255 * (Sqr(C) / Sqr(127)))
            Else
                G = CLng(255 * (Sqr(255 - C) / Sqr(127)))
            End If
            B = 255 * (1! - Sqr(C) / Sqr(255))
'
            TCol(C) = RGB(R, G, B)
        Next C
'
        Case 1280
        ' Prepare the table in 1280 colors:
        C = 0
        For C1 = 0 To 4
            For C2 = 0 To 255
                R = Switch(C1 = 0, 255 - C2 _
                         , C1 = 1, 0 _
                         , C1 = 2, 0 _
                         , C1 = 3, C2 _
                         , C1 = 4, 255)
                G = Switch(C1 = 0, 0 _
                         , C1 = 1, C2 _
                         , C1 = 2, 255 _
                         , C1 = 3, 255 _
                         , C1 = 4, 255 - C2)
                B = Switch(C1 = 0, 255 _
                         , C1 = 1, 255 _
                         , C1 = 2, 255 - C2 _
                         , C1 = 3, 0 _
                         , C1 = 4, 0)
'
                TCol(C) = RGB(R, G, B)
                C = C + 1
            Next C2
        Next C1
'
        Case 1792
        ' Prepare the table at 1792 colors:
        For C1 = 0 To 6
            For C2 = 0 To 255
                R = Switch(C1 = 0, 255 _
                         , C1 = 1, 255 - C2 _
                         , C1 = 2, 0 _
                         , C1 = 3, 0 _
                         , C1 = 4, C2 _
                         , C1 = 5, 255 _
                         , C1 = 6, 255 - C2)
                G = Switch(C1 = 0, 255 - C2 _
                         , C1 = 1, 0 _
                         , C1 = 2, C2 _
                         , C1 = 3, 255 _
                         , C1 = 4, 255 _
                         , C1 = 5, 255 - C2 _
                         , C1 = 6, 0)
                B = Switch(C1 = 0, 255 _
                         , C1 = 1, 255 _
                         , C1 = 2, 255 _
                         , C1 = 3, 255 - C2 _
                         , C1 = 4, 0 _
                         , C1 = 5, 0 _
                         , C1 = 6, 0)
'
                TCol(C) = RGB(R, G, B)
                C = C + 1
            Next C2
        Next C1
    End Select
'
    ColorTable = TCol()
'
'
'
End Function

Public Function RandU(ByVal V_Min!, ByVal V_Max!) As Single
'
'   Returns a real random variable with uniform distribution:
'
    RandU = V_Min + (V_Max - V_Min) * Rnd
'
'
'
End Function

Public Function Decima(dV#(), Optional ByRef lNV&) As Double()
'
'   Return a vector with the elements of dV () but without duplicates.
'   lNV is the number of elements left:
'
    Dim I&, J&, Il&, Iu&, K&, dVT#()
'
    Il = LBound(dV)
    Iu = UBound(dV)
    I = Il - 1
    dVT() = dV()
'
    Do
        I = I + 1
        J = I + 1
        Do While J <= Iu
            If dVT(I) = dVT(J) Then
                For K = J To Iu - 1
                    dVT(K) = dVT(K + 1)
                Next K
                Iu = Iu - 1
            Else
                J = J + 1
            End If
        Loop
    Loop While I < Iu
'
    lNV = Iu - Il + 1
    ReDim Preserve dVT(Il To Iu)
'
    Decima = dVT()
'
'
'
End Function

Public Sub QuickSort(ByRef ValTab() As Double, ByVal Low&, ByVal High&, _
    Optional ByVal OrderDir& = -1)
'
'   Routine QuickSort:
'    ValTab():  Vector that you want to order.
'    Low:       Initial position of the area to be ordered.
'    High:      Final position of the area to be ordered.
'    OrderDir:  Direction of the order:
'                > 0 -> from the minor to the major.
'                = 0 -> no sorting.
'                < 0 -> from major to minor.
'
    Dim RandIndex&, I&, J&, M$
    Dim ValTemp As Double   ' Type of the carrier that you want to order.
    Dim Part As Double      ' Type of sorting key.
'
    On Error GoTo QuickSort_ERR
    If OrderDir = 0 Then Exit Sub
'
    If Low < High Then
'
        If High - Low = 1 Then
            ' Only two elements in this subdivision; swap them
            ' if they are out of order, then end recursive calls:
            If ((OrderDir > 0) And (ValTab(Low) > ValTab(High))) _
            Or ((OrderDir < 0) And (ValTab(Low) < ValTab(High))) Then
                'SWAP ValTab(Low), ValTab(High)
                ValTemp = ValTab(Low)
                ValTab(Low) = ValTab(High)
                ValTab(High) = ValTemp
            End If
'
        Else
            ' Pick a pivot element, then move it to the end:
            RandIndex = (High + Low) / 2
            'SWAP ValTab(High), ValTab(RandIndex)
            ValTemp = ValTab(High)
            ValTab(High) = ValTab(RandIndex)
            ValTab(RandIndex) = ValTemp
            Part = ValTab(High)
'
            ' Move in from both sides towards the pivot element:
            Do
                I = Low: J = High
                Do While ((OrderDir > 0) And (I < J) And (ValTab(I) <= Part)) _
                Or ((OrderDir < 0) And (I < J) And (ValTab(I) >= Part))
                    I = I + 1
                Loop
                Do While ((OrderDir > 0) And (J > I) And (ValTab(J) >= Part)) _
                Or ((OrderDir < 0) And (J > I) And (ValTab(J) <= Part))
                    J = J - 1
                Loop
'
                If I < J Then
                    ' We haven't reached the pivot element; it means that two
                    ' elements on either side are out of order, so swap them:
                    'SWAP ValTab(I), ValTab(J)
                    ValTemp = ValTab(I)
                    ValTab(I) = ValTab(J)
                    ValTab(J) = ValTemp
                End If
'
            Loop While I < J
            ' Move the pivot element back to its proper place in the array:
            'SWAP ValTab(I), ValTab(High)
            ValTemp = ValTab(I)
            ValTab(I) = ValTab(High)
            ValTab(High) = ValTemp
'
            ' Recursively call the QuickSort procedure (pass the smaller
            ' subdivision first to use less stack space):
            If (I - Low) < (High - I) Then
                QuickSort ValTab(), Low, I - 1, OrderDir
                QuickSort ValTab(), I + 1, High, OrderDir
            Else
                QuickSort ValTab(), I + 1, High, OrderDir
                QuickSort ValTab(), Low, I - 1, OrderDir
            End If
        End If
    End If
'
'
QuickSort_ERR:
    If (Err <> 0) Then
        M$ = "Error " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " QuickSort"
    End If
'
'
'
End Sub

Public Function CMDialog_Files(ByVal CMDialog As CommonDialog, ByVal Oper$, _
    ByVal Tipo$, ByVal Ext$, Optional ByVal DirNome$ = "", _
    Optional ByVal FileNome$ = "", Optional ByVal Title$ = "") As String
'
'   Imposta i valori di una finestra per la gestione dei Files
'   e ritorna il nome completo del File scelto.
'   La gestione degli errori (e.g. cdlCancel) va' fatta nella
'   routine chiamante.
'
'   Parametri:
'    CMDialog:  e' il controllo Common Dialog che si intende usare.
'    Oper$:     operazione da eseguire (solo "Save" o "Open").
'    Tipo$:     tipo dei files da proporre (e.g. "Filtri IIR").
'    Ext$:      estensioni dei files da proporre (e.g. "*.dat[;*.txt]").
'    DirNome$:  Folder di default.
'    FileNome$: nome del File di default.
'    Title$:   titolo della finestra.
'
    Dim Filter$
'
    If Oper$ <> "Open" And Oper$ <> "Save" Then Err.Raise 5
'
    ' Prepara ed imposta tipo ed estensioni dei
    ' files da proporre:
    Filter$ = Tipo$ & " (" & Ext$ & ")|" & Ext$ & "|"
    Filter$ = Filter$ & "Tutti i Files (*.*)|*.*"
    CMDialog.Filter = Filter$
    CMDialog.FilterIndex = 1
'
    ' Imposta il Folder di default:
    If DirNome$ <> "" Then
        CMDialog.InitDir = DirNome$
    Else
        CMDialog.InitDir = App.Path
    End If
'
    ' Imposta il File Name di default:
    If FileNome$ <> "" Then
        CMDialog.FileName = FileNome$
    Else
        If Oper$ = "Open" Then
            CMDialog.FileName = ""
        ElseIf Oper$ = "Save" Then
            CMDialog.FileName = Left$(Ext$, 5)
        End If
    End If
'
    ' Scrive il titolo della finestra:
    CMDialog.DialogTitle = " " & Title$
'
    ' Controlla l' esistenza del File, chiede conferma
    ' se File Already Exists, nasconde la casella Read Only:
    CMDialog.Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt _
                   + cdlOFNHideReadOnly
    ' e permette il Cancel:
    CMDialog.CancelError = True
'
    ' Apre la finestra con
    ' l' operazione richiesta:
    If Oper$ = "Open" Then
        CMDialog.ShowOpen
    ElseIf Oper$ = "Save" Then
        CMDialog.ShowSave
    End If
'
    CMDialog_Files = CMDialog.FileName
'
'
'
End Function
Public Function DMIN1(ParamArray vD() As Variant) As Double
'
'   Implementa la funzione DMIN1(D1, D2, ...DN) del FORTRAN:
'
    Dim J&, J1&, J2&, vDMin As Variant
'
    J1 = LBound(vD)
    J2 = UBound(vD)
    vDMin = vD(J1)
    For J = J1 + 1 To J2
        If vDMin > vD(J) Then vDMin = vD(J)
    Next J
'
    DMIN1 = CDbl(vDMin)
'
'
'
End Function
Public Function DMAX1(ParamArray vD() As Variant) As Double
'
'   Implementa la funzione DMAX1(D1, D2, ...DN) del FORTRAN:
'
    Dim J&, J1&, J2&, vDMax As Variant
'
    J1 = LBound(vD)
    J2 = UBound(vD)
    vDMax = vD(J1)
    For J = J1 + 1 To J2
        If vDMax < vD(J) Then vDMax = vD(J)
    Next J
'
    DMAX1 = CDbl(vDMax)
'
'
'
End Function
Public Function Quadro(ByVal Foglio As PictureBox, _
    ByVal X0!, ByVal Xn!, ByVal Y0!, ByVal Yn!, _
    Optional ByVal FormatVX$ = "#0.0##", _
    Optional ByVal FormatVY$ = "#0.0##", _
    Optional ByVal Npx& = 1, Optional PxN_X!, Optional PxN_Y!, _
    Optional ByVal Title$ = "", _
    Optional ByVal UnitaX$ = "", _
    Optional ByVal UnitaY$ = "", _
    Optional ByVal AutoRed As Boolean = False) As Boolean
'
'   Routine, di uso generale, per la scalatura di un foglio
'   adatto a rappresentare un grafico y = f(x).
'    Foglio:    PictureBox da scalare.
'    X0:        Valore minimo di ascissa da rappresentare.
'    Xn:        Valore massimo di ascissa da rappresentare.
'               Deve essere X0 < Xn.
'    Y0:        Valore minimo di ordinata da rappresentare.
'    Yn:        Valore massimo di ordinata da rappresentare.
'               Deve essere Y0 <= Yn.
'    FormatVX$: Stringa di formato dei valori sull' asse X.
'    FormatVY$: Stringa di formato dei valori sull' asse Y.
'    Npx:       N� di Pixels di cui si vuole conoscere
'    PxN_X:      larghezza in [vbUser] e
'    PxN_Y:      altezza in [vbUser].
'    Title$:   Titolo del grafico.
'    UnitaX$:   Unita' (o titolo) dell' asse X.
'    UnitaY$:   Unita' (o titolo) dell' asse Y.
'    AutoRed:   Stato di Foglio.AutoRedraw dopo il disegno del quadro.
'
    Dim I&, XI!, D_X!, rrx!, YI!, D_Y!, rry!, Tx$
    Dim QxMin!, QxMax!, QyMin!, QyMax!, B0!, Bn!, TxW!
    Dim TitL!, TitT!, TitW!, TitH!, Po4_X!, Po4_Y!
    Const Log10! = 2.30258509299405 ' Log(10#)
    Const DYMin! = 0.0001           ' Ampiezza min. della scala Y.
'
    On Error GoTo Quadro_ERR
    ' Verifica la correttezza delle scale:
    If X0 >= Xn Then Err.Raise 1001, "Quadro", "Error scale X."
    If Y0 > Yn Then Err.Raise 1001, "Quadro", "Error scale Y."
'
    ' Imposta i dati di Font dei valori
    ' degli assi:
    Foglio.FontName = "MS Sans Serif"
    Foglio.FontSize = 8
    Foglio.FontBold = False
'
    ' Calcola la spaziatura dei valori scritti
    ' sull' asse X: la sequenza e' 1, 2, 2.5 e 5:
    D_X = Xn - X0
    rrx = 10! ^ Ceil(Log(D_X / 20!) / Log10)
    Do While D_X / rrx < 5!
        rrx = rrx / 2!
    Loop
    If D_X / rrx > 10! Then rrx = rrx * 2!
    X0 = rrx * Int(Round(X0 / rrx, 3))
    Xn = rrx * Ceil(Round(Xn / rrx, 3))
    D_X = Xn - X0
'
    ' Imposta una scala minima
    ' per l' asse Y:
    If Yn - Y0 < DYMin Then
        Y0 = Y0 - DYMin / 2!
        Yn = Yn + DYMin / 2!
    End If
'
    ' Calcola la spaziatura dei valori scritti
    ' sull' asse Y: la sequenza e' 1, 2, 2.5 e 5:
    D_Y = Yn - Y0
    rry = 10! ^ Ceil(Log(D_Y / 20!) / Log10)
    Do While D_Y / rry < 5!
        rry = rry / 2!
    Loop
    If D_Y / rry > 10! Then rry = rry * 2!
    Y0 = rry * Int(Round(Y0 / rry, 3))
    Yn = rry * Ceil(Round(Yn / rry, 3))
    D_Y = Yn - Y0
'
    ' Il bordo a destra dipende dalla
    ' presenza, o meno, di un' etichetta:
    If UnitaX$ = "" Then
        Bn = D_X / 20!
    Else
        Bn = D_X / 10!
    End If
'
    ' Il bordo a sinistra deve essere sufficiente
    ' a contenere il valore Y piu' largo:
    TxW = Foglio.TextWidth(Format$(Y0, FormatVY$) & " ")
    If TxW < Foglio.TextWidth(Format$(Yn, FormatVY$) & " ") Then
        TxW = Foglio.TextWidth(Format$(Yn, FormatVY$) & " ")
    End If
    B0 = TxW * (D_X + Bn) / (Foglio.ScaleWidth - TxW)
    If B0 < D_X / 10! Then B0 = D_X / 10!
'
    ' Imposta i bordi orizzontali
    ' e verticali:
    QxMin = X0 - B0
    QxMax = Xn + Bn
    QyMin = Y0 - D_Y / 10!
    QyMax = Yn + D_Y / 7!
'
    ' Cancella il Foglio e imposta la scala:
    Foglio.Picture = LoadPicture("")
    Foglio.Scale (QxMin, QyMax)-(QxMax, QyMin)
    ' La scalatura deve essere permanente:
    Foglio.AutoRedraw = True
    ' Calcola larghezza ed altezza di Npx pixels:
    PxN_X = Abs(Foglio.ScaleX(Npx, vbPixels, vbUser))
    PxN_Y = Abs(Foglio.ScaleY(Npx, vbPixels, vbUser))
    ' Calcola larghezza ed altezza di 4 points:
    Po4_X = Foglio.ScaleX(4, vbPoints, vbUser)
    Po4_Y = Foglio.ScaleY(4, vbPoints, vbUser)
'
    Foglio.DrawMode = vbCopyPen
    Foglio.DrawWidth = 1
    Foglio.DrawStyle = vbDash
    Foglio.ForeColor = vbGreen
    ' Traccia la griglia verticale e scrive
    ' i valori dell' asse X:
    For XI = X0 To Xn + 0.1 * rrx Step rrx
        Foglio.Line (XI, Y0)-(XI, Yn), vbGreen
        Tx$ = Format$(XI, FormatVX$)
        ' Verifica che il formato scelto non
        ' induca ad errori di rappresentazione:
        If Abs(XI - Val(Tx$)) < rrx / 10 Then
            Foglio.CurrentX = XI - Foglio.TextWidth(Tx$) / 2!
            Foglio.CurrentY = Y0 - D_Y / 70!
            Foglio.Print Tx$;
        End If
    Next XI
    ' Scrive l' etichetta dell' asse X:
    If UnitaX$ <> "" Then
        ' Etichetta tutta a destra:
        'Foglio.CurrentX = QxMax - Foglio.TextWidth(UnitaX$ & " ")
        ' Etichetta in centro tra l' ultimo valore ed il bordo a destra:
        Foglio.CurrentX = (Foglio.CurrentX + QxMax - Foglio.TextWidth(UnitaX$)) / 2!
        Foglio.Print UnitaX$;
    End If
    ' Traccia l' asse Y:
    If (X0 <= 0!) And (0! <= Xn) Then
        Foglio.DrawStyle = vbSolid
        Foglio.Line (0!, Y0)-(0!, QyMax - D_Y / 30!), vbGreen
        Foglio.Line (0!, QyMax - D_Y / 30!) _
                   -(-Po4_X / 2!, Po4_Y + QyMax - D_Y / 30!), vbGreen
        Foglio.Line (0!, QyMax - D_Y / 30!) _
                   -(Po4_X / 2!, Po4_Y + QyMax - D_Y / 30!), vbGreen
    End If
'
    Foglio.DrawStyle = vbDash
    ' Traccia la griglia orizzontale e scrive
    ' i valori dell' asse Y:
    For YI = Y0 To Yn + 0.1 * rry Step rry
        Foglio.Line (X0, YI)-(Xn, YI), vbGreen
        Tx$ = Format$(YI, FormatVY$)
        Foglio.CurrentX = QxMin
        Foglio.CurrentY = YI - Foglio.TextHeight(Tx$) / 2!
        Foglio.Print Tx$;
    Next YI
    ' Scrive l' etichetta dell' asse Y:
    If UnitaY$ <> "" Then
        Foglio.CurrentX = QxMin
        Foglio.CurrentY = QyMax
        Foglio.Print UnitaY$;
    End If
    ' Traccia l' asse X:
    If (Y0 <= 0!) And (0! <= Yn) Then
        Foglio.DrawStyle = vbSolid
        Foglio.Line (X0, 0!)-(QxMax - D_X / 30!, 0!), vbGreen
        Foglio.Line (QxMax - D_X / 30!, 0!) _
                   -(QxMax - D_X / 30! - Po4_X, -Po4_Y / 2!), vbGreen
        Foglio.Line (QxMax - D_X / 30!, 0!) _
                   -(QxMax - D_X / 30! - Po4_X, Po4_Y / 2!), vbGreen
    End If
'
    ' Scrive il titolo del grafico:
    If Title$ <> "" Then
        Foglio.FontSize = 12
        Foglio.FontBold = True
        Foglio.ForeColor = vbRed
'
        TitW = Foglio.TextWidth(Title$)
        TitH = Foglio.TextHeight(Title$)
        ' Verifica che il titolo stia tutto nel Foglio:
        If TitW <= Foglio.ScaleWidth Then
            TitL = (QxMin + QxMax - TitW) / 2!
        ' e se no' lo taglia:
        Else
            TitL = Foglio.ScaleLeft
            Tx$ = " . . . ."
            Title$ = Left$(Title$, Int(Len(Title$) * _
            (Foglio.ScaleWidth - Foglio.TextWidth(Tx$)) / TitW)) & Tx$
        End If
        TitT = QyMax
        ' Cancella l' area su cui andra' scritto il titolo:
        'Foglio.Line (TitL, TitT)-(TitL + TitW, TitT + TitH), Foglio.BackColor, BF
        Foglio.CurrentX = TitL
        Foglio.CurrentY = TitT
        Foglio.Print Title$
    End If
'
    Foglio.DrawStyle = vbSolid
    Foglio.AutoRedraw = AutoRed
'
'
Quadro_ERR:
    Quadro = (Err = 0)
    If Err <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
'
'
'
End Function
Public Function Ceil(ByVal X As Double) As Double
'
'   Funzione di arrotondamento, per numeri reali,
'   all' intero uguale o immediatamente superiore:
'
    If X = Int(X) Then
        Ceil = X
    Else
        Ceil = Int(X) + 1#
    End If
'
'
'
End Function
Public Function MAX0(ParamArray vD() As Variant) As Long
'
'   Implementa la funzione MAX0(K1, K2, ...KN) del FORTRAN:
'
    Dim J&, J1&, J2&, vDMax As Variant
'
    J1 = LBound(vD)
    J2 = UBound(vD)
    vDMax = vD(J1)
    For J = J1 + 1 To J2
        If vDMax < vD(J) Then vDMax = vD(J)
    Next J
'
    MAX0 = CLng(vDMax)
'
'
'
End Function
Public Function MIN0(ParamArray vD() As Variant) As Long
'
'   Implementa la funzione MIN0(K1, K2, ...KN) del FORTRAN:
'
    Dim J&, J1&, J2&, vDMin As Variant
'
    J1 = LBound(vD)
    J2 = UBound(vD)
    vDMin = vD(J1)
    For J = J1 + 1 To J2
        If vDMin > vD(J) Then vDMin = vD(J)
    Next J
'
    MIN0 = CLng(vDMin)
'
'
'
End Function
Public Function BreakDown(ByVal Full$, Optional ByRef PName$, _
    Optional ByRef FName$, Optional ByRef Ext$) As Boolean
'
'   Scompone un nome di File completo di Path nelle sue parti:
'    Full$  = Nome completo del File.
'    PName$ = Path del File (con \ finale).
'    FName$ = Nome del File con Estensione.
'    Ext$   = .Estensione del File.
'
'   Se il File non esiste scompone il nome e ritorna False.
'
    Dim Sloc&, Dot&
'
    BreakDown = FileExists(Full$)
'
    If InStr(Full$, "\") Then
        FName$ = Full$
        PName$ = ""
        Sloc = InStr(FName$, "\")
        Do While Sloc <> 0
            PName$ = PName$ & Left$(FName$, Sloc)
            FName$ = Mid$(FName$, Sloc + 1)
            Sloc = InStr(FName$, "\")
        Loop
    Else
        PName$ = ""
        FName$ = Full$
    End If
'
    Dot = InStr(Full$, ".")
    If Dot <> 0 Then
        Ext$ = Mid$(Full$, Dot)
    Else
        Ext$ = ""
    End If
'
'
'
End Function
Private Function FileExists(ByVal FileN$) As Boolean
'
'   Verifica l' esistenza del file FileN$:
'
    Dim res&
'
    On Error Resume Next
'
    res = FileLen(FileN$)
    FileExists = (Err = 0)
    Err.Clear
'
'
'
End Function
Public Function Atan2(ByVal Y#, ByVal X#) As Double
'
'   Ritorna il Valore dell' ArcoTangente di y/x su 4 Quadranti.
'   E':    0 <= Atan2 < 2 * PI.
'
    Select Case X
        Case Is > 0#
        If Y >= 0# Then
            Atan2 = Atn(Y / X)
        Else
            Atan2 = Atn(Y / X) + PI2
        End If
'
        Case Is < 0#
        Atan2 = Atn(Y / X) + PI
'
        Case Is = 0#
        Atan2 = Sgn(Y) * PI / 2#
    End Select
'
'
'
End Function

Public Function DMINVAL(dVet#()) As Double
'
'   Ritorna il valore minimo contenuto nel vettore dVet().
'   Implementa l' equivalente funzione del FORTRAN.
'
    Dim J&, J1&, J2&, dDMin#
'
    J1 = LBound(dVet)
    J2 = UBound(dVet)
    dDMin = dVet(J1)
    For J = J1 + 1 To J2
        If dDMin > dVet(J) Then dDMin = dVet(J)
    Next J
'
    DMINVAL = dDMin
'
'
'
End Function
Public Function DMAXVAL(dVet#()) As Double
'
'   Ritorna il valore massimo contenuto nel vettore dVet().
'   Implementa l' equivalente funzione del FORTRAN.
'
    Dim J&, J1&, J2&, dDMax#
'
    J1 = LBound(dVet)
    J2 = UBound(dVet)
    dDMax = dVet(J1)
    For J = J1 + 1 To J2
        If dDMax < dVet(J) Then dDMax = dVet(J)
    Next J
'
    DMAXVAL = dDMax
'
'
'
End Function
