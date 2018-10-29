VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSurFit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Surface Fitting"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9510
   Icon            =   "frmSurFit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkValoriLivelli 
      Alignment       =   1  'Right Justify
      Caption         =   "&Valori dei Livelli:"
      Height          =   255
      Left            =   1980
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame zFrame2 
      Caption         =   "Funzioni di prova:"
      Height          =   795
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdProva 
         Caption         =   "&Prova"
         Height          =   255
         Left            =   3600
         TabIndex        =   0
         Top             =   360
         Width           =   795
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&6"
         Height          =   315
         Index           =   6
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&5"
         Height          =   315
         Index           =   5
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&4"
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&3"
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&2"
         Height          =   315
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&1"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   435
      End
      Begin VB.OptionButton optZxy 
         Caption         =   "&7"
         Height          =   315
         Index           =   7
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame zFrame1 
      Caption         =   "Interpolazione"
      Height          =   795
      Left            =   4800
      TabIndex        =   21
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkGradiente 
         Caption         =   "&Gradiente"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optKTB2D 
         Caption         =   "&KTB2D"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optQSHEP2D 
         Caption         =   "Q&SHEP2D"
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optMASUB 
         Caption         =   "&MASUB"
         Height          =   255
         Left            =   1020
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdGrigliaSurFit 
      Caption         =   "Griglia"
      Height          =   255
      Left            =   8460
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdGrigliaOrg 
      Caption         =   "Griglia"
      Height          =   255
      Left            =   3780
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox picSurFit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   4800
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Click per vista 3D della Superficie "
      Top             =   1020
      Width           =   4575
   End
   Begin VB.PictureBox picOrg 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Click per vista 3D dei Punti "
      Top             =   1020
      Width           =   4575
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   3840
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label lblYMax 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7380
      TabIndex        =   30
      Top             =   5880
      Width           =   795
   End
   Begin VB.Label zLabel06 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Max:"
      Height          =   195
      Left            =   7380
      TabIndex        =   29
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblXMax 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6540
      TabIndex        =   28
      Top             =   5880
      Width           =   795
   End
   Begin VB.Label zLabel05 
      BackStyle       =   0  'Transparent
      Caption         =   "X Max:"
      Height          =   195
      Left            =   6540
      TabIndex        =   27
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblYMin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   26
      Top             =   5880
      Width           =   795
   End
   Begin VB.Label zLabel04 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Min:"
      Height          =   195
      Left            =   5640
      TabIndex        =   25
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblXMin 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   5880
      Width           =   795
   End
   Begin VB.Label zLabel03 
      BackStyle       =   0  'Transparent
      Caption         =   "X Min:"
      Height          =   195
      Left            =   4800
      TabIndex        =   23
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblNAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   900
      TabIndex        =   20
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label zLabel02 
      BackStyle       =   0  'Transparent
      Caption         =   "N Add:"
      Height          =   195
      Left            =   900
      TabIndex        =   19
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label zLabel01 
      BackStyle       =   0  'Transparent
      Caption         =   "ND:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblND 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLeggiDati 
         Caption         =   "&Leggi file dati"
      End
      Begin VB.Menu zSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalvaInterpolati 
         Caption         =   "&Salva dati interpolati"
         Enabled         =   0   'False
      End
      Begin VB.Menu zSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "&Esci"
      End
      Begin VB.Menu zSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecenti 
         Caption         =   "Files Recenti:"
         Index           =   0
      End
   End
   Begin VB.Menu mnuImpostazioni 
      Caption         =   "&Impostazioni"
   End
   Begin VB.Menu mnuQM 
      Caption         =   "&?"
      Begin VB.Menu mnuIstruzioni 
         Caption         =   "I&struzioni"
      End
      Begin VB.Menu zSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformazioni 
         Caption         =   "&Informazioni su SurFit"
      End
   End
End
Attribute VB_Name = "frmSurFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
' Descrizione.....: Form di prova per le routines di "Surface
'                   Fitting".
' Nome dei Files..: frmSurFit.frm, frmSurFit.frx
'                   frmImpostazioni.frm, frmImpostazioni.frx
'                   frm3D.frm, frm3D.frx
'                   frmIstr.frm, frmIstr.frx
'                   InfoCr.frm, InfoCr.frx
'                   modKTB2D.bas, modMASUB.bas, modQSHEP2D
'                   modGradiente2D.bas, Livelli.bas
'                   modUtilita.bas
' Data............: 21/9/2001
' Versione........: 1.0 a 32 bits.
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco ®
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.virgilio.it/flanguasco/
'                   http://www.flanguasco.org
'=============================================================
'
'   Nota:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1 (ZCol() escluso).
'
Option Explicit
'
Dim ND&             ' N° di dati nei vettori.
Dim XD#()           ' Vettori dei valori
Dim YD#()           ' dati della superficie
Dim ZD#()           ' da interpolare.
'
Dim Xs#(), Ys#()    ' Coordinate della griglia dei punti dati.
'
Dim NXI&, NYI&      ' N° di colonne e di righe nella
                    ' griglia dei punti interpolati.
Dim XI#(), YI#()    ' Coordinate della griglia dei punti interpolati.
Dim ZI#()           ' Superficie interpolata.
Dim ZC#()           ' Superficie calcolata.
Dim Grad() As Grad_Type   ' Matrice del gradiente.
'
' Impostazioni per KTB2D:
Dim Par As ParType, IER&
'
' Impostazioni per MASUB:
Dim TP#
'
' Impostazioni per QSHEP2D:
Dim NQ&, NW&, NR&
'
Dim ZCol&()         ' Tavola dei colori.
Const NTCol& = 1280 ' Numero di colori disponibili in ZCol().
Dim NLiv&           ' Numero di livelli da tracciare.
'
Dim FolderN$        ' Folder dei files dati.
'
Dim lZxy&           ' Indice della funzione di prova.
'
Dim bScriviVal As Boolean   ' Scrive i valori delle linee di livello.
Dim bDisegnaGD As Boolean   ' Disegna le freccette del gradiente.
Dim bDisegnaZC As Boolean   ' Disegna la superficie calcolata.
Dim Titolo$                 ' Titolo del quadro picOrg.
'
Const Me_W& = 9600          ' Dimensioni di questo
Const Me_H& = 7200 - 255    ' Form [Twips].

Private Sub Prova_KTB2D()
'
'   Interpolazione, con "kriging",  di una superficie
'   con punti dati nei vettori XD(), YD(), ZD():
'
    Dim A#, B#, C#, D#, HX#, HY#, Px3!, Py3!
    Dim IER&
'
    ' Prepara i vettori XI() ed YI() con le
    ' coordinate della griglia di interpolazione:
    GrigliaDiInterpolazione A, B, C, D
'
    HX = (B - A) / CDbl(NXI - 1)
    HY = (D - C) / CDbl(NYI - 1)
'
    IER = 0 ' Niente file di debug.
    Call KTB2D(ND, XD(), YD(), ZD(), Par, NXI, A, HX, NYI, C, HY, ZI(), IER)
    If IER <> 0 Then
        MsgBox "Errore " & IER & " in KTB2D", vbCritical
        Exit Sub
    End If
'
    DisegnaLivelli A, B, C, D, Px3, Py3
    If bDisegnaGD Then
        Call Gradiente_2D(XI(), YI(), ZI(), NXI, NYI, Grad())
        DisegnaGradiente Px3, Py3
    End If
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub
Private Sub Prova_QSHEP2D()
'
'   Interpolazione di una superficie con punti dati nei
'   vettori XD(), YD(), ZD() con il metodo quadratico
'   di Shepard:
'
    Dim I&, J&, A#, B#, C#, D#, Px3!, Py3!
    Dim IER&, XMin#, YMin#, DX#, DY#, RMAX#
    Dim LCELL&(), LNEXT&(), RSQ#(), ASh#()
'
    ' Prepara i vettori XI() ed YI() con le
    ' coordinate della griglia di interpolazione:
    GrigliaDiInterpolazione A, B, C, D
'
    ' Impostazione dei parametri per QSHEP2:
    NQ = MIN0(MAX0(5, NQ), MIN0(40, ND - 1))    ' 5 <= NQ <= MIN(40,ND-1)
    NW = MIN0(MAX0(1, NW), MIN0(40, ND - 1))    ' 1 <= NW <= MIN(40,ND-1)
    ReDim LCELL(1 To NR, 1 To NR), LNEXT(1 To ND)
    ReDim RSQ(1 To ND), ASh(1 To 5, 1 To ND)
'
    ' La chiamata alla routine QSHEP2 prepara valori e
    ' vettori necessari all' interpolazione successiva
    ' da fare con QS2GRD o QS2VAL:
    Call QSHEP2(ND, XD(), YD(), ZD(), NQ, NW, NR, LCELL(), LNEXT(), _
                XMin, YMin, DX, DY, RMAX, RSQ(), ASh(), IER)
    If IER <> 0 Then
        MsgBox "Errore " & IER & " in QSHEP2D", vbCritical
        Exit Sub
    End If
'
    If bDisegnaGD Then
        ' Calcola la superficie interpolata ed il gradiente:
        For J = 1 To NYI
            For I = 1 To NXI
                Call QS2GRD(XI(I), YI(J), ND, XD(), YD(), ZD(), NR, LCELL(), LNEXT(), _
                            XMin, YMin, DX, DY, RMAX, RSQ(), ASh(), _
                            ZI(I, J), Grad(I, J).DX, Grad(I, J).DY, IER)
                If IER <> 0 Then
                    MsgBox "Errore " & IER & " in QS2GRD", vbCritical
                    Exit Sub
                End If
            Next I
        Next J
'
    Else
        ' Calcola solo la superficie interpolata:
        For J = 1 To NYI
            For I = 1 To NXI
                ZI(I, J) = QS2VAL(XI(I), YI(J), ND, XD(), YD(), ZD(), NR, LCELL(), LNEXT(), _
                                  XMin, YMin, DX, DY, RMAX, RSQ(), ASh())
            Next I
        Next J
    End If
'
    DisegnaLivelli A, B, C, D, Px3, Py3
    If bDisegnaGD Then DisegnaGradiente Px3, Py3
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub
Private Function Zxy(ByVal x1#, ByVal x2#) As Double
'
'   Funzioni di prova a due parametri:
'
    Dim A#, B#
'
    Select Case lZxy
        Case 1
        ' Cerchi concentrici:
        A = Sqr(x1 ^ 2 + x2 ^ 2) - 6#
        B = 1#
'
        Case 2
        x1 = x1 / 10#
        x2 = x2 / 10#
        A = Abs(Cos((x1 - 0.1) ^ 2 + x2 ^ 2))
        B = Abs(x1 - 0.1) + Abs(x2) + 3#
        'B = Abs(x2) + 3#
'
        Case 3
        ' Funzione di prova #1 per "Genetic Algorithms":
        A = Cos(Sqr((x1 + 1#) ^ 2 + (x2 - 1#) ^ 2))
        B = Sqr((x1 - 1#) ^ 2 + (x2 + 1#) ^ 2) + 2#
'
        Case 4
        ' Funzione di prova #2 per "Genetic Algorithms":
        A = Cos(Sqr((x1) ^ 2 + x2 * x2))
        B = Sqr(0.005 * x1 ^ 2 + x2 * x2) + 2
        B = Sqr(x2 * x2) + 2#
'
        Case 5
        ' Funzione di prova #3 per "Genetic Algorithms":
        A = Cos(Sqr((x1 + 1) ^ 2 + (x2 - 1) ^ 2))
        B = Sqr((x1 + 1) ^ 2 + x2 * x2) + 2#
'
        Case 6
        ' Rosembrook:
        A = -100 * (x2 - x1 ^ 2) ^ 2 + (1 - x1) ^ 2
        B = 1#
'
        Case 7
        ' CONREC Esempio 1:
        x1 = x1 / 20#
        x2 = x2 / 20#
        A = 1#
        B = ((x1 ^ 2 + (x2 - 0.842) * (x2 + 0.842))) ^ 2 _
          + (x1 * (x2 - 0.842) + x1 * (x2 - 0.842)) ^ 2
'
        Case 8
        ' Caleido:
        A = Cos((x1 * x2) ^ 2) + 1# / (0.01 * Abs(x1 * x2) + 0.2) ^ 0.2
        B = 1#
    End Select
'
    Zxy = A / B
'
'
'
End Function
Private Sub chkGradiente_Click()
'
'
    bDisegnaGD = (chkGradiente.Value = vbChecked)
'
    Screen.MousePointer = vbHourglass
'
    ' Chiama la routine di interpolazione:
    If optKTB2D Then
        Prova_KTB2D
    ElseIf optMASUB Then
        Prova_MASUB
    ElseIf optQSHEP2D Then
        Prova_QSHEP2D
    End If
'
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub chkValoriLivelli_Click()
'
'
    bScriviVal = (chkValoriLivelli.Value = vbChecked)
'
    Screen.MousePointer = vbHourglass
'
    ' Chiama la routine di interpolazione:
    If optKTB2D Then
        Prova_KTB2D
    ElseIf optMASUB Then
        Prova_MASUB
    ElseIf optQSHEP2D Then
        Prova_QSHEP2D
    End If
'
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub cmdGrigliaOrg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    DisegnaGriglia picOrg, Xs(), Ys()
'
'
'
End Sub
Private Sub cmdGrigliaOrg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    picOrg.Cls
'
'
'
End Sub
Private Sub cmdGrigliaSurFit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    DisegnaGriglia picSurFit, XI(), YI()
'
'
'
End Sub
Private Sub cmdGrigliaSurFit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    picSurFit.Cls
'
'
'
End Sub
Private Sub cmdProva_Click()
'
'
    Dim I&, N&
'
    cmdProva.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
'
    ND = CLng(RandU(6, 200))
    ReDim XD#(1 To ND), YD#(1 To ND), ZD#(1 To ND)
    For N = 1 To ND
        ' Ascisse dei punti dati:
        XD(N) = RandU(-10, 10)
        ' Ordinate dei punti dati:
        YD(N) = RandU(-10, 10)
    Next N
'
    Call ParametriDiDefault
'
    ' Calcola il valore della superficie
    ' ai punti dati:
    For N = 1 To ND
        ZD(N) = Zxy(XD(N), YD(N))
    Next N
'
    ' Prepara una griglia
    ' corrispondente ai punti dati:
    GrigliaPuntiDati XD(), YD(), Xs(), Ys()
'
    ' Chiama la routine di interpolazione:
    bDisegnaZC = True
    Titolo$ = ND & " Punti casuali"
    If optKTB2D Then
        Prova_KTB2D
    ElseIf optMASUB Then
        Prova_MASUB
    ElseIf optQSHEP2D Then
        Prova_QSHEP2D
    End If
'
    Screen.MousePointer = vbDefault
    cmdProva.Enabled = True
'
'
'
End Sub
Private Sub Form_Load()
'
'
    Dim Me_L&, Me_T&
'
    LeggiPosizioneForm Me, Me_L, Me_T
    Me.Move Me_L, Me_T, Me_W, Me_H
'
    LeggiFilesRecenti Me.mnuRecenti, "PuntiDati"
'
    ZCol() = TabellaColori(NTCol)
'
    NXI = 50
    NYI = 50
    NLiv = 10   ' Numero di livelli da tracciare.
'
    optZxy(1).Value = True
    cmdProva_Click
'
'
'
End Sub
Private Sub Prova_MASUB()
'
'   Interpolazione di una superficie con punti dati
'   nei vettori XD(), YD(), ZD():
'
    Dim A#, B#, C#, D#, Px3!, Py3!
    Dim IC&, IEX&
'
    ' Prepara i vettori XI() ed YI() con le
    ' coordinate della griglia di interpolazione:
    GrigliaDiInterpolazione A, B, C, D, 0.1
'
    ' Impostazione parametri per MASUB:
    IC = 1      ' Prima ed unica chiamata.
    IEX = 1     ' E' richiesta l' estrapolazione.
'
    If Not MASUB(IC, IEX, ND, XD(), YD(), ZD(), TP _
               , NXI, NYI, XI(), YI(), ZI()) Then
        MsgBox "Errore in MASUB", vbCritical
        Exit Sub
    End If
    lblNAdd = UBound(XD) - ND   ' Punti aggiunti per estrapolazione.
'
    DisegnaLivelli A, B, C, D, Px3, Py3
    If bDisegnaGD Then
        Call Gradiente_2D(XI(), YI(), ZI(), NXI, NYI, Grad())
        DisegnaGradiente Px3, Py3
    End If
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub
Private Sub DisegnaGriglia(Quadro As PictureBox, dXI#(), dYI#())
'
'   Disegna una griglia alle coordinate definite
'   nei vettori dX() e dY():
'
    Dim I&, J&, XMin#, XMax#, YMin#, YMax#
'
    XMin = dXI(1)
    XMax = dXI(UBound(dXI))
    YMin = dYI(1)
    YMax = dYI(UBound(dYI))
'
    For I = 1 To UBound(dXI)
        Quadro.Line (dXI(I), YMin)-(dXI(I), YMax), vbYellow
    Next I
'
    For J = 1 To UBound(dYI)
        Quadro.Line (XMin, dYI(J))-(XMax, dYI(J)), vbYellow
    Next J
'
'
'
End Sub
Public Sub GrigliaPuntiDati(XD#(), YD#(), XGD#(), YGD#())
'
'   Prepara i vettori delle coordinate dei punti dati,
'   eliminando i valori doppi ed ordinandoli in verso
'   crescente:
'
    XGD() = XD()
    YGD() = YD()
'
    QuickSort XGD(), 1, UBound(XGD), 1
    QuickSort YGD(), 1, UBound(YGD), 1
'
    XGD() = Decima(XGD())
    YGD() = Decima(YGD())
'
'
'
End Sub
Private Sub DisegnaLivelli(ByVal A#, ByVal B#, ByVal C#, ByVal D#, _
    ByRef Px3!, ByRef Py3!)
'
'   Visualizzazione delle curve di livello:
'
    Dim I&, J&, K&, N&, ZMin#, ZMax#, ZDif#, GMax2#, Msg$
    ReDim ZLin(1 To NLiv) As LineaLivello_Type
'
    ' Imposta la grafica:
    Quadro picOrg, A, B, C, D, , , 3, Px3, Py3, "Punti dati: " & Titolo$, _
                  "x", "y", True
    Quadro picSurFit, A, B, C, D, , , , , , "Superficie interpolata", "x", "y", True
'
    ' Visualizza i punti dati:
    For N = 1 To ND
        picOrg.Circle (XD(N), YD(N)), Px3, vbRed
    Next N
'
    ' Visualizza i punti aggiunti (solo per MASUB):
    For N = ND + 1 To UBound(XD)
        picOrg.Circle (XD(N), YD(N)), Px3, vbYellow
    Next N
'
    ' Trova i Max. e Min. della superficie:
    ZMin = ZI(1, 1)
    ZMax = ZI(1, 1)
    For I = 1 To NXI
        For J = 1 To NYI
            ZMin = DMIN1(ZMin, ZI(I, J))
            ZMax = DMAX1(ZMax, ZI(I, J))
        Next J
    Next I
'
    ' Prepara il vettore dei livelli:
    For K = 1 To NLiv
        ZLin(K).dLinVal = ZMin + (K - 1) * (ZMax - ZMin) / (NLiv - 1)
        ZLin(K).lLinCol = ZCol((K - 1) * (NTCol - 1) / (NLiv - 1))
        ZLin(K).lLinSps = 1
        'ZLin(k).lLinSps = IIf(k Mod 3 = 0, 2, 1)
        If bScriviVal Then
            ZLin(K).sLblTes = Format$(ZLin(K).dLinVal, "#0.000")
        Else
            ZLin(K).sLblTes = ""
        End If
        ZLin(K).lLblCol = ZLin(K).lLinCol
        ZLin(K).lLblFSz = 10
    Next K
'
    ' Visualizza la superficie interpolata:
    CONREC_pK picSurFit, ZI(), XI(), YI(), ZLin(), Msg$
    If Msg$ <> "" Then
        MsgBox Msg$, vbCritical, " CONREC di ZI()"
    End If
'
    If bDisegnaZC Then
        ' Calcola il valore della superficie
        ' su tutti i punti della griglia:
        For I = 1 To NXI
            For J = 1 To NYI
                ZC(I, J) = Zxy(XI(I), YI(J))
            Next J
        Next I
'
        ' Visualizza la superficie calcolata:
        CONREC_pK picOrg, ZC(), XI(), YI(), ZLin(), Msg$
        If Msg$ <> "" Then
            MsgBox Msg$, vbCritical, " CONREC di ZC()"
        End If
    End If
'
    ' Disegna la superficie in 3D:
    frm3D.Superficie XI#(), YI#(), ZI#(), Titolo$
'
'
'
End Sub
Private Sub DisegnaGradiente(ByVal Px3!, ByVal Py3!)
'
'   Disegna le freccette del gradiente:
'
    Dim I&, J&, PlFQ#, PxF#, PyF#, QX#, QY#, M$
    Dim GMax2#, GradScalaX#, GradScalaY#, GAMMA#
    Const GRis& = 5, LnF! = 10!
'
    ' Trova il Max. del gradiente:
    GMax2 = 0#
    For I = 1 To NXI
        For J = 1 To NYI
            GMax2 = DMAX1(GMax2, (Grad(I, J).DY ^ 2 + Grad(I, J).DX ^ 2))
        Next J
    Next I
    If GMax2 = 0 Then
        M$ = "La superficie e' piana." & vbNewLine
        M$ = M$ & "Non e' possibile disegnare il gradiente."
        MsgBox M$, vbInformation, " Disegno del gradiente"
        Exit Sub
    End If
    
'
    ' Con questa scala il gradiente massimo
    ' corrisponde ad una freccia di 3 * LnF [Pixels].
    ' Vale solo per scala orizzontale uguale alla
    ' scala verticale:
    GradScalaX = LnF * CDbl(Px3) / Sqr(GMax2)
    GradScalaY = LnF * CDbl(Py3) / Sqr(GMax2)
'
    ' Angolo e lunghezza delle punte delle frecce:
    Const ApF# = PI_2 / 3#  ' 30°
    PxF = 1.9 * Px3
    PyF = 1.9 * Py3
    PlFQ = 1# * (PxF ^ 2 + PyF ^ 2) ' Lunghezza minima rappresentabile
                                    ' dell' asta della freccia.
'
    For I = 1 To NXI Step GRis
        For J = 1 To NYI Step GRis
            If ((Grad(I, J).DX <> 0#) Or (Grad(I, J).DY <> 0#)) Then
                ' Asta della freccia:
                QX = XI(I) + Grad(I, J).DX * GradScalaX
                QY = YI(J) + Grad(I, J).DY * GradScalaY
                If ((QX - XI(I)) ^ 2 + (QY - YI(J)) ^ 2) > PlFQ Then
                    picSurFit.Line (QX, QY)-(XI(I), YI(J)), vbWhite
                End If
'
                ' Punta della freccia:
                GAMMA = Atan2(Grad(I, J).DY, Grad(I, J).DX) + ApF
                picSurFit.Line (QX, QY) _
                              -(QX - PxF * Cos(GAMMA), QY - PyF * Sin(GAMMA)), vbWhite
                GAMMA = GAMMA - 2# * ApF
                picSurFit.Line (QX, QY) _
                              -(QX - PxF * Cos(GAMMA), QY - PyF * Sin(GAMMA)), vbWhite
'
            End If
            If J = 1 Then J = J - 1
        Next J
        If I = 1 Then I = I - 1
    Next I
'
'
'
End Sub
Private Sub Form_Unload(Cancel As Integer)
'
'
    If IsLoaded(frmIstruzioni) Then Unload frmIstruzioni
    If IsLoaded(frm3D) Then Unload frm3D
'
    SalvaFilesRecenti Me.mnuRecenti, "PuntiDati"
'
    SalvaPosizioneForm Me
'
'
'
End Sub

Private Sub mnuEsci_Click()
'
'
    Unload Me
'
'
'
End Sub

Private Sub mnuImpostazioni_Click()
'
'
    If frmImpostazioni.Apri(ND, NXI, NYI, NLiv, Par, NQ, NW, NR, TP) Then
        Screen.MousePointer = vbHourglass
    '
        ' Chiama la routine di interpolazione
        ' con i parametri modificati:
        If optKTB2D Then
            Prova_KTB2D
        ElseIf optMASUB Then
            Prova_MASUB
        ElseIf optQSHEP2D Then
            Prova_QSHEP2D
        End If
    '
        Screen.MousePointer = vbDefault
    End If
'
'
'
End Sub

Private Sub mnuInformazioni_Click()
'
'
    frmCopyright.Apri Me
'
'
End Sub
Private Sub mnuIstruzioni_Click()
'
'
    frmIstruzioni.Apri App.Path & "\Istr.rtf"
'
'
'
End Sub
Private Sub mnuLeggiDati_Click()
'
'
    Dim FN_Temp$, M$
'
    On Error GoTo mnuLeggi_Click_ERR
'
    FN_Temp$ = CMDialog_Files(CMDialog1, "Open", "Files dati", "*.dat;*.txt", _
                              FolderN$, , " Dati della superficie da interpolare")
'
'
    If BreakDown(FN_Temp$, FolderN$, Titolo$) Then
        ElaboraFileDati FN_Temp$
    End If
'
'
mnuLeggi_Click_ERR:
    If Err <> 0 And Err <> cdlCancel Then
        M$ = "Errore " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " mnuLeggi/" & Err.Source
    End If
'
'
'
End Sub

Private Sub mnuRecenti_Click(INDEX As Integer)
'
'
    Dim FN_Temp$
'
    If INDEX = 0 Then Exit Sub
'
'
    FN_Temp$ = mnuRecenti(INDEX).Caption
    FN_Temp$ = Right$(FN_Temp$, Len(FN_Temp$) - 3)
'
    If BreakDown(FN_Temp$, FolderN$, Titolo$) Then
        ElaboraFileDati FN_Temp$
    End If
'
'
'
End Sub

Private Sub mnuSalvaInterpolati_Click()
'
'
    Dim FF%, I&, J&, FileNome$, M$
'
    On Error GoTo mnuSalvaInterpolati_Click_ERR
'
    FileNome$ = CMDialog_Files(CMDialog1, "Save", "Files dati", "*.dat;*.txt", _
                               FolderN$, , " Dati interpolati")
'
    FF = FreeFile
    Open FileNome$ For Output As #FF
'
    Print #FF, NXI, NYI
    For I = 1 To NXI
        For J = 1 To NYI
            Print #FF, XI(I), YI(J), ZI(I, J)
        Next J
    Next I
    Close FF
'
'
mnuSalvaInterpolati_Click_ERR:
    If Err <> 0 And Err <> cdlCancel Then
        M$ = "Errore " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " mnuSalva/" & Err.Source
    End If
'
'
'
End Sub
Private Sub optKTB2D_Click()
'
'
    Screen.MousePointer = vbHourglass
    Prova_KTB2D
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub optMASUB_Click()
'
'
    Screen.MousePointer = vbHourglass
    Prova_MASUB
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub optQSHEP2D_Click()
'
'
    Screen.MousePointer = vbHourglass
    Prova_QSHEP2D
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub GrigliaDiInterpolazione(ByRef A#, ByRef B#, ByRef C#, ByRef D#, _
    Optional ByVal est# = 0)
'
'   Prepara i vettori contenenti ascisse ed ordinate della griglia
'   di interpolazione.  Calcola anche gli estremi delle coordinate
'   di interpolazione, eventualmente allargandole del fattore Est:
'   da usare, principalmente, per MASUB che va' facilmente in errore
'   quando gli estremi di interpolazione coincidono con gli estremi
'   dei punti dati.
'
    Dim I&, J&, HX#, HY#
'
    ' Per eliminare i punti aggiunti da
    ' un' eventuale chiamata precedente a MASUB:
    ReDim Preserve XD(1 To ND)
    ReDim Preserve YD(1 To ND)
    lblND = ND
    lblNAdd = "--"
'
    ' Trova le coordinate Max. e Min.
    ' dei punti dati:
    A = DMINVAL(XD())   ' Minimo X.
    B = DMAXVAL(XD())   ' Massimo X.
    C = DMINVAL(YD())   ' Minimo Y.
    D = DMAXVAL(YD())   ' Massimo Y.
    ' e le allarga del fattore Est:
    HX = (B - A)
    A = A - est * HX
    B = B + est * HX
    HY = (D - C)
    C = C - est * HY
    D = D + est * HY
    lblXMin = Format$(A, "#0.000")
    lblXMax = Format$(B, "#0.000")
    lblYMin = Format$(C, "#0.000")
    lblYMax = Format$(D, "#0.000")
'
    ReDim XI(1 To NXI), YI#(1 To NYI)   ' Coordinate della griglia dei punti interpolati.
    ReDim ZI(1 To NXI, 1 To NYI)        ' Superficie interpolata.
    ReDim ZC(1 To NXI, 1 To NYI)        ' Superficie calcolata.
    ReDim Grad(1 To NXI, 1 To NYI)      ' Matrice del gradiente.
'
    ' Ascisse della griglia dei punti interpolati:
    HX = (B - A) / CDbl(NXI - 1)
    For I = 1 To NXI
        XI(I) = A + (I - 1) * HX
    Next I
'
    ' Ordinate della griglia dei punti interpolati:
    HY = (D - C) / CDbl(NYI - 1)
    For J = 1 To NYI
        YI(J) = C + (J - 1) * HY
    Next J
'
'
'
End Sub
Private Sub optZxy_Click(INDEX As Integer)
'
'
    lZxy = INDEX
'
'
'
End Sub


Private Sub ParametriDiDefault()
'
'   Attribuisce i valori di default ai parametri delle
'   routines di interpolazione.  Questa routine viene
'   richiamata ogni volta che si generano nuovi dati
'   casuali o si leggono i dati di un file.
'
    ' Tabella dei parametri di KTB2D:
    Par.tmin = -1E+21   ' Par.tmin < Par.tmax
    Par.tmax = 1E+21    ' Par.tmin < Par.tmax
    Par.nxdis = 1       ' 1 <= Par.nxdis <= 64/Par.nydis
    Par.nydis = 1       ' 1 <= Par.nydis <= 64/Par.nxdis
    Par.ndmin = 4       ' 0 <= Par.ndmin
    Par.ndmax = 8       ' Par.ndmax <= 120
    Par.radius = Sqr((DMAXVAL(XD()) - DMINVAL(XD())) ^ 2 _
                   + (DMAXVAL(YD()) - DMINVAL(YD())) ^ 2)   ' 0 < Par.radius
    Par.ktype = 1       ' 0/1 (0=No, 1=Yes)
    Par.skmean = 2.302
    Par.c0 = 2#
    Par.Nst = 1         ' 1 <= Par.Nst <= 4
    Par.It(1) = 1
    Par.cc(1) = 8#
    Par.ang(1) = 0#     ' 0 <= Par.ang < 360
    Par.AA(1) = 1#      ' Se Par.It(I) = 4 deve essere 0 < Par.AA(I) < 2
    Par.a2(1) = 1#
'
    ' Impostazione parametri per MASUB:
    'IEX = 1     ' E' richiesta l' estrapolazione.
    TP = 10#    ' Tensione della superficie (TP >= 0).
'
    ' Impostazione dei parametri per QSHEP2:
    NQ = 13  ' 5 <= NQ <= MIN(40,ND-1)
    NW = 19  ' 1 <= NW <= MIN(40,ND-1)
    NR = MAX0(1, Sqr(ND / 3))       ' 1 <= NR
'
'
'
End Sub

Private Sub picOrg_Click()
'
'
    frm3D.Punti XD#(), YD#(), ZD#(), Titolo$
'
'
'
End Sub
Private Sub picSurFit_Click()
'
'
    frm3D.Superficie XI#(), YI#(), ZI#(), Titolo$
'
'
'
End Sub
Private Sub ElaboraFileDati(ByVal FileN$)
'
'
    Dim FF%
'
    On Error GoTo ElaboraFileDati_ERR
'
    mnuFile.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
'
    AggiornaFilesRecenti Me.mnuRecenti, FileN$
'
    FF = FreeFile
    Open FileN$ For Input As #FF
'
    ' Legge i punti dati dal file:
    ND = 0
    Do While Not EOF(FF)
        ND = ND + 1
        ReDim Preserve XD(1 To ND), YD(1 To ND), ZD(1 To ND)
        Input #FF, XD(ND), YD(ND), ZD(ND)
    Loop
'
    Call ParametriDiDefault
'
    ' Prepara una griglia
    ' corrispondente ai punti dati:
    GrigliaPuntiDati XD(), YD(), Xs(), Ys()
'
    ' Chiama la routine di interpolazione:
    bDisegnaZC = False
    If optKTB2D Then
        Prova_KTB2D
    ElseIf optMASUB Then
        Prova_MASUB
    ElseIf optQSHEP2D Then
        Prova_QSHEP2D
    End If
    mnuSalvaInterpolati.Enabled = True
'
'
ElaboraFileDati_ERR:
    Close FF
    Screen.MousePointer = vbDefault
    mnuFile.Enabled = True
'
    If (Err <> 0) Then
        MsgBox Err.Description, vbCritical, Err.Source
    End If
'
'
'
End Sub

