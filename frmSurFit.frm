VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSurFit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Surface Fitting"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9510
   Icon            =   "frmSurFit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFilterEnable 
      Alignment       =   1  'Right Justify
      Caption         =   "&Filter Enable:"
      Height          =   255
      Left            =   2990
      TabIndex        =   31
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1208
   End
   Begin VB.CheckBox chkValuesLevels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Val. of Levels:"
      Height          =   255
      Left            =   2900
      TabIndex        =   12
      Top             =   5880
      Width           =   1298
   End
   Begin VB.Frame zFrame2 
      Caption         =   "Test functions:"
      Height          =   795
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
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
      Caption         =   "Interpolation"
      Height          =   795
      Left            =   4800
      TabIndex        =   21
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkGradient 
         Caption         =   "&Gradient"
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
      Caption         =   "Grid"
      Height          =   255
      Left            =   8975
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   400
   End
   Begin VB.CommandButton cmdGrigliaOrg 
      Caption         =   "Grid"
      Height          =   255
      Left            =   4295
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   400
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
      ToolTipText     =   "Click for 3D view of the Surface "
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
      ToolTipText     =   "Click to view 3D points "
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
   Begin VB.Label zLabel08 
      BackStyle       =   0  'Transparent
      Caption         =   "Offset Y"
      Height          =   195
      Left            =   740
      TabIndex        =   41
      Top             =   5640
      Width           =   610
   End
   Begin VB.Label lblOffsetY 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   740
      TabIndex        =   40
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label zLabel09 
      BackStyle       =   0  'Transparent
      Caption         =   "Offset X"
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   5640
      Width           =   610
   End
   Begin VB.Label lblOffsetX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label lblRadius 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.000"
      Height          =   255
      Left            =   1360
      TabIndex        =   37
      Top             =   5880
      Width           =   570
   End
   Begin VB.Label zLabel07 
      BackStyle       =   0  'Transparent
      Caption         =   "Radius"
      Height          =   195
      Left            =   1360
      TabIndex        =   36
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblYMid 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   7930
      TabIndex        =   35
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Mid:"
      Height          =   195
      Left            =   7930
      TabIndex        =   34
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblXMid 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   6040
      TabIndex        =   33
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X Mid:"
      Height          =   195
      Left            =   6040
      TabIndex        =   32
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblYMax 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   7310
      TabIndex        =   30
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label zLabel06 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Max:"
      Height          =   195
      Left            =   7310
      TabIndex        =   29
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblXMax 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   5420
      TabIndex        =   28
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label zLabel05 
      BackStyle       =   0  'Transparent
      Caption         =   "X Max:"
      Height          =   195
      Left            =   5420
      TabIndex        =   27
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblYMin 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   6690
      TabIndex        =   26
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label zLabel04 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Min:"
      Height          =   195
      Left            =   6690
      TabIndex        =   25
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblXMin 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.000"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   5880
      Width           =   610
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
      Caption         =   "---"
      Height          =   255
      Left            =   2390
      TabIndex        =   20
      Top             =   5880
      Width           =   350
   End
   Begin VB.Label zLabel02 
      BackStyle       =   0  'Transparent
      Caption         =   "N +:"
      Height          =   195
      Left            =   2390
      TabIndex        =   19
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label zLabel01 
      BackStyle       =   0  'Transparent
      Caption         =   "ND:"
      Height          =   195
      Left            =   1940
      TabIndex        =   18
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblND 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
      Height          =   255
      Left            =   1940
      TabIndex        =   17
      Top             =   5880
      Width           =   440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadData 
         Caption         =   "&Load data files"
      End
      Begin VB.Menu zSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveInterpolated 
         Caption         =   "&Save interpolated data"
         Enabled         =   0   'False
      End
      Begin VB.Menu zSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEsci 
         Caption         =   "&Exit"
      End
      Begin VB.Menu zSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "Recent Files:"
         Index           =   0
      End
   End
   Begin VB.Menu mnuImpostazioni 
      Caption         =   "&Settings"
   End
   Begin VB.Menu mnuQM 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "I&nstructions"
      End
      Begin VB.Menu zSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&About SurFit"
      End
   End
End
Attribute VB_Name = "frmSurFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================
' Description......: Test form for "Surface Fitting" routines.
' Name of the Files: frmSurFit.frm, frmSurFit.frx
'                    frmSettings.frm, frmSettings.frx
'                    frm3D.frm, frm3D.frx
'                    frmInstr.frm, frmInstr.frx
'                    InfoCr.frm, InfoCr.frx
'                    modKTB2D.bas, modMASUB.bas, modQSHEP2D
'                    modGradiente2D.bas, modLayers.bas
'                    modUtility.bas
' Date.............: 21/9/2001
' Version..........: 1.0 at 32 bits.
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.virgilio.it/flanguasco/
'                    http://www.flanguasco.org
'=============================================================
'
'   Note:   All the vectors and matrices of these routines start from
'            index 1 (ZCol() excluded).
'
Option Explicit
'
Const XDMin# = -25# ' Min of X grid
Const XDMax# = 25#  ' Max of X
Const YDMin# = -25# ' Min of Y
Const YDMax# = 25#  ' Max of Y
Const ZDMin# = 0#   ' Min of Z
Const ZDMax# = 50#  ' Max of Z
'
Dim ND&             ' Number of data in the vectors.
Dim PhiD#()         ' Angle Phi data.
Dim ThetaD#()       ' Angle Theta data.
Dim XD#()           ' Vector data values
Dim YD#()           '  of the surface
Dim ZD#()           '  to be interpolated.
Dim OXD#, OYD#      ' X/Y offset values.
Dim RD#             ' Radius of data.
'
Dim Xs#(), Ys#()    ' Coordinates of the data point grid.
'
Dim NXI&, NYI&      ' Number of columns and rows in
                    ' the interpolated points grid.
Dim PhiI#(), ThetaI#()    ' Coordinates of the interpolated points grid.
Dim XI#(), YI#()    ' Coordinates of the interpolated points grid.
Dim ZI#()           ' Interpolated surface.
Dim ZI_default#     ' Default value of Interpolated surface.
Dim ZC#()           ' Calculated surface.
Dim Grad() As Grad_Type   ' Gradient matrix.
'
' Impostazioni per KTB2D:
Dim Par As ParType, IER&
'
' Settings for MASUB:
Dim TP#
'
' Impostazioni per QSHEP2D:
Dim NQ&, NW&, NR&
'
Dim ZCol&()         ' Table of colors.
Const NTCol& = 1280 ' Number of colors available in ZCol ().
Dim NLiv&           ' Number of levels to trace.
'
Dim FolderN$        ' Folder dei files dati.
'
Dim lZxy&           ' Indice della funzione di prova.
'
Dim bFilterEnabled As Boolean   ' Enable filter on data file proces.
Dim bDrawUpVal As Boolean       ' Draw up the values of the level lines.
Dim bDrawGD As Boolean          ' Draw the darts of the gradient.
Dim bDrawZC As Boolean          ' Draw the calculated surface.
Dim Title$                      ' Title of the picOrg picture.
'
Const Me_W& = 9600          ' Dimensioni di questo
Const Me_H& = 7200 - 255    ' Form [Twips].

Private Sub Test_KTB2D()
'
'   Interpolazione, con "kriging",  di una superficie
'   con punti dati nei vettori XD(), YD(), ZD():
'
    Dim A#, B#, C#, D#, HX#, HY#, Px3!, Py3!
    Dim IER&
'
    ' Prepara i vettori XI() ed YI() con le
    ' coordinate della griglia di interpolazione:
    GridForInterpolation A, B, C, D
'
    HX = (B - A) / CDbl(NXI - 1)
    HY = (D - C) / CDbl(NYI - 1)
'
    IER = 0 ' Niente file di debug.
    Call KTB2D(ND, XD(), YD(), ZD(), Par, NXI, A, HX, NYI, C, HY, ZI(), IER)
    If IER <> 0 Then
        MsgBox "Error " & IER & " in KTB2D", vbCritical
        Exit Sub
    End If
'
    DrawLevels A, B, C, D, Px3, Py3
    If bDrawGD Then
        Call Gradient_2D(XI(), YI(), ZI(), NXI, NYI, Grad())
        DrawGradient Px3, Py3
    End If
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub

Private Sub Test_QSHEP2D()
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
    GridForInterpolation A, B, C, D
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
        MsgBox "Error " & IER & " in QSHEP2D", vbCritical
        Exit Sub
    End If
'
    If bDrawGD Then
        ' Calcola la superficie interpolata ed il gradiente:
        For J = 1 To NYI
            For I = 1 To NXI
                Call QS2GRD(XI(I), YI(J), ND, XD(), YD(), ZD(), NR, LCELL(), LNEXT(), _
                            XMin, YMin, DX, DY, RMAX, RSQ(), ASh(), _
                            ZI(I, J), Grad(I, J).DX, Grad(I, J).DY, IER)
                If IER <> 0 Then
                    MsgBox "Error " & IER & " in QS2GRD", vbCritical
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
    DrawLevels A, B, C, D, Px3, Py3
    If bDrawGD Then DrawGradient Px3, Py3
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub

Private Function Zxy(ByVal x1#, ByVal x2#) As Double
'
'   Two parameter test functions:
'
    Dim A#, B#
'
    Select Case lZxy
        Case 1
        ' Concentric circles:
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
        ' Test function #1 for "Genetic Algorithms":
        A = Cos(Sqr((x1 + 1#) ^ 2 + (x2 - 1#) ^ 2))
        B = Sqr((x1 - 1#) ^ 2 + (x2 + 1#) ^ 2) + 2#
'
        Case 4
        ' Test function #2 for "Genetic Algorithms":
        A = Cos(Sqr((x1) ^ 2 + x2 * x2))
        B = Sqr(0.005 * x1 ^ 2 + x2 * x2) + 2
        B = Sqr(x2 * x2) + 2#
'
        Case 5
        ' Test function #3 for "Genetic Algorithms":
        A = Cos(Sqr((x1 + 1) ^ 2 + (x2 - 1) ^ 2))
        B = Sqr((x1 + 1) ^ 2 + x2 * x2) + 2#
'
        Case 6
        ' Rosembrook:
        A = -100 * (x2 - x1 ^ 2) ^ 2 + (1 - x1) ^ 2
        B = 1#
'
        Case 7
        ' CONREC Example 1:
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

Private Sub chkFilterEnable_Click()
'
'
    bFilterEnabled = (chkFilterEnable.Value = vbChecked)
'
'
End Sub

Private Sub chkGradient_Click()
'
'
    bDrawGD = (chkGradient.Value = vbChecked)
'
    Screen.MousePointer = vbHourglass
'
    ' Call the interpolation routine:
    If optKTB2D Then
        Test_KTB2D
    ElseIf optMASUB Then
        Test_MASUB
    ElseIf optQSHEP2D Then
        Test_QSHEP2D
    End If
'
    Screen.MousePointer = vbDefault
'
'
'
End Sub

Private Sub chkValuesLevels_Click()
'
'
    bDrawUpVal = (chkValuesLevels.Value = vbChecked)
'
    Screen.MousePointer = vbHourglass
'
    ' Call the interpolation routine:
    If optKTB2D Then
        Test_KTB2D
    ElseIf optMASUB Then
        Test_MASUB
    ElseIf optQSHEP2D Then
        Test_QSHEP2D
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

Private Sub cmdTest_Click()
'
'
    Dim N&
'
    cmdTest.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
'
    ND = CLng(RandU(6, 200))
    'D = CLng(89 * 89)
    'ND = CLng(46 * 46)
    ReDim PhiD#(1 To ND), ThetaD#(1 To ND)
    ReDim XD#(1 To ND), YD#(1 To ND), ZD#(1 To ND)
    For N = 1 To ND
        ' Abscissas of data points:
        XD(N) = RandU(-10, 10)
        ' Ordinates of data points:
        YD(N) = RandU(-10, 10)
    Next N
'
    Call DefaultParameters
'
    ' Calculate the value of the surface at the data points:
    For N = 1 To ND
        ZD(N) = Zxy(XD(N), YD(N))
    Next N
'
    ' Prepare a grid corresponding to data points:
    GridPointsData XD(), YD(), Xs(), Ys()
'
    ' Call the interpolation routine:
    bDrawZC = True
    Title$ = ND & " Random points"
    If optKTB2D Then
        Test_KTB2D
    ElseIf optMASUB Then
        Test_MASUB
    ElseIf optQSHEP2D Then
        Test_QSHEP2D
    End If
'
    Screen.MousePointer = vbDefault
    cmdTest.Enabled = True
'
'
'
End Sub

Private Sub Form_Load()
'
'
    Dim Me_L&, Me_T&
'
    LoadFormsPositions Me, Me_L, Me_T
    Me.Move Me_L, Me_T, Me_W, Me_H
'
    LoadRecentFiles Me.mnuRecent, "PointsData"
'
    ZCol() = ColorTable(NTCol)
'
    NXI = 50
    NYI = 50
    NLiv = 20   ' Number of levels to trace.
'
    bFilterEnabled = (chkFilterEnable.Value = vbChecked)
'
    optZxy(1).Value = True
'
'    cmdTest_Click
'
    Dim FN_Temp$
'
    If mnuRecent.UBound = 0 Then
        mnuLoadData_Click
    Else
        FN_Temp$ = mnuRecent(1).Caption
        FN_Temp$ = Right$(FN_Temp$, Len(FN_Temp$) - 3)
'
        If BreakDown(FN_Temp$, FolderN$, Title$) Then
            ProcessDataFile FN_Temp$
        End If
    End If
'
'
'
End Sub

Private Sub Test_MASUB()
'
'   Interpolation of a surface with data points in
'   the XD (), YD (), ZD () vectors:
'
    Dim A#, B#, C#, D#, Px3!, Py3!
    Dim IC&, IEX&
'
    ' Prepare the PhiI(), ThetaI(), XI () and YI () vectors with
    '  the coordinates of the interpolation grid:
    GridForInterpolation A, B, C, D, 0.1
'
    ' Parameter setting for MASUB:
    IC = 1      ' First and only call.
    'IEX = 1     ' Extrapolation is required.
    ZI_default = 0
'
    If Not MASUB(IC, IEX, ND, XD(), YD(), ZD(), TP _
               , NXI, NYI, XI(), YI(), ZI(), ZI_default) Then
        MsgBox "Error in MASUB", vbCritical
        Exit Sub
    End If
    If IEX = 1 Then
        lblNAdd = UBound(XD) - ND   ' Points added for extrapolation.
    End If
'
    DrawLevels A, B, C, D, Px3, Py3
    If bDrawGD Then
        Call Gradient_2D(XI(), YI(), ZI(), NXI, NYI, Grad())
        DrawGradient Px3, Py3
    End If
'
    picOrg.AutoRedraw = False
    picSurFit.AutoRedraw = False
'
'
'
End Sub

Private Sub DisegnaGriglia(Painting As PictureBox, dXI#(), dYI#())
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
        Painting.Line (dXI(I), YMin)-(dXI(I), YMax), vbYellow
    Next I
'
    For J = 1 To UBound(dYI)
        Painting.Line (XMin, dYI(J))-(XMax, dYI(J)), vbYellow
    Next J
'
'
'
End Sub

Public Sub GridPointsData(XD#(), YD#(), XGD#(), YGD#())
'
'   Prepare the vector vectors of the data points, eliminating
'   the double values and ordering them in increasing direction:
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

Private Sub DrawLevels(ByVal A#, ByVal B#, ByVal C#, ByVal D#, _
    ByRef Px3!, ByRef Py3!)
'
'   Displaying level curves:
'
    Dim I&, J&, K&, N&, ZMin#, ZMax#, ZDif#, Msg$
    ReDim ZLin(1 To NLiv) As LineaLivello_Type
'
    ' Set the graphic:
    Painting picOrg, A, B, C, D, , , 3, Px3, Py3, "Data points: " & Title$, _
                  "x", "y", True
    Painting picSurFit, A, B, C, D, , , , , , "Interpolated surface", "x", "y", True
'
    ' Displaying data points:
    For N = 1 To ND
        picOrg.Circle (XD(N), YD(N)), Px3, vbRed
    Next N
'
    ' Displaying added points (MASUB only):
    For N = ND + 1 To UBound(XD)
        picOrg.Circle (XD(N), YD(N)), Px3, vbYellow
    Next N
'
    ' Find the Max. and Min. of the surface:
    ZMin = ZI(1, 1)
    ZMax = ZI(1, 1)
    For I = 1 To NXI
        For J = 1 To NYI
            ZMin = DMIN1(ZMin, ZI(I, J))
            ZMax = DMAX1(ZMax, ZI(I, J))
        Next J
    Next I
'
    ' Prepare the layer vector:
    For K = 1 To NLiv
        ZLin(K).dLinVal = ZMin + (K - 1) * (ZMax - ZMin) / (NLiv - 1)
        ZLin(K).lLinCol = ZCol((K - 1) * (NTCol - 1) / (NLiv - 1))
        ZLin(K).lLinSps = 1
        'ZLin(k).lLinSps = IIf(k Mod 3 = 0, 2, 1)
        If bDrawUpVal Then
            ZLin(K).sLblTes = Format$(ZLin(K).dLinVal, "#0.000")
        Else
            ZLin(K).sLblTes = ""
        End If
        ZLin(K).lLblCol = ZLin(K).lLinCol
        ZLin(K).lLblFSz = 10
    Next K
'
    ' Displaying the interpolated surface:
    CONREC_pK picSurFit, ZI(), XI(), YI(), ZLin(), Msg$
    If Msg$ <> "" Then
        MsgBox Msg$, vbCritical, " CONREC of ZI()"
    End If
'
    If bDrawZC Then
        ' Calculate the value of the surface on all points of the grid:
        For I = 1 To NXI
            For J = 1 To NYI
                ZC(I, J) = Zxy(XI(I), YI(J))
            Next J
        Next I
'
        ' Displaying the calculated area:
        CONREC_pK picOrg, ZC(), XI(), YI(), ZLin(), Msg$
        If Msg$ <> "" Then
            MsgBox Msg$, vbCritical, " CONREC of ZC()"
        End If
    End If
'
    ' Draw the surface in 3D:
    'frm3D.Surface XI#(), YI#(), ZI#(), Title$
    frm3D.Surface PhiD#(), ThetaD#(), XI#(), YI#(), ZI#(), Title$, False, XDMin, XDMax, YDMin, YDMax, ZDMin, ZDMax
'
'
'
End Sub

Private Sub DrawGradient(ByVal Px3!, ByVal Py3!)
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
        M$ = "The surface is flat." & vbNewLine
        M$ = M$ & "It is not possible to draw the gradient."
        MsgBox M$, vbInformation, " Gradient design"
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
    Const ApF# = PI_2 / 3#  ' 30?
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
    If IsLoaded(frmInstructions) Then Unload frmInstructions
    If IsLoaded(frm3D) Then Unload frm3D
'
    SaveRecentFiles Me.mnuRecent, "PointsData"
'
    SaveFormsPositions Me
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
    If frmSettings.OpenForm(ND, NXI, NYI, NLiv, Par, NQ, NW, NR, TP) Then
        Screen.MousePointer = vbHourglass
    '
        ' Call the interpolation routine with the modified parameters:
        If optKTB2D Then
            Test_KTB2D
        ElseIf optMASUB Then
            Test_MASUB
        ElseIf optQSHEP2D Then
            Test_QSHEP2D
        End If
    '
        Screen.MousePointer = vbDefault
    End If
'
'
'
End Sub

Private Sub mnuInformation_Click()
'
'
    frmCopyright.OpenForm Me
'
'
End Sub

Private Sub mnuInstructions_Click()
'
'
    frmInstructions.OpenForm App.Path & "\Instr.rtf"
'
'
'
End Sub

Private Sub mnuLoadData_Click()
'
'
    Dim FN_Temp$, M$
'
    On Error GoTo mnuLoadData_Click_ERR
'
    FN_Temp$ = CMDialog_Files(CMDialog1, "Open", "Data files", "*.dat;*.txt", _
                              FolderN$, , " Surface data to be interpolated")
'
'
    If BreakDown(FN_Temp$, FolderN$, Title$) Then
        ProcessDataFile FN_Temp$
    End If
'
'
mnuLoadData_Click_ERR:
    If Err <> 0 And Err <> cdlCancel Then
        M$ = "Error " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " mnuLoadData/" & Err.Source
    End If
'
'
'
End Sub

Private Sub mnuRecent_Click(INDEX As Integer)
'
'
    Dim FN_Temp$
'
    If INDEX = 0 Then Exit Sub
'
'
    FN_Temp$ = mnuRecent(INDEX).Caption
    FN_Temp$ = Right$(FN_Temp$, Len(FN_Temp$) - 3)
'
    If BreakDown(FN_Temp$, FolderN$, Title$) Then
        ProcessDataFile FN_Temp$
    End If
'
'
'
End Sub

Private Sub mnuSaveInterpolated_Click()
'
'
    Dim FF%, I&, J&, FileNome$, M$
'
    On Error GoTo mnuSaveInterpolated_Click_ERR
'
    FileNome$ = CMDialog_Files(CMDialog1, "Save", "Data files", "*.dat;*.txt", _
                               FolderN$, , " Interpolated data")
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
mnuSaveInterpolated_Click_ERR:
    If Err <> 0 And Err <> cdlCancel Then
        M$ = "Error " & Str$(Err.Number) & vbNewLine
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
    Test_KTB2D
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub optMASUB_Click()
'
'
    Screen.MousePointer = vbHourglass
    Test_MASUB
    Screen.MousePointer = vbDefault
'
'
'
End Sub
Private Sub optQSHEP2D_Click()
'
'
    Screen.MousePointer = vbHourglass
    Test_QSHEP2D
    Screen.MousePointer = vbDefault
'
'
'
End Sub

Private Sub GridForInterpolation(ByRef A#, ByRef B#, ByRef C#, ByRef D#, _
    Optional ByVal est# = 0)
'
'   Prepare the vectors containing abscissa and order the interpolation grid.
'   It also calculates the extremes of the interpolation coordinates,
'   possibly extending them to the Est factor:
'   to be used, mainly, for MASUB which is easily mistaken when the
'   interpolation extremes coincide with the ends of the data points.
'
    Dim I&, J&, HX#, HY#
'
    ' To delete points added by a previous call to MASUB:
    ReDim Preserve XD(1 To ND)
    ReDim Preserve YD(1 To ND)
    lblND = ND
    'lblNAdd = "--"
'
    ' Find the Max. and Min. coordinates of the data points:
    A = DMINVAL(XD())   ' Minimum X.
    B = DMAXVAL(XD())   ' Maximum X.
    C = DMINVAL(YD())   ' Minimum Y.
    D = DMAXVAL(YD())   ' Maximum Y.
    lblXMin = Format$(A, "#0.000")
    lblXMax = Format$(B, "#0.000")
    lblXMid = Format$((B + A) / 2#, "#0.000")
    lblYMin = Format$(C, "#0.000")
    lblYMax = Format$(D, "#0.000")
    lblYMid = Format$((D + C) / 2#, "#0.000")
'
    ' And widens the Est factor:
    HX = (B - A)
    A = A - est * HX
    B = B + est * HX
    HY = (D - C)
    C = C - est * HY
    D = D + est * HY
'
    ReDim PhiI(1 To NXI), ThetaI#(1 To NYI) ' Angle phi and theta of the interpolated points grid.
    ReDim XI(1 To NXI), YI#(1 To NYI)   ' Coordinates of the interpolated points grid.
    ReDim ZI(1 To NXI, 1 To NYI)        ' Interpolated surface.
    ReDim ZC(1 To NXI, 1 To NYI)        ' Calculated surface.
    ReDim Grad(1 To NXI, 1 To NYI)      ' Gradient matrix.
'
    ' Abscissas of the grid of the interpolated points:
    HX = (B - A) / CDbl(NXI - 1)
    For I = 1 To NXI
        XI(I) = A + (I - 1) * HX
    Next I
'
    ' Ordinates of the grid of the interpolated points:
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


Private Sub DefaultParameters()
'
'   It assigns the default values to the parameters of interpolation routines.
'   This routine is invoked whenever you generate new random data or read data
'    from a file.
'
    ' Table of KTB2D parameters:
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
    Par.AA(1) = 1#      ' If Par.It(I) = 4 must be 0 < Par.AA(I) < 2
    Par.a2(1) = 1#
'
    ' Parameter setting for MASUB:
    'IEX = 1     ' Extrapolation is required.
    TP = 10#    ' Surface tension (TP >= 0).
'
    ' Setting the parameters for QSHEP2:
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
    'frm3D.Points XD#(), YD#(), ZD#(), Title$
    frm3D.Points PhiD#(), ThetaD#(), XD#(), YD#(), ZD#(), Title$, False, XDMin, XDMax, YDMin, YDMax, ZDMin, ZDMax
'
'
'
End Sub

Private Sub picSurFit_Click()
'
'
    'frm3D.Surface XI#(), YI#(), ZI#(), Title$
    frm3D.Surface PhiD#(), ThetaD#(), XI#(), YI#(), ZI#(), Title$, False, XDMin, XDMax, YDMin, YDMax, ZDMin, ZDMax
'
'
'
End Sub

Private Sub ProcessDataFile(ByVal FileN$)
'
'
    Dim FF%
    Dim lND&        ' Number of data in the vectors.
    Dim lPhiD#()    ' Angle Phi data.
    Dim lThetaD#()  ' Angle Theta data.
    Dim lXD#()      ' Vector data values
    Dim lYD#()      ' of the surface
    Dim lZD#()      ' to be interpolated.
    Dim lZDAvg#     ' Average of ZD().
    Dim lZDMin#     ' Min of ZD().
    Dim lDSkip() As Boolean   ' Flag data will skip.
    Dim I%, J%
'
    On Error GoTo ProcessDataFile_ERR
'
    mnuFile.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
'
    UpdateRecentFiles Me.mnuRecent, FileN$
'
    FF = FreeFile
    Open FileN$ For Input As #FF
'
    ' Read the offset X/Y data from the file:
    If (Not EOF(FF)) Then
        Input #FF, OXD, OYD, RD
        lblOffsetX = OXD
        lblOffsetY = OYD
        lblRadius = RD
    End If
'
    If (bFilterEnabled = False) Then
        ' Read the data points from the file:
        ND = 0
        Do While Not EOF(FF)
            ND = ND + 1
            ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), XD(1 To ND), YD(1 To ND), ZD(1 To ND)
            Input #FF, PhiD(ND), ThetaD(ND), XD(ND), YD(ND), ZD(ND)
            If (XD(ND) = 0 And YD(ND) = 0 And ZD(ND) = 50) Then
                ND = ND - 1
            Else
                XD(ND) = XD(ND) - OXD
                YD(ND) = YD(ND) - OYD
            End If
        Loop
        lblNAdd = 0
    Else ' Else of If (bFilterEnabled = False) Then
        ' Read the data points from the file:
        lND = 0
        Do While Not EOF(FF)
            lND = lND + 1
            ReDim Preserve lPhiD(1 To lND), lThetaD(1 To lND), _
                            lXD(1 To lND), lYD(1 To lND), lZD(1 To lND)
            Input #FF, lPhiD(lND), lThetaD(lND), lXD(lND), lYD(lND), lZD(lND)
            If (lXD(lND) = 0 And lYD(lND) = 0 And lZD(lND) = 50) Then
                lND = lND - 1
                GoTo Continue
            End If
            lXD(lND) = lXD(lND) - OXD
            lYD(lND) = lYD(lND) - OYD
            If (RD <> 0) And (Sqr(lXD(lND) ^ 2 + lYD(lND) ^ 2) > RD) Then
                lND = lND - 1
                GoTo Continue
            End If
Continue:
        Loop
'
        ' Sort the vectors so that the points with major Z remain behind:
        QuickSort5V lZD(), lXD(), lYD(), lPhiD(), lThetaD(), 1, lND, 1
'
        ' Skip point over Z axis.
        Dim lZOk As Boolean
        Dim lZOverCnt%
        ReDim Preserve lDSkip(1 To lND)
'
        ND = 0
        lZDAvg = 0
        lZDMin = ZDMax
        For I = 1 To lND
            lZOk = False
            lZOverCnt = 0
            For J = 1 To I - 1
                If (lDSkip(J) = False) Then
                    If (Abs(lXD(I) - lXD(J)) < 2#) And (Abs(lYD(I) - lYD(J)) < 2#) Then
                        lZOverCnt = lZOverCnt + 1
                        If (lZD(I) < lZD(J) + 2#) Then
                            lZOk = True
                            Exit For
                        End If
                    End If
                End If
            Next J
            If lZOverCnt = 0 Or lZOk = True Then
                ND = ND + 1
                ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), _
                                XD(1 To ND), YD(1 To ND), ZD(1 To ND)
                PhiD(ND) = lPhiD(I)
                ThetaD(ND) = lThetaD(I)
                XD(ND) = lXD(I)
                YD(ND) = lYD(I)
                ZD(ND) = lZD(I)
                lZDAvg = lZDAvg + lZD(I)
                If (lZD(I) < lZDMin) Then lZDMin = lZD(I)
            Else
                lDSkip(I) = True
            End If
        Next I
'
        lZDAvg = lZDAvg / ND
'
        ' Fill data for BIN shape.
        Dim BinX#, BinY#, Distance#, DistanceMin#, DistanceMinZ#
        Dim AddBin As Boolean
        Const DistanceRangeMin# = 2#
        Dim DistanceRangeMax#
        DistanceRangeMax = RD
'
        lND = ND
        lblNAdd = 0
'
        For I = 0 To 360 - 1 Step 10
            BinX = RD * Sin(I * PI / 180#)
            BinY = RD * Cos(I * PI / 180#)
            AddBin = True
            DistanceMin = DMAX1(XDMax, YDMax)
            DistanceMinZ = lZDAvg
            For J = 1 To lND
                If (Abs(BinX - XD(J)) < DistanceRangeMin) _
                    And (Abs(BinY - YD(J)) < DistanceRangeMin) Then
                    AddBin = False
                    Exit For
                End If
                Distance = Sqr((BinX - XD(J)) ^ 2 + (BinY - YD(J)) ^ 2)
                'If (Distance < (lZDAvg - lZDMin)) Then
                If (Distance < DistanceRangeMin) Then
                    AddBin = False
                    Exit For
                End If
                If (Distance < DistanceMin) Then
                    DistanceMin = Distance
                    DistanceMinZ = ZD(J)
                End If
            Next J
            If (AddBin = True) Then
                ND = ND + 1
                ReDim Preserve PhiD(1 To ND), ThetaD(1 To ND), _
                                XD(1 To ND), YD(1 To ND), ZD(1 To ND)
                PhiD(ND) = I
                ThetaD(ND) = DistanceMin
                XD(ND) = BinX
                YD(ND) = BinY
                If (DistanceMin > DistanceRangeMax#) Then
                    ZD(ND) = lZDMin
                Else
                    ZD(ND) = _
                        (lZDMin - DistanceMinZ) _
                        / (DistanceRangeMax - 0) _
                        * DistanceMin _
                        + DistanceMinZ
                End If
            End If
        Next I
'
        lblNAdd = ND - lND
    End If ' End of If (bFilterEnabled = False) Then
'
    Call DefaultParameters
'
    ' Prepare a grid corresponding to data points:
    GridPointsData XD(), YD(), Xs(), Ys()
'
    ' Call the interpolation routine:
    bDrawZC = False
    If optKTB2D Then
        Test_KTB2D
    ElseIf optMASUB Then
        Test_MASUB
    ElseIf optQSHEP2D Then
        Test_QSHEP2D
    End If
    mnuSaveInterpolated.Enabled = True
'
'
ProcessDataFile_ERR:
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

