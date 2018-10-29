VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmImpostazioni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Settings"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmImpostazioni.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "&Close"
      Height          =   315
      Left            =   3600
      TabIndex        =   50
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2700
      TabIndex        =   49
      Top             =   3960
      Width           =   675
   End
   Begin VB.Frame zFrame4 
      Caption         =   "Interpolation grid"
      Height          =   1635
      Left            =   4800
      TabIndex        =   48
      Top             =   2640
      Width           =   1935
      Begin VB.TextBox txtNLiv 
         Height          =   285
         Left            =   780
         TabIndex        =   44
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtNXI 
         Height          =   285
         Left            =   780
         TabIndex        =   40
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtNYI 
         Height          =   285
         Left            =   780
         TabIndex        =   42
         Top             =   780
         Width           =   915
      End
      Begin VB.Label zLabel25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NLi&v:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label zLabel22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NX&I:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   420
         Width           =   555
      End
      Begin VB.Label zLabel23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NY&I:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   825
         Width           =   555
      End
   End
   Begin VB.Frame zFrame3 
      Caption         =   "QSHEP2D:"
      Height          =   1635
      Left            =   4800
      TabIndex        =   47
      Top             =   960
      Width           =   1935
      Begin VB.TextBox txtNR 
         Height          =   285
         Left            =   780
         TabIndex        =   38
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtNW 
         Height          =   285
         Left            =   780
         TabIndex        =   36
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtNQ 
         Height          =   285
         Left            =   780
         TabIndex        =   34
         Top             =   360
         Width           =   915
      End
      Begin VB.Label zLabel21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N&R:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1230
         Width           =   555
      End
      Begin VB.Label zLabel20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N&W:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   825
         Width           =   555
      End
      Begin VB.Label zLabel19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N&Q:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.Frame zFrame2 
      Caption         =   "MASUB:"
      Height          =   795
      Left            =   4800
      TabIndex        =   46
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtTP 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Top             =   360
         Width           =   915
      End
      Begin VB.Label zLabel18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T&P:"
         Height          =   255
         Left            =   60
         TabIndex        =   31
         Top             =   405
         Width           =   555
      End
   End
   Begin VB.Frame zFrame1 
      Caption         =   "KTB2D"
      Height          =   3735
      Left            =   180
      TabIndex        =   45
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtEditPunti 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   30
         ToolTipText     =   "Enter to confirm "
         Top             =   2220
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdNstOK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         Picture         =   "frmImpostazioni.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   270
      End
      Begin VB.TextBox txtNst 
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         ToolTipText     =   "OK or Enter to confirm "
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtC0 
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtSkmean 
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtKtype 
         Height          =   285
         Left            =   780
         TabIndex        =   15
         Top             =   3300
         Width           =   915
      End
      Begin VB.TextBox txtRadius 
         Height          =   285
         Left            =   780
         TabIndex        =   13
         Top             =   2880
         Width           =   915
      End
      Begin VB.TextBox txtNdmax 
         Height          =   285
         Left            =   780
         TabIndex        =   11
         Top             =   2460
         Width           =   915
      End
      Begin VB.TextBox txtNdmin 
         Height          =   285
         Left            =   780
         TabIndex        =   9
         Top             =   2040
         Width           =   915
      End
      Begin VB.TextBox txtNydis 
         Height          =   285
         Left            =   780
         TabIndex        =   7
         Top             =   1620
         Width           =   915
      End
      Begin VB.TextBox txtNxdis 
         Height          =   285
         Left            =   780
         TabIndex        =   5
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtTMax 
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txtTMin 
         Height          =   285
         Left            =   780
         TabIndex        =   1
         Top             =   360
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid grdNstPar 
         Height          =   1815
         Left            =   2880
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Click to change "
         Top             =   1620
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   6
         FixedCols       =   0
         AllowBigSelection=   0   'False
         GridLines       =   2
         ScrollBars      =   1
         MousePointer    =   99
         MouseIcon       =   "frmImpostazioni.frx":067C
      End
      Begin VB.Label zLabel12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "n:"
         Height          =   255
         Left            =   2220
         TabIndex        =   23
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label zLabel17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "a2(n):"
         Height          =   255
         Left            =   2220
         TabIndex        =   28
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label zLabel16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "aa(n):"
         Height          =   255
         Left            =   2220
         TabIndex        =   27
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label zLabel15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ang(n):"
         Height          =   255
         Left            =   2220
         TabIndex        =   26
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label zLabel14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "cc(n):"
         Height          =   255
         Left            =   2220
         TabIndex        =   25
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label zLabel13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "It(n):"
         Height          =   255
         Left            =   2220
         TabIndex        =   24
         Top             =   1920
         Width           =   555
      End
      Begin VB.Line zLine3 
         X1              =   2520
         X2              =   2220
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line zLine2 
         X1              =   2220
         X2              =   2220
         Y1              =   1320
         Y2              =   1800
      End
      Begin VB.Line zLine1 
         X1              =   2400
         X2              =   2220
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label zLabel11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ns&t:"
         Height          =   255
         Left            =   2220
         TabIndex        =   20
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label zLabel10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "c&0:"
         Height          =   255
         Left            =   2220
         TabIndex        =   18
         Top             =   825
         Width           =   555
      End
      Begin VB.Label zLabel09 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&skmean:"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   405
         Width           =   615
      End
      Begin VB.Label zLabel08 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&ktype:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3345
         Width           =   555
      End
      Begin VB.Label zLabel07 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "radi&us:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2925
         Width           =   555
      End
      Begin VB.Label zLabel06 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ndm&ax:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2505
         Width           =   555
      End
      Begin VB.Label zLabel05 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "n&dmin:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label zLabel04 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "n&ydis:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1665
         Width           =   555
      End
      Begin VB.Label zLabel03 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "n&xdis:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label zLabel02 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "t&max:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   825
         Width           =   555
      End
      Begin VB.Label zLabel01 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "tmi&n:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   405
         Width           =   555
      End
   End
   Begin VB.Label zLabel24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ND:"
      Height          =   255
      Left            =   420
      TabIndex        =   52
      Top             =   3975
      Width           =   435
   End
   Begin VB.Label lblND 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   960
      TabIndex        =   51
      Top             =   3960
      Width           =   915
   End
End
Attribute VB_Name = "frmImpostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Descrizione.....: Form per l' impostazione dei parametri delle
'                   routines di "Surface Fitting".
' Nome dei Files..: frmSurFit.frm, frmSurFit.frx
'                   frmImpostazioni.frm, frmImpostazioni.frx
'                   modUtilita.bas
' Data............: 21/9/2001
' Versione........: 1.0 a 32 bits.
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco 
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.it/flanguasco/
'                   http://www.flanguasco.org
'===============================================================
'
Option Explicit
'
Dim ND&             ' N?di dati da interpolare.
Dim NXI&, NYI&      ' N?di colonne e di righe nella
                    ' griglia dei punti interpolati.
Dim NLiv&           ' N?di livelli per CONREC.
'
Dim Par As ParType  ' Parametri di KTB2D.
'
Dim NQ&, NW&, NR&   ' Parametri di QSHEP2D.
'
Dim TP#             ' Parametri di MASUB.
'
Dim RS1&            ' Posizionamenti
Dim CS1&            ' per l' "editing"
Dim RS1_O&          ' dei valori dei punti
Dim CS1_O&          ' nella tabella.
'
Dim grdNstPar_Left& ' Posizioni all' interno
Dim grdNstPar_Top&  ' della tabella grdNst.
'
Dim fOK As Boolean
Dim grdMouseDown As Boolean
Friend Function Apri(ByVal ND_I&, ByRef NXI_I&, ByRef NYI_I&, ByRef NLiv_I&, _
    Par_I As ParType, ByRef NQ_I&, ByRef NW_I&, ByRef NR_I&, ByRef TP_I#) As Boolean
'
'
    ND = ND_I
'
    ' Per griglia:
    NXI = NXI_I
    NYI = NYI_I
    NLiv = NLiv_I
'
    ' Per routines di interpolazione:
    Par = Par_I ' KTB2D.
    NQ = NQ_I   ' QSHEP2D.
    NW = NW_I   '    "
    NR = NR_I   '    "
    TP = TP_I   ' MASUB.
'
    Me.Show vbModal
'
    Apri = fOK
    If fOK Then
        ' Ritorna i parametri modificati:
        NXI_I = NXI
        NYI_I = NYI
        NLiv_I = NLiv
        Par_I = Par
        NQ_I = NQ
        NW_I = NW
        NR_I = NR
        TP_I = TP
    End If
'
'
'
End Function
Private Sub cmdChiudi_Click()
'
'
    Unload Me
'
'
'
End Sub

Private Sub cmdNstOK_Click()
'
'
    Dim C&
'
    If Val(txtNst) < 1 Then txtNst = 1
    If Val(txtNst) > 4 Then txtNst = 4
'
    Par.Nst = Val(txtNst)
'
    AggiornaTabellaNst
'
'
'
End Sub

Private Sub cmdOK_Click()
'
'
    fOK = Verifica
    If fOK Then Unload Me
'
'
'
End Sub
Private Sub Form_Load()
'
'
    Dim I&, C&, R&
    Dim Me_L&, Me_T&
'
    LeggiPosizioneForm Me, Me_L, Me_T
    Me.Move Me_L, Me_T ', Me_W, Me_H
'
    lblND = ND
'
    ' Griglia:
    txtNXI = NXI
    txtNYI = NYI
    txtNLiv = NLiv
'
    ' KTB2D:
    txtTMin = Par.tmin
    txtTMax = Par.tmax
    txtNxdis = Par.nxdis
    txtNydis = Par.nydis
    txtNdmin = Par.ndmin
    txtNdmax = Par.ndmax
    txtRadius = Par.radius
    txtKtype = Par.ktype
    txtSkmean = Par.skmean
    txtC0 = Par.c0
'
    txtNst = Par.Nst
    AggiornaTabellaNst
'
    grdNstPar_Left = grdNstPar.Left + 45
    grdNstPar_Top = grdNstPar.Top + 45
'
    ' QSHEP2D:
    txtNQ = NQ
    txtNW = NW
    txtNR = NR
'
    ' MASUB:
    txtTP = TP
'
    fOK = False
'
'
'
End Sub
Private Sub Form_Unload(Cancel As Integer)
'
'
    SalvaPosizioneForm Me
'
'
'
End Sub


Private Sub grdNstPar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    If Button <> vbLeftButton Then Exit Sub
    
    With grdNstPar
        RS1 = .Row
        CS1 = .Col
        txtEditPunti.Visible = False
    End With
'
    grdMouseDown = True
'
'
'
End Sub
Private Sub grdNstPar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    With grdNstPar
        If Button = vbLeftButton And grdMouseDown Then
            .Row = RS1
            .Col = CS1
            txtEditPunti.Text = .Text
'
            txtEditPunti.Left = .ColPos(CS1) + grdNstPar_Left
            txtEditPunti.Top = .RowPos(RS1) + grdNstPar_Top
            txtEditPunti.Width = .ColWidth(CS1) - 15
            txtEditPunti.Height = .RowHeight(RS1) - 15
            txtEditPunti.Visible = True
            txtEditPunti.SetFocus
        Else
            txtEditPunti.Visible = False
        End If
        grdMouseDown = False
'
        RS1_O = RS1
        CS1_O = CS1
    End With
'
'
'
End Sub
Private Sub txtC0_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii)
'
'
'
End Sub
Private Sub txtEditPunti_Change()
'
'
    If Not txtEditPunti.Visible Then Exit Sub
'
'
'
End Sub

Private Sub txtEditPunti_KeyPress(KeyAscii As Integer)
'
'
    Select Case KeyAscii
        Case vbKeyEscape
        KeyAscii = 0
        txtEditPunti.Visible = False
        txtEditPunti.Text = ""
'
        Case vbKeyReturn
        KeyAscii = 0
        ' Aggiorna i valori con l' evento LostFocus:
        txtEditPunti.Visible = False
'
        Case vbKeyTab
        KeyAscii = 0
        ' Aggiorna i valori con l' evento LostFocus:
        txtEditPunti.Visible = False
'
        Case Else
        If RS1 = 1 Then
            KeyAscii = KAscNumInteri(KeyAscii)
        Else
            KeyAscii = KAscNumReali(KeyAscii)
        End If
    End Select
'
'
'
End Sub
Private Sub txtEditPunti_LostFocus()
'
'
    If Trim$(txtEditPunti.Text) <> "" Then
        Select Case RS1_O
            Case 1
            Par.It(CS1_O + 1) = Val(txtEditPunti)
            Case 2
            Par.cc(CS1_O + 1) = Val(Replace(txtEditPunti, ",", "."))
            Case 3
            Par.ang(CS1_O + 1) = Val(Replace(txtEditPunti, ",", "."))
            Case 4
            Par.AA(CS1_O + 1) = Val(Replace(txtEditPunti, ",", "."))
            Case 5
            Par.a2(CS1_O + 1) = Val(Replace(txtEditPunti, ",", "."))
        End Select
'
        AggiornaTabellaNst
    End If
'
    txtEditPunti.Visible = False
    txtEditPunti.Text = ""
'
'
'
End Sub
Private Sub txtKtype_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNdmax_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNdmin_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNLiv_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNQ_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNR_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNst_KeyPress(KeyAscii As Integer)
'
'
    If KeyAscii = vbKeyReturn Then
        cmdNstOK_Click
    ElseIf KeyAscii = vbKeyEscape Then
        txtNst = Par.Nst
    End If
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNW_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNxdis_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNXI_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNydis_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtNYI_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumInteri(KeyAscii)
'
'
'
End Sub
Private Sub txtRadius_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii)
'
'
'
End Sub
Private Sub txtSkmean_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii)
'
'
'
End Sub
Private Sub txtTMax_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub
Private Sub txtTMin_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii, True)
'
'
'
End Sub
Private Sub txtTP_KeyPress(KeyAscii As Integer)
'
'
    KeyAscii = KAscNumReali(KeyAscii)
'
'
'
End Sub
Private Function Verifica() As Boolean
'
'   Verifica i valori scritti nelle TextBoxes
'   e li assegna alle variabili di ritorno:
'
    Dim N&, M$
'
    On Error Resume Next
'
    ' Controlli per i parametri della griglia:
    If CLng(txtNXI) < 3 Then
        M$ = M$ & "Must be NXI >= 3" & vbNewLine
    Else
        NXI = CLng(txtNXI)
    End If
'
    If CLng(txtNYI) < 3 Then
        M$ = M$ & "Must be NYI >= 3" & vbNewLine
    Else
        NYI = CLng(txtNYI)
    End If
'
    If CLng(txtNLiv) <= 0 Then
        M$ = M$ & "Must be NLiv > 0" & vbNewLine
    Else
        NLiv = CLng(txtNLiv)
    End If
'
    ' Controlli per i parametri di KTB2D:
    If CDbl(txtTMax) <= CDbl(txtTMin) Then
        M$ = M$ & "Must be tmin < tmax" & vbNewLine
    Else
        Par.tmin = CDbl(txtTMin)
        Par.tmax = CDbl(txtTMax)
    End If
'
    If (CLng(txtNxdis) < 1) Or (64 / CLng(txtNydis) < CLng(txtNxdis)) _
    Or (CLng(txtNydis) < 1) Or (64 / CLng(txtNxdis) < CLng(txtNydis)) Then
        M$ = M$ & "Must be 1 <= Nxdis <= 64/nydis" & vbNewLine _
                & "and     1 <= Nydis <= 64/nxdis" & vbNewLine
    Else
        Par.nxdis = CLng(txtNxdis)
        Par.nydis = CLng(txtNydis)
    End If
'
    If (CLng(txtNdmin) < 1) Or (CLng(txtNdmax) <= CLng(txtNdmin)) _
    Or (120 < CLng(txtNdmax)) Then
        M$ = M$ & "Must be 1 <= Ndmin < Ndmax" & vbNewLine _
                & "and       Ndmin <  Ndmax <= 120" & vbNewLine
    Else
        Par.ndmin = CLng(txtNdmin)
        Par.ndmax = CLng(txtNdmax)
    End If
'
    If CDbl(txtRadius) <= 0 Then
        M$ = M$ & "Must be radius > 0" & vbNewLine
    Else
        Par.radius = CDbl(txtRadius)
    End If
'
    If (CLng(txtKtype) <> 0) And (CLng(txtKtype) <> 1) Then
        M$ = M$ & "Must be ktype = 0 o ktype = 1" & vbNewLine
    Else
        Par.ktype = CLng(txtKtype)
    End If
'
    If CDbl(txtSkmean) < 0 Then
        M$ = M$ & "Must be skmean >= 0" & vbNewLine
    Else
        Par.skmean = CDbl(txtSkmean)
    End If
'
    If CDbl(txtC0) < 0 Then
        M$ = M$ & "Must be c0 >= 0" & vbNewLine
    Else
        Par.c0 = CDbl(txtC0)
    End If
'
    If (Par.Nst < 1) Or (4 < Par.Nst) Then
        M$ = M$ & "Must be 1 <= nst <= 4" & vbNewLine
'
    Else
        For N = 1 To Par.Nst
            If (Par.It(N) < 1) Or (4 < Par.It(N)) Then
                M$ = M$ & "Must be 1 <= It(" & N & ") <= 4" & vbNewLine
            End If
'
            If Par.cc(N) <= 0 Then
                M$ = M$ & "Must be cc(" & N & ") > 0" & vbNewLine
            End If
'
            If (Par.ang(N) < 0) Or (360 < Par.ang(N)) Then
                M$ = M$ & "Must be 0 <= ang(" & N & ") <= 360" & vbNewLine
            End If
'
            If (Par.It(N) = 4) And ((Par.AA(N) <= 0) Or (2 <= Par.AA(N))) Then
                M$ = M$ & "Must be 0 < aa(" & N & ") < 2" & vbNewLine
            ElseIf (Par.AA(N) <= 0) Then
                M$ = M$ & "Must be 0 < aa(" & N & ")" & vbNewLine
            End If
'
            If Par.a2(N) <= 0 Then
                M$ = M$ & "Must be a2(" & N & ") > 0" & vbNewLine
            End If
        Next N
    End If
'
    ' Controlli per i parametri di MASUB:
    If CDbl(txtTP) < 0 Then
        M$ = M$ & "Must be TP >= 0" & vbNewLine
    Else
        TP = CDbl(txtTP)
    End If
'
    ' Controlli per i parametri di QSHEP2D:
    If (CLng(txtNQ) < 5) Or (MIN0(40, ND - 1) < CLng(txtNQ)) Then
        M$ = M$ & "Must be 5 <= NQ <= " & MIN0(40, ND - 1) _
                & " = MIN(40, ND - 1)" & vbNewLine
    Else
        NQ = CLng(txtNQ)
    End If
'
    If (CLng(txtNW) < 1) Or (MIN0(40, ND - 1) < CLng(txtNW)) Then
        M$ = M$ & "Must be 1 <= NW <= " & MIN0(40, ND - 1) _
                & " = MIN(40, ND - 1)" & vbNewLine
    Else
        NW = CLng(txtNW)
    End If
'
    If (CLng(txtNR) < 1) Then
        M$ = M$ & "Must be NR > 0" & vbNewLine
    Else
        NR = CLng(txtNR)
    End If
'
'
    Verifica = (M$ = "")
    If M$ <> "" Then
        MsgBox M$, vbCritical, " Verifica"
    End If
'
'
'
End Function
Private Sub AggiornaTabellaNst()
'
'
    Dim C&
'
    With grdNstPar
        .Cols = Par.Nst
        For C = 1 To Par.Nst
            .ColWidth(C - 1) = 825
            .FixedAlignment(C - 1) = flexAlignCenterCenter
            .ColAlignment(C - 1) = flexAlignRightCenter
            .Col = C - 1
'
            .Row = 0
            .Text = Format$(C)
'
            .Row = 1
            .Text = Format$(Par.It(C))
'
            .Row = 2
            .Text = Format$(Par.cc(C))
'
            .Row = 3
            .Text = Format$(Par.ang(C))
'
            .Row = 4
            .Text = Format$(Par.AA(C))
'
            .Row = 5
            .Text = Format$(Par.a2(C))
        Next C
    End With
'
'
'
End Sub
