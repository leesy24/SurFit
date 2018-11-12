VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm3D 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " 3D isometric view"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frm3D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Begin VB.Frame zFrame2 
      Caption         =   "Rotation:"
      Height          =   615
      Left            =   2160
      TabIndex        =   23
      Top             =   0
      Width           =   3435
      Begin VB.CommandButton cmdPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "&Rotate"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.UpDown updTheta 
         Height          =   255
         Left            =   2895
         TabIndex        =   5
         Top             =   240
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   450
         _Version        =   393216
         BuddyControl    =   "lblTheta"
         BuddyDispid     =   196613
         OrigLeft        =   312
         OrigTop         =   12
         OrigRight       =   341
         OrigBottom      =   28
         Max             =   361
         Min             =   -1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.Label zLabel06 
         Alignment       =   1  'Right Justify
         Caption         =   "[Degree]:"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   270
         Width           =   915
      End
      Begin VB.Label lblTheta 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame zFrame1 
      Caption         =   "Views:"
      Height          =   615
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   1755
      Begin VB.CommandButton cmdZY 
         Caption         =   "Z - &Y"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Side view "
         Top             =   240
         Width           =   435
      End
      Begin VB.CommandButton cmdZX 
         Caption         =   "&Z - X"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   660
         TabIndex        =   1
         ToolTipText     =   "Front view "
         Top             =   240
         Width           =   435
      End
      Begin VB.CommandButton cmdXY 
         Caption         =   "&X - Y"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Plan view "
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   420
      Top             =   900
   End
   Begin VB.PictureBox pic3D 
      BackColor       =   &H00000000&
      Height          =   5700
      Left            =   240
      MouseIcon       =   "frm3D.frx":014A
      MousePointer    =   2  'Cross
      ScaleHeight     =   376
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   9030
      Begin VB.CommandButton cmdCopiaGrafico 
         Height          =   285
         Left            =   8340
         MaskColor       =   &H000000FF&
         MousePointer    =   1  'Arrow
         Picture         =   "frm3D.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Copy the image to the clipboard "
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton optBN 
         BackColor       =   &H00000000&
         Caption         =   "&B/N"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8280
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   60
         Width           =   615
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00000000&
         Caption         =   "&Col."
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8280
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Shape shpInd 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         Height          =   315
         Left            =   180
         Shape           =   3  'Circle
         Top             =   780
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.Label lblZ 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2820
      TabIndex        =   21
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label zLabel05 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Z:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2580
      TabIndex        =   20
      Top             =   6525
      Width           =   195
   End
   Begin VB.Label lblX 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   19
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label zLabel03 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   6525
      Width           =   195
   End
   Begin VB.Label lblY 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   17
      Top             =   6510
      Width           =   855
   End
   Begin VB.Label zLabel04 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   6525
      Width           =   195
   End
   Begin VB.Label zlblAutore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F. Languasco  fecit. 2001"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7620
      TabIndex        =   15
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Label lblIstruzioni 
      BackStyle       =   0  'Transparent
      Caption         =   "The view can be changed by moving (with the left mouse button pressed) to the end of the Y axis."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   14
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label zLabel01 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "[Degree]:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4980
      TabIndex        =   13
      Top             =   6525
      Width           =   1095
   End
   Begin VB.Label zLabel02 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RAyx:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3780
      TabIndex        =   12
      Top             =   6525
      Width           =   495
   End
   Begin VB.Label lblRAyx 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   11
      Top             =   6510
      Width           =   795
   End
   Begin VB.Label lblAlfa 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6120
      TabIndex        =   10
      Top             =   6510
      Width           =   795
   End
End
Attribute VB_Name = "frm3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
' Descrizione.....: Form, per la rappresentazione in 3D, di
'                   punti o di superfici (vista assonometrica).
' Nome dei Files..: frm3D.frm, fr3D.frx
'                   modUtilita.bas
' Data............: 10/12/2001
' Aggiornamento...: 1/2/2002  (aggiunta la rappresentazione a punti).
' Aggiornamento...: 17/3/2002 (sistemate alcune incongruenze di
'                              rappresentazione).
' Aggiornamento...: 21/3/2002 (aggiunta la rotazione).
' Versione........: 1.1 a 32 bits (provvisoria, in via di sviluppo).
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco 
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.virgilio.it/flanguasco/
'                   http://www.flanguasco.org
'===================================================================
'
'   Routines di ingresso:
'
'    frm3D.Points XD#(), YD#(), ZD#() [, Title$ = ""] [, IC& = 1] _
'                [, bAutoScale as Boolean = True] _
'                [, XMin#, XMax#, YMin#, YMax#, ZMin#, ZMax#]
'     XD():       vettore contenente le ascisse  dei punti da rappresentare.
'     YD():          "        "       " ordinate  "    "    "      ".
'     ZD():          "        "       " altezze   "    "    "      ".
'     Title$:    titolo del quadro (opzionale).
'     IC:         se IC <= 1 il Form viene messo in primo piano (opzionale).
'     bAutoScale: if False must also pass the values Min and Max to be used
'                  as extremes of the three axes.
'
'    frm3D.Surface XI#(), YI#(), ZI#() [, Title$ = ""] [, IC& = 1]
'     XI():     vettore contenente le ascisse  della superficie da rappresentare.
'     YI():        "        "       " ordinate   "       "       "      ".
'     ZI():     matrice     "       i livelli    "       "       "      ".
'     Title$:  titolo del quadro (opzionale).
'     IC:       se IC <= 1 il Form viene messo in primo piano (opzionale).
'
'   Nota:   Per Sub Points:
'            i vettori XD(1 to ND), YD(1 to ND) e ZD(1 To ND) devono avere
'            le stesse dimensioni; in caso contrario viene usata la dimensione
'            piu' piccola.
'           Per Sub Surface:
'            se le dimensioni dei vettori sono XI(1 to NXI) e YI(1 to NYI),
'            la matrice deve essere dimensionata come ZI(1 to NXI, 1 to NYI).
'
'   Nota:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1 (TCol() escluso).
'
Option Explicit
'
Dim XV#(), YV#()    ' Vettori delle ascisse e delle ordinate.
Dim ZV#()           ' Vettore delle altezze dei punti o
                    ' matrice dei livelli della superficie.
Dim Title$
'
Dim fPoints As Boolean  ' If True, draw points;
                        ' if False, draw a surface.
'
Dim NV&             ' N?di valori nei vettori XV(), YV() e ZV().
'
Dim NXV&, NYV&      ' N?di valori nei vettori XV(), YV() e
                    ' di colonne e righe nella matrice ZV().
'
Dim XMin!, XMax!    ' Valori minimi
Dim YMin!, YMax!    ' e massimi
Dim ZMin!, ZMax!    ' dei dati in ingresso.
'
Dim XRMin!, XRMax!  ' Valori minimi e massimi su gli assi del quadro:
Dim YRMin!, YRMax!  ' servono ad evitare, in questa particolare applicazione, che
Dim ZRMin!, ZRMax!  ' chiamate successive a Picture3D cambino le scale degli assi.
                    ' Inoltre il cambio vista pilotato dal Mouse, richiede i veri
                    ' valori di XRMin, XRMax e ZRMin.
Dim AsseX!          ' XRMax - XRMin.
Dim sUX$            ' Label of the units of the X axis.
Dim XEsp&           ' X scale reduction factor
Dim sUY$            ' Label of the units of the Y axis.
Dim YEsp&           ' Y scale reduction factor
Dim sUZ$            ' Label of the units of the Z axis.
Dim ZEsp&           ' Z scale reduction factor
'
Dim Ax!, Bx!        ' Coefficienti di
Dim Ay!, By!        ' conversione scale
Dim Az!, Bz!        ' da vbUser a Pixels.
'
Dim RAyx!           ' Rapporto lunghezza asse Y in [Pixels] su lunghezza asse X in [Pixels].
Dim ALFA!           ' Angolo dell' asse Y rispetto all' asse X [Rad].
Dim SinA!, CosA!    ' Sin(ALFA) e Cos(ALFA) calcolati solo dopo un cambio vista.
'
' Costanti e variabili per la rotazione:
Dim THETA!              ' Angolo corrente di rotazione [Radianti].
Const dth! = PI / 180#  ' Incremento dell' angolo di rotazione [Radianti].
Dim X0r#, Y0r#          ' Coordinate del centro di rotazione.
Dim SemiAsseX!          ' Misure degli assi del disegno.
Dim SemiAsseY!          '   "      "     "   "     "
Dim AsseZ!              '   "      "     "   "     "
Dim TrRotX!, TrRotY!    ' Coeff. di trasformazione delle coordinate durante la rotazione.
Dim bPause As Boolean   ' Flag for rotation paused.
'Const RFCL& = &H8000&   ' Colore di assi e griglie di riferimento durante la rotazione.
Const RFCL& = &H404080  ' Colore di assi e griglie di riferimento durante la rotazione.
'
Dim TCol&()         ' Tavola dei colori.
Const NTCol& = 1280 ' Numero di colori disponibili in TCol().
Dim ZCol&()         ' Vettore o matrice dei colori da usare.
'
Const Log10# = 2.30258509299405
Const RadToGrd# = 180# / PI ' Fattore di conversione da [Rad] a [Grd].
'
Private Type POINTAPI
     X As Long          ' [Pixels].
     Y As Long          '     "
End Type
'
' Variabili per la Sub DrawPoints:
Dim PRv() As POINTAPI       ' Vettore dei punti o matrice dei vertici dei
                            ' quadrilateri proiettati sul piano di rappresentazione
                            ' (e' usata anche dalle Subs DisegnaSup_BN e DrawSurface).
Const lRP& = 4              ' Raggio del disegno dei punti [Pixels].
'
' Variabili per la Sub DisegnaSup_BN:
Dim NPoli&                  ' N?di quadrilateri in una riga.
Dim lpPoint() As POINTAPI   ' Vettore dei vertici di una riga.
Dim lpVertici&()            ' Vettore del N?di vertici di ogni poligono.
'
' Variabili per la Sub DrawSurface:
Dim lpPoint_C() As POINTAPI ' Vettore dei vertici di un quadrilatero.
'
' Costanti per la ricerca della posizione 3D:
Const shpIndOffx& = lRP + 2 ' Offset orizzontale e verticale del cerchio
Const shpIndOffy& = lRP + 2 ' di evidenziazione.
Const PCHL& = &HC0FFFF      ' Colore di evidenza per le etichette di posizione cursore.
'
Dim bRotate As Boolean      ' Flag for Rotation in progress.
Dim bLoaded As Boolean      ' Flag di Form inizializzato.
'
'-------------------------------------------------------------------------------------
'   API grafiche:
'
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, _
    ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
    lpPoint As POINTAPI, ByVal nCount As Long) As Long
'
Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, _
    lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
'
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
    ByVal nWidth As Long, ByVal crColor As Long) As Long
'
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
'
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
    lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" _
    (ByVal X As Long, ByVal Y As Long) As Long

Private Sub Rotate(ByVal X0#, ByVal Y0#, ByVal Rot#, ByVal CRx#, ByVal CRy#, Xr#, Yr#)
'
'   Routines for rotating a point:
'
'   X0, Y0:     coordinates of the point to rotate.
'   Rot:        rotation of the point in [radians].
'   CRx, CRy:   coordinates of the rotation center.
'   Xr, Yr:     final coordinates of the rotated point.
'
    ' Rotation:
    Xr = (X0 - CRx) * Cos(Rot) - (Y0 - CRy) * Sin(Rot) + CRx
    Yr = (X0 - CRx) * Sin(Rot) + (Y0 - CRy) * Cos(Rot) + CRy
'
'
'
End Sub

Private Function Picture3D(ByVal Foglio As PictureBox, _
    ByRef X0!, ByRef Xn!, ByRef Y0!, ByRef Yn!, ByRef Z0!, ByRef Zn!, _
    Optional ByVal ALFA! = PI / 6!, Optional ByRef RAyx! = 1!, _
    Optional ByRef Ax!, Optional ByRef Bx!, _
    Optional ByRef Ay!, Optional ByRef By!, _
    Optional ByRef Az!, Optional ByRef Bz!, _
    Optional ByVal FormatVX$ = "#0.0##", _
    Optional ByVal FormatVY$ = "#0.0##", _
    Optional ByVal FormatVZ$ = "#0.0##", _
    Optional ByVal Npx& = 1, Optional ByRef PxN_X!, Optional ByRef PxN_Z!, _
    Optional ByVal Title$ = "", _
    Optional ByVal UnitaX$ = "", _
    Optional ByVal UnitaY$ = "", _
    Optional ByVal UnitaZ$ = "", _
    Optional ByVal RifCol& = vbGreen, _
    Optional ByVal AutoRed As Boolean = False) As Boolean
'
'   Routine, of general use, for the preparation of a sheet suitable
'   to represent, in axonometry, a graph z = f (x, y).
'    Input parameters:
'     Foglio:    PictureBox to climb.
'     X0:        Minimum value of the abscissa to be represented.
'     Xn:        Maximum abscissa value to be represented.
'                It must be X0 < Xn.
'     Y0:        Minimum value of the ordinate to be represented.
'     Yn:        Maximum value of the ordinate to be represented.
'                It must be Y0 < Yn.
'     Z0:        Minimum value of elevation to be represented.
'     Zn:        Maximum value of elevation to be represented.
'                It must be Z0 <= Zn.
'     Alfa:      Angle formed by the Y axis with the X axis [Rad].
'     RAyx:      Ratio between the length of the Y axis and that of the X axis.
'                It must be RAyx <= 1000.
'     FormatVX$: Value format string on the X axis.
'     FormatVY$: Value format string on the Y axis.
'     FormatVZ$: Value format string on the Z axis.
'     Npx:       Number of Pixels whose width and height
'                you want to know in [vbUser].
'     Title$:    Title of the graph.
'     UnitaX$:   Unit (or title) of the X axis.
'     UnitaY$:   Unit (or title) of the Y axis.
'     UnitaZ$:   Unit (or title) of the Z axis.
'     RifCol:    Color of the reference axes and grids.
'     AutoRed:   State of Foglio.AutoRedraw after drawing the painting.
'    Output parameters:
'     X0:        Minimum abscissa value represented.
'     Xn:        Maximum abscissa value represented.
'     Y0:        Minimum ordinate value shown.
'     Yn:        Maximum value of the ordinate shown.
'     Z0:        Minimum value of elevation represented.
'     Zn:        Maximum value of elevation shown.
'     RAyx:      Ratio used between the length of the Y axis and that of the X axis.
'     Ax, Bx:    Transformation coefficients from the vbUser scale,
'     Ay, By:     set by this routine, to the "Logical Coordinates"
'     Az, Bz:     required by the graphic APIs.
'     PxN_X:     Width in [vbUser] and height
'     PxN_Z:      in [vbUser] of Npx [Pixels].
'
    Dim I&, XI!, rrx!, YI!, D_Y!, rry!, ZI!, rrz!
    Dim CosA!, SinA!, Ryx!, LAx!, LAy!, LAz$
    Dim QxMin!, QxMax!, QzMin!, QzMax!
    Dim Px1_X!, Px1_Z!, TxWs!, TxWd!, TxH!, Tx$
    Dim bVlX As Boolean, bVlY As Boolean, bVlZ As Boolean
    
    Dim t0!
    t0 = Timer
    
'
    On Error GoTo Picture3D_ERR
    ' Check the correctness of the stairs:
    If X0 >= Xn Then Err.Raise 1001, "Picture3D", " Error scale X."
    If Y0 >= Yn Then Err.Raise 1001, "Picture3D", " Error scale Y."
    If Z0 > Zn Then Err.Raise 1001, "Picture3D", " Error scale Z."
'
'-------------------------------------------------------------------------------------
'   Calculation of the grating step of the three axes.
'
    Dim DZMin!                      ' Ampiezza min. della scala Z.
    Const Log10! = 2.30258509299405 ' Log(10#)
'
    ' Calculates the spacing of the values written on the X axis:
    '  the sequence is 1, 2, 2.5 and 5:
    LAx = Xn - X0
    rrx = 10! ^ Ceil(Log(LAx / 20!) / Log10)
    Do While LAx / rrx < 5!
        rrx = rrx / 2!
    Loop
    If LAx / rrx > 10! Then rrx = rrx * 2!
    X0 = rrx * Int(Round(X0 / rrx, 3))
    Xn = rrx * Ceil(Round(Xn / rrx, 3))
    LAx = Xn - X0
'
    ' Calculates the spacing of the values written on the Y axis:
    '  the sequence is 1, 2, 2.5 and 5:
    D_Y = Yn - Y0
    rry = 10! ^ Ceil(Log(D_Y / 20!) / Log10)
    Do While D_Y / rry < 5!
        rry = rry / 2!
    Loop
    If D_Y / rry > 10! Then rry = rry * 2!
    Y0 = rry * Int(Round(Y0 / rry, 3))
    Yn = rry * Ceil(Round(Yn / rry, 3))
    D_Y = Yn - Y0
    If RAyx > 1000! Then RAyx = 1000!
    LAy = RAyx * LAx
'
    ' Set a minimum scale for the Z axis:
    DZMin! = 0.0001
    If (Zn - Z0) < DZMin Then
        Do While (Z0 - DZMin / 20!) = Z0
            DZMin = 2! * DZMin
        Loop
'
        Z0 = Z0 - DZMin
        Zn = Zn + DZMin
    End If
'
    ' Calculates the spacing of the values written on the Y axis:
    '  the sequence is 1, 2, 2.5 and 5:
    LAz = Zn - Z0
    rrz = 10! ^ Ceil(Log(LAz / 20!) / Log10)
    Do While LAz / rrz < 5!
        rrz = rrz / 2!
    Loop
    If LAz / rrz > 10! Then rrz = rrz * 2!
    Z0 = rrz * Int(Round(Z0 / rrz, 3))
    Zn = rrz * Ceil(Round(Zn / rrz, 3))
    LAz = Zn - Z0
'
'-------------------------------------------------------------------------------------
'   Calculate the width and height of the edges.
'
    Dim Bl!, Br!, BB!, Bt!, BDen!
    Dim DT_X!, DT_Z!, DPz!, DDz!, TxHt!, TxHb!
'
    ' Set Font data of axis and title values:
    Foglio.FontName = "MS Sans Serif"
    Foglio.FontBold = False
'
    CosA = Cos(ALFA)
    SinA = Sin(ALFA)
'
    'Cancel the previous stairs:
    Foglio.ScaleMode = vbPixels
'
    ' The left edge must be sufficient to contain the largest Z value:
    Foglio.FontSize = 8
    TxWs = MAX0(Foglio.TextWidth(Format$(-Abs(Z0), FormatVZ$) & "W"), _
               Foglio.TextWidth(Format$(-Abs(Zn), FormatVZ$) & "W"), _
               Foglio.TextWidth(UnitaZ$ & "W"))
'
    ' The border on the right must be sufficient to contain
    '  the Xn value and the UnitaX $ label:
    TxWd = Foglio.TextWidth(Format$(-Abs(Xn), FormatVZ$) & "W") _
         + Foglio.TextWidth(UnitaX$ & "W")
'
    ' The borders on the left and on the right are:
    DT_X = LAx * (1! + RAyx * CosA)
    BDen = DT_X / (Foglio.ScaleWidth - TxWs - TxWd)
    Bl = TxWs * BDen
    Br = TxWd * BDen
'
    ' The border below is 2 times the height of the values:
    TxHb = 2! * Abs(Foglio.TextHeight("W"))
'
    ' The edge above is 2 times the height of the values plus 2 times the height of the title:
    Foglio.FontSize = 12
    TxHt = TxHb + 2! * Abs(Foglio.TextHeight(Title$))
'
    ' The edges above and below are:
    DDz = Foglio.ScaleWidth * LAx * RAyx * SinA / (Bl + DT_X + Br)
    DPz = Abs(Foglio.ScaleHeight) - DDz - TxHt - TxHb
    If DPz <= 0 Then DPz = 0.001
    DT_Z = LAz * (1! + DDz / DPz)
    BDen = DT_Z / (Abs(Foglio.ScaleHeight) + TxHb + TxHt)
    BB = TxHb * BDen
    Bt = TxHt * BDen
'
'-------------------------------------------------------------------------------------
'   Set the scale and calculate the common values.
'
    Dim TaccheX!, TaccheZ!  ' Length of the notches on the axes.
    Dim EstAx!, EstAz!      ' Extension of the X and Z axes.
    Dim LyCosA!, LySinA!    ' Projections of the Y axis.
'
    ' Set horizontal and vertical borders:
    QxMin = X0 - Bl
    QxMax = X0 + DT_X + Br
    QzMin = Z0 - BB
    QzMax = Z0 + DT_Z + Bt
'
    ' Set the scale and delete the Sheet:
    'Foglio.Picture = LoadPicture("")
    Foglio.Scale (QxMin, QzMax)-(QxMax, QzMin)
    Foglio.Line (QxMin, QzMin)-(QxMax, QzMax), Foglio.BackColor, BF ' This is faster than
                                                                    '  Foglio.Cls.
    ' The drawing of the painting must be permanent:
    Foglio.AutoRedraw = True
'
    ' Width and height of 1 pixel in [vbUser]:
    Px1_X = Abs(Foglio.ScaleX(1, vbPixels, vbUser))
    Px1_Z = Abs(Foglio.ScaleY(1, vbPixels, vbUser))
    Ryx = Px1_Z / Px1_X ' Y/X scale ratio.
'
    ' Calculate width and height of Npx pixels in [vbUser]:
    PxN_X = Npx * Px1_X
    PxN_Z = Npx * Px1_Z
'
    ' Precalculation of some frequently used values:
    TaccheX = 4! * Px1_X
    TaccheZ = 4! * Px1_Z
    EstAx = 15! * Px1_X
    EstAz = 15! * Px1_Z
    LyCosA = LAy * CosA
    LySinA = Ryx * LAy * SinA
'
    If Foglio.ScaleY(LAz, vbUser, vbPixels) > 0 Then
    End If
'
'-------------------------------------------------------------------------------------
'   Draw axes, grids and write scale values.
'
    Foglio.FontSize = 8
    Foglio.DrawWidth = 1
    Foglio.ForeColor = RifCol
    Foglio.DrawMode = vbCopyPen
'
    ' Check the separation of the labels:
    Dim TxW!
    TxW = DMAX1(Foglio.TextWidth(Format$(X0, FormatVX$)), _
               Foglio.TextWidth(Format$(Xn, FormatVX$)))
    bVlX = (LAx / rrx) * TxW < LAx
'
    TxW = DMAX1(Foglio.TextWidth(Format$(Y0, FormatVY$)), _
               Foglio.TextWidth(Format$(Yn, FormatVY$)))
    TxH = Abs(Foglio.TextHeight("W"))
    bVlY = ((Yn - Y0) / rry) * TxW < LyCosA
    bVlY = bVlY Or (((Yn - Y0) / rry) * TxH < LySinA)
'
    bVlZ = (LAz / rrz) * TxH < LAz
'
    ' Draw the X axis:
    Foglio.DrawStyle = vbSolid
    Foglio.Line (X0, Z0)-(Xn + EstAx, Z0)
    If bVlX Then
        Foglio.Line (Xn + EstAx, Z0) _
                   -(Xn + EstAx - TaccheX, Z0 + TaccheZ / 2!)
        Foglio.Line (Xn + EstAx, Z0) _
                   -(Xn + EstAx - TaccheX, Z0 - TaccheZ / 2!)
        ' and write the label of the X axis:
        If Len(UnitaX$) > 0 Then
            Foglio.CurrentX = Xn + Foglio.TextWidth(Xn & "W")
            Foglio.Print UnitaX$;
        End If
    End If
'
    ' Draw the Y axis:
    Foglio.Line (X0, Z0)-(X0 + (LAy + EstAx) * CosA, _
                          Z0 + (LAy + EstAx) * SinA * Ryx)
    If bVlY Then
        ' and write the Y axis label:
        If Len(UnitaY$) > 0 Then
            Foglio.CurrentY = Foglio.CurrentY - Foglio.TextHeight("W")
            Foglio.Print UnitaY$;
        End If
    End If
'
    ' Draw the Z axis:
    Foglio.Line (X0, Z0)-(X0, Zn + EstAz)
    If bVlZ Then
        Foglio.Line (X0, Zn + EstAz) _
                   -(X0 - TaccheX / 2!, Zn + EstAz - TaccheZ)
        Foglio.Line (X0, Zn + EstAz) _
                   -(X0 + TaccheX / 2!, Zn + EstAz - TaccheZ)
        ' and write the Z axis label:
        If Len(UnitaZ$) > 0 Then
            Foglio.CurrentX = QxMin
            Foglio.CurrentY = Zn + EstAz - Foglio.TextHeight("W") / 2!
            Foglio.Print UnitaZ$;
        End If
    End If
'
    ' Draw the vertical grid on the Z-X plane,
    '  the plane on the X-Y plane and write the values of the X axis:
    Dim rrx_10!
'
    rrx_10 = rrx / 10!
    Foglio.DrawStyle = vbDot
    For XI = X0 To Xn + rrx_10 Step rrx
        Foglio.Line (XI + LyCosA, Z0 + LySinA + LAz)-(XI + LyCosA, Z0 + LySinA)
        Foglio.Line -(XI, Z0)
        Foglio.Line -(XI, Z0 - TaccheZ)
        If bVlX Then
            Tx$ = Format$(XI, FormatVX$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(XI - Val(Tx$)) < rrx_10) Then
                Foglio.CurrentX = XI - Foglio.TextWidth(Tx$) / 2!
                Foglio.Print Tx$;
            End If
        End If
    Next XI
'
    ' Draw the horizontal grid on the X-Y plane,
    '  the vertical grid on the Z-Y plane and
    '  write the values of the Y axis:
    Dim LyCosA_Y!, LySinA_Y!, Yx!, Yz!, rry_10!
'
    LyCosA_Y = LyCosA / D_Y
    LySinA_Y = LySinA / D_Y
    rry_10 = rry / 10!
    For YI = Y0 To Yn + rry_10 Step rry
        Yx = Xn + (YI - Y0) * LyCosA_Y
        Yz = Z0 + (YI - Y0) * LySinA_Y
        Foglio.Line (Yx, Yz)-(Yx - LAx, Yz)
        Foglio.Line -(Foglio.CurrentX, Yz + LAz)
        If bVlY Then
            Tx$ = Format$(YI, FormatVY$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(YI - Val(Tx$)) < rry_10) Then
                ' The positions of the Y labels depend on
                '  the presence of those Z:
                If bVlZ Then
                    Foglio.Line -(Foglio.CurrentX, Foglio.CurrentY + EstAz)
                    Foglio.CurrentX = Foglio.CurrentX - Foglio.TextWidth(Tx$) / 2!
                    Foglio.CurrentY = Foglio.CurrentY - Foglio.TextHeight(Tx$)
                Else
                    Foglio.Line -(Foglio.CurrentX - TaccheX, Foglio.CurrentY)
                    Foglio.CurrentX = Foglio.CurrentX - Foglio.TextWidth(Tx$)
                    Foglio.CurrentY = Foglio.CurrentY - Foglio.TextHeight(Tx$) / 2!
                End If
                Foglio.Print Tx$;
            End If
        End If
    Next YI
'
    ' Draw the horizontal grid on the Z-Y plane,
    '  the horizontal grid on the Z-X plane and
    '  write the values of the Z axis:
    Dim rrz_10!
'
    rrz_10 = rrz / 10!
    For ZI = Z0 To Zn + rrz_10 Step rrz
        Foglio.Line (X0 - TaccheX, ZI)-(X0, ZI)
        Foglio.Line -(X0 + LyCosA, ZI + LySinA)
        Foglio.Line -(Foglio.CurrentX + LAx, Foglio.CurrentY)
        If bVlZ Then
            Tx$ = Format$(ZI, FormatVZ$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(ZI - Val(Tx$)) < rrz_10) Then
                Foglio.CurrentX = QxMin
                Foglio.CurrentY = ZI - Foglio.TextHeight(Tx$) / 2!
                Foglio.Print Tx$;
            End If
        End If
    Next ZI
'
'-------------------------------------------------------------------------------------
'   Write the chart title:
'
    Dim TitL!, TitT!, TitW!, TitH!
'
    If Len(Title$) > 0 Then
        Foglio.FontSize = 12
        Foglio.FontBold = True
        Foglio.ForeColor = vbRed
'
        TitW = Foglio.TextWidth(Title$)
        TitH = Foglio.TextHeight(Title$)
        ' Verify that the title is all in the Sheet:
        If TitW <= Foglio.ScaleWidth Then
            TitL = (QxMin + QxMax - TitW) / 2!
        Else
            ' and if not, cut it:
            TitL = Foglio.ScaleLeft
            Tx$ = " . . . ."
            Title$ = Left$(Title$, Int(Len(Title$) * _
            (Foglio.ScaleWidth - Foglio.TextWidth(Tx$)) / TitW)) & Tx$
        End If
        TitT = QzMax
        ' Delete the area on which the title will be written:
        'Foglio.Line (TitL, TitT)-(TitL + TitW, TitT + TitH), Foglio.BackColor, BF
        Foglio.CurrentX = TitL
        Foglio.CurrentY = TitT
        Foglio.Print Title$
    End If
'
'-------------------------------------------------------------------------------------
'   Calculates the transformation coefficients from vbUser to Pixels:
'
    Dim C0_Px!, Cn_Px!
'
    C0_Px = Foglio.ScaleX(X0 - Foglio.ScaleLeft, vbUser, vbPixels)
    Cn_Px = Foglio.ScaleX(Xn - Foglio.ScaleLeft, vbUser, vbPixels)
    Ax = (Cn_Px - C0_Px) / LAx
    Bx = (C0_Px * Xn - Cn_Px * X0) / LAx
'
    Ay = Foglio.ScaleX(LAy / D_Y, vbUser, vbPixels)
    By = -Ay * Y0
'
    C0_Px = Foglio.ScaleY(Z0 - Foglio.ScaleTop, vbUser, vbPixels)
    Cn_Px = Foglio.ScaleY(Zn - Foglio.ScaleTop, vbUser, vbPixels)
    Az = (Cn_Px - C0_Px) / LAz
    Bz = (C0_Px * Zn - Cn_Px * Z0) / LAz
'
    ' And leave the Sheet set:
    Foglio.DrawStyle = vbSolid
    Foglio.AutoRedraw = AutoRed
'
'
Picture3D_ERR:
    Picture3D = (Err = 0)
    If Err <> 0 Then
        MsgBox Err.Description, vbCritical, " Picture3D/" & Err.Source
    End If
'
'
'
End Function

Public Sub Points(XD_I#(), YD_I#(), ZD_I#(), Optional ByVal Title_I$ = "", _
    Optional ByVal bAutoScale As Boolean = True, _
    Optional ByVal XMin_I#, Optional ByVal XMAx_I#, _
    Optional ByVal YMin_I#, Optional ByVal YMAx_I#, _
    Optional ByVal ZMin_I#, Optional ByVal ZMax_I#)
'
'   Input routines for representing points in 3D space:
'
    If (Not bLoaded) Then
        Me.Show
        Me.ZOrder vbBringToFront
    End If
'
    XV() = XD_I()
    YV() = YD_I()
    ZV() = ZD_I()
    Title$ = Title_I$
    If (Not bAutoScale) Then
        XMin = XMin_I
        XMax = XMAx_I
        YMin = YMin_I
        YMax = YMAx_I
        ZMin = ZMin_I
        ZMax = ZMax_I
    End If
'
    fPoints = True
    Settings bAutoScale
'
    Draw True
    MeasureSpace3D
'
'
'
End Sub

Public Sub Surface(XI_I#(), YI_I#(), ZI_I#(), Optional ByVal Title_I$ = "", _
    Optional ByVal bAutoScale As Boolean = True, _
    Optional ByVal XMin_I#, Optional ByVal XMAx_I#, _
    Optional ByVal YMin_I#, Optional ByVal YMAx_I#, _
    Optional ByVal ZMin_I#, Optional ByVal ZMax_I#)
'
'   Input routines for the 3D representation of a surface:
'
    On Error GoTo 0
    If (Not bLoaded) Then
        Me.Show
        Me.ZOrder vbBringToFront
    End If
'
    XV() = XI_I()
    YV() = YI_I()
    ZV() = ZI_I()
    Title$ = Title_I$
    If (Not bAutoScale) Then
        XMin = XMin_I
        XMax = XMAx_I
        YMin = YMin_I
        YMax = YMAx_I
        ZMin = ZMin_I
        ZMax = ZMax_I
    End If
'
    fPoints = False
    Settings bAutoScale
'
    Draw True
    MeasureSpace3D
'
'
'
End Sub

Private Sub DrawSurface(ByVal bCol As Boolean)
'
'   Disegna, con API, i quadrilateri a colori o in B/N.  Disegnando le righe
'   dall' ultima indietro fino alla prima (i.e. muovendosi verso l' osservatore),
'   ogni quadrilatero nasconde la parte di immagine da lui coperta.
'   Se bCol = True ai quadrilateri viene assegnato un colore proporzionale al
'   valor medio delle coordinate Z dei quattro vertici.
'
'   Schema del quadrilatero utilizzato dal vettore lpPoint_C():
'    2 _____ 3      Vertice 1 -> lpPoint_C(1) = PRv(I, J)
'     |     |       Vertice 2 -> lpPoint_C(2) = PRv(I, J + 1)
'     |     |       Vertice 3 -> lpPoint_C(3) = PRv(I + 1, J + 1)
'     |_____|       Vertice 4 -> lpPoint_C(4) = PRv(I + 1, J)
'    1       4      con: 1 <= I <= NXV - 1, 1 <= J <= NYV - 1
'
    Dim I&, J&, hPen&, hPen_O&, hBrush&, hBrush_O&, lR1&
    Dim Xr#, Yr#, Quadrante&
'
    hPen = CreatePen(vbSolid, 1, vbWhite)
    hPen_O = SelectObject(pic3D.hdc, hPen)
    hBrush = CreateSolidBrush(&H808080)         ' Imposta il colore dei
    hBrush_O = SelectObject(pic3D.hdc, hBrush)  ' quadrilateri per disegno in B/N.
'
    ' Calcolo la matrice dei vertici dei quadrilateri
    ' proiettati sul piano di rappresentazione:
    For J = 1 To NYV
        For I = 1 To NXV
            If bRotate Then
                Rotate XV(I), YV(J), CDbl(THETA), X0r, Y0r, Xr, Yr
                PRv(I, J).X = CLng((Ax * Xr + Bx) + (Ay * Yr + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * Yr + By) * SinA)
            Else
                PRv(I, J).X = CLng((Ax * XV(I) + Bx) + (Ay * YV(J) + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * YV(J) + By) * SinA)
            End If
        Next I
    Next J
'
    ' Disegno i quadrilateri. In funzione di
    ' THETA disegno per primi quelli piu'
    ' lontani dall' osservatore:
    Quadrante = CLng(Int(THETA / PI_2))
'
    Select Case Quadrante
        Case 0  ' 1?quadrante.
        For J = NYV - 1 To 1 Step -1
            For I = 1 To NXV - 1
                lpPoint_C(1).X = PRv(I, J).X
                lpPoint_C(1).Y = PRv(I, J).Y
                lpPoint_C(2).X = PRv(I, J + 1).X
                lpPoint_C(2).Y = PRv(I, J + 1).Y
                lpPoint_C(3).X = PRv(I + 1, J + 1).X
                lpPoint_C(3).Y = PRv(I + 1, J + 1).Y
                lpPoint_C(4).X = PRv(I + 1, J).X
                lpPoint_C(4).Y = PRv(I + 1, J).Y
'
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(I, J))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Polygon pic3D.hdc, lpPoint_C(1), 4
            Next I
        Next J
'
        Case 1  ' 2?quadrante.
        For I = NXV - 1 To 1 Step -1
            For J = NYV - 1 To 1 Step -1
                lpPoint_C(1).X = PRv(I, J).X
                lpPoint_C(1).Y = PRv(I, J).Y
                lpPoint_C(2).X = PRv(I, J + 1).X
                lpPoint_C(2).Y = PRv(I, J + 1).Y
                lpPoint_C(3).X = PRv(I + 1, J + 1).X
                lpPoint_C(3).Y = PRv(I + 1, J + 1).Y
                lpPoint_C(4).X = PRv(I + 1, J).X
                lpPoint_C(4).Y = PRv(I + 1, J).Y
'
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(I, J))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Polygon pic3D.hdc, lpPoint_C(1), 4
            Next J
        Next I
'
        Case 2  ' 3?quadrante.
        For J = 1 To NYV - 1
            For I = NXV - 1 To 1 Step -1
                lpPoint_C(1).X = PRv(I, J).X
                lpPoint_C(1).Y = PRv(I, J).Y
                lpPoint_C(2).X = PRv(I, J + 1).X
                lpPoint_C(2).Y = PRv(I, J + 1).Y
                lpPoint_C(3).X = PRv(I + 1, J + 1).X
                lpPoint_C(3).Y = PRv(I + 1, J + 1).Y
                lpPoint_C(4).X = PRv(I + 1, J).X
                lpPoint_C(4).Y = PRv(I + 1, J).Y
'
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(I, J))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Polygon pic3D.hdc, lpPoint_C(1), 4
            Next I
        Next J
'
        Case 3  ' 4?quadrante.
        For I = 1 To NXV - 1
            For J = 1 To NYV - 1
                lpPoint_C(1).X = PRv(I, J).X
                lpPoint_C(1).Y = PRv(I, J).Y
                lpPoint_C(2).X = PRv(I, J + 1).X
                lpPoint_C(2).Y = PRv(I, J + 1).Y
                lpPoint_C(3).X = PRv(I + 1, J + 1).X
                lpPoint_C(3).Y = PRv(I + 1, J + 1).Y
                lpPoint_C(4).X = PRv(I + 1, J).X
                lpPoint_C(4).Y = PRv(I + 1, J).Y
'
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(I, J))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Polygon pic3D.hdc, lpPoint_C(1), 4
            Next J
        Next I
    End Select
'
    lR1 = SelectObject(pic3D.hdc, hPen_O)
    lR1 = SelectObject(pic3D.hdc, hBrush_O)
    lR1 = DeleteObject(hPen)
    lR1 = DeleteObject(hBrush)
'
'
'
End Sub
Private Sub DisegnaSup_BN()
'
'   NON USATA.
'
'   Disegna, con API, le righe di quadrilateri.  Disegnando le righe dall' ultima
'   indietro fino alla prima (i.e. muovendosi verso l' osservatore), ogni quadri-
'   latero nasconde la parte di immagine da lui coperta.
'
'   Schema dei quadrilateri utilizzati dal vettore lpPoint():
'  2 _____ 3  6 _____ 7
'   |     |    |     | .... In tutte le righe coincidono i vertici 3 con 6,
'   |     |    |     | .... 4 con 5, 7 con 10, 8 con 9, etc...
'   |_____|    |_____| .... Nelle righe J < NYV coincidono i vertici
'  1       4  5       8     2 con 1 della riga J + 1, 4 con 3 della riga J + 1 , etc...
'
'   Questa routine e' riportata solo per curiosita': La Sub DrawSurface,
'   infatti, offre la stessa funzionalita'.  La curiosita' sta' nel fatto che,
'   usando l' API PolyPolygon, i poligoni successivi, definiti in lpPoint(),
'   NON nascondono quelli precedentemente definiti nello stesso vettore.
'   Disegnando i quadrilateri riga per riga l' effetto non e' molto evidente:
'   se invece si volessero mettere in  lpPoint() TUTTI i quadrilateri che
'   formano la superficie il risultato sarebbe disastroso.
'
    Dim I&, J&, NP&, hPen&, hPen_O&, hBrush&, hBrush_O&, lR1&
    Dim Xr#, Yr#
'
    hPen = CreatePen(vbSolid, 1, vbWhite)
    hPen_O = SelectObject(pic3D.hdc, hPen)
    hBrush = CreateSolidBrush(&H808080) 'pic3D.BackColor)
    hBrush_O = SelectObject(pic3D.hdc, hBrush)
'
    ' Calcolo la matrice dei vertici dei quadrilateri
    ' proiettati sul piano di rappresentazione:
    For J = 1 To NYV
        For I = 1 To NXV
            If bRotate Then
                Rotate XV(I), YV(J), THETA, X0r, Y0r, Xr, Yr
                PRv(I, J).X = CLng((Ax * X0r + Bx) + (Ay * Y0r + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * Y0r + By) * SinA)
            Else
            PRv(I, J).X = CLng((Ax * XV(I) + Bx) + (Ay * YV(J) + By) * CosA)
            PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * YV(J) + By) * SinA)
            End If
        Next I
    Next J
'
    ' Disegno le righe di quadrilateri:
    For J = NYV - 1 To 1 Step -1
        For I = 1 To NXV - 1
            NP = 4 * (I - 1)
            lpPoint(NP + 1).X = PRv(I, J).X
            lpPoint(NP + 1).Y = PRv(I, J).Y
            lpPoint(NP + 2).X = PRv(I, J + 1).X
            lpPoint(NP + 2).Y = PRv(I, J + 1).Y
            lpPoint(NP + 3).X = PRv(I + 1, J + 1).X
            lpPoint(NP + 3).Y = PRv(I + 1, J + 1).Y
            lpPoint(NP + 4).X = PRv(I + 1, J).X
            lpPoint(NP + 4).Y = PRv(I + 1, J).Y
        Next I
        PolyPolygon pic3D.hdc, lpPoint(1), lpVertici(1), NPoli
    Next J
'
    lR1 = SelectObject(pic3D.hdc, hPen_O)
    lR1 = SelectObject(pic3D.hdc, hBrush_O)
    lR1 = DeleteObject(hPen)
    lR1 = DeleteObject(hBrush)
'
'
'
End Sub
Private Sub cmdCopiaGrafico_Click()
'
'
    Clipboard.Clear
    Clipboard.SetData pic3D.Image, vbCFDIB
'
'
'
End Sub

Private Sub cmdPause_Click()
'
'
    bPause = Not bPause
'
    cmdPause.Caption = IIf(bPause, "Co&nt.", "&Pause")
    cmdRotate.Enabled = Not bPause
    updTheta.Enabled = bPause
'
    Timer1.Enabled = Not bPause
'
'
'
End Sub

Private Sub cmdRotate_Click()
'
'
    bRotate = Not bRotate
    cmdRotate.Caption = IIf(bRotate, "&Stop", "&Rotate")
'
    THETA = 0!
    lblTheta = Format(RadToGrd * THETA, "#0.0")
'
    cmdPause.Enabled = bRotate
    shpInd.Visible = False
    UpdateCursorPositions lblX, "", lblY, "", lblZ, ""
'
    If bRotate Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
        Draw
    End If
'
'
'
End Sub

Private Sub cmdXY_Click()
'
'   Imposta la vista in pianta:
'
    Dim DT_X!, DT_Xp&
'
    ' DT_X e' la larghezza totale del grafico meno
    ' la larghezza dei bordi. Per ALFA = PI / 2 e'
    ' anche la lunghezza dell' asse X:
    DT_X = AsseX * (1! + RAyx * Cos(ALFA))
    DT_Xp = pic3D.ScaleX(DT_X, vbUser, vbPixels)
'
    ' Con questo rapporto l' asse Y sara' l' 80%
    ' dell' altezza totale del grafico:
    RAyx = 0.8 * pic3D.Height / CSng(DT_Xp)
    ALFA = CSng(PI / 2#)
'
    lblRAyx = Format$(RAyx, "#0.000")
    lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    Draw True
'
'
'
End Sub
Private Sub cmdZX_Click()
'
'
    RAyx = 0!
    ALFA = CSng(PI / 2#)
'
    lblRAyx = Format$(RAyx, "#0.000")
    lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    Draw True
'
'
'
End Sub
Private Sub cmdZY_Click()
'
'
    RAyx = 1000!
    ALFA = 0!
'
    lblRAyx = Format$(RAyx, "#0.000")
    lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    Draw True
'
'
'
End Sub

Private Sub Form_Load()
'
'
    Dim Me_L&, Me_T&
'
    LoadPositionForm Me, Me_L, Me_T
    Me.Move Me_L, Me_T ', Me_W, Me_H
'
    TCol() = ColorTable(NTCol)
'
    ' Highlight circle size:
    shpInd.Width = 2 * shpIndOffx
    shpInd.Height = 2 * shpIndOffy
'
    ' Initial values:
    RAyx = 1!       ' Y axis length ratio to X axis length.
    ALFA = PI / 6!  ' Angle of the Y axis with respect to the X axis: 30 [Grd].
'
    lblRAyx = Format$(RAyx, "#0.000")
    lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    bRotate = False
    bPause = False
    lblTheta = "0.0"
'
    bLoaded = True
'
'
'
End Sub

Private Sub Settings(Optional ByVal bAutoScale As Boolean = True)
'
'   Find and calculate commonly used settings and variables:
'
    Dim I&, J&, N&, ZnCol!, ZMed!
    Dim AMax#, Rid#
'
    If bAutoScale Then
        ' Find the minimum and maximum values of the abscissas:
        XMin = DMINVAL(XV())
        XMax = DMAXVAL(XV())
'
        ' and of the ordinates:
        YMin = DMINVAL(YV())
        YMax = DMAXVAL(YV())
    End If
'
    If fPoints Then
        ' Point design settings:
        NV = MIN0(UBound(XV), UBound(YV), UBound(ZV))
        ReDim Preserve XV(1 To NV)  ' Resize vectors
        ReDim Preserve YV(1 To NV)  '  all to the
        ReDim Preserve ZV(1 To NV)  '  same length.
        ReDim PRv(1 To NV)
'
        ' Sorts the vectors so that the points with major Y remain behind:
        QuickSort3V YV(), XV(), ZV(), 1, NV
'
        If bAutoScale Then
            ' Find the minimum and maximum values of heights:
            ZMin = DMINVAL(ZV())
            ZMax = DMAXVAL(ZV())
        End If
'
        ' Step between colors:
        If (ZMax > ZMin) Then
            ZnCol = CSng(NTCol - 1) / (ZMax - ZMin)
        End If
'
        ' Prepare the color vector to be assigned to the points:
        ReDim ZCol(1 To NV)
        For N = 1 To NV
            ' Calculation of the corresponding color:
            ZCol(N) = TCol(CLng((ZV(N) - ZMin) * ZnCol))
        Next N
'
    Else
        ' Settings for drawing a surface:
        NXV = UBound(XV)
        NYV = UBound(YV)
'
        If bAutoScale Then
            ' Find the minimum and maximum levels values:
            ZMin = ZV(1, 1)
            ZMax = ZV(1, 1)
            For J = 1 To NYV
                For I = 1 To NXV
                    If ZMin > ZV(I, J) Then ZMin = ZV(I, J)
                    If ZMax < ZV(I, J) Then ZMax = ZV(I, J)
                Next I
            Next J
        End If
        
        AMax = DMAX1(Abs(XMin), Abs(XMax))
        If AMax > 1000# Then
            ' Reduces the scale of X values:
            XEsp = 3 * Int((Log(AMax) / Log10) / 3#)
            sUX$ = "x[10^" & XEsp & "]"
            Rid = 10# ^ XEsp
            XMin = XMin / Rid
            XMax = XMax / Rid
            For I = 1 To NXV
                XV(I) = XV(I) / Rid
            Next I
        Else
            XEsp = 0
            sUX$ = "x"
        End If
    
        AMax = DMAX1(Abs(YMin), Abs(YMax))
        If AMax > 1000# Then
            ' Reduces the scale of Y values:
            YEsp = 3 * Int((Log(AMax) / Log10) / 3#)
            sUY$ = "y[10^" & YEsp & "]"
            Rid = 10# ^ YEsp
            YMin = YMin / Rid
            YMax = YMax / Rid
            For I = 1 To NYV
                YV(I) = YV(I) / Rid
            Next I
        Else
            YEsp = 0
            sUY$ = "y"
        End If

        AMax = DMAX1(Abs(ZMin), Abs(ZMax))
        If AMax > 1000# Then
            ' Reduces the scale of Z values:
            ZEsp = 3 * Int((Log(AMax) / Log10) / 3#)
            sUZ$ = "z[10^" & ZEsp & "]"
            Rid = 10# ^ ZEsp
            ZMin = ZMin / Rid
            ZMax = ZMax / Rid
            For J = 1 To NYV
                For I = 1 To NXV
                    ZV(I, J) = ZV(I, J) / Rid
                Next I
            Next J
        Else
            ZEsp = 0
            sUZ$ = "z"
        End If
'
        ' Settings for the Sub DrawSup_BN:
        NPoli = NXV - 1
        ReDim PRv(1 To NXV, 1 To NYV)   ' Also for Sub DrawSurface.
        ReDim lpPoint(1 To 4 * NPoli)
        ReDim lpVertici(1 To NPoli)
        For I = 1 To NPoli
            lpVertici(I) = 4
        Next I
    '
        ' Settings for the Sub DrawSurface:
        ReDim lpPoint_C(1 To 4)
'
        ' Step between colors:
        If (ZMax > ZMin) Then
            ZnCol = CSng(NTCol - 1) / (ZMax - ZMin)
        End If
'
        ' Prepare the array of colors to assign to quadrilaterals:
        ReDim ZCol(1 To NXV - 1, 1 To NYV - 1)
        For J = 1 To NYV - 1
            For I = 1 To NXV - 1
                ' Calculation of the mean value of the Z coordinates of the four vertices:
                ZMed = CSng(ZV(I, J) + ZV(I, J + 1) + ZV(I + 1, J + 1) + ZV(I + 1, J)) / 4!
                ' and of the corresponding color:
                ZCol(I, J) = TCol(CLng((ZMed - ZMin) * ZnCol))
            Next I
        Next J
    End If
'
'
'
End Sub

Private Sub QuickSort3V(ByRef ValTab#(), ByRef ValTab1#(), ByRef ValTab2#(), _
    ByVal Low&, ByVal High&, Optional ByVal OrderDir& = -1)
'
'   Routine QuickSort3V:
'    ValTab():  Vettore che si vuole ordinare.
'    ValTab1(): Primo vettore associato.
'    ValTab2(): Secondo vettore associato.
'    Low:       Posizione iniziale della zona da ordinare.
'    High:      Posizione finale della zona da ordinare.
'    OrderDir:     Direzione dell' ordinamento:
'                > 0 -> dal minore al maggiore.
'                = 0 -> nessun ordinamento.
'                < 0 -> dal maggiore al minore.
'
    Dim RandIndex&, I&, J&, M$
    Dim ValTemp As Double   ' Tipo del vettore che si vuole ordinare.
    Dim Part As Double      ' Tipo della chiave di ordinamento.
'
    On Error GoTo QuickSort3V_ERR
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
                ' Vettore principale:
                ValTemp = ValTab(Low)
                ValTab(Low) = ValTab(High)
                ValTab(High) = ValTemp
                ' Primo vettore associato:
                ValTemp = ValTab1(Low)
                ValTab1(Low) = ValTab1(High)
                ValTab1(High) = ValTemp
                ' Secondo vettore associato:
                ValTemp = ValTab2(Low)
                ValTab2(Low) = ValTab2(High)
                ValTab2(High) = ValTemp
            End If
'
        Else
            ' Pick a pivot element, then move it to the end:
            RandIndex = (High + Low) / 2
            'SWAP ValTab(High), ValTab(RandIndex)
            ' Vettore principale:
            ValTemp = ValTab(High)
            ValTab(High) = ValTab(RandIndex)
            ValTab(RandIndex) = ValTemp
            Part = ValTab(High)
            ' Primo vettore associato:
            ValTemp = ValTab1(High)
            ValTab1(High) = ValTab1(RandIndex)
            ValTab1(RandIndex) = ValTemp
            ' Secondo vettore associato:
            ValTemp = ValTab2(High)
            ValTab2(High) = ValTab2(RandIndex)
            ValTab2(RandIndex) = ValTemp
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
                    ' Vettore principale:
                    ValTemp = ValTab(I)
                    ValTab(I) = ValTab(J)
                    ValTab(J) = ValTemp
                    ' Primo vettore associato:
                    ValTemp = ValTab1(I)
                    ValTab1(I) = ValTab1(J)
                    ValTab1(J) = ValTemp
                    ' Secondo vettore associato:
                    ValTemp = ValTab2(I)
                    ValTab2(I) = ValTab2(J)
                    ValTab2(J) = ValTemp
                End If
'
            Loop While I < J
            ' Move the pivot element back to its proper place in the array:
            'SWAP ValTab(I), ValTab(High)
            ' Vettore principale:
            ValTemp = ValTab(I)
            ValTab(I) = ValTab(High)
            ValTab(High) = ValTemp
            ' Primo vettore associato:
            ValTemp = ValTab1(I)
            ValTab1(I) = ValTab1(High)
            ValTab1(High) = ValTemp
            ' Secondo vettore associato:
            ValTemp = ValTab2(I)
            ValTab2(I) = ValTab2(High)
            ValTab2(High) = ValTemp
'
            ' Recursively call the QuickSort3V procedure (pass the smaller
            ' subdivision first to use less stack space):
            If (I - Low) < (High - I) Then
                QuickSort3V ValTab(), ValTab1(), ValTab2(), Low, I - 1, OrderDir
                QuickSort3V ValTab(), ValTab1(), ValTab2(), I + 1, High, OrderDir
            Else
                QuickSort3V ValTab(), ValTab1(), ValTab2(), I + 1, High, OrderDir
                QuickSort3V ValTab(), ValTab1(), ValTab2(), Low, I - 1, OrderDir
            End If
        End If
    End If
'
'
QuickSort3V_ERR:
    If (Err <> 0) Then
        M$ = "Error " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " QuickSort3V"
    End If
'
'
'
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    shpInd.Visible = False
    UpdateCursorPositions lblX, "", lblY, "", lblZ, ""
'
'
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'
    Timer1.Enabled = False
    bRotate = False
    bPause = False
    bLoaded = False
    DoEvents
'
    SalvaPosizioneForm Me
'
'
'
End Sub

Private Sub optBN_Click()
'
'
    Draw
'
'
'
End Sub
Private Sub optCol_Click()
'
'
    Draw
'
'
'
End Sub
Private Sub pic3D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    If Button = vbLeftButton Then
        pic3D.MousePointer = vbCustom
'
        shpInd.Visible = False
        UpdateCursorPositions lblX, "", lblY, "", lblZ, ""
'
        lblRAyx.BackColor = PCHL
        lblAlfa.BackColor = PCHL
    End If
'
'
'
End Sub
Private Sub pic3D_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    Dim I&, J&, N&, LAxPx!, LxPx!, LyPx!, LPx As POINTAPI
'
    If (Button = vbLeftButton) Then
        ' Sposta la vista:
        LAxPx = Ax * AsseX      ' Lunghezza asse X [Pixels].
        LxPx = Ax * (X - XRMin) ' Posizione orizzontale del cursore [Pixels dall' asse Z].
        LyPx = Az * (ZRMin - Y) ' Posizione verticale del cursore [Pixels dall' asse X].
'
        ALFA = DATAN2(LyPx, LxPx)
        RAyx = Sqr(LxPx * LxPx + LyPx * LyPx) / LAxPx
'
        shpInd.Visible = False
        Draw True
'
        lblRAyx = Format$(RAyx, "#0.000")
        lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    Else
        MeasureSpace3D
    End If
'
'
'
End Sub
Private Sub pic3D_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    pic3D.MousePointer = vbCrosshair
'
    lblRAyx.BackColor = vbButtonFace
    lblAlfa.BackColor = vbButtonFace
'
'
'
End Sub
Private Sub DrawPoints(ByVal bCol As Boolean)
'
'   Disegna, con API, i punti specificati nei vettori XV(), YV() e ZV().
'   Se bCol = True ai punti viene assegnato un colore proporzionale alla
'   loro altezza:
'
    Dim N&, hPen&, hPen_O&, hBrush&, hBrush_O&, lR1&
    Dim Xr#, Yr#
'
    hPen = CreatePen(vbSolid, 1, vbWhite)
    hPen_O = SelectObject(pic3D.hdc, hPen)
    hBrush = CreateSolidBrush(&H808080) 'pic3D.BackColor)
    hBrush_O = SelectObject(pic3D.hdc, hBrush)
'
    ' Disegno i punti proiettati sul piano di rappresentazione:
    For N = 1 To NV
        If bRotate Then
            Rotate XV(N), YV(N), CDbl(THETA), X0r, Y0r, Xr, Yr
            PRv(N).X = CLng((Ax * Xr + Bx) + (Ay * Yr + By) * CosA)
            PRv(N).Y = CLng((Az * ZV(N) + Bz) - (Ay * Yr + By) * SinA)
        Else
            PRv(N).X = CLng((Ax * XV(N) + Bx) + (Ay * YV(N) + By) * CosA)
            PRv(N).Y = CLng((Az * ZV(N) + Bz) - (Ay * YV(N) + By) * SinA)
        End If
'
        If bCol Then
            lR1 = DeleteObject(hBrush)
            hBrush = CreateSolidBrush(ZCol(N))
            lR1 = SelectObject(pic3D.hdc, hBrush)
        End If
'
        Ellipse pic3D.hdc, PRv(N).X - lRP, PRv(N).Y - lRP, _
                           PRv(N).X + lRP, PRv(N).Y + lRP
    Next N
'
    lR1 = SelectObject(pic3D.hdc, hPen_O)
    lR1 = SelectObject(pic3D.hdc, hBrush_O)
    lR1 = DeleteObject(hPen)
    lR1 = DeleteObject(hBrush)
'
'
'
End Sub

Private Sub Draw(Optional ByVal bChangeView As Boolean = False)
'
'   Draw, with the required parameters, the data points or the surface:
'
    Dim lGridCol&, RLyx!, Rzx!
    
    Dim t0!
    t0 = Timer
'
    XRMin = XMin: XRMax = XMax ' Minimum and maximum values
    YRMin = YMin: YRMax = YMax ' on the axes of the switchboard.
    ZRMin = ZMin: ZRMax = ZMax '
'
    ' Set the graphic:
    lGridCol = IIf(bRotate, RFCL, vbGreen)
    Picture3D pic3D, XRMin, XRMax, YRMin, YRMax, ZRMin, ZRMax, _
             ALFA, RAyx, Ax, Bx, Ay, By, Az, Bz, , , , , , , _
             Title$, sUX$, sUY$, sUZ$, lGridCol, True
'
    If bChangeView Then
        ' Precalculation of translations for the Subs DisegnaXXX:
        SinA = Sin(ALFA)
        CosA = Cos(ALFA)
'
        ' Position on the X-Y plane of the rotation center:
        X0r = CDbl(XRMin + XRMax) / 2#
        Y0r = CDbl(YRMin + YRMax) / 2#
'
        ' Parameters for drawing rotation references:
        AsseX = XRMax - XRMin
        SemiAsseX = AsseX / 2!
        SemiAsseY = (YRMax - YRMin) / 2!
        AsseZ = ZRMax - ZRMin
'
        ' Z/X scale ratio:
        Rzx = -Ax / Az
        ' Coordinate rotations:
        RLyx = AsseX / (YRMax - YRMin)
        TrRotX = RAyx * RLyx * CosA
        TrRotY = RAyx * RLyx * SinA * Rzx
    End If
'
    If fPoints Then
        DrawPoints optCol
    Else
        DrawSurface optCol
    End If
'
    If bRotate Then DrawAxisRot
'
    pic3D.Refresh
'
'
'
End Sub

Private Sub CercaVertice(ByVal LPx&, ByVal lPy&, ByRef Iu&, ByRef Ju&)
'
'   Cerca, nella matrice PRv(), il vertice piu' vicino alle
'   coordinate proiettate lPx, lPy e ne ritorna gli indici
'   di colonna e riga.  In questa matrice i vertici in primo
'   piano sono quelli della prima riga e vengono favoriti:
'
    Dim I&, J&, DisY&, DisQ&, DisQMin&
'
    DisQMin = 2147483647
    For J = 1 To NYV
        DisY = (YV(J) - YRMin)  ' Favorisce, nella ricerca,
                                ' i vertici con Y minore.
        For I = 1 To NXV
            DisQ = (PRv(I, J).X - LPx) * (PRv(I, J).X - LPx) _
                 + (PRv(I, J).Y - lPy) * (PRv(I, J).Y - lPy) + DisY
            If DisQMin > DisQ Then
                Iu = I
                Ju = J
                DisQMin = DisQ
            End If
        Next I
    Next J
'
'
'
End Sub

Private Sub UpdateCursorPositions(ByVal lblVal1 As Label, ByVal Val1$ _
                                   , ByVal lblVal2 As Label, ByVal Val2$ _
                                   , ByVal lblVal3 As Label, ByVal Val3$)
'
'   Update the values of the cursor position windows
'    and set the highlight color:
'
    lblVal1 = Val1$
    lblVal2 = Val2$
    lblVal3 = Val3$
'
    lblVal1.BackColor = IIf(Len(Val1$) = 0, vbButtonFace, PCHL)
    lblVal2.BackColor = IIf(Len(Val2$) = 0, vbButtonFace, PCHL)
    lblVal3.BackColor = IIf(Len(Val3$) = 0, vbButtonFace, PCHL)
'
'
'
End Sub

Private Function CercaPunto(ByVal LPx&, ByVal lPy&) As Long
'
'   Cerca, nel vettore PRv(), il punto piu' vicino alle
'   coordinate proiettate lPx, lPy e ne ritorna l' indice.
'   Il vettore viene esplorato all' indietro per trovare
'   prima i punti in primo piano:
'
    Dim N&, Nu&, DisQ&, DisQMin&
'
    DisQMin = 2147483647
    For N = NV To 1 Step -1
        DisQ = (PRv(N).X - LPx) * (PRv(N).X - LPx) _
             + (PRv(N).Y - lPy) * (PRv(N).Y - lPy)
        If DisQMin > DisQ Then
            Nu = N
            DisQMin = DisQ
        End If
    Next N
'
    CercaPunto = Nu
'
'
'
End Function

Private Sub MeasureSpace3D()
'
'   Measure 3D space:
'
    Dim I&, J&, N&, LPx As POINTAPI
'
    If bRotate Then Exit Sub
'
    GetCursorPos LPx
    If WindowFromPoint(LPx.X, LPx.Y) = pic3D.hWnd Then
        ScreenToClient pic3D.hWnd, LPx
'
        If fPoints Then
            N = CercaPunto(CLng(LPx.X), CLng(LPx.Y))
'            pic3D.ToolTipText = ""
'            pic3D.ToolTipText = " X = " & Format$(XV(N), "#0.000 ") & _
'                                " Y = " & Format$(YV(N), "#0.000 ") & _
'                                " Z = " & Format$(ZV(N), "#0.000 ")
            UpdateCursorPositions lblX, Format$(XV(N), "#0.000 "), _
                                     lblY, Format$(YV(N), "#0.000 "), _
                                     lblZ, Format$(ZV(N), "#0.000 ")
            shpInd.Left = pic3D.ScaleX(PRv(N).X - shpIndOffx, vbPixels, vbUser) _
                          + pic3D.ScaleLeft
            shpInd.Top = pic3D.ScaleY(PRv(N).Y - shpIndOffy, vbPixels, vbUser) _
                         + pic3D.ScaleTop
        Else
            CercaVertice CLng(LPx.X), CLng(LPx.Y), I, J
'            pic3D.ToolTipText = ""
'            pic3D.ToolTipText = " X = " & Format$(XV(I), "#0.000 ") & _
'                                " Y = " & Format$(YV(J), "#0.000 ") & _
'                                " Z = " & Format$(ZV(I, J), "#0.000 ")
            UpdateCursorPositions lblX, Format$(XV(I) * 10 ^ XEsp, "#0.000 "), _
                                     lblY, Format$(YV(J) * 10 ^ YEsp, "#0.000 "), _
                                     lblZ, Format$(ZV(I, J) * 10 ^ ZEsp, "#0.000 ")
            shpInd.Left = pic3D.ScaleX(PRv(I, J).X - shpIndOffx, vbPixels, vbUser) _
                          + pic3D.ScaleLeft
            shpInd.Top = pic3D.ScaleY(PRv(I, J).Y - shpIndOffy, vbPixels, vbUser) _
                         + pic3D.ScaleTop
        End If
'
        shpInd.Visible = True
'
    Else
        shpInd.Visible = False
    End If
'
'
'
End Sub

Private Sub Timer1_Timer()
'
'
    THETA = THETA + dth
    If THETA >= PI2 Then THETA = 0
    lblTheta = Format(RadToGrd * THETA, "#0.0")
'
    Draw
'
'
'
End Sub
Private Sub updTheta_Change()
'
'
    If lblTheta >= 360 Then lblTheta = 0
    If lblTheta < 0 Then lblTheta = 359
'
    THETA = CDbl(lblTheta) / RadToGrd
    lblTheta = Format(RadToGrd * THETA, "#0.0")
'
    Draw
'
'
'
End Sub
Private Sub DrawAxisRot()
'
'
    Dim X0!, x1!, x2!, Y0!, y1!, y2!
'
    pic3D.ForeColor = vbGreen
'
    ' Coordinate, in [vbUser], del centro di rotazione:
    X0 = X0r + (Y0r - YRMin) * TrRotX
    Y0 = ZRMin + (Y0r - YRMin) * TrRotY
'
    ' Coordinate, in [vbUser], dell' estremita' dell' asse di rotazione:
    y1 = Y0 + AsseZ
    pic3D.DrawStyle = vbDashDot
    pic3D.Line (X0, Y0)-(X0, y1)
'
    ' Coordinate, in [vbUser], dell' estremita' dell' asse X in rotazione:
    x2 = X0r + SemiAsseX * Cos(THETA)
    y2 = Y0r + SemiAsseY * Sin(THETA)
    x1 = x2 + (y2 - YRMin) * TrRotX
    y1 = ZRMin + (y2 - YRMin) * TrRotY
    pic3D.DrawStyle = vbSolid
    pic3D.Line (X0, Y0)-(x1, y1)
    pic3D.Print "x"
'
'
'
End Sub
