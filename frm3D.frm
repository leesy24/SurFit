VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm3D 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " 3D isometric view"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frm3D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   582
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Begin VB.Frame zFrame2 
      Caption         =   "Rotation:"
      Height          =   615
      Left            =   2160
      TabIndex        =   24
      Top             =   45
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin MSComCtl2.UpDown updTheta 
         Height          =   255
         Left            =   2895
         TabIndex        =   6
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame zFrame1 
      Caption         =   "Views:"
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   45
      Width           =   1755
      Begin VB.CommandButton cmdDef 
         Caption         =   "&Def"
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
         Left            =   60
         TabIndex        =   0
         ToolTipText     =   "Default view "
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdZY 
         Caption         =   "Z-&Y"
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
         Left            =   1320
         TabIndex        =   3
         ToolTipText     =   "Side view "
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdZX 
         Caption         =   "&Z-X"
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
         Left            =   900
         TabIndex        =   2
         ToolTipText     =   "Front view "
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdXY 
         Caption         =   "&X-Y"
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
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Plan view "
         Top             =   240
         Width           =   375
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
      Height          =   7560
      Left            =   240
      MouseIcon       =   "frm3D.frx":014A
      MousePointer    =   2  'Cross
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   9060
      Begin VB.CommandButton cmdCopiaGrafico 
         Height          =   285
         Left            =   0
         MaskColor       =   &H000000FF&
         MousePointer    =   1  'Arrow
         Picture         =   "frm3D.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copy the image to the clipboard "
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton optBW 
         BackColor       =   &H00000000&
         Caption         =   "&B/W"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   8160
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optCol 
         BackColor       =   &H00000000&
         Caption         =   "&Col."
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7440
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   60
         Value           =   -1  'True
         Width           =   735
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
   Begin VB.Label zLabel3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Theta:"
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
      Left            =   1200
      TabIndex        =   30
      Top             =   8400
      Width           =   555
   End
   Begin VB.Label lblATheta 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.00"
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
      Left            =   1800
      TabIndex        =   29
      Top             =   8370
      Width           =   615
   End
   Begin VB.Label lblAPhi 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-00.00"
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
      Left            =   600
      TabIndex        =   28
      Top             =   8370
      Width           =   615
   End
   Begin VB.Label zLabel1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phi:"
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
      TabIndex        =   27
      Top             =   8400
      Width           =   315
   End
   Begin VB.Label lblZ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.000"
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
      TabIndex        =   22
      Top             =   8370
      Width           =   615
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
      Left            =   4080
      TabIndex        =   21
      Top             =   8400
      Width           =   195
   End
   Begin VB.Label lblX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.000"
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
      Left            =   2640
      TabIndex        =   20
      Top             =   8370
      Width           =   615
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
      Left            =   2400
      TabIndex        =   19
      Top             =   8400
      Width           =   195
   End
   Begin VB.Label lblY 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00.000"
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
      Left            =   3480
      TabIndex        =   18
      Top             =   8370
      Width           =   615
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
      Left            =   3240
      TabIndex        =   17
      Top             =   8400
      Width           =   195
   End
   Begin VB.Label zlblAutore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DASAN Info Tek 2018"
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
      Left            =   7560
      TabIndex        =   16
      Top             =   8460
      Width           =   1620
   End
   Begin VB.Label lblstructions 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   15
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
      Left            =   6240
      TabIndex        =   14
      Top             =   8400
      Width           =   615
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
      Left            =   4920
      TabIndex        =   13
      Top             =   8400
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
      Left            =   5400
      TabIndex        =   12
      Top             =   8370
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
      Left            =   6840
      TabIndex        =   11
      Top             =   8370
      Width           =   795
   End
End
Attribute VB_Name = "frm3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================
' Description......: Form, for the 3D representation of points or
'                     surfaces (axonometric view).
' Name of the Files: frm3D.frm, fr3D.frx
'                    modUtility.bas
' Date.............: 10/12/2001
' Update...........: 1/2/2002  (added point representation).
' Update...........: 17/3/2002 (arrange some inconsistencies of
'                     representation).
' Update...........: 21/3/2002 (rotation added).
' Version..........: 1.1 at 32 bits (provisional, developing).
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.virgilio.it/flanguasco/
'                    http://www.flanguasco.org
'===================================================================
'
'   Entry routines:
'
'    frm3D.Points XD#(), YD#(), ZD#() [, Title$ = ""] [, IC& = 1] _
'                [, bAutoScale as Boolean = True] _
'                [, XMin#, XMax#, YMin#, YMax#, ZMin#, ZMax#]
'     XD():       vector containing the abscissas of the points to represent.
'     YD():         "        "       "  ordinates "   "    "      ".
'     ZD():         "        "       "  altitude  "   "    "      ".
'     Title$:     title of the panel (optional).
'     IC:         if IC <= 1 the Form is put in the foreground (optional).
'     bAutoScale: if False must also pass the values Min and Max to be used
'                  as extremes of the three axes.
'
'    frm3D.Surface XI#(), YI#(), ZI#() [, Title$ = ""] [, IC& = 1]
'     XI():   vector containing the abscissas of the surface to be represented.
'     YI():     "        "       "  ordinates "         "       "      ".
'     ZI():   matrix     "       "  levels    "         "       "      ".
'     Title$: title of the panel (optional).
'     IC:     if IC <= 1 the Form is put in the foreground (optional).
'
'   Note:   For Sub Points:
'            the XD (1 to ND), YD (1 to ND) and ZD (1 To ND) vectors must have
'             the same dimensions; otherwise the smaller size is used.
'           For Sub Surface:
'            if the dimensions of the vectors are XI(1 to NXI) and YI(1 to NYI),
'            the matrix must be sized as ZI (1 to NXI, 1 to NYI).
'
'   Note:   All vectors and matrices of these routines start from index 1
'            (TCol() excluded).
'
Option Explicit
'
Dim PhiV#(), ThetaV#()    ' Vectors of the angle phi and theta.
Dim XV#(), YV#()    ' Vectors of the abscissas and ordinates.
Dim ZV#()           ' Vectors of the heights of the points
                    '  or matrix of the levels of the surface.
Dim Title$
'
Dim fPoints As Boolean  ' If True, draw points;
                        ' if False, draw a surface.
'
Dim NV&             ' Number of values in the XV(), YV() and ZV() vectors.
'
Dim NXV&, NYV&      ' Number of values in the XV(), YV(), and column and row
                    '  vectors in the ZV () array.
'
Dim XMin!, XMax!    ' Minimum
Dim YMin!, YMax!    '  and maximum values
Dim ZMin!, ZMax!    '  of incoming data.
'
Dim XRMin!, XRMax!  ' Minimum and maximum values on the axis of the frame:
Dim YRMin!, YRMax!  ' they are used to avoid, in this particular application,
Dim ZRMin!, ZRMax!  '  that subsequent calls to Picture3D change the axis scales.
                    ' In addition, the view change driven by the mouse requires
                    '  the true values of XRMin, XRMax and ZRMin.
Dim AsseX!          ' XRMax - XRMin.
Dim sUX$            ' Label of the units of the X axis.
Dim XEsp&           ' X scale reduction factor
Dim sUY$            ' Label of the units of the Y axis.
Dim YEsp&           ' Y scale reduction factor
Dim sUZ$            ' Label of the units of the Z axis.
Dim ZEsp&           ' Z scale reduction factor
'
Dim Ax!, Bx!        ' Scale conversion
Dim Ay!, By!        '  coefficients from
Dim Az!, Bz!        '  vbUser to Pixels.
'
Dim RAyx!           ' Y-axis length ratio in [Pixels] to X-axis length in [Pixels].
Dim ALFA!           ' Angle of the Y axis with respect to the X axis [Rad].
Dim SinA!, CosA!    ' Sin(ALFA) and Cos(ALFA) calculated only after a change of view.
'
' Constants and variables for rotation:
Dim THETA!              ' Rotation current angle [Radians].
Const dth! = PI / 180#  ' Increment of the rotation angle [Radians].
Dim X0r#, Y0r#          ' Coordinates of the rotation center.
Dim SemiAsseX!          ' Measurements of the axes of the drawing.
Dim SemiAsseY!          '      "       "     "     "     "
Dim AsseZ!              '      "       "     "     "     "
Dim TrRotX!, TrRotY!    ' Coeff. transformation of the coordinates during rotation.
Dim bPause As Boolean   ' Flag for rotation paused.
'Const RFCL& = &H8000&   ' Color of axes and reference grids during rotation.
Const RFCL& = &H404080  ' Color of axes and reference grids during rotation.
'
Dim TCol&()         ' Table of colors.
Const NTCol& = 1280 ' Number of colors available in TCol().
Dim ZCol&()         ' Vector or array of colors to be used.
'
Const Log10# = 2.30258509299405
Const RadToGrd# = 180# / PI ' Conversion factor from [Rad] to [Grd].
'
Private Type POINTAPI
     X As Long          ' [Pixels].
     Y As Long          '     "
End Type
'
' Variables for the Sub DrawPoints:
Dim PRv() As POINTAPI       ' Vector of the points or matrix of the vertices of the
                            '  quadrilaterals projected on the representation plane
                            '  (it is also used by the Subs DrawSup_BN and DrawSurface).
Const lRP& = 4              ' Dot drawing radius [Pixels].
'
' Variables for the Sub DrawSup_BN:
Dim NPoli&                  ' Number of quadrilaterals in a row.
Dim lpPoint() As POINTAPI   ' Vector of the vertices of a row.
Dim lpVertici&()            ' Vector of the number of vertices of each polygon.
'
' Variables for the Sub DrawSurface:
Dim lpPoint_C() As POINTAPI ' Vector of the vertices of a quadrilateral.
'
' Constants for finding the 3D position:
Const shpIndOffx& = lRP + 2 ' Horizontal and vertical offset
Const shpIndOffy& = lRP + 2 '  of the highlight circle.
Const PCHL& = &HC0FFFF      ' Color of evidence for the cursor position labels.
'
Dim bRotate As Boolean      ' Flag for Rotation in progress.
Dim bLoaded As Boolean      ' Flag for Form initialized.
'
Dim Mouse_X_Prev As Single
Dim Mouse_Y_Prev As Single
'
'-------------------------------------------------------------------------------------
'   Graphic APIs:
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

Private Sub ROTATE(ByVal X0#, ByVal Y0#, ByVal Rot#, ByVal CRx#, ByVal CRy#, ByRef Xr#, ByRef Yr#)
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

Private Function Picture3D(ByVal Sheet As PictureBox, _
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
'     Sheet:    PictureBox to climb.
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
'     AutoRed:   State of Sheet.AutoRedraw after drawing the painting.
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
    Dim DZMin!                      ' Min. Width of the Z scale
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
    Sheet.FontName = "MS Sans Serif"
    Sheet.FontBold = False
'
    CosA = Cos(ALFA)
    SinA = Sin(ALFA)
'
    'Cancel the previous stairs:
    Sheet.ScaleMode = vbPixels
'
    ' The left edge must be sufficient to contain the largest Z value:
    Sheet.FontSize = 8
    TxWs = MAX0(Sheet.TextWidth(Format$(-Abs(Z0), FormatVZ$) & "W"), _
               Sheet.TextWidth(Format$(-Abs(Zn), FormatVZ$) & "W"), _
               Sheet.TextWidth(UnitaZ$ & "W"))
'
    ' The border on the right must be sufficient to contain
    '  the Xn value and the UnitaX $ label:
    TxWd = Sheet.TextWidth(Format$(-Abs(Xn), FormatVZ$) & "W") _
         + Sheet.TextWidth(UnitaX$ & "W")
'
    ' The borders on the left and on the right are:
    DT_X = LAx * (1! + RAyx * CosA)
    BDen = DT_X / (Sheet.ScaleWidth - TxWs - TxWd)
    Bl = TxWs * BDen
    Br = TxWd * BDen
'
    ' The border below is 2 times the height of the values:
    TxHb = 2! * Abs(Sheet.TextHeight("W"))
'
    ' The edge above is 2 times the height of the values plus 2 times the height of the title:
    Sheet.FontSize = 12
    TxHt = TxHb + 2! * Abs(Sheet.TextHeight(Title$))
'
    ' The edges above and below are:
    DDz = Sheet.ScaleWidth * LAx * RAyx * SinA / (Bl + DT_X + Br)
    DPz = Abs(Sheet.ScaleHeight) - DDz - TxHt - TxHb
    If DPz <= 0 Then DPz = 0.001
    DT_Z = LAz * (1! + DDz / DPz)
    BDen = DT_Z / (Abs(Sheet.ScaleHeight) + TxHb + TxHt)
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
    'Sheet.Picture = LoadPicture("")
    Sheet.Scale (QxMin, QzMax)-(QxMax, QzMin)
    Sheet.Line (QxMin, QzMin)-(QxMax, QzMax), Sheet.BackColor, BF ' This is faster than
                                                                    '  Sheet.Cls.
    ' The drawing of the painting must be permanent:
    Sheet.AutoRedraw = True
'
    ' Width and height of 1 pixel in [vbUser]:
    Px1_X = Abs(Sheet.ScaleX(1, vbPixels, vbUser))
    Px1_Z = Abs(Sheet.ScaleY(1, vbPixels, vbUser))
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
    If Sheet.ScaleY(LAz, vbUser, vbPixels) > 0 Then
    End If
'
'-------------------------------------------------------------------------------------
'   Draw axes, grids and write scale values.
'
    Sheet.FontSize = 8
    Sheet.DrawWidth = 1
    Sheet.ForeColor = RifCol
    Sheet.DrawMode = vbCopyPen
'
    ' Check the separation of the labels:
    Dim TxW!
    TxW = DMAX1(Sheet.TextWidth(Format$(X0, FormatVX$)), _
               Sheet.TextWidth(Format$(Xn, FormatVX$)))
    bVlX = (LAx / rrx) * TxW < LAx
'
    TxW = DMAX1(Sheet.TextWidth(Format$(Y0, FormatVY$)), _
               Sheet.TextWidth(Format$(Yn, FormatVY$)))
    TxH = Abs(Sheet.TextHeight("W"))
    bVlY = ((Yn - Y0) / rry) * TxW < LyCosA
    bVlY = bVlY Or (((Yn - Y0) / rry) * TxH < LySinA)
'
    bVlZ = (LAz / rrz) * TxH < LAz
'
    ' Draw the X axis:
    Sheet.DrawStyle = vbSolid
    Sheet.Line (X0, Z0)-(Xn + EstAx, Z0)
    If bVlX Then
        Sheet.Line (Xn + EstAx, Z0) _
                   -(Xn + EstAx - TaccheX, Z0 + TaccheZ / 2!)
        Sheet.Line (Xn + EstAx, Z0) _
                   -(Xn + EstAx - TaccheX, Z0 - TaccheZ / 2!)
        ' and write the label of the X axis:
        If Len(UnitaX$) > 0 Then
            Sheet.CurrentX = Xn + Sheet.TextWidth(Xn & "W")
            Sheet.Print UnitaX$;
        End If
    End If
'
    ' Draw the Y axis:
    Sheet.Line (X0, Z0)-(X0 + (LAy + EstAx) * CosA, _
                          Z0 + (LAy + EstAx) * SinA * Ryx)
    If bVlY Then
        ' and write the Y axis label:
        If Len(UnitaY$) > 0 Then
            Sheet.CurrentY = Sheet.CurrentY - Sheet.TextHeight("W")
            Sheet.Print UnitaY$;
        End If
    End If
'
    ' Draw the Z axis:
    Sheet.Line (X0, Z0)-(X0, Zn + EstAz)
    If bVlZ Then
        Sheet.Line (X0, Zn + EstAz) _
                   -(X0 - TaccheX / 2!, Zn + EstAz - TaccheZ)
        Sheet.Line (X0, Zn + EstAz) _
                   -(X0 + TaccheX / 2!, Zn + EstAz - TaccheZ)
        ' and write the Z axis label:
        If Len(UnitaZ$) > 0 Then
            Sheet.CurrentX = QxMin
            Sheet.CurrentY = Zn + EstAz - Sheet.TextHeight("W") / 2!
            Sheet.Print UnitaZ$;
        End If
    End If
'
    ' Draw the vertical grid on the Z-X plane,
    '  the plane on the X-Y plane and write the values of the X axis:
    Dim rrx_10!
'
    rrx_10 = rrx / 10!
    Sheet.DrawStyle = vbDot
    For XI = X0 To Xn + rrx_10 Step rrx
        Sheet.Line (XI + LyCosA, Z0 + LySinA + LAz)-(XI + LyCosA, Z0 + LySinA)
        Sheet.Line -(XI, Z0)
        Sheet.Line -(XI, Z0 - TaccheZ)
        If bVlX Then
            Tx$ = Format$(XI, FormatVX$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(XI - Val(Tx$)) < rrx_10) Then
                Sheet.CurrentX = XI - Sheet.TextWidth(Tx$) / 2!
                Sheet.Print Tx$;
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
        Sheet.Line (Yx, Yz)-(Yx - LAx, Yz)
        Sheet.Line -(Sheet.CurrentX, Yz + LAz)
        If bVlY Then
            Tx$ = Format$(YI, FormatVY$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(YI - Val(Tx$)) < rry_10) Then
                ' The positions of the Y labels depend on
                '  the presence of those Z:
                If bVlZ Then
                    Sheet.Line -(Sheet.CurrentX, Sheet.CurrentY + EstAz)
                    Sheet.CurrentX = Sheet.CurrentX - Sheet.TextWidth(Tx$) / 2!
                    Sheet.CurrentY = Sheet.CurrentY - Sheet.TextHeight(Tx$)
                Else
                    Sheet.Line -(Sheet.CurrentX - TaccheX, Sheet.CurrentY)
                    Sheet.CurrentX = Sheet.CurrentX - Sheet.TextWidth(Tx$)
                    Sheet.CurrentY = Sheet.CurrentY - Sheet.TextHeight(Tx$) / 2!
                End If
                Sheet.Print Tx$;
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
        Sheet.Line (X0 - TaccheX, ZI)-(X0, ZI)
        Sheet.Line -(X0 + LyCosA, ZI + LySinA)
        Sheet.Line -(Sheet.CurrentX + LAx, Sheet.CurrentY)
        If bVlZ Then
            Tx$ = Format$(ZI, FormatVZ$)
            ' Verify that the chosen format does not lead
            '  to representation errors:
            If (Abs(ZI - Val(Tx$)) < rrz_10) Then
                Sheet.CurrentX = QxMin
                Sheet.CurrentY = ZI - Sheet.TextHeight(Tx$) / 2!
                Sheet.Print Tx$;
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
        Sheet.FontSize = 12
        Sheet.FontBold = True
        Sheet.ForeColor = vbRed
'
        TitW = Sheet.TextWidth(Title$)
        TitH = Sheet.TextHeight(Title$)
        ' Verify that the title is all in the Sheet:
        If TitW <= Sheet.ScaleWidth Then
            TitL = (QxMin + QxMax - TitW) / 2!
        Else
            ' and if not, cut it:
            TitL = Sheet.ScaleLeft
            Tx$ = " . . . ."
            Title$ = Left$(Title$, Int(Len(Title$) * _
            (Sheet.ScaleWidth - Sheet.TextWidth(Tx$)) / TitW)) & Tx$
        End If
        TitT = QzMax
        ' Delete the area on which the title will be written:
        'Sheet.Line (TitL, TitT)-(TitL + TitW, TitT + TitH), Sheet.BackColor, BF
        Sheet.CurrentX = TitL
        Sheet.CurrentY = TitT
        Sheet.Print Title$
    End If
'
'-------------------------------------------------------------------------------------
'   Calculates the transformation coefficients from vbUser to Pixels:
'
    Dim C0_Px!, Cn_Px!
'
    C0_Px = Sheet.ScaleX(X0 - Sheet.ScaleLeft, vbUser, vbPixels)
    Cn_Px = Sheet.ScaleX(Xn - Sheet.ScaleLeft, vbUser, vbPixels)
    Ax = (Cn_Px - C0_Px) / LAx
    Bx = (C0_Px * Xn - Cn_Px * X0) / LAx
'
    Ay = Sheet.ScaleX(LAy / D_Y, vbUser, vbPixels)
    By = -Ay * Y0
'
    C0_Px = Sheet.ScaleY(Z0 - Sheet.ScaleTop, vbUser, vbPixels)
    Cn_Px = Sheet.ScaleY(Zn - Sheet.ScaleTop, vbUser, vbPixels)
    Az = (Cn_Px - C0_Px) / LAz
    Bz = (C0_Px * Zn - Cn_Px * Z0) / LAz
'
    ' And leave the Sheet set:
    Sheet.DrawStyle = vbSolid
    Sheet.AutoRedraw = AutoRed
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

Public Sub Points(PhiD_I#(), ThetaD_I#(), XD_I#(), YD_I#(), ZD_I#(), Optional ByVal Title_I$ = "", _
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
    PhiV() = PhiD_I()
    ThetaV() = ThetaD_I()
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

Public Sub Surface(PhiD_I#(), ThetaD_I#(), XI_I#(), YI_I#(), ZI_I#(), Optional ByVal Title_I$ = "", _
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
    PhiV() = PhiD_I()
    ThetaV() = ThetaD_I()
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
'   Draw, with API, the quadrilaterals in color or in B/W. Drawing the lines
'    from the last back to the first (moving towards the observer), each
'    quadrilateral hides the part of the image he covers.
'   If bCol = True the quadrilaterals is assigned a color proportional to the
'    mean value of the Z coordinates of the four vertices.
'
'   Diagram of the quadrilateral used by the lpPoint_C() vector:
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
    hBrush = CreateSolidBrush(&H808080)         ' Set the color of the
    hBrush_O = SelectObject(pic3D.hdc, hBrush)  ' quadrilaterals for drawing in B/W.
'
    ' Calculating the matrix of the vertices of the quadrilaterals projected
    '  on the representation plane:
    For J = 1 To NYV
        For I = 1 To NXV
            If bRotate Then
                ROTATE XV(I), YV(J), CDbl(THETA), X0r, Y0r, Xr, Yr
                PRv(I, J).X = CLng((Ax * Xr + Bx) + (Ay * Yr + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * Yr + By) * SinA)
            Else
                PRv(I, J).X = CLng((Ax * XV(I) + Bx) + (Ay * YV(J) + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * YV(J) + By) * SinA)
            End If
        Next I
    Next J
'
    ' Drawing the quadrilaterals. As a function of THETA,
    '  first design those more 'far from the' observer:
    If (THETA + ALFA) < 0# And (THETA + ALFA) >= (-1#) * PI_2 Then
        Quadrante = CLng(Int((THETA + ALFA) / PI_2) + 4) Mod 4
    ElseIf (THETA + ALFA) < (-1#) * PI_2 And (THETA + ALFA) > (-1#) * PI Then
        Quadrante = CLng(Int((THETA + ALFA) / PI_2) + 2) Mod 4
    Else
        Quadrante = CLng(Int((THETA + ALFA) / PI_2)) Mod 4
    End If
'
    Select Case Quadrante
        Case 0  ' 1st Quadrant. 0 ~ 90
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
        Case 1  ' 2nd Quadrant. 90 ~ 180
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
        Case 2  ' 3rd Quadrant. 180 ~ 270
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
        Case 3  ' 4th Quadrant. 270 ~ 360
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

Private Sub DrawSurface_BW()
'
'   NOT USED.
'
'  Draw, with API, the rows of quadrilaterals. Drawing the lines from the last
'   back to the first (moving towards the observer), each picture - latero
'   hides the part of the image he covers.
'
'   Schema of quadrilaterals used by the lpPoint() vector:
'  2 _____ 3  6 _____ 7
'   |     |    |     | .... In all the lines the vertices 3 coincide with 6,
'   |     |    |     | ....  4 with 5, 7 with 10, 8 with 9, etc ...
'   |_____|    |_____| .... In lines J <NYV, vertices 2 coincide with 1 of
'  1       4  5       8      line J + 1, 4 with 3 of line J + 1, etc ...
'
'   This routine is reported only for curiosity: The Sub DrawSurface, in fact,
'    offers the same functionality. The curiosity lies in the fact that, using
'    the PolyPolygon API, the following polygons, defined in lpPoint(),
'   DO NOT hide those previously defined in the same vector.
'   Drawing the quadrilaterals line by line the effect is not very evident:
'   if instead you want to put in lpPoint() ALL the quadrilaterals that form
'    the surface the result would be disastrous.
'
    Dim I&, J&, NP&, hPen&, hPen_O&, hBrush&, hBrush_O&, lR1&
    Dim Xr#, Yr#
'
    hPen = CreatePen(vbSolid, 1, vbWhite)
    hPen_O = SelectObject(pic3D.hdc, hPen)
    hBrush = CreateSolidBrush(&H808080) 'pic3D.BackColor)
    hBrush_O = SelectObject(pic3D.hdc, hBrush)
'
    ' Calculating the matrix of the vertices of the quadrilaterals projected on
    '  the representation plane:
    For J = 1 To NYV
        For I = 1 To NXV
            If bRotate Then
                ROTATE XV(I), YV(J), THETA, X0r, Y0r, Xr, Yr
                PRv(I, J).X = CLng((Ax * X0r + Bx) + (Ay * Y0r + By) * CosA)
                PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * Y0r + By) * SinA)
            Else
            PRv(I, J).X = CLng((Ax * XV(I) + Bx) + (Ay * YV(J) + By) * CosA)
            PRv(I, J).Y = CLng((Az * ZV(I, J) + Bz) - (Ay * YV(J) + By) * SinA)
            End If
        Next I
    Next J
'
    ' Draw the lines of quadrilaterals:
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
    UpdateCursorPositions lblAPhi, "", lblATheta, "", lblX, "", lblY, "", lblZ, ""
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
'
    RAyx = 1!
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

Private Sub cmdDef_Click()

'
'
    RAyx = 0.5      ' Y axis length ratio to X axis length.
    ALFA = PI / 3!  ' Angle of the Y axis with respect to the X axis: 60 [Grd].
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
    LoadFormsPositions Me, Me_L, Me_T
    Me.Move Me_L, Me_T ', Me_W, Me_H
'
    TCol() = ColorTable(NTCol)
'
    ' Highlight circle size:
    shpInd.Width = 2 * shpIndOffx
    shpInd.Height = 2 * shpIndOffy
'
    ' Initial values:
    'RAyx = 1!       ' Y axis length ratio to X axis length.
    'ALFA = PI / 6!  ' Angle of the Y axis with respect to the X axis: 30 [Grd]
    RAyx = 0.5      ' Y axis length ratio to X axis length.
    ALFA = PI / 3!  ' Angle of the Y axis with respect to the X axis: 60 [Grd].
'
    lblRAyx = Format$(RAyx, "#0.000")
    lblAlfa = Format$(RadToGrd * ALFA, "#0.000")
'
    bRotate = False
    bPause = False
    lblTheta = "0.0"
    lblstructions = "Rotate by left mouse buttion at center." & vbCrLf & _
                    "View change by right mouse button at" & vbCrLf & _
                    "the end of Y axis."
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
    Dim I&, J&, N&, ZnCol!, ZMed!, ZMinCol!, ZMaxCol!
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
        NV = MIN0(NV, UBound(PhiV), UBound(ThetaV))
        ReDim Preserve XV(1 To NV)      ' Resize vectors
        ReDim Preserve YV(1 To NV)      '  all to the
        ReDim Preserve ZV(1 To NV)      '  same length.
        ReDim Preserve PhiV(1 To NV)    '
        ReDim Preserve ThetaV(1 To NV)  '
        ReDim PRv(1 To NV)
'
        ' Sort the vectors so that the points with major Y remain behind:
        QuickSort5V YV(), XV(), ZV(), PhiV(), ThetaV(), 1, NV
'
        If bAutoScale Then
            ' Find the minimum and maximum values of heights:
            ZMin = DMINVAL(ZV())
            ZMax = DMAXVAL(ZV())
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
            ' Find the minimum and maximum levels values for color table:
            ZMinCol = ZMax
            ZMaxCol = ZMin
            For N = 1 To NV
                If ZV(N) <> ZMin Then
                    If ZMinCol > ZV(N) Then ZMinCol = ZV(N)
                    If ZMaxCol < ZV(N) Then ZMaxCol = ZV(N)
                End If
            Next N
'
            ' Step between colors:
            If (ZMaxCol > ZMinCol) Then
                ZnCol = CSng(NTCol - 1) / (ZMaxCol - ZMinCol)
            End If
'
            ' Prepare the color vector to be assigned to the points:
            ReDim ZCol(1 To NV)
            For N = 1 To NV
                If ZV(N) >= ZMinCol Then
                    ' Calculation of the corresponding color:
                    ZCol(N) = TCol(CLng((ZV(N) - ZMinCol) * ZnCol))
                Else
                    ZCol(N) = &H808080 ' Grey.
                End If
            Next N
'
        End If
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
        Else
            ' Find the minimum and maximum levels values for color table:
            ZMinCol = ZMax
            ZMaxCol = ZMin
            For J = 1 To NYV
                For I = 1 To NXV
                    If ZV(I, J) <> ZMin Then
                        If ZMinCol > ZV(I, J) Then ZMinCol = ZV(I, J)
                        If ZMaxCol < ZV(I, J) Then ZMaxCol = ZV(I, J)
                    End If
                Next I
            Next J
        End If
'
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
'
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
'
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
        If bAutoScale Then
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
                    If ZMed >= ZMin Then
                        ZCol(I, J) = TCol(CLng((ZMed - ZMin) * ZnCol))
                    Else
                        ZCol(I, J) = TCol(0)
                    End If
                Next I
            Next J
        Else
            ' Step between colors:
            If (ZMaxCol > ZMinCol) Then
                ZnCol = CSng(NTCol - 1) / (ZMaxCol - ZMinCol)
            End If
'
            ' Prepare the array of colors to assign to quadrilaterals:
            ReDim ZCol(1 To NXV - 1, 1 To NYV - 1)
            For J = 1 To NYV - 1
                For I = 1 To NXV - 1
                    ' Calculation of the mean value of the Z coordinates of the four vertices:
                    ZMed = CSng(ZV(I, J) + ZV(I, J + 1) + ZV(I + 1, J + 1) + ZV(I + 1, J)) / 4!
                    ' and of the corresponding color:
                    If ZMed < ZMinCol _
                        Or ZMed = ZMin _
                        Or CSng(ZV(I, J)) = ZMin _
                        Or CSng(ZV(I, J + 1)) = ZMin _
                        Or CSng(ZV(I + 1, J + 1)) = ZMin _
                        Or CSng(ZV(I + 1, J)) = ZMin _
                        Then
                        ZCol(I, J) = &H808080 ' Grey.
                    Else
                        ZCol(I, J) = TCol(CLng((ZMed - ZMinCol) * ZnCol))
                    End If
                Next I
            Next J
        End If
    End If
'
'
'
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
    shpInd.Visible = False
    UpdateCursorPositions lblAPhi, "", lblATheta, "", lblX, "", lblY, "", lblZ, ""
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
    SaveFormsPositions Me
'
'
'
End Sub

Private Sub optBW_Click()
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
    If Button = vbLeftButton Or Button = vbRightButton Then
        If Button = vbLeftButton Then
            Mouse_X_Prev = X
            Mouse_Y_Prev = Y
        End If
'
        pic3D.MousePointer = vbCustom
'
        shpInd.Visible = False
        UpdateCursorPositions lblAPhi, "", lblATheta, "", lblX, "", lblY, "", lblZ, ""
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
    If (Button = vbRightButton) Then
        If (X < XRMin) Then
            X = XRMin
        End If
        If (Y < ZRMin) Then
            Y = ZRMin
        End If
'
        ' Move the view:
        LAxPx = Ax * AsseX      ' X axis length [Pixels].
        LxPx = Ax * (X - XRMin) ' Horizontal position of the cursor [Pixels from the Z axis].
        LyPx = Az * (ZRMin - Y) ' Vertical position of the cursor [Pixels from the X axis].
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
    ElseIf (Button = vbLeftButton) Then
        If (bRotate = False) Or (bPause = False) Then
            bRotate = True
            bPause = True
            
            cmdRotate.Enabled = False
            cmdRotate.Caption = "&Stop"
            cmdPause.Enabled = True
            cmdPause.Caption = "Co&nt."
            updTheta.Enabled = True
'
            Timer1.Enabled = False
        End If
'
        If Y > (ZRMin + (Y0r - YRMin) * TrRotY) Then
            If (X - Mouse_X_Prev) < 0 Then
                THETA = THETA + dth * 4
            ElseIf (X - Mouse_X_Prev) > 0 Then
                THETA = THETA - dth * 4
            End If
        Else
            If (X - Mouse_X_Prev) > 0 Then
                THETA = THETA + dth * 4
            ElseIf (X - Mouse_X_Prev) < 0 Then
                THETA = THETA - dth * 4
            End If
        End If
'
        If THETA >= PI2 Then
            THETA = THETA - PI2
        ElseIf THETA < 0 Then
            THETA = PI2 + THETA
        End If
'
        Mouse_X_Prev = X
        Mouse_Y_Prev = Y
'
        shpInd.Visible = False
        Draw
'
        lblTheta = Format(RadToGrd * THETA, "#0.0")
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

Private Sub QuickSort1Double1POINTAPI1Long( _
    ByRef ValTab#(), ByRef ValTab1() As POINTAPI, ByRef ValTab2&(), _
    ByVal Low&, ByVal High&, Optional ByVal OrderDir& = -1)
'
'   Routine QuickSort1Double1POINTAPI1Long:
'    ValTab():  Vector that you want to order.
'    ValTab1(): First associated vector.
'    ValTab2(): Second associated vector.
'    Low:       Initial position of the area to be ordered.
'    High:      Final position of the area to be ordered.
'    OrderDir:     Direction of the sorting:
'                > 0 -> from the minor to the major.
'                = 0 -> no sorting.
'                < 0 -> from the major to the minor.
'
    Dim RandIndex&, I&, J&, M$
    Dim DoubleValTemp As Double   ' Type of the vector that you want to order.
    Dim POINTAPIValTemp As POINTAPI   ' Type of the vector that you want to order.
    Dim LongValTemp As Long   ' Type of the vector that you want to order.
    Dim Part As Double      ' Type of sorting key.
'
    On Error GoTo QuickSort1Double1POINTAPI1Long_ERR
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
                ' Main Vector:
                DoubleValTemp = ValTab(Low)
                ValTab(Low) = ValTab(High)
                ValTab(High) = DoubleValTemp
                ' First associated vector:
                POINTAPIValTemp = ValTab1(Low)
                ValTab1(Low) = ValTab1(High)
                ValTab1(High) = POINTAPIValTemp
                ' Second associated vector:
                LongValTemp = ValTab2(Low)
                ValTab2(Low) = ValTab2(High)
                ValTab2(High) = LongValTemp
            End If
'
        Else
            ' Pick a pivot element, then move it to the end:
            RandIndex = (High + Low) / 2
            'SWAP ValTab(High), ValTab(RandIndex)
            ' Main Vector:
            DoubleValTemp = ValTab(High)
            ValTab(High) = ValTab(RandIndex)
            ValTab(RandIndex) = DoubleValTemp
            Part = ValTab(High)
            ' First associated vector:
            POINTAPIValTemp = ValTab1(High)
            ValTab1(High) = ValTab1(RandIndex)
            ValTab1(RandIndex) = POINTAPIValTemp
            ' Second associated vector:
            LongValTemp = ValTab2(High)
            ValTab2(High) = ValTab2(RandIndex)
            ValTab2(RandIndex) = LongValTemp
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
                    '  elements on either side are out of order, so swap them:
                    'SWAP ValTab(i), ValTab(J)
                    ' Main Vector:
                    DoubleValTemp = ValTab(I)
                    ValTab(I) = ValTab(J)
                    ValTab(J) = DoubleValTemp
                    ' First associated vector:
                    POINTAPIValTemp = ValTab1(I)
                    ValTab1(I) = ValTab1(J)
                    ValTab1(J) = POINTAPIValTemp
                    ' Second associated vector:
                    LongValTemp = ValTab2(I)
                    ValTab2(I) = ValTab2(J)
                    ValTab2(J) = LongValTemp
                End If
'
            Loop While I < J
            ' Move the pivot element back to its proper place in the array:
            'SWAP ValTab(i), ValTab(High)
            ' Main Vector:
            DoubleValTemp = ValTab(I)
            ValTab(I) = ValTab(High)
            ValTab(High) = DoubleValTemp
            ' First associated vector:
            POINTAPIValTemp = ValTab1(I)
            ValTab1(I) = ValTab1(High)
            ValTab1(High) = POINTAPIValTemp
            ' Second associated vector:
            LongValTemp = ValTab2(I)
            ValTab2(I) = ValTab2(High)
            ValTab2(High) = LongValTemp
'
            ' Recursively call the QuickSort1Double1POINTAPI1Long procedure (pass the smaller
            '  subdivision first to use less stack space):
            If (I - Low) < (High - I) Then
                QuickSort1Double1POINTAPI1Long ValTab(), ValTab1(), ValTab2(), Low, I - 1, OrderDir
                QuickSort1Double1POINTAPI1Long ValTab(), ValTab1(), ValTab2(), I + 1, High, OrderDir
            Else
                QuickSort1Double1POINTAPI1Long ValTab(), ValTab1(), ValTab2(), I + 1, High, OrderDir
                QuickSort1Double1POINTAPI1Long ValTab(), ValTab1(), ValTab2(), Low, I - 1, OrderDir
            End If
        End If
    End If
'
'
QuickSort1Double1POINTAPI1Long_ERR:
    If (Err <> 0) Then
        M$ = "Error " & Str$(Err.Number) & vbNewLine
        M$ = M$ & Err.Description
        MsgBox M$, vbCritical, " QuickSort1Double1POINTAPI1Long"
    End If
'
'
'
End Sub

Private Sub DrawPoints(ByVal bCol As Boolean)
'
' Draw, with API, the points specified in the XV (), YV () and ZV () vectors.
' If bCol = True the points are assigned a color proportional to their height:
'
    Dim N&, hPen&, hPen_O&, hBrush&, hBrush_O&, lR1&
    Dim Xr#, Yr#
    Dim Quadrante&
'
    If bRotate Then
        ' Caculate the rotate vectors:
        For N = 1 To NV
            ROTATE XV(N), YV(N), CDbl(THETA), X0r, Y0r, Xr, Yr
'
            ' Caculate the points projected on the representation plane:
            PRv(N).X = CLng((Ax * Xr + Bx) + (Ay * Yr + By) * CosA)
            PRv(N).Y = CLng((Az * ZV(N) + Bz) - (Ay * Yr + By) * SinA)
        Next N
    Else
        ' Caculate the points projected on the representation plane:
        For N = 1 To NV
            PRv(N).X = CLng((Ax * XV(N) + Bx) + (Ay * YV(N) + By) * CosA)
            PRv(N).Y = CLng((Az * ZV(N) + Bz) - (Ay * YV(N) + By) * SinA)
        Next N
    End If
'
    hPen = CreatePen(vbSolid, 1, vbWhite)
    hPen_O = SelectObject(pic3D.hdc, hPen)
    hBrush = CreateSolidBrush(&H808080) 'pic3D.BackColor)
    hBrush_O = SelectObject(pic3D.hdc, hBrush)
'
    ' Drawing the quadrilaterals. As a function of THETA,
    '  first design those more 'far from the' observer:
    Quadrante = CLng(Int(THETA / PI_2))
'
    Select Case Quadrante
        Case 0, 1   ' 1st Quadrant. 0 ~ 90 and 2nd Quadrant. 90 ~ 180
            ' Drawing the points projected on the representation plane:
            For N = 1 To NV
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(N))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Ellipse pic3D.hdc, PRv(N).X - lRP, PRv(N).Y - lRP, _
                                   PRv(N).X + lRP, PRv(N).Y + lRP
            Next N
        Case 2, 3   ' 3rd Quadrant. 180 ~ 270 and 4th Quadrant. 270 ~ 360
            ' Drawing the points projected on the representation plane:
            For N = NV To 1 Step -1
                If bCol Then
                    lR1 = DeleteObject(hBrush)
                    hBrush = CreateSolidBrush(ZCol(N))
                    lR1 = SelectObject(pic3D.hdc, hBrush)
                End If
'
                Ellipse pic3D.hdc, PRv(N).X - lRP, PRv(N).Y - lRP, _
                                   PRv(N).X + lRP, PRv(N).Y + lRP
            Next N
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

Private Sub Draw(Optional ByVal bChangeView As Boolean = False)
'
'   Draw, with the required parameters, the data points or the surface:
'
    Dim lGridCol&, RLyx!, Rzx!
    
    Dim t0!
    t0 = Timer
'
    XRMin = XMin: XRMax = XMax ' Minimum and maximum values
    YRMin = YMin: YRMax = YMax '  on the axes of the switchboard.
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

Private Sub SearchVertex(ByVal LPx&, ByVal lPy&, ByRef Iu&, ByRef Ju&)
'
'   Search, in the PRv () matrix, the vertex closest to the projected
'    coordinates lPx, lPy and return the column and row indexes.
'   In this matrix the vertices in the foreground are those of the first row
'    and are favored:
'
    Dim I&, J&, DisY&, DisQ&, DisQMin&
'
    DisQMin = 2147483647
    For J = 1 To NYV
        DisY = (YV(J) - YRMin)  ' It favors, in the research,
                                '  the vertices with Y minor.
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
                                   , ByVal lblVal3 As Label, ByVal Val3$ _
                                   , ByVal lblVal4 As Label, ByVal Val4$ _
                                   , ByVal lblVal5 As Label, ByVal Val5$)
'
'   Update the values of the cursor position windows
'    and set the highlight color:
'
    lblVal1 = Val1$
    lblVal2 = Val2$
    lblVal3 = Val3$
    lblVal4 = Val4$
    lblVal5 = Val5$
'
    lblVal1.BackColor = IIf(Len(Val1$) = 0, vbButtonFace, PCHL)
    lblVal2.BackColor = IIf(Len(Val2$) = 0, vbButtonFace, PCHL)
    lblVal3.BackColor = IIf(Len(Val3$) = 0, vbButtonFace, PCHL)
    lblVal4.BackColor = IIf(Len(Val4$) = 0, vbButtonFace, PCHL)
    lblVal5.BackColor = IIf(Len(Val5$) = 0, vbButtonFace, PCHL)
'
'
'
End Sub

Private Function SearchPoint(ByVal LPx&, ByVal lPy&) As Long
'
'   Search, in the vector PRv (), the point closest to the projected
'    coordinates lPx, lPy and returns the index. The vector is explored
'    backwards to find the foreground points first:
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
    SearchPoint = Nu
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
            N = SearchPoint(CLng(LPx.X), CLng(LPx.Y))
'            pic3D.ToolTipText = ""
'            pic3D.ToolTipText = " X = " & Format$(XV(N), "#0.000 ") & _
'                                " Y = " & Format$(YV(N), "#0.000 ") & _
'                                " Z = " & Format$(ZV(N), "#0.000 ")
            UpdateCursorPositions lblAPhi, Format$(PhiV(N), "#0.00 "), _
                                    lblATheta, Format$(ThetaV(N), "#0.00 "), _
                                    lblX, Format$(XV(N), "#0.000 "), _
                                    lblY, Format$(YV(N), "#0.000 "), _
                                    lblZ, Format$(ZV(N), "#0.000 ")
            shpInd.Left = pic3D.ScaleX(PRv(N).X - shpIndOffx, vbPixels, vbUser) _
                          + pic3D.ScaleLeft
            shpInd.Top = pic3D.ScaleY(PRv(N).Y - shpIndOffy, vbPixels, vbUser) _
                         + pic3D.ScaleTop
        Else
            SearchVertex CLng(LPx.X), CLng(LPx.Y), I, J
'            pic3D.ToolTipText = ""
'            pic3D.ToolTipText = " X = " & Format$(XV(I), "#0.000 ") & _
'                                " Y = " & Format$(YV(J), "#0.000 ") & _
'                                " Z = " & Format$(ZV(I, J), "#0.000 ")
            UpdateCursorPositions lblAPhi, Format$(PhiV(J), "#0.00 "), _
                                    lblATheta, Format$(ThetaV(I), "#0.00 "), _
                                    lblX, Format$(XV(I) * 10 ^ XEsp, "#0.000 "), _
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
    ' Coordinates, in [vbUser], of the rotation center:
    X0 = X0r + (Y0r - YRMin) * TrRotX
    Y0 = ZRMin + (Y0r - YRMin) * TrRotY
'
    ' Coordinates, in [vbUser], of the end of the rotation axis:
    y1 = Y0 + AsseZ
    pic3D.DrawStyle = vbDashDot
    pic3D.Line (X0, Y0)-(X0, y1)
'
    ' Coordinates, in [vbUser], of the end of the rotating X axis:
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
