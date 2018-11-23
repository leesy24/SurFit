VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInstructions 
   AutoRedraw      =   -1  'True
   Caption         =   " Instructions:"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "frmInstr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   7680
   Begin RichTextLib.RichTextBox rtfIstruzioni 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmInstr.frx":0442
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
' Description......: Form, of general use, for displaying an
'                     instruction file.
' Name of the Files: frmInstr.frm, frmInstr.frx
' Date.............: 15/9/1999
' Version..........: 1.0 at 32 bits.
' System...........: Visual Basic 6.0 under Windows NT 4.0
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.it/flanguasco/
'                    http://www.flanguasco.org
'================================================================
'
'   Nel Form di origine usare il codice:
'    frmInstructions.Apri FileIstruzioni$ [, Title$] [, Posizione&] [, SForm as Form]
'
'     FileIstruzioni$:    Nome, completo di path e di tipo .txt o .rtf,
'                         con le istruzioni da visualizzare.
'     Title$:            Caption di questo Form [opzionale].
'     Posizione:          Posizione iniziale di questo Form [opzionale]
'                         e puo' essere solo:
'                          0 = in alto, a destra, del Form di origine.
'                          1 = in centro al Form di origine (vbStartUpOwner).
'                          2 = in centro allo schermo o al Form MDI
'                              ospitante (default) (vbStartUpScreen).
'                          3 = in alto, a sinistra dello schermo o del
'                              Form MDI ospitante (vbStartUpWindowsDefault).
'                          Se Posizione = 0  o Posizione = 1
'                          deve essere anche presente:
'     SForm:              Form da cui e' stata lanciata la richiesta.
'
Option Explicit
'
Dim FileIstruzioni$
Dim SForm As Form
Dim Posizione&
'
Private Type RECT_Type
    Left  As Long
    Top  As Long
    Right  As Long
    Bottom  As Long
End Type
'
Private Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long _
, lpRect As RECT_Type) As Long
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

Private Sub Form_Activate()
'
'
    If FileExists(FileIstruzioni$) Then
        rtfIstruzioni.LoadFile FileIstruzioni$
    Else
        MsgBox "Manca il file " & FileIstruzioni$, vbCritical _
        , " Error in Instructions"
        Unload Me
    End If
'
'
'
End Sub
Private Sub Form_Load()
'
'
    Dim X&, Y&, PForm_Width&, PForm_Height&
    Dim Rect As RECT_Type
'
    On Error Resume Next
'
    Select Case Posizione
        Case 0
        ' Posizionamento in alto a destra del Form di origine:
        X = SForm.Left + SForm.Width - Me.Width
        Y = SForm.Top
'
        Case vbStartUpOwner
        ' Posizionamento in centro al Form di origine:
        X = SForm.Left + (SForm.Width - Me.Width) / 2
        Y = SForm.Top + (SForm.Height - Me.Height) / 2
'
        Case vbStartUpScreen
        ' Posizionamento in centro allo schermo o all' MDI:
        If Me.MDIChild Then
            GetWindowRect GetParent(Me.hWnd), Rect
            PForm_Width = Abs(Rect.Right - Rect.Left) _
            * Screen.TwipsPerPixelX
            PForm_Height = Abs(Rect.Bottom - Rect.Top) _
            * Screen.TwipsPerPixelY
            X = (PForm_Width - Me.Width) / 2
            Y = (PForm_Height - Me.Height) / 2
        Else
            X = (Screen.Width - Me.Width) / 2
            Y = (Screen.Height - Me.Height) / 2
        End If
'
        Case vbStartUpWindowsDefault
        ' Posizionamento in alto a sinistra dello schermo
        ' o dell' MDI:
        X = 0
        Y = 0
    End Select
'
    Me.Move X, Y
'
'
'
End Sub

Private Sub Form_Resize()
'
'
    Dim rtfIsW&, rtfIsH&
'
    If Me.WindowState <> vbMinimized Then
        rtfIsW = Me.Width - 120
        If rtfIsW < 1 Then rtfIsW = 1
        rtfIsH = Me.Height - 480
        If rtfIsH < 1 Then rtfIsH = 1
'
        rtfIstruzioni.Move 0, 0, rtfIsW, rtfIsH
    End If
'
'
'
End Sub



Public Sub Apri(ByVal NFileIstruzioni$, Optional ByVal Title$ = "Instructions", _
    Optional ByVal NPosizione& = 2, Optional ByVal NSform As Form)
'
'
    FileIstruzioni$ = NFileIstruzioni$
    Posizione = NPosizione
    Set SForm = NSform
    Me.Caption = " " & Title$
'
    Me.Show
'
'
'
End Sub
