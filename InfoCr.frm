VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmCopyright 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Informazioni su ..."
   ClientHeight    =   3870
   ClientLeft      =   2325
   ClientTop       =   1635
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "InfoCr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4080
      Top             =   240
   End
   Begin VB.CommandButton cmdInformazioniSistema 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   3300
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "web: http://www.flanguasco.org"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2940
      Width           =   3015
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "InfoCr.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblProgramma 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  ?
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblEMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "e-Mail:  MC7061@mclink.it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2700
      Width           =   2175
   End
   Begin VB.Label lblTelefono 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Roma  -  Tel: 5449322"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2460
      Width           =   2175
   End
   Begin VB.Label lblAutore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Franco Languasco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblDescrizione 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione...  (App.FileDescription)"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersione 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Versione: "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblCommenti 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App.Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   1410
   End
End
Attribute VB_Name = "frmCopyright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================
' Description......: Form, of general use, for displaying
'                     information on the program and on the
'                     system.
' Name of the Files: InfoCr.frm, InfoCr.frx
' Date.............: 27/2/1998
' Version..........: 2.0 at 32 bits
' System...........: Visual Basic 6.0 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.virgilio.it/flanguasco/
'                    http://www.flanguasco.org
'==============================================================
'
'   Nel Form di origine usare il codice:
'    frmCopyright.OpenForm [IconForm as Form] [,Stile] [,TApp]
'     IconForm: Form da cui si vuole copiare
'               l' Icona di programma [opzionale].
'     Stile:    Modo di rappresentazione di frmCopyright;
'               mettere Stile = vbModeless per usare
'               frmCopyright come Splash Form.
'               [opzionale; default = vbModal].
'     TApp:     Da usare solo con Stile = vbModeless
'               per stabilire la durata di visione del Form [s].
'               [opzionale; default = 0].
'
'   Le informazioni, usate per popolare il Form,
'   vengono prese dalle Project Properties / Make.
'
Option Explicit
'
Dim MeStile&    ' Modo di rappresentazione del Form.
                ' Puo' essere solo:
                '  0 = vbModeless
                '  1 = vbModal
Dim MeTApp&     ' Tempo di visione del Form quando
                ' usato come Splash Form [s].
'
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
              KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
              KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1        ' Unicode nul terminated string
Const REG_DWORD = 4     ' 32-bit number
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
'
Const HWND_TOPMOST& = -1
Const SWP_NOSIZE& = &H1
Const SWP_NOMOVE& = &H2
Const SWP_NOACTIVATE& = &H10
Const SWP_DRAWFRAME& = &H20
Const SWP_SHOWWINDOW& = &H40
'
Private Declare Function GetEnvironmentVariable Lib "kernel32" _
    Alias "GetEnvironmentVariableA" (ByVal lpname As String, _
    ByVal lpBuffer As String, ByVal nSize As Long) As Long
'
Private Declare Function RegOpenKeyEx Lib "advapi32" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, _
    ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" _
    (ByVal hKey As Long) As Long
'
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'
'
    Dim I As Long           ' Loop Counter
    Dim rc As Long          ' Return Code
    Dim hKey As Long        ' Handle To An Open Registry Key
    Dim hDepth As Long      '
    Dim KeyValType As Long  ' Data Type Of A Registry Key
    Dim tmpVal As String    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    ' Open Registry Key:
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
'
    ' Handle Error...
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
'
    tmpVal = String$(1024, 0)   ' Allocate Variable Space
    KeyValSize = 1024           ' Mark Variable Size
'
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
    KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
'
    ' Handle Errors:
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
'
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then   ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)       ' Null Found, Extract From String
    Else                                            ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)           ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" & KeyVal)                     ' Convert Double Word To String
    End Select
'
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
'
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
'
'
'
End Function

Private Sub StartSysInfo()
'
'
    Dim rc As Long
    Dim SysInfoPath As String
'
    On Error GoTo SysInfoErr
'
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO _
    , gREGVALSYSINFO, SysInfoPath) Then
'
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC _
    , gREGVALSYSINFOLOC, SysInfoPath) Then
'
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
'
    Shell SysInfoPath, vbNormalFocus
    
    Exit Sub
'
'
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
'
'
'
End Sub

Private Function GetSystemInfo$(ByVal Info$)
'
'
    Dim dl&, BUF$
'
    On Error GoTo GetSystemInfo_ERR
'
    BUF$ = String$(2048, 0)
    dl = GetEnvironmentVariable(Info$, BUF$, 2047)
    If dl > 0 Then
        GetSystemInfo$ = Left$(BUF$, dl)
    Else
        GetSystemInfo$ = ""
    End If
'
    Exit Function
'
GetSystemInfo_ERR:
    GetSystemInfo$ = ""
'
'
'
End Function

Private Sub cmdInformazioniSistema_Click()
'
'
    StartSysInfo
'
'
'
End Sub

Private Sub cmdOK_Click()
'
'
    Unload Me
'
'
'
End Sub

Private Sub Form_Load()
'
'
    Dim X&, Y&, R&
'
    On Error GoTo Form_Load_ERR
'
    If MeStile = vbModeless Then
        X = Me.Left
        Y = Me.Top
        R = SetWindowPos(Me.hWnd, HWND_TOPMOST _
        , X, Y, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
'
        If MeTApp > 0 Then
            Timer1.Interval = 1000& * MeTApp
            Timer1.Enabled = True
        End If
    End If
'
    Me.Caption = " Informazioni su " & App.EXEName
    lblProgramma = App.ProductName & lblProgramma
    lblCommenti = App.Comments
    lblVersione = lblVersione & App.Major & "." & App.Minor
    lblDescrizione = App.FileDescription
    lblAutore = App.LegalCopyright
'
    MMControl1.Command = "close"
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = GetSystemInfo$("windir") + "\Media\Tada.wav"
    MMControl1.Command = "open"
    MMControl1.Command = "play"
'
'
Form_Load_ERR:
'
'
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
'
    MMControl1.Command = "close"
    Set frmCopyright = Nothing
'
'
'
End Sub

Public Sub OpenForm(Optional ByVal IconForm As Form = Nothing _
    , Optional ByVal Stile& = vbModal, Optional ByVal TApp& = 0)
'
'
    MeStile = Stile
    MeTApp = TApp
'
    If Not IconForm Is Nothing Then
        imgIcon.Picture = IconForm.Icon
    End If
'
    Me.Show Stile
'
'
'
End Sub

Private Sub Timer1_Timer()
'
'
    Unload Me
'
'
'
End Sub
