Attribute VB_Name = "modGradiente"
'=============================================================
' Descrizione.....: Routine di calcolo, con il metodo delle
'                   differenze finite, del gradiente di una
'                   superficie.
' Nome dei Files..: modGradiente.bas
' Data............: 21/9/2001
' Versione........: 1.0 a 32 bits.
' Sistema.........: VB6 sotto Windows NT.
' Scritto da......: F. Languasco 
' E-Mail..........: MC7061@mclink.it
' DownLoads a.....: http://members.xoom.it/flanguasco/
'                   http://www.flanguasco.org
'=============================================================
'
'   Nota:   Tutti i vettori e le matrici di queste routines
'           iniziano dall' indice 1.
'
Option Explicit
'
Public Type Grad_Type   ' Componenti
    DX As Double        ' orizzontali
    DY As Double        ' e verticali
End Type                ' del gradiente.

Public Sub Gradient_2D(XI#(), YI#(), ZI#(), ByVal NXI&, ByVal NYI&, Grad() As Grad_Type)
'
'   Calcola il gradiente di una superficie.
'   Parametri in ingresso:
'    XI(1 To NXI):              vettore delle ascisse della superficie.
'    YI(1 To NYI):              vettore delle ordinate della superficie.
'    ZI(1 To NXI, 1 To NYI):    matrice dei valori della superficie.
'    NXI:                       N?di colonne nella griglia di ZI() (NXI >= 3).
'    NYI:                       N?di righe nella griglia di ZI() (NYI >= 3).
'   Parametri in uscita:
'    Grad(1 To NXI, 1 To NYI):  matrice delle componenti orizzontali
'                               e verticali del gradiente.
'
'   Nota: usare questa routine SOLO per griglie di ZI() con ascisse XI()
'         ed ordinate YI() equispaziate.
'
    Dim I&, J&, dDenx1#, dDeny1#, dDenx2#, dDeny2#
'
    ' Calcola le differenze su righe e colonne:
    dDenx1 = XI(2) - XI(1)
    dDenx2 = XI(3) - XI(1)
    dDeny1 = YI(2) - YI(1)
    dDeny2 = YI(3) - YI(1)
'
    ' Calcola il gradiente della parte centrale:
    For J = 2 To NYI - 1
        For I = 2 To NXI - 1
            Grad(I, J).DX = (ZI(I + 1, J) - ZI(I - 1, J)) / dDenx2
            Grad(I, J).DY = (ZI(I, J + 1) - ZI(I, J - 1)) / dDeny2
        Next I
    Next J
'
    ' Calcola il gradiente delle due righe in alto ed in basso:
    For I = 2 To NXI - 1
        Grad(I, 1).DX = (ZI(I + 1, 1) - ZI(I - 1, 1)) / dDenx2
        Grad(I, 1).DY = (ZI(I, 2) - ZI(I, 1)) / dDeny1
'
        Grad(I, NYI).DX = (ZI(I + 1, NYI) - ZI(I - 1, NYI)) / dDenx2
        Grad(I, NYI).DY = (ZI(I, NYI) - ZI(I, NYI - 1)) / dDeny1
    Next I
'
    ' Calcola il gradiente delle due colonne a destra ed a sinistra:
    For J = 2 To NYI - 1
        Grad(1, J).DX = (ZI(2, J) - ZI(1, J)) / dDenx1
        Grad(1, J).DY = (ZI(1, J + 1) - ZI(1, J - 1)) / dDeny2
'
        Grad(NXI, J).DX = (ZI(NXI, J) - ZI(NXI - 1, J)) / dDenx1
        Grad(NXI, J).DY = (ZI(NXI, J + 1) - ZI(NXI, J - 1)) / dDeny2
    Next J
'
    ' Calcola il gradiente ai quattro angoli:
    Grad(1, 1).DX = (ZI(2, 1) - ZI(1, 1)) / dDenx1                  ' Angolo in basso
    Grad(1, 1).DY = (ZI(1, 2) - ZI(1, 1)) / dDeny1                  ' a sinistra.
'
    Grad(NXI, 1).DX = (ZI(NXI, 1) - ZI(NXI - 1, 1)) / dDenx1        ' Angolo in basso
    Grad(NXI, 1).DY = (ZI(NXI, 2) - ZI(NXI, 1)) / dDeny1            ' a destra.
'
    Grad(NXI, NYI).DX = (ZI(NXI, NYI) - ZI(NXI - 1, NYI)) / dDenx1  ' Angolo in alto
    Grad(NXI, NYI).DY = (ZI(NXI, NYI) - ZI(NXI, NYI - 1)) / dDeny1  ' a destra.
'
    Grad(1, NYI).DX = (ZI(2, NYI) - ZI(1, NYI)) / dDenx1            ' Angolo in alto
    Grad(1, NYI).DY = (ZI(1, NYI) - ZI(1, NYI - 1)) / dDeny1        ' a sinistra.
'
'
'
End Sub
