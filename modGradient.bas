Attribute VB_Name = "modGradient"
'=============================================================
' Description......: Calculation routine, with the finite
'                     difference method, of the gradient of a
'                     surface.
' Name of the Files: modGradient.bas
' Date.............: 21/9/2001
' Version..........: 1.0 at 32 bits.
' System...........: VB6 under Windows NT.
' Written by.......: F. Languasco
' E-Mail...........: MC7061@mclink.it
' Download by......: http://members.xoom.it/flanguasco/
'                    http://www.flanguasco.org
'=============================================================
'
'   Note:   All the vectors and matrices of these routines start from index 1.
'
Option Explicit
'
Public Type Grad_Type   ' Horizontal and
    DX As Double        '  vertical
    DY As Double        '  components of
End Type                '  the gradient.

Public Sub Gradient_2D(XI#(), YI#(), ZI#(), ByVal NXI&, ByVal NYI&, Grad() As Grad_Type)
'
'   Calculate the gradient of a surface.
'   Input parameters:
'    XI(1 To NXI):              vector of the abscissas of the surface.
'    YI(1 To NYI):              vector of the ordinates of the surface.
'    ZI(1 To NXI, 1 To NYI):    matrix of the values of the surface.
'    NXI:                       Number of columns in the grid ZI() (NXI> = 3).
'    NYI:                       Number of rows in the grid ZI() (NYI> = 3).
'   Output parameters:
'    Grad(1 To NXI, 1 To NYI):  matrix of the horizontal and vertical
'                                components of the gradient.
'
'   Note: use this routine ONLY for grids of ZI() with abscissae XI() and
'          order YI() equispaced.
'
    Dim I&, J&, dDenx1#, dDeny1#, dDenx2#, dDeny2#
'
    ' Calculate the differences on rows and columns:
    dDenx1 = XI(2) - XI(1)
    dDenx2 = XI(3) - XI(1)
    dDeny1 = YI(2) - YI(1)
    dDeny2 = YI(3) - YI(1)
'
    ' Calculate the gradient of the central part:
    For J = 2 To NYI - 1
        For I = 2 To NXI - 1
            Grad(I, J).DX = (ZI(I + 1, J) - ZI(I - 1, J)) / dDenx2
            Grad(I, J).DY = (ZI(I, J + 1) - ZI(I, J - 1)) / dDeny2
        Next I
    Next J
'
    ' Calculate the gradient of the two rows at the top and bottom:
    For I = 2 To NXI - 1
        Grad(I, 1).DX = (ZI(I + 1, 1) - ZI(I - 1, 1)) / dDenx2
        Grad(I, 1).DY = (ZI(I, 2) - ZI(I, 1)) / dDeny1
'
        Grad(I, NYI).DX = (ZI(I + 1, NYI) - ZI(I - 1, NYI)) / dDenx2
        Grad(I, NYI).DY = (ZI(I, NYI) - ZI(I, NYI - 1)) / dDeny1
    Next I
'
    ' Calculate the gradient of the two columns on the right and on the left:
    For J = 2 To NYI - 1
        Grad(1, J).DX = (ZI(2, J) - ZI(1, J)) / dDenx1
        Grad(1, J).DY = (ZI(1, J + 1) - ZI(1, J - 1)) / dDeny2
'
        Grad(NXI, J).DX = (ZI(NXI, J) - ZI(NXI - 1, J)) / dDenx1
        Grad(NXI, J).DY = (ZI(NXI, J + 1) - ZI(NXI, J - 1)) / dDeny2
    Next J
'
    ' Calculate the gradient at the four corners:
    Grad(1, 1).DX = (ZI(2, 1) - ZI(1, 1)) / dDenx1                  ' Bottom left
    Grad(1, 1).DY = (ZI(1, 2) - ZI(1, 1)) / dDeny1                  '  corner.
'
    Grad(NXI, 1).DX = (ZI(NXI, 1) - ZI(NXI - 1, 1)) / dDenx1        ' Bottom right
    Grad(NXI, 1).DY = (ZI(NXI, 2) - ZI(NXI, 1)) / dDeny1            '  corner.
'
    Grad(NXI, NYI).DX = (ZI(NXI, NYI) - ZI(NXI - 1, NYI)) / dDenx1  ' Top right
    Grad(NXI, NYI).DY = (ZI(NXI, NYI) - ZI(NXI, NYI - 1)) / dDeny1  '  corner.
'
    Grad(1, NYI).DX = (ZI(2, NYI) - ZI(1, NYI)) / dDenx1            ' Top left
    Grad(1, NYI).DY = (ZI(1, NYI) - ZI(1, NYI - 1)) / dDeny1        '  corner.
'
'
'
End Sub
