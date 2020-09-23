Attribute VB_Name = "mInercia"
Option Explicit
Option Base 1
Private momInercia(1 To 3) As Double
Private vInercia(1 To 3, 1 To 3) As Double


' +-------------------------------------------------------------------
' | TENSOR DE INERCIA
' +-------------------------------------------------------------------

Public Sub TensorDeInercia()
Dim i As Integer, j As Integer, k As Integer
Dim Tensor(1 To 3, 1 To 3) As Double
Dim a(1 To 3, 1 To 3) As Double
Dim xterm As Double, yterm As Double, zterm As Double
Dim dot As Double
Dim xx As Double, yy As Double, zz As Double
Dim xy As Double, xz As Double, yz As Double

Dim masa As Double
For i = 1 To nAtoms
    With atoms(i)
        masa = .numeroAtomico 'Hay algun problema?
        xterm = .x
        yterm = .y
        zterm = .Z
    End With
    xx = xx + xterm * xterm * masa
    xy = xy + xterm * yterm * masa
    xz = xz + xterm * zterm * masa
    yy = yy + yterm * yterm * masa
    yz = yz + yterm * zterm * masa
    zz = zz + zterm * zterm * masa
Next i
    Tensor(1, 1) = yy + zz
    Tensor(2, 1) = -xy
    Tensor(3, 1) = -xz
    Tensor(1, 2) = -xy
    Tensor(2, 2) = xx + zz
    Tensor(3, 2) = -yz
    Tensor(1, 3) = -xz
    Tensor(2, 3) = -yz
    Tensor(3, 3) = xx + yy

' +-------------------------------------------------------------------
    Jacobi 3, Tensor(), momInercia(), vInercia()
' +-------------------------------------------------------------------

For i = 1 To 2
    For j = 1 To nAtoms
        With atoms(j)
            xterm = vInercia(1, i) * (.x)
            yterm = vInercia(2, i) * (.y)
            zterm = vInercia(3, i) * (.Z)
        End With
        dot = xterm + yterm + zterm
        If dot < 0# Then
            For k = 1 To 3
                vInercia(k, i) = -vInercia(k, i)
            Next k
        End If
        If dot <> 0# Then GoTo 10
    Next j
10:
Next i

' +-------------------------------------------------------------------
xterm = vInercia(1, 1) * (vInercia(2, 2) * vInercia(3, 3) - vInercia(2, 3) * vInercia(3, 2))
yterm = vInercia(2, 1) * (vInercia(1, 3) * vInercia(3, 2) - vInercia(1, 2) * vInercia(3, 3))
zterm = vInercia(3, 1) * (vInercia(1, 2) * vInercia(2, 3) - vInercia(1, 3) * vInercia(2, 2))

dot = xterm + yterm + zterm

If dot < 0# Then
    For j = 1 To 3
        vInercia(j, 3) = -vInercia(j, 3)
    Next j
End If

For i = 1 To 3
    For j = 1 To 3
        a(i, j) = vInercia(j, i)
    Next j
Next i

' copiamos el vector momento de inercia a unos ejes estandar
vMomentoInercia1.x = vInercia(1, 1)
vMomentoInercia1.y = vInercia(1, 2)
vMomentoInercia1.Z = vInercia(1, 3)
vMomentoInercia2.x = vInercia(2, 1)
vMomentoInercia2.y = vInercia(2, 2)
vMomentoInercia2.Z = vInercia(2, 3)
vMomentoInercia3.x = vInercia(3, 1)
vMomentoInercia3.y = vInercia(3, 2)
vMomentoInercia3.Z = vInercia(3, 3)
'trasladar al origen
    For i = 1 To nAtoms
        With atoms(i)

            xterm = .x
            yterm = .y
            zterm = .Z

            .x = a(1, 1) * xterm + a(1, 2) * yterm + a(1, 3) * zterm
            .y = a(2, 1) * xterm + a(2, 2) * yterm + a(2, 3) * zterm
            .Z = a(3, 1) * xterm + a(3, 2) * yterm + a(3, 3) * zterm

        End With
    Next i

End Sub

'**********************************************************
' Aquesta rutina ens centra la molecula en el seu centre
' de masses.
' faig proporcional la massa al numero atomic (!)
'**********************************************************
Public Sub CentratCDM()
Dim i As Integer
Dim masaTotal As Double
    xcm = 0: ycm = 0: zcm = 0
    masaTotal = 0
' calculem el baricentre de la molecula
    For i = 1 To nAtoms
        With atoms(i)
            xcm = xcm + .numeroAtomico * .x
            ycm = ycm + .numeroAtomico * .y
            zcm = zcm + .numeroAtomico * .Z
            masaTotal = masaTotal + .numeroAtomico
        End With
    Next i
   
' calculem el centre de masses
    xcm = xcm / masaTotal
    ycm = ycm / masaTotal
    zcm = zcm / masaTotal
    
' centrem les coordenades
    For i = 1 To nAtoms
        With atoms(i)
            .x = .x - xcm
            .y = .y - ycm
            .Z = .Z - zcm
        End With
    Next i
End Sub

Public Function baricentro(atomos() As Integer) As TVECTOR
Dim atomo 'variant
Dim atomoSel As TATOMO
Dim nAtom As Integer
Dim xcm As Double, ycm As Double, zcm As Double
xcm = 0: ycm = 0: zcm = 0
nAtom = UBound(atomos) 'numero de atomos en la matriz
For Each atomo In atomos()
    If atomo <> 0 Then
        atomoSel = atoms(atomo)
        With atomoSel
                xcm = xcm + .x
                ycm = ycm + .y
                zcm = zcm + .Z
        End With
    End If
Next
baricentro.x = xcm / nAtom
baricentro.y = ycm / nAtom
baricentro.Z = zcm / nAtom
End Function
