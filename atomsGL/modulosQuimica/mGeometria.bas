Attribute VB_Name = "mGeometria"
Option Explicit
Option Base 1

'modulo de calculo de planos por minimos cuadrados

'Public puntos(10) As TVECTOR
'Public numeroDeDatos As Integer
'**********************************************************
'
'**********************************************************
Public Function planoMinCuad() As TPlano

Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double
Dim a5 As Double, a6 As Double, a7 As Double, a8 As Double
Dim a9 As Double, b1 As Double, b2 As Double, b3 As Double
Dim dt As Double, dtx As Double, dty As Double, dtz As Double
Dim dt1 As Double, dt2 As Double, dt3 As Double
Dim ak1 As Double

Dim i As Integer
a1 = UBound(atmsSelecds())
If a1 < 3 Then Exit Function
Dim atmSelcc
For Each atmSelcc In atmsSelecds()
    With atoms(atmSelcc)
        a2 = a2 + .x
        a4 = a2
        a5 = a5 + .x ^ 2
        a6 = a6 + .x * .y
        a8 = a6
        a3 = a3 + .y
        a7 = a3
        a9 = a9 + .y ^ 2
        b1 = b1 + .Z
        b2 = b2 + .Z * .x
        b3 = b3 + .Z * .y
    End With
Next

     dt = a1 * (a5 * a9 - a6 * a8) - a2 * (a4 * a9 - a6 * a7) + a3 * (a4 * a8 - a5 * a7)
    dtx = b1 * (a5 * a9 - a6 * a8) - a2 * (b2 * a9 - a6 * b3) + a3 * (b2 * a8 - a5 * b3)
    dty = a1 * (b2 * a9 - a6 * b3) - b1 * (a4 * a9 - a6 * a7) + a3 * (a4 * b3 - b2 * a7)
    dtz = a1 * (a5 * b3 - b2 * a8) - a2 * (a4 * b3 - b2 * a7) + b1 * (a4 * a8 - a5 * a7)
dt1 = dtx / dt
dt2 = dty / dt
dt3 = dtz / dt

ak1 = Sqr(1 + dt2 ^ 2 + dt3 ^ 2)

planoMinCuad.a = -dt2 / ak1
planoMinCuad.b = -dt3 / ak1
planoMinCuad.c = 1 / ak1
planoMinCuad.d = dt1 / ak1

End Function

'**********************************************************
' Dist2Plane
'**********************************************************
Public Function Dist2Plane(punto As TATOMO, plano As TPlano) As Double
    Dist2Plane = ((plano.a * punto.x + _
                          plano.b * punto.y + _
                          plano.c * punto.Z) - plano.d) / _
                          Sqr(plano.a ^ 2 + plano.b ^ 2 + plano.c ^ 2)
End Function
'**********************************************************
'
'**********************************************************
Public Function distanciaAB(p1 As TATOMO, p2 As TATOMO) As Double
distanciaAB = Sqr((p1.x - p2.x) ^ 2 + _
                  (p1.y - p2.y) ^ 2 + _
                  (p1.Z - p2.Z) ^ 2)
End Function
'**********************************************************
'
'**********************************************************
Public Function distanciaABV(p1 As TVECTOR, p2 As TVECTOR) As Double
distanciaABV = Sqr((p1.x - p2.x) ^ 2 + _
                  (p1.y - p2.y) ^ 2 + _
                  (p1.Z - p2.Z) ^ 2)
End Function

'**********************************************************
' angle entre tres punts
'**********************************************************
Public Function angleABC(p1 As TATOMO, p2 As TATOMO, p3 As TATOMO)
Dim a As TATOMO, b As TATOMO
a = VectorDif(p1, p2)
b = VectorDif(p3, p2)

If modulo(a) = 0 Or modulo(b) = 0 Then Exit Function

angleABC = Arccos(dot(a, b) / (modulo(a) * modulo(b))) 'uso del modeulo private arccos
angleABC = angleABC * 180 / PI
End Function

'**********************************************************
'           ANGULO(diedro())
'
'       el vector del plano que pasa
'       por tres puntos es igual al producto vectorial
'       de los vectores diferencia de los puntos:
'       a-b, b-c  --->  Vplano = (a-b) ^ (b-c)
'       calculo el angulo entre los vectores de los planos
'       correspondientes
'**********************************************************

Public Function AngleDiedre(p1 As TATOMO, p2 As TATOMO, _
                            p3 As TATOMO, p4 As TATOMO) As Double
                            
Dim a As TATOMO, b As TATOMO, c As TATOMO
Dim aa As TATOMO, bb As TATOMO

a = VectorDif(p1, p2)
b = VectorDif(p3, p2)
c = VectorDif(p3, p4)

aa = VectorProdVect(a, b)
bb = VectorProdVect(b, c)

If modulo(aa) = 0 Or modulo(bb) = 0 Then Exit Function

AngleDiedre = Arccos(dot(aa, bb) / (modulo(aa) * modulo(bb)))
AngleDiedre = AngleDiedre * 180 / PI

End Function

'**********************************************************
'   Angle entre dos plans
'   entrada pla: p1, pla: p2
'  sortida angle en graus
'**********************************************************
Public Function angleEntrePlans(p1 As TPlano, p2 As TPlano) As Double

Dim a As TATOMO, b As TATOMO
a.x = p1.a: a.y = p1.b: a.Z = p1.c
b.x = p2.a: b.y = p2.b: b.Z = p2.c
If modulo(a) = 0 Or modulo(b) = 0 Then Exit Function
angleEntrePlans = Arccos(dot(a, b) / (modulo(a) * modulo(b))) 'uso del modeulo private arccos
angleEntrePlans = angleEntrePlans * 180 / PI
'If angleEntrePlans > 90 Then angleEntrePlans = angleEntrePlans - 90
End Function

'**********************************************************
' entrada vector v1, vector v2
' sortida VectorDif
'**********************************************************
Private Function VectorDif(a As TATOMO, b As TATOMO) As TATOMO
    VectorDif.x = a.x - b.x
    VectorDif.y = a.y - b.y
    VectorDif.Z = a.Z - b.Z
End Function

'**********************************************************
'
'**********************************************************
Private Function VectorProdVect(a As TATOMO, b As TATOMO) As TATOMO
    VectorProdVect.x = a.y * b.Z - a.Z * b.y
    VectorProdVect.y = a.x * b.Z - a.Z * b.x
    VectorProdVect.Z = a.x * b.y - a.y * b.x
End Function

'**********************************************************
'
'**********************************************************
Public Function dot(a As TATOMO, b As TATOMO) As Double
    dot = a.x * b.x + a.y * b.y + a.Z * b.Z
End Function

'**********************************************************
'
'**********************************************************
Public Function VectorNorm(vector As TATOMO) As TATOMO
Dim r As Double, m As Double
    m = modulo(vector)
    VectorNorm.x = vector.x / m
    VectorNorm.y = vector.y / m
    VectorNorm.Z = vector.Z / m

End Function

'**********************************************************
'
'**********************************************************
Private Function modulo(vector As TATOMO)
modulo = Sqr(vector.x ^ 2 + vector.y ^ 2 + vector.Z ^ 2)
End Function

' +-------------------------------------------------------------------
' | Arccos (salida en radianes)
' +-------------------------------------------------------------------
Public Function Arccos(x As Single) As Single
Select Case x
    Case Is >= 1
        Arccos = 0
    Case Is <= -1
        Arccos = PI
    Case Else
        Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Select
End Function
'**********************************************************
'
'**********************************************************
Public Sub frac2cart()
'Clacularemos las coordenadas cartesianas
Dim i As Integer
Dim cosAlpha As Double, cosBeta As Double
Dim cosGamma As Double, sinGamma As Double
Dim gg1 As Double, gg2 As Double
cosAlpha = Cos(celda.alfa * PI / 180)
cosBeta = Cos(celda.beta * PI / 180)
cosGamma = Cos(celda.gamma * PI / 180)
sinGamma = Sin(celda.gamma * PI / 180)
gg1 = (cosAlpha - cosBeta * cosGamma) / sinGamma
gg2 = Sqr(1 - cosBeta * cosBeta - gg1 * gg1)
For i = 1 To nAtoms
    With atoms(i)
        .x = celda.a * .fracX + celda.b * .fracY * cosGamma + celda.c * .fracZ * cosBeta
        .y = celda.b * .fracY * sinGamma + celda.c * gg1 * .fracZ
        .Z = celda.c * gg2 * atoms(i).fracZ
    End With
Next i
End Sub

'**********************************************************
'
'**********************************************************
Public Function planoMinCuad2() As TPlano
Dim i As Integer, j As Integer, k As Integer
Dim Tensor(1 To 3, 1 To 3) As Double
Dim a(1 To 3, 1 To 3) As Double
Dim dx As Double, dy As Double, dz As Double
Dim dot As Double
Dim xx As Double, yy As Double, zz As Double
Dim xy As Double, xz As Double, yz As Double
Dim eigenValues(1 To 3) As Double
Dim eigenVectors(1 To 3, 1 To 3) As Double
Dim atmSelcc
For Each atmSelcc In atmsSelecds()
    With atoms(atmSelcc)
        dx = .x - baricentroPlano.x
        dy = .y - baricentroPlano.y
        dz = .Z - baricentroPlano.Z
    End With
    
    Tensor(1, 1) = Tensor(1, 1) + dx * dx
    Tensor(2, 1) = Tensor(2, 1) + dx * dy
    Tensor(3, 1) = Tensor(1, 1) + dx * dz
    Tensor(1, 2) = Tensor(1, 2) + dx * dy
    Tensor(2, 2) = Tensor(2, 2) + dy * dy
    Tensor(3, 2) = Tensor(3, 2) + dy * dz
    Tensor(1, 3) = Tensor(1, 3) + dx * dz
    Tensor(2, 3) = Tensor(2, 3) + dy * dz
    Tensor(3, 3) = Tensor(3, 3) + dz * dz
Next
' +-------------------------------------------------------------------
    Jacobi 3, Tensor(), eigenValues(), eigenVectors()
' +-------------------------------------------------------------------
'    Debug.Print "EigenValues"
'    Call printVector(eigenValues())
'    Debug.Print "EigenVectors"
'    Call printMatrix(eigenVectors())
    eigsrt eigenValues(), eigenVectors(), 3, 3
'    Debug.Print "#######################"
'    Debug.Print "EigenValues"
'    Call printVector(eigenValues())
'    Debug.Print "EigenVectors"
'    Call printMatrix(eigenVectors())
    planoMinCuad2.a = eigenVectors(1, 3)
    planoMinCuad2.b = eigenVectors(2, 3)
    planoMinCuad2.c = eigenVectors(3, 3)
    planoMinCuad2.d = baricentroPlano.x * eigenVectors(1, 3) + _
                      baricentroPlano.y * eigenVectors(2, 3) + _
                      baricentroPlano.Z * eigenVectors(3, 3)
End Function

' +-------------------------------------------------------------------
' + Ordenar valores propios y vectores propios
' +-------------------------------------------------------------------

Private Sub eigsrt(d() As Double, v() As Double, n As Integer, np As Integer)
Dim i, j, k, p
For i = 1 To n - 1
    k = i
    p = d(i)
    For j = i + 1 To n
        If d(j) >= p Then
            k = j
            p = d(j)
        End If
    Next
    If k <> i Then
        d(k) = d(i)
        d(i) = p
        For j = 1 To n
            p = v(j, i)
            v(j, i) = v(j, k)
            v(j, k) = p
        Next
    End If
Next
End Sub
'**********************************************************
'
'**********************************************************
Public Sub BorrarPlano(ByRef p As TPlano)
p.a = 0: p.b = 0: p.c = 0: p.d = 0
End Sub


