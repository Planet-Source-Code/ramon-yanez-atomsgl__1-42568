Attribute VB_Name = "mSimetria"
Option Explicit
Const tolDist = 0.1
Public GrupPuntual As String

Public Type TEJESSIM
    'coordenadas del vector que define el eje
    x As Double
    y As Double
    Z As Double
    tipo As String * 5
    visible As Boolean
End Type

Public ElemSim() As TEJESSIM
Public ejesP() As TEJESSIM
Public nElemSim As Integer
Public nEjesP As Integer

' +-------------------------------------------------------------------
' |        MODULO SIMETRIA
' |
' +-------------------------------------------------------------------
Public Function simetria()
Dim tol As Double, vtol As Double
GrupPuntual = ""
For tol = 0.1 To 0.5 Step 0.1
    For vtol = 0.01 To 0.2 Step 0.01
        buscarSimetria tol, vtol
        If GrupPuntual <> "" Then Exit Function
    Next
Next
End Function
' +-------------------------------------------------------------------
' |        MODULO SIMETRIA
' |
' +-------------------------------------------------------------------
Public Sub buscarSimetria(tol As Double, vtol As Double)
Dim i As Integer, j As Integer, k As Integer
Dim distAtomoI As Double, distAtomoJ As Double, distAtomoK As Double
Dim v1 As TVECTOR, v2 As TVECTOR
Dim v3 As TVECTOR, v4 As TVECTOR
Dim v5 As TVECTOR, v6 As TVECTOR
nElemSim = 0
nEjesP = 0
Open "datos.out" For Output As #1
'+------------------------------------------------------------
'| #0
'|primero usamos los vectores de inercia como elementos de simetria
'+------------------------------------------------------------

guardarEjePotencial vMomentoInercia1, vtol
printVector vMomentoInercia1
guardarEjePotencial vMomentoInercia2, vtol
printVector vMomentoInercia2
guardarEjePotencial vMomentoInercia3, vtol
printVector vMomentoInercia3
Print #1, "ejes potenciales= " & nEjesP
'+------------------------------------------------------------
'| #1
'| usamos los atomos como ejes de simetria
'+------------------------------------------------------------
Print #1, "############# ATOMOS #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
For i = 1 To nAtoms
    If Not atomCentral(i) Then
        v1.x = atoms(i).x: v1.y = atoms(i).y: v1.Z = atoms(i).Z
        v1 = normVp(v1)
        printVector v1
        guardarEjePotencial v1, vtol
    End If
Next i
Print #1, "ejes potenciales= " & nEjesP
'+------------------------------------------------------------
'| #1bis
'| buscamos los enlaces como elementos de simetria
'+------------------------------------------------------------
Print #1, "############# Enlaces #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
For i = 1 To nAtoms
    For j = i + 1 To nAtoms
        v1.x = atoms(i).x - atoms(j).x
        v1.y = atoms(i).y - atoms(j).y
        v1.Z = atoms(i).Z - atoms(j).Z
        v1 = normVp(v1)
        printVector v1
        guardarEjePotencial v1, vtol
    Next
Next
Print #1, "ejes potenciales= " & nEjesP
'+------------------------------------------------------------
'| #2
'| segundo, tomamos como posibles elementos de simetría
'| la bisectriz entre dos atomos equivalentes
'+------------------------------------------------------------
Print #1, "############# BISECTRIZ #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"

For i = 1 To nAtoms
    distAtomoI = distancia(i)
    For j = i + 1 To nAtoms
        'han de ser atomos del mismo tipo
        If atoms(i).simbol = atoms(j).simbol Then
            distAtomoJ = distancia(j)
            If (distAtomoJ - distAtomoI) <= tolDist Then ' son iguales y estan a la misma distancia
                With v1
                    .x = (atoms(i).x + atoms(j).x) / 2
                    .y = (atoms(i).y + atoms(j).y) / 2
                    .Z = (atoms(i).Z + atoms(j).Z) / 2
                End With
                If Not CentroMolecula(v1) Then
                    v1 = normVp(v1)
                    guardarEjePotencial v1, vtol
                End If
            End If
        End If
    Next j
Next i
Print #1, "ejes potenciales= " & nEjesP

'+------------------------------------------------------------
'| #3
'| tomamos los vectores perpendiculares a un par de atomos
'| que estan a la misma distancia del centro
'| lo que equivale al vector perpendicular al plano que forman
'| el centro de la molecula y los dos atomos
'+------------------------------------------------------------
Print #1, "############# V PERP #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"

For i = 1 To nAtoms
    distAtomoI = distancia(i)
    For j = i + 1 To nAtoms
        If atoms(i).simbol = atoms(j).simbol Then
            distAtomoJ = distancia(j)
            If (distAtomoJ - distAtomoI) < tolDist Then ' son iguales y estan a la misma distancia
                Print #1, i & atoms(i).simbol & ", " & j & atoms(j).simbol; " "
                ' calculamos el producto vectorial de los
                ' dos vectores normalizados
                Let v1 = normV(i)
                Let v2 = normV(j)
                Let v3 = prodVect(v1, v2)
                Print #1, "VP ";
                printVector v3
                If Not CentroMolecula(v3) Then
                    v3 = normVp(v3)
                    guardarEjePotencial v3, vtol
                End If
            End If
        End If
    Next j
Next i
Print #1, "ejes potenciales= " & nEjesP
'+------------------------------------------------------------
'| #4
'| vectores perpendiculares a un plano que pasa por 3 puntos
'+------------------------------------------------------------
Print #1, "############# 3 PUNTS #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"

For i = 1 To nAtoms
    distAtomoI = distancia(i)
    For j = i + 1 To nAtoms
        If atoms(i).simbol = atoms(j).simbol Then
            distAtomoJ = distancia(j)
            If (distAtomoJ - distAtomoI) < tolDist Then ' son iguales y estan a la misma distancia
                For k = j + 1 To nAtoms
                    If atoms(j).simbol = atoms(k).simbol Then
                        distAtomoK = distancia(k)
                        If (distAtomoK - distAtomoJ) < tolDist Then ' son iguales y estan a la misma distancia
                            Print #1, i & atoms(i).simbol & ", " & j & atoms(j).simbol & ", " & k & atoms(k).simbol; " "
                            ' calculamos el producto vectorial de los
                            ' dos vectores normalizados
                            Let v1 = normV(i)
            '                printVector v1
                            Let v2 = normV(j)
            '                printVector v2
                            Let v3 = normV(k)
                            Let v4 = difV(v2, v1)
                            Let v5 = difV(v3, v1)
                            Let v6 = prodVect(v4, v5)
                            If Not CentroMolecula(v6) Then
                                v6 = normVp(v6)
                                guardarEjePotencial v6, vtol
                            End If
                            printVector v6
                        End If
                    End If
                Next k
            End If
        End If
    Next j
Next i
Print #1, "ejes potenciales= " & nEjesP

'+------------------------------------------------------------
'| #4
'| busquemos poligonos
'+------------------------------------------------------------
Print #1, "############# POLIGONS #################################"
Print #1, "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"

Dim di As Double, dj As Double, dk As Double
Dim dij As Double, djk As Double
For i = 1 To nAtoms
    di = distancia(i)
    If Round(di, 2) <> 0# Then
        For j = i + 1 To nAtoms
            If j <> i Then
                dj = distancia(j)
                If atoms(i).simbol = atoms(j).simbol And _
                    Abs(di - dj) < tolDist Then
                    dij = distanciaIJ(i, j)
                    For k = j + 1 To nAtoms
                        If k <> j Then
                            dk = distancia(k)
                            djk = distanciaIJ(j, k)
                            If atoms(j).simbol = atoms(k).simbol And _
                            Abs(dj - dk) < tolDist And _
                            Abs(dij - djk) < tolDist Then
                            End If
                        End If
                    Next k
                End If
            End If
        Next j
    End If
Next i
Print #1, "ejes potenciales= " & nEjesP
printEjespotenciales
buscar_ElementosAsociados (tol)
Print #1, "_________________ELEMENTOS DE SIMETRIA___________________"
Print #1,
ordElmSim
PrintElementosDeSimetria
Dim elementos As String
Dim gp As String
elementos = ListaElementosDeSimetria
gp = grupoP(elementos)
Print #1, "_________________ GRUPO PUNTUAL _________________________"
Print #1,
Print #1, "GRUPO PUNTUAL: "; gp
If gp <> "" Then
    GrupPuntual = gp
End If
Close #1
End Sub 'END SIMETRIA!

'**********************************************************
'
'**********************************************************
Function distanciaIJ(i As Integer, j As Integer) As Double
distanciaIJ = Sqr((atoms(i).x - atoms(j).x) ^ 2 + (atoms(i).y - atoms(j).y) ^ 2 + (atoms(i).Z - atoms(j).Z) ^ 2)
End Function

'**********************************************************
'
'**********************************************************
Private Function difV(v1 As TVECTOR, v2 As TVECTOR) As TVECTOR
With difV
    .x = v2.x - v1.x
    .y = v2.y - v1.y
    .Z = v2.Z - v1.Z
End With
End Function
' +-------------------------------------------------------------------
' |       Impresion
' |     Vector, Eje, Punto
' +-------------------------------------------------------------------
Private Sub printVector(v As TVECTOR)
Print #1, Format(Format(v.x, formato), "@@@@@@@"); " // "; Format(Format(v.y, formato), "@@@@@@@"); " // "; Format(Format(v.Z, formato), "@@@@@@@")
End Sub

'**********************************************************
'
'**********************************************************
Private Function normV(i As Integer) As TVECTOR
Dim d As Double
normV.x = atoms(i).x
normV.y = atoms(i).y
normV.Z = atoms(i).Z
d = Sqr(normV.x ^ 2 + normV.y ^ 2 + normV.Z ^ 2)
If Round(d, 2) <> 0# Then
    normV.x = normV.x / d
    normV.y = normV.y / d
    normV.Z = normV.Z / d
End If
End Function

'**********************************************************
'
'**********************************************************
Private Function normVp(v As TVECTOR) As TVECTOR
Dim d As Double
d = Sqr(v.x ^ 2 + v.y ^ 2 + v.Z ^ 2)
normVp.x = v.x / d
normVp.y = v.y / d
normVp.Z = v.Z / d
End Function

'**********************************************************
'
'**********************************************************
Private Function prodVect(u As TVECTOR, v As TVECTOR) As TVECTOR
With prodVect
    .x = u.y * v.Z - u.Z * v.y
    .y = u.x * v.Z - u.Z * v.x
    .Z = u.x * v.y - u.y * v.x
End With
End Function

' +-------------------------------------------------------------------
' |        InvertirZ
' |     Invierte un punto paralelamente al eje Z
' +-------------------------------------------------------------------
Private Sub InvertirZ(a() As TATOMO)
Dim i As Integer
For i = 1 To nAtoms
    a(i).Z = -a(i).Z
Next i
End Sub

' +-------------------------------------------------------------------
' |        Invertirsion
' |
' +-------------------------------------------------------------------
Private Sub Inversion(a() As TATOMO)
Dim i As Integer
For i = 1 To nAtoms
    a(i).x = -a(i).x
    a(i).y = -a(i).y
    a(i).Z = -a(i).Z
Next i
End Sub
' +-------------------------------------------------------------------
' |        girarAlrededorY
' |     Gira alrededor del eje Y , angulo
' +-------------------------------------------------------------------
Public Sub girarAlrededorY(a() As TATOMO, angulo As Double)
Dim xg As Double, zg As Double
Dim i As Integer
For i = 1 To nAtoms
    xg = a(i).x * Cos(angulo) - a(i).Z * Sin(angulo)
    zg = a(i).x * Sin(angulo) + a(i).Z * Cos(angulo)
    a(i).x = xg
    a(i).Z = zg
Next i
End Sub

' +-------------------------------------------------------------------
' |        girarAlrededorZ
' |     Gira alrededor del eje Z , angulo
' +-------------------------------------------------------------------
Public Sub girarAlrededorZ(a() As TATOMO, angulo As Double)
Dim i As Integer
Dim xg As Double, yg As Double
For i = 1 To nAtoms
    xg = a(i).x * Cos(angulo) + a(i).y * Sin(angulo)
    yg = -a(i).x * Sin(angulo) + a(i).y * Cos(angulo)
    a(i).x = xg
    a(i).y = yg
Next i
End Sub
' +-------------------------------------------------------------------
' |        Comparar_moleculas
' |   entrada: molecula a y molecula b
' |   resultado: true si las dos molecula son iguales
' | El factor de tolerancia ha de ser mayor cuanto mas alejados esten
' | los atomos del centro de la molecula
' +-------------------------------------------------------------------
Public Function comparar_Moleculas(a() As TATOMO, b() As TATOMO, mtol) As Boolean
Dim i As Integer, j As Integer
Dim encontrado As Boolean
Dim d As Double
comparar_Moleculas = False
For i = 1 To nAtoms
    encontrado = False
    For j = 1 To nAtoms
        If a(i).simbol = b(j).simbol Then
            d = Sqr((a(i).x - b(j).x) ^ 2 + (a(i).y - b(j).y) ^ 2 + (a(i).Z - b(j).Z) ^ 2)
            If d <= mtol Then
                'hemos encontrado un atomo identico en la misma posicion
                encontrado = True
                Exit For
            End If
        End If
    Next j
    If encontrado = False Then Exit Function
Next i
    comparar_Moleculas = True
End Function
Private Function distancia2Atomos(i, j) As Double
distancia2Atomos = Sqr((atoms(i).x - atoms(j).x) ^ 2 + (atoms(i).y - atoms(j).y) ^ 2 + (atoms(i).Z - atoms(j).Z) ^ 2)
End Function

' +-------------------------------------------------------------------
' | Arccos
' | sacado del control opengl
' +-------------------------------------------------------------------
Private Function Arccos(x As Double) As Double
Select Case x
    Case Is >= 1
        Arccos = 0
    Case Is <= -1
        Arccos = PI
    Case Else
    Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Select
End Function

' +-------------------------------------------------------------------
' | rad2Deg
' | convierte angulos en radianes a grados
' +-------------------------------------------------------------------
Private Function rad2deg(rad As Double) As Double
rad2deg = rad * 180 / PI
End Function

' +-------------------------------------------------------------------
' | printAtoms
' | imprime las coordenadas (solo para depurar)
' +-------------------------------------------------------------------

Private Function printAtoms(a() As TATOMO)
Dim i As Integer
For i = 1 To nAtoms
    Print #1, i & Format(a(i).x, formato), Format(a(i).y, formato), Format(a(i).Z, formato)
Next i
End Function

' +-------------------------------------------------------------------
' | distancia
' | busca la distancia de un atomo al centro de gravedad
' +-------------------------------------------------------------------
Private Function distancia(i As Integer) As Double
distancia = Sqr(atoms(i).x ^ 2 + atoms(i).y ^ 2 + atoms(i).Z ^ 2)
End Function

' +-------------------------------------------------------------------
' |        theta
' |     calcula el angulo del vector i respecto Z
' +-------------------------------------------------------------------
Public Function theta(i As Integer) As Double
theta = Arccos(ejesP(i).Z)
End Function

' +-------------------------------------------------------------------
' |        phi
' |     calcula el angulo del eje i con el eje X
' +-------------------------------------------------------------------
Public Function phi(i As Integer) As Double

Dim d As Double
Dim cosphi As Double
Dim sinphi As Double
'On Error GoTo fin
d = Sqr(ejesP(i).x ^ 2 + ejesP(i).y ^ 2 + ejesP(i).Z ^ 2)
If d = 0 Then GoTo fin
cosphi = ejesP(i).x / d
sinphi = ejesP(i).y / d
'
' Per calcular l'arctangent necesitem calcular la tangent
' amb el cocient sinphi/cosphi
' per tant si tinguessim cosphi=0 provocariem un error
'
If cosphi = 0 Then ' vector sobre l'eix vertical Y
    If sinphi > 0 Then
        phi = PI / 2
    Else
        phi = 3 * PI / 2
    End If
Else
' Calculem l 'arctangent amb la funcio del Vbasic ATN()
' el transformem en graus mitjançant la funcio propia
' rad2deg() que es trova en aquest modul
' finalment calculem el seu valor absolut amb la funcio
' ABS() del Vbasic
    phi = Abs(Atn(sinphi / cosphi))
'Busquem els quadrants
    Select Case sinphi
        Case Is > 0  'quadrants 1 i 2
            If cosphi > 0 Then
                phi = phi       ' quadrant 1
            Else
                phi = PI - phi ' quadrant 2
            End If
        Case Is < 0  'quadrants 3 i 4
            If cosphi > 0 Then
                phi = 2 * PI - phi ' quadrant 4
            Else
                phi = phi + PI ' quadrant 3
            End If
    End Select 'sinphi
End If 'cosphi=0
fin:
End Function

'**********************************************************
'
'**********************************************************
Private Sub buscar_ElementosAsociados(tol As Double)
Const c10 = 2 * PI / 10
Const c8 = PI / 4               '45
Const c6 = PI / 3               '60º
Const c5 = 2 * PI / 5           '72º
Const c4 = PI / 2               '90º
Const c3 = 2 * PI / 3           '120º
Const c2 = PI                   '180º
Dim a() As TATOMO
Dim b() As TATOMO
Dim i As Integer
For i = 1 To nEjesP
'**************** C2 *****************************
    Let a() = atoms() 'inicializamos a()
'*************************************************
' proyectamos El eje de simetria junto con la molecula sobre el eje Z
' despues aplicamos la operacion de simetria y comprobamos finalmente
' si tenemos la misma molecula
'*************************************************
    proyectarZ a(), i 'proyectamos sobreZ
    Let b() = a()     'guardamos la molecula proyectada
    girarAlrededorZ a(), c2
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "C2"
    End If
'**************** C3 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c3
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "C3"
    End If
'**************** C4 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c4
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "C4"
    End If
'**************** C5 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c5
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "C5"
    End If
'**************** C6 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c6
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "C6"
    End If
'**************** SIGMA *****************************
    Let a() = b() 'restauramos la molecula
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "sigma"
    End If
'**************** S3 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c3
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S3"
    End If
'**************** S4 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c4
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S4"
    End If
'**************** S5 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c5
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S5"
    End If
'**************** S6 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c6
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S6"
    End If
'**************** S8 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c8
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S8"
    End If
'**************** S10 *****************************
    Let a() = b() 'restauramos la molecula
    girarAlrededorZ a(), c10
    InvertirZ a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria i, "S10"
    End If
Next i

'**************** I *****************************
    Let a() = b() 'restauramos la molecula
    Inversion a()
    If comparar_Moleculas(a(), b(), tol) Then
        GuardarElementoSimetria 0, "I"
    End If
End Sub

'**********************************************************
'
'**********************************************************
Private Sub GuardarElementoSimetria(j As Integer, tipo As String)
Dim i As Integer
If j = 0 Then 'centro de inversion
    nElemSim = nElemSim + 1
    ReDim Preserve ElemSim(1 To nElemSim)
    ElemSim(nElemSim).x = 0
    ElemSim(nElemSim).y = 0
    ElemSim(nElemSim).Z = 0
    ElemSim(nElemSim).tipo = tipo
    ElemSim(nElemSim).visible = True
Else
    nElemSim = nElemSim + 1
    ReDim Preserve ElemSim(1 To nElemSim)
    ElemSim(nElemSim).x = ejesP(j).x
    ElemSim(nElemSim).y = ejesP(j).y
    ElemSim(nElemSim).Z = ejesP(j).Z
    ElemSim(nElemSim).tipo = tipo
    ElemSim(nElemSim).visible = True
End If
End Sub

'**********************************************************
'
'**********************************************************
Private Function atomCentral(i As Integer) As Boolean
    If (Abs(Round(atoms(i).x, 2)) = 0# And _
        Abs(Round(atoms(i).y, 2)) = 0# And _
        Abs(Round(atoms(i).Z, 2)) = 0#) Then
        atomCentral = True
    End If
End Function
Private Function CentroMolecula(v As TVECTOR) As Boolean
    If (Abs(Round(v.x, 2)) = 0# And _
        Abs(Round(v.y, 2)) = 0# And _
        Abs(Round(v.Z, 2)) = 0#) Then
        CentroMolecula = True
    End If
End Function

'**********************************************************
'
'**********************************************************
Private Sub guardarEjePotencial(eje As TVECTOR, ejestol As Double)  ' eje ha de ser normalizado
Dim i As Integer
For i = 1 To nEjesP
    DoEvents
    If Abs(eje.x - ejesP(i).x) <= ejestol And _
       Abs(eje.y - ejesP(i).y) <= ejestol And _
       Abs(eje.Z - ejesP(i).Z) <= ejestol Or _
       Abs(eje.x + ejesP(i).x) <= ejestol And _
       Abs(eje.y + ejesP(i).y) <= ejestol And _
       Abs(eje.Z + ejesP(i).Z) <= ejestol Then
       Exit Sub
    End If
Next i
nEjesP = nEjesP + 1
ReDim Preserve ejesP(1 To nEjesP)
ejesP(nEjesP).x = eje.x
ejesP(nEjesP).y = eje.y
ejesP(nEjesP).Z = eje.Z
End Sub

'**********************************************************
'
'**********************************************************
Private Sub proyectarZ(ByRef a() As TATOMO, i As Integer)
Dim aPhi As Double, aTheta As Double
    aPhi = phi(i)
    aTheta = theta(i)
    girarAlrededorZ a(), aPhi
    girarAlrededorY a(), aTheta
End Sub

' +-------------------------------------------------------------------
' |       Impresion
' |     Elementos de Simetria
' +-------------------------------------------------------------------
Private Sub PrintElementosDeSimetria()
Dim i
For i = 1 To nElemSim
    Print #1, ElemSim(i).tipo & " // " & _
                formatNumero(ElemSim(i).x) & " // " & _
                formatNumero(ElemSim(i).y) & " // " & _
                formatNumero(ElemSim(i).Z)
Next i
End Sub

' +-------------------------------------------------------------------
' |       Impresion
' |     Ejes Potenciales
' +-------------------------------------------------------------------
Private Sub printEjespotenciales()
Dim i
For i = 1 To nEjesP
    Print #1, formatNumero(ejesP(i).x) & " // " & _
              formatNumero(ejesP(i).y) & " // " & _
              formatNumero(ejesP(i).Z)
Next i
End Sub

' +-------------------------------------------------------------------
' |       Impresion
' |     Eje Potenciales
' +-------------------------------------------------------------------
Private Sub printElemento(i As Integer)
    Print #1, formatNumero(ejesP(i).x) & " // " & _
              formatNumero(ejesP(i).y) & " // " & _
              formatNumero(ejesP(i).Z)
End Sub

' +-------------------------------------------------------------------
' |       borrar
' |     Eje Potenciales, Simetria
' +-------------------------------------------------------------------
Public Sub BorrarSimetria()
nEjesP = 0
nElemSim = 0
bListaSimetria = GL_FALSE
glFlush
End Sub

' +-------------------------------------------------------------------
' |     Contar ejes para determinar el posible grupo puntual
' |     Salida Texto con el numero de elementos de simetria
' +-------------------------------------------------------------------

Private Function ListaElementosDeSimetria() As String
Dim i As Integer
Dim e As String
Dim texto As String
Dim c2, c3, c4, c5, c6, s3, s4, s5, s6, s8, s10, ci, sigma
For i = 1 To nElemSim
    e = ElemSim(i).tipo
    If e = "sigma" Then                 ' SIGMA
        sigma = sigma + 1
        ElseIf Left(e, 1) = "C" Then    ' ejes C
            Select Case Mid(e, 2, 1)
                Case "2"
                    c2 = c2 + 1
                Case "3"
                    c3 = c3 + 1
                Case "4"
                    c4 = c4 + 1
                Case "5"
                    c5 = c5 + 1
                Case "6"
                    c6 = c6 + 1
            End Select
            ElseIf Left(e, 1) = "S" Then    'ejes S
                Select Case Mid(e, 2, 1)
                    Case "3"
                        s3 = s3 + 1
                    Case "4"
                        s4 = s4 + 1
                    Case "5"
                        s5 = s5 + 1
                    Case "6"
                        s6 = s6 + 1
                    Case "8"
                        s8 = s8 + 1
                    Case "1"
                        s10 = s10 + 1
                End Select
                ElseIf Left(e, 1) = "I" Then
                    ci = ci + 1
    End If
Next
    texto = ""
    If ci <> 0 Then texto = texto & "i" & " "
    If c6 <> 0 Then texto = texto & c6 & "(C6)" & " "
    If c5 <> 0 Then texto = texto & c5 & "(C5)" & " "
    If c4 <> 0 Then texto = texto & c4 & "(C4)" & " "
    If c3 <> 0 Then texto = texto & c3 & "(C3)" & " "
    If c2 <> 0 Then texto = texto & c2 & "(C2)" & " "
    If s10 <> 0 Then texto = texto & s10 & "(S10)" & " "
    If s8 <> 0 Then texto = texto & s8 & "(S8)" & " "
    If s6 <> 0 Then texto = texto & s6 & "(S6)" & " "
    If s4 <> 0 Then texto = texto & s4 & "(S4)" & " "
    If s3 <> 0 Then texto = texto & s3 & "(S3)" & " "
    If sigma <> 0 Then texto = texto & sigma & "(sigma)"
ListaElementosDeSimetria = texto
End Function

'**********************************************************
' Entrada: Texto que contiene los elementos de simetria
' Salida: GrupoPuntual
'**********************************************************
Private Function grupoP(texto As String) As String
Select Case RTrim(texto)
    Case "1(C2)"
        grupoP = "C2"
    Case "1(sigma)"
        grupoP = "Cs"
    Case "i"
        grupoP = "Ci"
    Case " "
        grupoP = "C1"
    Case "1(C3)"
        grupoP = "(C3)"
    Case "1(C4) 1(C2)"
        grupoP = "C4"
    Case "1(C5)"
        grupoP = "C5"
    Case "1(C6) 1(C3) 1(C2)"
        grupoP = "C6"
    Case "1(C8) 1(C4) 1(C2)"
        grupoP = "C8"
    Case "1(C2) 2(sigma)"
        grupoP = "C2v"
    Case "1(C3) 3(sigma)"
        grupoP = "C3v"
    Case "1(C4) 4(sigma)"
        grupoP = "C4v"
    Case "1(C5) 5(sigma)"
        grupoP = "C5v"
    Case "1(C6) 1(C3) 1(C2) 6(sigma)"
        grupoP = "C6v"
    Case "i 1(C2) 1(sigma)"
        grupoP = "C2h"
    Case "1(C3) 1(S3) 1(sigma)"
        grupoP = "C3h"
    Case "i 1(C4) 1(C2) 1(S4) 1(sigma)"
        grupoP = "C4h"
    Case "1(C5) 1(S5) 1(sigma)"
        grupoP = "C5h"
    Case "i 1(C6) 1(C3) 1(C2) 1(S6) 1(S3) 1(sigma)"
        grupoP = "C6h"
    Case "i 3(C2) 3(sigma)"
        grupoP = "D2h"
    Case "1(C3) 3(C2) 1(S3) 4(sigma)"
        grupoP = "D3h"
    Case "i 1(C4) 5(C2) 1(S4) 5(sigma)"
        grupoP = "D4h"
    Case "1(C5) 5(C2) 1(S5) 6(sigma)"
        grupoP = "D5h"
    Case "i 1(C6) 1(C3) 7(C2) 1(S6) 1(S3) 7(sigma)"
        grupoP = "D6h"
    Case "3(C2) 1(S4) 2(sigma)"
        grupoP = "D2d"
    Case "i 1(C3) 3(C2) 1(S6) 3(sigma)"
        grupoP = "D3d"
    Case "1(C4) 5(C2) 1(S8) 4(sigma)"
        grupoP = "D4d"
    Case "i 1(C5) 5(C2) 1(S10) 5(sigma)"
        grupoP = "D5d"
    Case "1(C2) 1(S4)"
        grupoP = "S4"
    Case "1(C4) 1(C2) 1(S8)"
        grupoP = "S8"
    Case "3(C2)"
        grupoP = "D2"
    Case "1(C3) 3(C2)"
        grupoP = "D3"
    Case "1(C4) 4(C2)"
        grupoP = "D4"
    Case "1(C5) 5(C2)"
        grupoP = "D5"
    Case "1(C6) 3(C2)"
        grupoP = "D6"
    Case "4(C3) 3(C2) 3(S4) 6(sigma)"
        grupoP = "Td"
    Case "i 3(C4) 4(C3) 9(C2) 4(S6) 3(S4) 9(sigma)"
        grupoP = "Oh"
    Case "i 6(C5) 10(C3) 15(C2) 6(S10) 10(S6) 15(sigma)"
        grupoP = "Ih"
End Select
frmMain!StatusBar1.Panels(1).Text = grupoP & " >>> elements: " & texto & " "
End Function

'+---------------------------------------------------------------
'+ ordenar elementos de simetria
'+---------------------------------------------------------------
Private Sub ordElmSim()
Dim flag As Boolean
Dim i As Integer
flag = False
    Do While flag = False
        flag = True
        For i = 1 To nElemSim - 1
            If ElemSim(i).tipo < ElemSim(i + 1).tipo Then
                intcam i, i + 1
                flag = False
            End If
        Next
    Loop
End Sub

'+---------------------------------------------------------------
'+ intercambiar elementos de simetria
'+---------------------------------------------------------------

Private Sub intcam(i As Integer, j As Integer)
Dim dummy As TEJESSIM
Let dummy = ElemSim(i)
Let ElemSim(i) = ElemSim(j)
Let ElemSim(j) = dummy
End Sub
