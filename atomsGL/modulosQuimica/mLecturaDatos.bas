Attribute VB_Name = "mLecturaDatos"
Option Explicit
'Option Base 1

' +-------------------------------------------------------------------
'
' +-------------------------------------------------------------------
Public Sub LecturaDatos()
Dim extension As String
Dim center As Boolean
    center = True

    frmMain.CommonDialog1.InitDir = App.Path
    frmMain.CommonDialog1.ShowOpen
    fitxer = frmMain.CommonDialog1.Filename
    If fitxer = "" Then Exit Sub
    extension = UCase(Right(frmMain.CommonDialog1.FileTitle, 3))
'    fitxer = App.Path & "\estructuras\ab6.xyz"
'    extension = "XYZ"
    Select Case extension
        Case "CIF"
            Call LecturaCIF
            Call frac2cart
        Case "XYZ"
            Call LecturaXYZ
        Case "FRC"
            Call LecturaFRC
            Call frac2cart
        Case "INT"
            Call LecturaINT
        Case "PDB"
            Call LecturaPDB
    End Select
    BorrarSeleccion
    BorrarPlano miPlano
    BorrarSimetria
    CentratCDM
    TensorDeInercia
    calcularEnlaces
    CuboDeInclusionMolecula 'calculamos las medidas máximas de la molecula y escalamos la pantalla en cosecuencia
    displayListInited = GL_FALSE
    displaySelectionInited = GL_FALSE
End Sub

' +-------------------------------------------------------------------
' LECTURA CIF ESTANDAR
' +-------------------------------------------------------------------
Public Sub LecturaCIF()
Dim stext() As String
Dim strRepl As String
'On Error Resume Next
nAtoms = 0
Dim linea As String
'+-------------------------------------------
'+ tractament dels fitxers UNIX
'+-------------------------------------------
Open fitxer For Input As #1
'On Error GoTo adios
Line Input #1, linea
    If InStr(linea, vbLf) Then
        Close #1
        strRepl = Replace(linea, vbLf, vbCrLf)
        Open "temp.tmp" For Output As #2
        Print #2, strRepl
        Close #2
        Open "temp.tmp" For Input As #1
    Else
        Close #1
        Open fitxer For Input As #1
    End If
'+-------------------------------------------
'+ tractament dels fitxers UNIX
'+-------------------------------------------
Do While Not EOF(1)
Line Input #1, linea
If InStrRev(linea, "_cell_length_a") Then
    stext = Split(linea, Chr(32), 2)
    celda.a = Val(stext(1))
End If

If InStrRev(linea, "_cell_length_b") Then
    stext = Split(linea, Chr(32), 2)
    celda.b = Val(stext(1))
End If

If InStrRev(linea, "_cell_length_c") Then
    stext = Split(linea, Chr(32), 2)
    celda.c = Val(stext(1))
End If

If InStrRev(linea, "_cell_angle_alpha") Then
    stext = Split(linea, Chr(32), 2)
    celda.alfa = Val(stext(1))
End If

If InStrRev(linea, "_cell_angle_beta") Then
    stext = Split(linea, Chr(32), 2)
    celda.beta = Val(stext(1))
End If

If InStrRev(linea, "_cell_angle_gamma") Then
    stext = Split(linea, Chr(32), 2)
    celda.gamma = Val(stext(1))
End If

If InStrRev(linea, "_atom_site_type") Then
    Do While InStr(linea, "_atom_site")
        Line Input #1, linea
    Loop
    stext = Split(linea, Chr(32), 5)
    nAtoms = nAtoms + 1
    ReDim Preserve atoms(1 To nAtoms)
    atoms(nAtoms).label = stext(0)
    atoms(nAtoms).simbol = stext(1)
    atoms(nAtoms).fracX = Val(stext(2))
    atoms(nAtoms).fracY = Val(stext(3))
    atoms(nAtoms).fracZ = Val(stext(4))
    Do While InStr(linea, "loop_") = 0
        Line Input #1, linea
        stext = Split(linea, Chr(32), 5)
        If UBound(stext) > 1 Then
                nAtoms = nAtoms + 1
                ReDim Preserve atoms(1 To nAtoms)
                atoms(nAtoms).label = Trim(stext(0))
                atoms(nAtoms).simbol = Trim(stext(1))
                atoms(nAtoms).fracX = Val(stext(2))
                atoms(nAtoms).fracY = Val(stext(3))
                atoms(nAtoms).fracZ = Val(stext(4))
        End If
    Loop
End If
If InStrRev(linea, "_atom_site_aniso") Then
    Do While InStr(linea, "_atom_site")
        Line Input #1, linea
    Loop
    stext = NeatSplit(linea)
    nEllipses = nEllipses + 1
    ReDim Preserve UatomsAniso(1 To nEllipses)
    atoms(nEllipses).label = stext(0)
    UatomsAniso(nEllipses).u(1, 1) = Val(stext(2))
    UatomsAniso(nEllipses).u(2, 2) = Val(stext(3))
    UatomsAniso(nEllipses).u(3, 3) = Val(stext(4))
    UatomsAniso(nEllipses).u(2, 3) = Val(stext(5))
    UatomsAniso(nEllipses).u(1, 3) = Val(stext(5))
    UatomsAniso(nEllipses).u(1, 2) = Val(stext(6))
    Do While UBound(stext) < 1
        Line Input #1, linea
        stext = NeatSplit(linea, Chr(32), 5)
        If UBound(stext) > 1 Then
            nEllipses = nEllipses + 1
            ReDim Preserve UatomsAniso(1 To nEllipses)
            atoms(nEllipses).label = stext(0)
            UatomsAniso(nEllipses).u(1, 1) = Val(stext(2))
            UatomsAniso(nEllipses).u(2, 2) = Val(stext(3))
            UatomsAniso(nEllipses).u(3, 3) = Val(stext(4))
            UatomsAniso(nEllipses).u(2, 3) = Val(stext(5))
            UatomsAniso(nEllipses).u(1, 3) = Val(stext(5))
            UatomsAniso(nEllipses).u(1, 2) = Val(stext(6))
        End If
    Loop
End If
Loop
asignarU
make_bmat
Call parametros
adios: Close #1
Exit Sub
End Sub

' +-------------------------------------------------------------------
' LECTURA FRC ESTANDAR
' +-------------------------------------------------------------------

Public Sub LecturaFRC()
Dim dummy As String
Dim i As Integer
Dim stext() As String
nAtoms = 0
On Error Resume Next
Open fitxer For Input As #1
With celda
    Input #1, dummy
        stext = NeatSplit(dummy)
        .a = Val(stext(0))
        .b = Val(stext(1))
        .c = Val(stext(2))
    Input #1, dummy
        stext = NeatSplit(dummy)
        .alfa = Val(stext(0))
        .beta = Val(stext(1))
        .gamma = Val(stext(2))
    End With
Do While Not (EOF(1))
    nAtoms = nAtoms + 1
    ReDim Preserve atoms(1 To nAtoms)
    Input #1, dummy
    stext = NeatSplit(dummy)
    With atoms(nAtoms)
        .label = Trim(stext(0))
        .simbol = Simbolo2(.label)
        .fracX = Val(stext(1))
        .fracY = Val(stext(2))
        .fracZ = Val(stext(3))
'        Debug.Print .label, .fracX, .fracY, .fracZ
    End With
Loop
    Close #1
    Call parametros
End Sub

' +-------------------------------------------------------------------
' LECTURA XYZ ESTANDAR
' +-------------------------------------------------------------------

Public Sub LecturaXYZ()
Dim dummy As String
Dim i As Integer, j As Integer
Dim stext() As String
Dim longitudDummy As Integer
Dim letra As String
Dim lineaSinComas As String
nAtoms = 0
'On Error Resume Next
Open fitxer For Input As #1
Input #1, nAtoms
Input #1, dummy
For i = 1 To nAtoms
    Line Input #1, dummy
    longitudDummy = Len(dummy)
    lineaSinComas = ""
    For j = 1 To longitudDummy
        letra = Mid(dummy, j, 1)
    ' instr devuelve la posicion del caracter encontrado
        If InStr(1, ",", letra) = 0 Then
           lineaSinComas = lineaSinComas + letra
        End If
    Next j
    stext = NeatSplit(lineaSinComas)
    ReDim Preserve atoms(1 To i)
    With atoms(i)
        .label = Trim(stext(0))
        .simbol = Simbolo2(.label)
        .x = Val(stext(1))
        .y = Val(stext(2))
        .Z = Val(stext(3))
    End With
Next i
    Close #1
    Call parametros
End Sub

' +-------------------------------------------------------------------
' LECTURA PDB
' +-------------------------------------------------------------------
'    COLUMNS  DATA TYPE     FIELD       DEFINITION
'    ---------------------------------------------------------------------------------
'     1 -  6  Record name   "ATOM  "
'     7 - 11  Integer       serial      Atom serial number.
'    13 - 16  Atom          name        Atom name.
'    17       Character     altLoc      Alternate location indicator.
'    18 - 20  Residue name  resName     Residue name.
'    22       Character     chainID     Chain identifier.
'    23 - 26  Integer       resSeq      Residue sequence number.
'    27       AChar         iCode       Code for insertion of residues.
'    31 - 38  Real(8.3)     x           Orthogonal coordinates for X, Angstroms.
'    39 - 46  Real(8.3)     y           Orthogonal coordinates for Y, Angstroms.
'    47 - 54  Real(8.3)     z           Orthogonal coordinates for Z, Angstroms.
'    55 - 60  Real(6.2)     occupancy   Occupancy.
'    61 - 66  Real(6.2)     tempFactor  Temperature factor.
'    73 - 76  LString(4)    segID       Segment identifier, left-justified.
'    77 - 78  LString(2)    element     Element symbol, right-justified.
'    79 - 80  LString(2)    charge      Charge on the atom.
'    ---------------------------------------------------------------------------------
Public Sub LecturaPDB()
Dim dummy As String
Dim i As Integer, j As Integer
Dim stext() As String
Dim texto As String
Dim strRepl As String
Dim numF As Integer
nAtoms = 0
Open fitxer For Input As #1

Line Input #1, dummy
    If InStr(dummy, vbLf) Then
        Close #1
        strRepl = Replace(dummy, vbLf, vbCrLf)
        Open "temp.tmp" For Output As #2
        Print #2, strRepl
        Close #2
        Open "temp.tmp" For Input As #1
    Else
        Close #1
        Open fitxer For Input As #1
    End If
Do While Not EOF(1)
    Line Input #1, dummy
    If dummy = "" Then Exit Do
    stext = NeatSplit(dummy)
    texto = Trim(stext(0))
    If texto = "ATOM" Or texto = "HETATM" Then
        nAtoms = nAtoms + 1
        ReDim Preserve atoms(1 To nAtoms)
        With atoms(nAtoms)
            .label = Trim(stext(2)) ' ATOM NAME
            .simbol = Simbolo2(.label)
            texto = Mid(dummy, 31)
            stext = NeatSplit(texto)
            For i = 0 To UBound(stext())
'            Debug.Print Val(stext(i))
                If IsNumeric(stext(i)) Then
                    .x = Val(stext(i))
                    .y = Val(stext(i + 1))
                    .Z = Val(stext(i + 2))
                    Exit For
                End If
            Next
        End With
    End If
Loop
Close #1
Call parametros
End Sub

'*********************************************************************
' LECTURA INTERNAS ESTANDAR
'*********************************************************************
Public Sub LecturaINT()
Dim dummy As String
Dim i As Integer
Dim stext() As String
nAtoms = 0
Open fitxer For Input As #1
Input #1, nAtoms 'numero atomos
ReDim coorInt(nAtoms)
ReDim atoms(nAtoms)
Input #1, dummy 'titulo
For i = 1 To nAtoms
    Input #1, dummy
    stext = NeatSplit(dummy)
    With coorInt(i)
        .label = Trim(stext(0))
        .simbol = Simbolo2(.label)
        .r = Val(stext(1))
        .w = Val(stext(2))
        .t = Val(stext(3))
        .na = Val(stext(4))
        .nb = Val(stext(5))
        .nc = Val(stext(6))
    End With
Next i
    Close #1
For i = 1 To nAtoms
    With coorInt(i)
    atoms(i).simbol = coorInt(i).simbol
    atoms(i).label = coorInt(i).label
    Debug.Print .simbol, .r, .w, .t, .na; .nb; .nc
    End With
Next
    Call intCar
    Call parametros
End Sub
' +-------------------------------------------------------------------
'
' +-------------------------------------------------------------------
Private Function Simbolo2(s As String) As String
If InStr("0123456789", Mid(s, 2, 1)) Then
    Simbolo2 = Left(s, 1)
    Else
    Simbolo2 = Left(s, 2)
End If
End Function
' +-------------------------------------------------------------------
' asociamos los parametros
' +-------------------------------------------------------------------
Public Sub parametros()
Dim i As Integer, j As Integer
For i = 1 To nAtoms
    For j = 1 To nParametros
        If UCase(Trim(atomGL(j).simbol)) = UCase(Trim(atoms(i).simbol)) Then
        With atoms(i)
            .color = atomGL(j).color
            .radioCov = atomGL(j).radioCov
            .numeroAtomico = j
        End With
    End If
    Next j
Next i
End Sub

' +-------------------------------------------------------------------
' | calcularEnlaces
' +-------------------------------------------------------------------
Public Sub calcularEnlaces()
Dim i As Integer, j As Integer
nEnlaces = 0
Dim dist As Single
For i = 1 To nAtoms
    For j = i + 1 To nAtoms
            dist = distanciaAB(atoms(i), atoms(j))
            If dist < (atoms(i).radioCov + atoms(j).radioCov) * 1.2 Then
                nEnlaces = nEnlaces + 1
                ReDim Preserve enlace(nEnlaces)
                With enlace(nEnlaces)
                        .a = i
                        .ca = ConvertRGB(QBColor(atoms(i).color))
                        .b = j
                        .cb = ConvertRGB(QBColor(atoms(j).color))
                        .d = dist
                        .x0 = atoms(i).x
                        .y0 = atoms(i).y
                        .z0 = atoms(i).Z
                        .vx = (atoms(j).x - atoms(i).x) / .d
                        .vy = (atoms(j).y - atoms(i).y) / .d
                        .vz = (atoms(j).Z - atoms(i).Z) / .d
                        .angle = -Arccos(.vz) * 180 / PI 'radianes
                  End With
            End If
    Next j
Next i
End Sub

' +-------------------------------------------------------------------
'
' +-------------------------------------------------------------------
Public Sub CuboDeInclusionMolecula()
Dim i As Integer
xmax = 0: xmin = 100
ymax = 0: ymin = 100
zmax = 0: zmin = 100

For i = 1 To nAtoms
    With atoms(i)
        If .x > xmax Then xmax = .x
        If .x < xmin Then xmin = .x
        If .y > ymax Then ymax = .y
        If .y < ymin Then ymin = .y
        If .Z > zmax Then zmax = .Z
        If .Z < zmin Then zmin = .Z
    End With
Next i
If zmax > xmax Then maximus = zmax Else maximus = xmax
If ymax > maximus Then maximus = ymax
End Sub

' +-------------------------------------------------------------------
' Equivalente al Check_atom_label
' +-------------------------------------------------------------------
Private Sub asignarU()
Dim i As Integer, j As Integer
For i = 1 To nAtoms
    For j = 1 To nEllipses
        If atoms(i).label = UatomsAniso(j).label Then
            atoms(i).u(1, 1) = UatomsAniso(j).u(1, 1)
            atoms(i).u(2, 2) = UatomsAniso(j).u(2, 2)
            atoms(i).u(3, 3) = UatomsAniso(j).u(3, 3)
            atoms(i).u(2, 3) = UatomsAniso(j).u(2, 3)
            atoms(i).u(1, 3) = UatomsAniso(j).u(1, 3)
            atoms(i).u(1, 2) = UatomsAniso(j).u(1, 2)
        End If
    Next
Next
End Sub
Private Sub make_bmat()
Dim i As Integer, j As Integer
Dim rad As Double
Dim snal As Double, snbe As Double, snga As Double, csal As Double
Dim csbe As Double, csga As Double, vol As Double, det As Double
Dim lat_con(5) As Double
Dim b_mat(3, 3) As Double, ginv(3, 3) As Double
Dim rec_lat_con(5) As Double
Dim b_inv_mat(3, 3) As Double

rad = 45 / Atn(1)

For i = 0 To 2
    For j = 0 To 2
        b_mat(j, i) = 0#
    Next
Next

lat_con(0) = celda.a
lat_con(1) = celda.b
lat_con(2) = celda.c
lat_con(3) = celda.alfa
lat_con(4) = celda.beta
lat_con(5) = celda.gamma

b_mat(0, 0) = 1#

csal = Cos(lat_con(3) / rad)
snal = Sin(lat_con(3) / rad)
csbe = Cos(lat_con(4) / rad)
snbe = Sin(lat_con(4) / rad)
csga = Cos(lat_con(5) / rad)
snga = Sin(lat_con(5) / rad)

For i = 0 To 2
    ginv(i, i) = lat_con(i) * lat_con(i)
Next
    ginv(0, 1) = lat_con(0) * lat_con(1) * csga
    ginv(1, 0) = ginv(0, 1)
    ginv(0, 2) = lat_con(0) * lat_con(2) * csbe
    ginv(2, 0) = ginv(0, 2)
    ginv(1, 2) = lat_con(1) * lat_con(2) * csal
    ginv(2, 1) = ginv(1, 2)
    b_mat(0, 1) = csga
    b_mat(1, 1) = snga
    b_mat(0, 2) = csbe
    b_mat(1, 2) = (csal - csbe * csga) / snga
    b_mat(2, 2) = Sqr(1# - b_mat(0, 2) * b_mat(0, 2) - b_mat(1, 2) * b_mat(1, 2))
    
    For i = 0 To 2
        For j = 0 To 2
            b_mat(j, i) = b_mat(j, i) * lat_con(i)
        Next
    Next
    
    vol = b_mat(0, 0) * (b_mat(1, 1) * b_mat(2, 2) - b_mat(1, 2) * b_mat(2, 1)) _
        - b_mat(0, 1) * (b_mat(1, 0) * b_mat(2, 2) - b_mat(1, 2) * b_mat(2, 0)) _
        + b_mat(0, 2) * (b_mat(1, 0) * b_mat(2, 1) - b_mat(1, 1) * b_mat(2, 0))
            
    rec_lat_con(0) = lat_con(1) * lat_con(2) * snal / vol
    rec_lat_con(1) = lat_con(0) * lat_con(2) * snbe / vol
    rec_lat_con(2) = lat_con(1) * lat_con(0) * snga / vol
    
    det = determinant(b_mat)
    
    b_inv_mat(0, 0) = (b_mat(1, 1) * b_mat(2, 2) - b_mat(1, 2) * b_mat(2, 1)) / det
    b_inv_mat(1, 0) = (b_mat(1, 0) * b_mat(2, 2) - b_mat(1, 2) * b_mat(2, 0)) / det
    b_inv_mat(2, 0) = (b_mat(1, 0) * b_mat(2, 1) - b_mat(1, 1) * b_mat(2, 0)) / det
    b_inv_mat(0, 1) = (b_mat(0, 1) * b_mat(2, 2) - b_mat(0, 2) * b_mat(2, 1)) / det
    b_inv_mat(1, 1) = (b_mat(0, 0) * b_mat(2, 2) - b_mat(0, 2) * b_mat(2, 0)) / det
    b_inv_mat(2, 1) = (b_mat(0, 0) * b_mat(2, 1) - b_mat(0, 1) * b_mat(2, 0)) / det
    b_inv_mat(0, 2) = (b_mat(0, 1) * b_mat(1, 2) - b_mat(0, 2) * b_mat(1, 1)) / det
    b_inv_mat(1, 2) = (b_mat(0, 0) * b_mat(1, 2) - b_mat(0, 2) * b_mat(1, 0)) / det
    b_inv_mat(2, 2) = (b_mat(0, 0) * b_mat(1, 1) - b_mat(0, 1) * b_mat(1, 0)) / det
End Sub

Private Function determinant(rot() As Double) As Double
Dim temp As Double
temp = rot(0, 0) * (rot(1, 1) * rot(2, 2) - rot(1, 2) * rot(2, 1)) _
      - rot(1, 1) * (rot(0, 1) * rot(2, 2) - rot(0, 2) * rot(2, 1)) _
      + rot(2, 2) * (rot(0, 1) * rot(1, 2) - rot(0, 2) * rot(1, 1))
determinant = temp
End Function


