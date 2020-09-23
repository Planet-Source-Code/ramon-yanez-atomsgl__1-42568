Attribute VB_Name = "mRotate"
Option Explicit
Const cosAngulo = 0.999390827
Const sinAngulo = 0.034899496

'**********************************************************
'
'**********************************************************
Public Sub rotateXdown(pct As Object)
Dim xg As Single, yg As Single
Dim i As Integer

For i = 1 To nAtoms
    xg = atoms(i).dx * cosAngulo - atoms(i).dz * sinAngulo
    atoms(i).dz = atoms(i).dx * sinAngulo + atoms(i).dz * cosAngulo
    atoms(i).dx = xg
Next i
    If Form1!inercia.Checked Then
        For i = 1 To 3
            xg = vInercia(1, i) * cosAngulo - vInercia(3, i) * sinAngulo
            vInercia(3, i) = vInercia(1, i) * sinAngulo + vInercia(3, i) * cosAngulo
            vInercia(1, i) = xg
        Next i
    End If
Call dibujarMolecula(pct)
End Sub
'**********************************************************
'
'**********************************************************
Public Sub rotateXup(pct As Object)
Dim xg As Single
Dim i As Single

For i = 1 To nAtoms
    xg = atoms(i).dx * cosAngulo + atoms(i).dz * sinAngulo
    atoms(i).dz = -atoms(i).dx * sinAngulo + atoms(i).dz * cosAngulo
    atoms(i).dx = xg
Next i
    If Form1!inercia.Checked Then
        For i = 1 To 3
            xg = vInercia(1, i) * cosAngulo + vInercia(3, i) * sinAngulo
            vInercia(3, i) = -vInercia(1, i) * sinAngulo + vInercia(3, i) * cosAngulo
            vInercia(1, i) = xg
        Next i
    End If
Call dibujarMolecula(pct)
End Sub

'**********************************************************
'
'**********************************************************
Public Sub rotateYdown(pct As Object)
Dim yg As Single
Dim i As Integer

For i = 1 To nAtoms
    yg = atoms(i).dy * cosAngulo - atoms(i).dz * sinAngulo
    atoms(i).dz = atoms(i).dy * sinAngulo + atoms(i).dz * cosAngulo
    atoms(i).dy = yg
Next i
    If Form1!inercia.Checked Then
        For i = 1 To 3
            yg = vInercia(2, i) * cosAngulo - vInercia(3, i) * sinAngulo
            vInercia(3, i) = vInercia(2, i) * sinAngulo + vInercia(3, i) * cosAngulo
            vInercia(2, i) = yg
        Next i
    End If
Call dibujarMolecula(pct)
End Sub
'**********************************************************
'
'**********************************************************
Public Sub rotateYup(pct As Object)
Dim yg As Single
Dim i As Integer

For i = 1 To nAtoms
    yg = atoms(i).dy * cosAngulo + atoms(i).dz * sinAngulo
    atoms(i).dz = -atoms(i).dy * sinAngulo + atoms(i).dz * cosAngulo
    atoms(i).dy = yg
Next i
    If Form1!inercia.Checked Then
        For i = 1 To 3
            yg = vInercia(2, i) * cosAngulo + vInercia(3, i) * sinAngulo
            vInercia(3, i) = -vInercia(2, i) * sinAngulo + vInercia(3, i) * cosAngulo
            vInercia(2, i) = yg
        Next i
    End If
Call dibujarMolecula(pct)
End Sub




