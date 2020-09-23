Attribute VB_Name = "mPrintMatrix"
Option Explicit
Option Base 1
'Const MAXIMO = 3
'**********************************************************
'
'**********************************************************
Public Sub printMatrix(a() As Double)
Dim i As Integer, j As Integer
Dim MAXIMO
MAXIMO = UBound(a())
For i = 1 To MAXIMO
    For j = 1 To MAXIMO
        Debug.Print formatNumero(a(i, j)),
    Next j
    Debug.Print
Next i
End Sub
'**********************************************************
'
'**********************************************************
Public Sub printVector(a() As Double)
Dim i As Integer
Dim MAXIMO
MAXIMO = UBound(a())
For i = 1 To MAXIMO
    Debug.Print formatNumero(a(i)),
Next i
Debug.Print
End Sub
'**********************************************************
'
'**********************************************************
Public Sub printTVector(a As TATOMO)
Debug.Print "VECTOR"
Debug.Print Format(a.x, formato) & "X+ " & _
            Format(a.y, formato) & "Y+ " & _
            Format(a.Z, formato) & "Z+ "
Debug.Print
End Sub
'**********************************************************
'
'**********************************************************
Public Sub printTPlano(a As TPlano)
Debug.Print
Debug.Print "PLANO"
Debug.Print Format(a.a, formato) & "X+ " & _
            Format(a.b, formato) & "Y+ " & _
            Format(a.c, formato) & "Z+ " & _
            Format(a.d, formato)
Debug.Print
End Sub
'**********************************************************
'
'**********************************************************
Public Sub imprimirCoordenadas()
Dim i As Integer
    For i = 1 To nAtoms
        Debug.Print atoms(i).x; atoms(i).y, atoms(i).Z
    Next i
End Sub

