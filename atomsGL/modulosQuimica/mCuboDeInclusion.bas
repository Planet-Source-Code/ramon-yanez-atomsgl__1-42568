Attribute VB_Name = "mCuboDeInclusion"
Option Explicit
'**********************************************************
Public Sub CuboDeInclusion()
Dim i As Integer
Dim anchoX As Single, anchoY As Single, anchoZ As Single
Dim xmin As Single, ymin As Single, zmin As Single
Dim xmax As Single, ymax As Single, zmax As Single
xmin = atoms(1).X
ymin = atoms(1).Y
zmin = atoms(1).Z
For i = 1 To nAtoms - 1
        With atoms(i)
        If .X > xmax Then
                    xmax = .X
                    ElseIf .X < xmin Then xmin = .X
        End If
        
        If .Y > ymax Then
                    ymax = .Y
                    ElseIf .Y < ymin Then ymin = .Y
        End If
                    
        If .Z > zmax Then
                    zmax = .Z
                    ElseIf .Z < zmin Then zmin = .Z
        End If
                
        End With
        
Debug.Print "xmax= "; xmax, "ymax= "; ymax, "zmax= "; zmax
Debug.Print "xmin= "; xmin, "ymin= "; ymin, "zmin= "; zmin

Next i


anchoX = Abs(xmax) - Abs(xmin)
anchoY = Abs(ymax) - Abs(ymin)
anchoZ = Abs(zmax) - Abs(zmin)

Debug.Print "anchoX= "; anchoX, "anchoY= "; anchoY, "anchoZ= "; anchoZ

centroX = xmin + anchoX / 2
centroy = ymin + anchoY / 2
centroz = zmin + anchoZ / 2

Debug.Print centroX, centroy, centroz

For i = 1 To nAtoms
    atoms(i).X = atoms(i).X - centroX
    atoms(i).Y = atoms(i).Y - centroy
    atoms(i).Z = atoms(i).Z - centroz
Next i
End Sub
