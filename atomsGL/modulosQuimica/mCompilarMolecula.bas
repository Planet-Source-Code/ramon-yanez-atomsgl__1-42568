Attribute VB_Name = "MCompilarMolecula"
Option Explicit
Option Base 1
Private QuadObj&
Private quadobj1&
Private Const LargoCilindro = 6

' +-------------------------------------------------------------------
' | MiMolecula
' +-------------------------------------------------------------------

Public Sub MiMolecula()
Dim i As Integer, j As Integer
Dim cP As TColorRGB
    glNewList LISTA_MOLECULA, GL_COMPILE_AND_EXECUTE
    glPushMatrix
    glInitNames
    glPushName (0)
    QuadObj = gluNewQuadric
    For i = 1 To nAtoms
        glPushMatrix
        With atoms(i)
                cP = ConvertRGB(QBColor(.color))
                glColor3ub cP.r, cP.g, cP.b
                glTranslatef .x, .y, .Z
                glLoadName (i)
                gluSphere QuadObj, atoms(i).radioCov / 3.5, 25, 25
        End With
        glPopMatrix
    Next i
        glPopName

    gluDeleteQuadric QuadObj
    QuadObj = gluNewQuadric
    For i = 1 To nEnlaces
       With enlace(i)
            glPushMatrix
                'primer enlace
                glColor3f .ca.r, .ca.g, .ca.b
                glTranslatef .x0, .y0, .z0
                glRotatef .angle, .vy, -.vx, 0
                gluCylinder QuadObj, 0.05, 0.05, (.d) / 2, 20, 1
            glPopMatrix
            glPushMatrix
                'primer enlace
                glColor3f .cb.r, .cb.g, .cb.b
                glTranslatef .x0, .y0, .z0
                glRotatef .angle, .vy, -.vx, 0
                glTranslatef 0, 0, .d / 2
                gluCylinder QuadObj, 0.05, 0.05, (.d) / 2, 20, 1
            glPopMatrix

        End With
    Next i
    gluDeleteQuadric QuadObj
    glPopMatrix
    glEndList
End Sub

' +-------------------------------------------------------------------
' | Plano
' +-------------------------------------------------------------------
Sub DibujoPlano(nPlano As TPlano, bcnt As TVECTOR, Optional color As Single)
        Dim angle As Single
        Dim cP As TColorRGB
        cP = ConvertRGB(color)
        glPushMatrix
            glTranslatef bcnt.x, bcnt.y, bcnt.Z
            glEnable GL_BLEND
            glColor4ub cP.r, cP.g, cP.b, 128
            With nPlano
                angle = -Arccos(CSng(.c)) * 180 / PI
                angle = angle
                glRotatef angle, .b, -.a, 0
                glRectf -2#, 2#, 2#, -2#
            End With
        glPopMatrix
End Sub

' +-------------------------------------------------------------------
' | Lista Atomos Seleccionados
' +-------------------------------------------------------------------
Public Sub ListaatmsSelecds()
Dim atmSelcc
glNewList LISTA_ATOMOS_SELECCIONADOS, GL_COMPILE_AND_EXECUTE
    QuadObj = gluNewQuadric
    For Each atmSelcc In atmsSelecds()
        If atmSelcc <> 0 Then ' hay seleccionado
            glEnable GL_BLEND
            glColor3f 1#, 1#, 0# 'atomos seleccionados de color amarillo
            glPushMatrix
                With atoms(atmSelcc)
                    glTranslatef .x, .y, .Z
                    gluSphere QuadObj, atoms(atmSelcc).radioCov / 3.4, 25, 25
                End With
            glPopMatrix
        End If
    Next
    gluDeleteQuadric QuadObj
glEndList
End Sub

' +-------------------------------------------------------------------
'
' +-------------------------------------------------------------------
Public Sub ListaElementosSimetria()
Dim i
Dim angle As Single
Dim cP As TColorRGB
glNewList LISTA_SIMETRIA, GL_COMPILE_AND_EXECUTE
    glPushMatrix
    QuadObj = gluNewQuadric
    glEnable GL_BLEND
For i = 1 To nElemSim
    If ElemSim(i).tipo = "sigma" And ElemSim(i).visible Then
        glPushMatrix
'            glDisable GL_DEPTH_TEST
            glColor4f 1, 0, 0, 0.5
            angle = -Arccos(CSng(ElemSim(i).Z)) * 180 / PI
            glRotatef angle, ElemSim(i).y, -ElemSim(i).x, 0
            glRectf -2#, 2#, 2#, -2#
        glEnd
        glPopMatrix
    ElseIf InStr(ElemSim(i).tipo, "C") And ElemSim(i).visible Then
            glPushMatrix
                glColor3f 1, 0, 1
                angle = -Arccos(CSng(ElemSim(i).Z)) * 180 / PI
                glRotatef angle, ElemSim(i).y, -ElemSim(i).x, 0
                glTranslatef 0#, 0#, -4#
                gluCylinder QuadObj, 0.02, 0.02, 8, 20, 1
            glPopMatrix
        ElseIf InStr(ElemSim(i).tipo, "S") And ElemSim(i).visible Then
                glPushMatrix 'S
                    glColor4f 1, 1, 0, 0.5
                    angle = -Arccos(CSng(ElemSim(i).Z)) * 180 / PI
                    glRotatef angle, ElemSim(i).y, -ElemSim(i).x, 0
                    glTranslatef 0#, 0#, -4#
                    gluCylinder QuadObj, 0.04, 0.04, 8, 20, 1
                glPopMatrix
            ElseIf InStr(ElemSim(i).tipo, "I") And ElemSim(i).visible Then
                glPushMatrix
                    glColor4f 0#, 1#, 0#, 0.9
                    gluSphere QuadObj, 0.3, 25, 25
                glPopMatrix
        End If
Next i
gluDeleteQuadric QuadObj

glPopMatrix
glEndList
End Sub
