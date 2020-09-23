Attribute VB_Name = "GLgeometria"
'Module with some geometric object declarations
 
 Option Explicit
 Sub GLCubo(a As Double, b As Double, c As Double)
 
  glBegin GL_QUADS
 
    glNormal3f 0#, 0#, -1
    'glTexCoord2f 0#, 0#:
    glVertex3f 0#, 0#, 0#:
    'glTexCoord2f 1#, 0#:
    glVertex3f a#, 0#, 0#:
    'glTexCoord2f 1#, 1#:
    glVertex3f a#, b#, 0#:
    'glTexCoord2f 0#, 1#:
    glVertex3f 0#, b#, 0#:

    glNormal3f 0#, 0#, 1#:
    'glTexCoord2f 0#, 0#:
    glVertex3f 0#, 0#, c#:
    'glTexCoord2f 1#, 0#:
    glVertex3f a#, 0#, c#:
    'glTexCoord2f 1#, 1#:
    glVertex3f a#, b#, c#:
    'glTexCoord2f 0#, 1#:
    glVertex3f 0#, b#, c#:

    glNormal3f 0#, 1#, 0#:
    'glTexCoord2f 0#, 0#:
    glVertex3f 0#, b#, 0#:
    'glTexCoord2f 1#, 0#:
    glVertex3f a#, b#, 0#:
    'glTexCoord2f 1#, 1#:
    glVertex3f a#, b#, c#:
    'glTexCoord2f 0#, 1#:
    glVertex3f 0#, b#, c#:

    glNormal3f 0#, -1#, 0#:
    'glTexCoord2f 0#, 0#:
    glVertex3f 0#, 0#, 0#:
    'glTexCoord2f 1#, 0#:
    glVertex3f a#, 0#, 0#:
    'glTexCoord2f 1#, 1#:
    glVertex3f a#, 0#, c#:
    'glTexCoord2f 0#, 1#:
    glVertex3f 0#, 0#, c#:

    glNormal3f 1#, 0#, 0#:
    'glTexCoord2f 0#, 0#:
    glVertex3f a#, 0#, 0#:
    'glTexCoord2f 1#, 0#:
    glVertex3f a#, b#, 0#:
    'glTexCoord2f 1#, 1#:
    glVertex3f a#, b#, c#:
    'glTexCoord2f 0#, 1#:
    glVertex3f a#, 0#, c#:

    glNormal3f -1#, 0#, 0#:
    'glTexCoord2f 0#, 0#:
    glVertex3f 0#, 0#, 0#:
    'glTexCoord2f 1#, 0#:
    glVertex3f 0#, b#, 0#:
    'glTexCoord2f 1#, 1#:
    glVertex3f 0#, b#, c#:
    'glTexCoord2f 0#, 1#:
    glVertex3f 0#, 0#, c#:
    glEnd

 End Sub
'
'
' Sub Esfera(objekt As GLUquadric, radio, calidad)
'    gluSphere objekt, radio, calidad, calidad
' End Sub
'
' Sub cilidro(objekt As GLUquadric, podstava, delka, calidad)
'    gluCylinder objekt, podstava, podstava, delka, calidad, calidad
' End Sub
 
'Sub OtoceniKoordinat(u1, u2, u3)
'    glRotatef u1, 1, 0, 0
'    glRotatef u2, 0, 1, 0
'    glRotatef u3, 0, 0, 1
'End Sub

Sub Cubo(a As Double)
GLCubo a, a, a
End Sub
