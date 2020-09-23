Attribute VB_Name = "GLSeleccion"
Option Base 1
Option Explicit

' +-------------------------------------------------------------------
' |     Selection
' +-------------------------------------------------------------------
Sub Selection(x As Single, y As Single)
Dim hits As Long
Dim viewport(0 To 3) As Long ' array for viewport
Dim selectbuf(0 To 512) As GLuint ' array for the reultes of picking
Dim BUFSIZE As Integer
BUFSIZE = 512
    glGetIntegerv GL_VIEWPORT, viewport(0)
    glSelectBuffer BUFSIZE, selectbuf(0)
'    glRenderMode GL_SELECT
    glInitNames
    glPushName 0 'intializing the name stack
    'setting the same geomety as for normal view
    glMatrixMode GL_PROJECTION
    glPushMatrix 'saving the original matrix
        glRenderMode GL_SELECT
        glLoadIdentity
        '2*2 region near cursor
        gluPickMatrix x, viewport(3) - y, 2, 2, viewport(0)
        MiCamara
        glMatrixMode GL_MODELVIEW
        glCallList LISTA_MOLECULA
        glFlush
        hits = glRenderMode(GL_RENDER)   'hits = objects in hit
        processHits hits, selectbuf, x, y
        glMatrixMode GL_PROJECTION
    glPopMatrix
    glMatrixMode GL_MODELVIEW
End Sub

' +-------------------------------------------------------------------
' |     processHits
' +-------------------------------------------------------------------
Private Sub processHits(hits As Long, selectbuf, x As Single, y As Single)
Dim i As Integer
Dim near As Long
Dim dist As Single, helpvar As GLuint
Dim texto
    If hits = 0 Then
       borrarStatusBar
       BorrarSeleccion
       If bDibPla2 Then
         bDibPla1 = True
       Else
         bDibPla1 = False
         frmMain!chkPla1 = False
         BorrarPlano miPlano
         BorrarPlano plano1
       End If
       Exit Sub
    End If
    If hits > 0 Then
'       Beep
        near = 0
         dist = 1E+30 ' very big number :)
         For i = 1 To hits
             helpvar = selectbuf((i - 1) * 4 + 1)
             If (GLuint_to_Single(helpvar) < dist) Then 'The distance to picked objekt
                 dist = GLuint_to_Single(helpvar)
                 near = selectbuf((i - 1) * 4 + 3) ' near = number of picked objekt
             End If
         Next i
         Debug.Print "near= "; near
         frmMain!StatusBar1.Panels(1).Text = ""
         If near = 0 Then Exit Sub
         ' seleccion multiple?
         If SelectMul Then
             Dim atmSelcc
             
             'Hay algo seleccionado?
             If UBound(atmsSelecds()) < 1 Then Exit Sub
             
             'comprobamos si ya esta en la lista
             For Each atmSelcc In atmsSelecds()
                If atmSelcc = near Then Exit Sub
             Next
             ' como no esta en la lista lo añadimos
           
             seleccionados = seleccionados + 1
             
         Else 'no hay seleccion multiple
             seleccionados = 1
         End If
        If seleccionados > 1 Then
            ReDim Preserve atmsSelecds(seleccionados)
        Else
            ReDim atmsSelecds(1) '? no puedo poner preserve?
        End If
         If near <> 0 Then
            atmsSelecds(seleccionados) = near
            For Each atmSelcc In atmsSelecds()
                texto = texto & atoms(atmSelcc).simbol & atmSelcc & " "
            Next
         End If
         frmMain!StatusBar1.Panels(1).Text = texto
    End If
    displayListInited = GL_FALSE
    displaySelectionInited = GL_FALSE
End Sub

' +-------------------------------------------------------------------
' | Gluint_to_Single
' +-------------------------------------------------------------------
Private Function GLuint_to_Single(GLuint As GLuint) As Single
' Transformation of GLuint variable (UnsignedLongInteger in C , no equivalent in VB) to Single
Dim result As Single
    result = GLuint
    If result < 0 Then result = 2147483648# + (2147483648# + result)
    GLuint_to_Single = result
End Function

' +-------------------------------------------------------------------
' |
' +-------------------------------------------------------------------
Public Sub BorrarSeleccion()
    ReDim atmsSelecds(1)
    atmsSelecds(1) = 0
    seleccionados = 0
    displaySelectionInited = GL_FALSE
    ListaatmsSelecds
End Sub
