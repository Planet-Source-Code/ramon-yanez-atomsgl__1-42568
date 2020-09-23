Attribute VB_Name = "GlRender"
Option Explicit
Option Base 1

' +-------------------------------------------------------------------
' | render
' +-------------------------------------------------------------------
Public Sub render(Optional pctOpenGL As PictureBox)
    
    ' adition for trackball
    If newModel Then
        recalcModelView
    End If
    ' end of adition for trackball
' +-------------------------------------------------------------------
' | Borramos la Pantalla antes de dibujar nada
' +-------------------------------------------------------------------

glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

' +-------------------------------------------------------------------
' | If a display list has been created, use it.  Otherwise, create it.
' +-------------------------------------------------------------------
   
        If bsimetria Then
        If bListaSimetria = GL_TRUE Then
            glCallList LISTA_SIMETRIA
        Else
            ListaElementosSimetria
            bListaSimetria = GL_TRUE
        End If
    End If

    If displayListInited = GL_TRUE Then
        glCallList LISTA_MOLECULA
    Else
        MiMolecula  'primera inicializacion del objeto molecula
        displayListInited = GL_TRUE
    End If
    
    If displaySelectionInited = GL_TRUE Then
        glCallList LISTA_ATOMOS_SELECCIONADOS
    Else
        ListaatmsSelecds
        displaySelectionInited = GL_TRUE
    End If
    
    If bDibPla1 Or bDistPla1 Then
        If bDibPla2 Then
            DibujoPlano plano1, bcntPlano1, vbBlue
            DibujoPlano miPlano, baricentroPlano, vbGreen
        Else
            DibujoPlano miPlano, baricentroPlano, vbBlue
        End If
    End If
        glFlush
        SwapBuffers pctOpenGL.hDC
End Sub

' +-------------------------------------------------------------------
' |  recalcModelView
' +-------------------------------------------------------------------
Private Sub recalcModelView()
    Dim m(0 To 15) As GLfloat
    
    glPopMatrix
    glPushMatrix
    build_rotmatrix m, curquat
    glMultMatrixf m(0)
    If scalefactor = 1 Then
            glDisable GL_NORMALIZE
        Else
            glEnable GL_NORMALIZE
    End If
    
    glScalef scalefactor, scalefactor, scalefactor
    newModel = False

End Sub

