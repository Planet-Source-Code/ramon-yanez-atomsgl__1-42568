Attribute VB_Name = "GLinicializacion"
Option Explicit

' +-------------------------------------------------------------------
' | Algunas variables
' +-------------------------------------------------------------------
Public hPalette As Long
Public hGLRC As Long

Public doubleBuffer As GLboolean
Public displayListInited As GLboolean
Public displayPlaneInited As GLboolean
Public displaySelectionInited As GLboolean
Public bListaSimetria As GLboolean


Public Sub SetupPalette(ByVal lhDC As Long)
    
' +-------------------------------------------------------------------
' | Initialize the Win32 form pallete.
' +-------------------------------------------------------------------
    
    Dim PixelFormat As Long
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pPal As LOGPALETTE
    Dim PaletteSize As Long
    PixelFormat = GetPixelFormat(lhDC)
    DescribePixelFormat lhDC, PixelFormat, Len(pfd), pfd
    
    If (pfd.dwFlags And PFD_NEED_PALETTE) <> 0 Then
        PaletteSize = 2 ^ pfd.cColorBits
    Else
        Exit Sub
    End If
    
    pPal.palVersion = &H300
    pPal.palNumEntries = PaletteSize
    Dim redMask As Long
    Dim GreenMask As Long
    Dim BlueMask As Long
    Dim i As Long
    redMask = 2 ^ pfd.cRedBits - 1
    GreenMask = 2 ^ pfd.cGreenBits - 1
    BlueMask = 2 ^ pfd.cBlueBits - 1
    For i = 0 To PaletteSize - 1
        With pPal.palPalEntry(i)
            .peRed = i
            .peGreen = i
            .peBlue = i
            .peFlags = 0
        End With
    Next
    GetSystemPaletteEntries lhDC, 0, 256, VarPtr(pPal.palPalEntry(0))
    hPalette = CreatePalette(pPal)
    If hPalette <> 0 Then
        SelectPalette lhDC, hPalette, False
        RealizePalette lhDC
    End If
End Sub

Public Sub SetupPixelFormat(ByVal hDC As Long)

' +-------------------------------------------------------------------
' | Retrieve/set a Win32 pixel format for OpenGL modes with double-
' | buffering, and direct draw to window with RGBA color mode.
' | 16bit (65536 colors) depth is preferable.
' +-------------------------------------------------------------------

    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 16
    pfd.cDepthBits = 16
    pfd.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hDC, pfd)
    If PixelFormat = 0 Then FatalError "Could not retrieve pixel format!"
    SetPixelFormat hDC, PixelFormat, pfd
End Sub

Public Sub FatalError(ByVal strMessage As String)

' +-------------------------------------------------------------------
' | A simple exit handler should something NASTY happen.
' +-------------------------------------------------------------------

    MsgBox "Fatal Error: " & strMessage, vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, "Fatal Error In " & App.Title
    Unload frmMain
    Set frmMain = Nothing
    End
End Sub

' +-------------------------------------------------------------------
' |
' |
' +-------------------------------------------------------------------
Public Sub InitGL(pctOpenGL As Object)
    
    'Setting the variables for trackball
    spinning = False
    moving = False
    anim = False
    h = pctOpenGL.ScaleHeight
    w = pctOpenGL.ScaleWidth
    newModel = True
    scaling = False
    scalefactor = 1
    TrackBall curquat, 0, 0, 0, 0
   ' End of setings for trackball
      
    Dim hGLRC As Long
    doubleBuffer = GL_TRUE
   
    SetupPixelFormat pctOpenGL.hDC
    hGLRC = wglCreateContext(pctOpenGL.hDC) '?????
      
    ' bind the rendering context to the window
    wglMakeCurrent pctOpenGL.hDC, hGLRC
       
    ' ---------------------------------------------
    ' configure the OpenGL context for rendering
    '----------------------------------------------
    
    glEnable GL_DEPTH_TEST

    glDepthFunc GL_LESS
    glClearDepth 1
    ' Borramos el frame Buffer de color gris
    glClearColor 0.5, 0.5, 0.5, 0 '
    
    ' set up projection transform
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glFrustum -30, 30, -30, 30, -30, 30 ' FUERA DE ESTOS LIMITES NO SE VE NADA
    glViewport 0, 0, w, h
    gluPerspective 40, w / h, 1, 40
       
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    gluLookAt 0, 0, 10, 0, 0, 0, 0, 1, 0
 
    glEnable GL_DEPTH_TEST
    glEnable GL_DITHER
    glDepthFunc GL_LESS
    glClearDepth 1
    glEnable GL_COLOR_MATERIAL
    glEnable GL_NORMALIZE
    
End Sub

' +-------------------------------------------------------------------
' |
' |
' +-------------------------------------------------------------------
Public Sub ShowLights()

' set the lights and mat property for lights

Dim MatSpecular(3) As GLfloat
Dim MatShininess(0) As GLfloat
Dim LightPosition(3) As GLfloat
Dim LightColor(3) As GLfloat

    MatSpecular(0) = 1
    MatSpecular(1) = 1
    MatSpecular(2) = 1
    MatSpecular(3) = 1

    MatShininess(0) = 50

     LightPosition(0) = 10
    LightPosition(1) = 4
    LightPosition(2) = 10
    LightPosition(3) = 1
  
    LightColor(0) = 0.8
    LightColor(1) = 1
    LightColor(2) = 0.8
    LightColor(3) = 1
    
    glMaterialfv GL_FRONT, GL_SPECULAR, MatSpecular(0)
    glMaterialfv GL_FRONT, GL_SHININESS, MatShininess(0)

    glLightModeli GL_LIGHT_MODEL_LOCAL_VIEWER, 1
    glLightfv GL_LIGHT0, GL_POSITION, LightPosition(0)
    glLightfv GL_LIGHT0, GL_DIFFUSE, LightColor(0)
    glLightf GL_LIGHT0, GL_CONSTANT_ATTENUATION, 0.1
    glLightf GL_LIGHT0, GL_LINEAR_ATTENUATION, 0.05
    
    glEnable GL_LIGHTING
    glEnable GL_LIGHT0
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA 'permitir transparencias
End Sub

' +-------------------------------------------------------------------
' | MiProyeccion
' +-------------------------------------------------------------------
Public Sub MiProyeccion()
    glViewport 0, 0, w, h
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    MiCamara
End Sub

' +-------------------------------------------------------------------
' | MiCamara
' +-------------------------------------------------------------------
Public Sub MiCamara()
    gluPerspective 40, w / h, 1, 40
    gluLookAt 0, 0, 10, 0, 0, 0, 0, 1, 0
End Sub



