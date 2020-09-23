Attribute VB_Name = "mSaveBMP"
Option Explicit

'-------------------------------------------------------------------------
'Save a 24bit bmp to file. This routine can save a selction (a subimage).
'based in part on a routine in OpenGL SuperBible. See same for a routine to print.
'-------------------------------------------------------------------------
Public Function SaveBMP_RP(Filename$, ByVal w&, ByVal h&) As Boolean
Dim infoHeader As BITMAPINFOHEADER, pixels() As Byte
Dim fileHeader As BITMAPFILEHEADER
Dim fn&, s$
Dim i&, j&, n&
Dim temp As Byte
Dim viewport&(0 To 3), pad&
Dim Width&, ww&
On Error GoTo eh
    w = w + 1
    'check range
    glGetIntegerv GL_VIEWPORT, viewport(0)
    'save as 24bit bmp, 3 bytes per color
    Width = (w * 3)
    'align to 4 bytes
    pad = Width Mod 4
    If pad Then pad = 4 - pad
    Width = Width + pad
    ReDim pixels(0 To Width * h - 1)
    'may need to rerender if you have raised a dialog box
    render frmMain!pctOpenGL
    ' Read the pixels
    glPixelStorei GL_UNPACK_ALIGNMENT, 4  ' Force 4-byte alignment
    glPixelStorei GL_UNPACK_SKIP_ROWS, 0
    glPixelStorei GL_UNPACK_SKIP_PIXELS, 0
    glPixelStorei GL_UNPACK_ROW_LENGTH, 0 'full image
    glReadPixels 0, 0, w, h, GL_RGB, GL_UNSIGNED_BYTE, pixels(0)

    ' set the pixel data in the pic
    infoHeader.biSize = Len(infoHeader)
    infoHeader.biPlanes = 1
    infoHeader.biBitCount = 24
    infoHeader.biCompression = BI_RGB
    infoHeader.biSizeImage = Width * Abs(h)
    infoHeader.biXPelsPerMeter = 0
    infoHeader.biYPelsPerMeter = 0
    infoHeader.biClrUsed = 0
    infoHeader.biClrImportant = 0
    infoHeader.biWidth = w
    infoHeader.biHeight = h
    'kill existing file
    s = Dir(Filename)
    If Len(s) Then Kill Filename
    'Definimos lo que va en la cabecera del BMP
    fileHeader.bfSize = Len(fileHeader)
    fileHeader.bfType = &H4D42
    fileHeader.bfReserved1 = 0
    fileHeader.bfReserved2 = 0
    fileHeader.bfOffBits = Len(fileHeader) + Len(infoHeader)
' Swap red and blue for the bitmap...
    For i = 0 To infoHeader.biSizeImage - 1 Step 3
            temp = pixels(i)
            pixels(i) = pixels(i + 2)
            pixels(i + 2) = temp
    Next
    fn = FreeFile
    Open Filename For Binary As fn
        Put #fn, , fileHeader
        Put #fn, , infoHeader
        Put #fn, , pixels
    Close fn
    SaveBMP_RP = True
    Exit Function
eh:
    Debug.Assert 0
    Exit Function
    Resume Next
End Function
'
''-------------------------------------------------------------------------
'' Save a bmp using a Memory Device Context.
'' The following function supports arbitrary resolutions,
'' independent of the OpenGL window resolution.
'' based in part on code by matumot at 'OpenGL Paradise'.
''-------------------------------------------------------------------------
'Public Function SaveBMP_MemDC(X&, Y&, w&, h&) As Boolean
'Dim bits() As Byte
'Dim hDC&, memDC&, hmemrc&, file$
'Dim bpp&
'Dim bfh As BITMAPFILEHEADER, bi As BITMAPINFO
'Dim hBmp&, hBmpOld&, ptr&
'Dim dwSizeImage&, s$
'On Error GoTo eh
'    '
'    Screen.MousePointer = 11
'    hDC = frmMain!pctOpenGL.hDC
'    file = App.Path & "\savememdc.bmp"
'    s = Dir(file)
'    If Len(s) Then Kill file
'    ' Create Compatible Memory Device Context
'    memDC = CreateCompatibleDC(hDC)
'    If memDC = 0 Then
'        Debug.Print "CompatibleDC Error"
'        GoTo eh
'    End If
'    ' Make a DIB image which is a multiple of the size of GL window
'    ' and full color(24 bpp)
'    'This doesn't work when you try to make it larger, but smaller seems ok.
'    'Howevever, the header isn't quite right.
'    'w = w / 2
'    'h = h / 2
'    bpp = 24
'    dwSizeImage = w * h * bpp / 8 ' i.e. 3
'    ' set the pixel data in the pic
'    bi.bmiHeader.biSize = Len(bi.bmiHeader)
'    bi.bmiHeader.biWidth = w
'    bi.bmiHeader.biHeight = h
'    bi.bmiHeader.biPlanes = 1
'    bi.bmiHeader.biBitCount = 24
'    bi.bmiHeader.biCompression = BI_RGB
'    bi.bmiHeader.biSizeImage = dwSizeImage
'    bi.bmiHeader.biXPelsPerMeter = 2952
'    bi.bmiHeader.biYPelsPerMeter = 2952
'
'    ' Create a DIB surface
'    hBmp = CreateDIBSection(hDC, bi, DIB_RGB_COLORS, ptr, 0, 0)
'    If hBmp = 0 Then
'        Debug.Print "CreateDIBSection Error"
'        GoTo eh
'    End If
'    ' Select the DIB Surface into the dc
'    hBmpOld = SelectObject(memDC, hBmp)
'    If hBmpOld = 0 Then
'        Debug.Print "Select Object Error"
'        GoTo eh
'    End If
'    ' Set up a Pixel format for the DIB surface
'    If Not m_SetupPixelFormat(memDC, hmemrc) Then
'        Debug.Print "SetPixelFormat failed"
'        GoTo eh
'    End If
'    'replicate the setup of the current rendering context.
'    'this was developed by looking at the ocx setup code
'    'and the setup code in the CX class. In general, whatever you do
'    'to the context - in setup or later - you must reproduce here.
'    'Then call your Draw routine and draw as usual.
'    With gCtl
'        'from CXX setup
'        glClearColor 0.3, 0.3, 0.3, 0
'        'from the ocx
'        'depth
'        glClearDepth 1
'        glEnable GL_DEPTH_TEST
'        glEnable GL_LIGHTING
'        'from CXX setup
'        With gCtl.Lights.Item(liLight0)
'            .SetAmbient 0.1, 0.1, 0.1
'            .SetDiffuse 1, 1, 1
'            .SetPosition 90, 90, 150
'            .Enabled = True
'        End With
'        With gCtl.Camera
'            .FarPlane = 100
'            .NearPlane = 0.1
'            .FieldOfView = 45
'            .SetEyePos 0, 0, 10
'            .SetTargetPos 0, 0, 0
'        End With
'        'couldndn't use the ocx grids because of display list problems
'        CreateGrids
'        'from the ocx
'        'the following will issue the same commands they usually do,
'        'and luckily it affects this rc correctly.
'        .TrackBall.Resize w, h
'        .Camera.OnSize w, h, True
'        'paint. This stuff is done in ocx Render.
'        glClear clrColorBufferBit Or clrDepthBufferBit
'        glLoadIdentity
'        .Camera.Update
'        'do our ordinary drawing
'        cx.Draw
'        glFlush
'    End With
'    ' Prepare BMP file header information
'    bfh.bfType = &H4D42
'    bfh.bfSize = Len(bfh) + Len(bi) + bi.bmiHeader.biSizeImage
'    bfh.bfOffBits = Len(bfh) + Len(bi)
'    Dim fn&
'    fn = FreeFile
'    Open file For Binary As fn
'        Put #fn, , bfh
'        Put #fn, , bi
'        ReDim bits(0 To dwSizeImage - 1)
'        CopyMemory bits(0), ByVal ptr, dwSizeImage
'        Put #fn, , bits
'    Close fn
'
'    wglMakeCurrent 0, 0
'    wglDeleteContext hmemrc
'    hBmp = SelectObject(memDC, hBmpOld)
'    DeleteObject hBmp
'    Dim r&
'    r = DeleteDC(memDC)
'    SaveBMP_MemDC = True
'    Screen.MousePointer = 0
'    Exit Function
'eh:
'    Screen.MousePointer = 0
'    Debug.Assert 0
'    Exit Function
'    Resume Next
'End Function

