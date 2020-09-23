VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AtomsGL v1.5  (c) Ramón Yáñez 2002 UAB"
   ClientHeight    =   8205
   ClientLeft      =   405
   ClientTop       =   1290
   ClientWidth     =   11220
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   748
   Begin VB.Frame frameSinNombre 
      Height          =   7575
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton btnSalvarImagen 
         Caption         =   "Guardar Imatge"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   7200
         Width           =   1575
      End
      Begin VB.ListBox List2 
         Height          =   1035
         Left            =   240
         TabIndex        =   8
         Top             =   5760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   240
         TabIndex        =   7
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Planes"
         Height          =   1695
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1695
         Begin VB.TextBox txtAnglePlans 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkDistPlans 
            Caption         =   "Dist Atom Pla 1"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkPla1 
            Caption         =   "Plane #1"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkAnglePlans 
            Caption         =   "Plane Angle"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame frCalculs 
         Caption         =   "Measure"
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         Begin VB.TextBox txtDiedre 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtAngle 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtDist 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Torsion"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Angle"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Distance"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Left Mouse: Rotate         Right Mouse: Select        Shift+RM: Extended "
         Height          =   975
         Left            =   240
         TabIndex        =   20
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Dist. Atoms Pla 2"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Dist. Atoms Pla 1"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7830
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9843
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9843
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctOpenGL 
      Height          =   7215
      Left            =   120
      ScaleHeight     =   477
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Menu Archivo 
      Caption         =   "&File"
      Begin VB.Menu Abrir 
         Caption         =   "&Open"
      End
      Begin VB.Menu Sortir 
         Caption         =   "&Exit"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuSimetria 
      Caption         =   "&Simmetry"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) 2002 Ramon Yáñez López
' Universitat Autonoma de Barcelona
' Departament de Quimica
' 08193 Cerdanyola del Valles
' Barcelona
' SPAIN
' WWW: http:\\www.uab.es
'
' OpenGL selection code and quaterion rotation matrix
' ported by Michal Husak husakm@vscht.cz
'
' e-mail: ramon.yanez@uab.es
'
' You can modify and freely distribute
' this code, but please keep this header intact.

' +-------------------------------------------------------------------
' |autoredraw = false
' |scalemode pixels
' +-------------------------------------------------------------------

Option Explicit

' +-------------------------------------------------------------------
' | Menu
' +-------------------------------------------------------------------

' +-------------------------------------------------------------------
' | Abrir
' +-------------------------------------------------------------------
Private Sub Abrir_Click()
    borrarStatusBar
    Inicio
    Unload frmGrupPuntual
    Call LecturaDatos
    frmMain.Caption = "AtomsGL v1.5  (c) Ramón Yáñez 2002 UAB" & "|  |" & CommonDialog1.FileTitle
    bDibPla1 = False
    bDibPla2 = False
    render pctOpenGL
End Sub

Private Sub btnSalvarImagen_Click()
SaveBMP_RP "pepe.bmp", w, h
End Sub

' +-------------------------------------------------------------------
' | Simetria
' +-------------------------------------------------------------------
Private Sub mnuSimetria_Click()
If CommonDialog1.Filename = "" Then Exit Sub
BorrarSimetria 'borramos los elementos de simetria
simetria
bsimetria = True
render pctOpenGL
frmGrupPuntual.Caption = GrupPuntual
frmGrupPuntual.Show
End Sub

' +-------------------------------------------------------------------
' | Help
' +-------------------------------------------------------------------
Private Sub help_Click()
frmAbout.Show
End Sub

' +-------------------------------------------------------------------
' | Sortir
' +-------------------------------------------------------------------

Private Sub Sortir_Click()
End
End Sub

' +-------------------------------------------------------------------
' | Frame
' +-------------------------------------------------------------------

Private Sub chkAnglePlans_Click()
If chkAnglePlans Then
    bDibPla1 = True
    bDibPla2 = True
    chkPla1.Enabled = False
    chkDistPlans = False
    chkDistPlans.Enabled = False
    bDistPla1 = False
    Label2.visible = True
    txtAnglePlans.visible = True
    Let plano1 = miPlano 'guardem els atoms del pla1
    Let bcntPlano1 = baricentroPlano
    List2.visible = True
    List2.Clear
Else
    chkPla1.Enabled = True
    chkPla1.Value = False
    bDibPla1 = True
    bDibPla2 = False
    Label2.visible = False
    txtAnglePlans.visible = False
    chkDistPlans.Enabled = True
    List2.visible = False
    List2.Clear
End If
    BorrarSeleccion
    BorrarPlano miPlano
End Sub

Private Sub chkDistPlans_Click()
    If chkDistPlans Then
        bDibPla1 = True
        bDistPla1 = True
        chkPla1.Enabled = False
        chkAnglePlans.Enabled = False
        Let plano1 = miPlano 'guardem els atoms del pla1
        Let bcntPlano1 = baricentroPlano
    Else
        bDibPla1 = False
        bDistPla1 = False
        chkPla1.Enabled = True
        chkPla1.Value = False
        chkAnglePlans.Enabled = True
    End If
    BorrarSeleccion
    List2.Clear
End Sub

Private Sub chkPla1_Click()
If chkPla1 Then
    bDibPla1 = True
    bDibPla2 = False
    chkAnglePlans.visible = True
    chkDistPlans.visible = True
    Label1.visible = True
    List1.visible = True
    List1.Clear
    List2.Clear
    baricentroPlano = baricentro(atmsSelecds())
Else
    bDibPla1 = False
    bDibPla2 = False
    BorrarSeleccion
    chkAnglePlans.visible = False
    chkDistPlans.visible = False
    Label1.visible = False
    List1.visible = False
    List1.Clear
    BorrarPlano miPlano
End If
    displaySelectionInited = GL_FALSE
End Sub


Private Sub Form_Terminate()

'Necessary to avoid infinite Do Loop
    anim = False
    StopFlag = True
' +-------------------------------------------------------------------
' | Release OpenGL if we decide to quit.
' +-------------------------------------------------------------------
    If hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC
    End If
    
    If hPalette <> 0 Then
        DeleteObject hPalette
    End If
Set frmMain = Nothing
End Sub


' +-------------------------------------------------------------------
' | MouseDown
' +-------------------------------------------------------------------
Private Sub pctOpenGL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'On Error Resume Next
    If Shift Then
        If bDistPla1 Then
            SelectMul = False
        Else
            SelectMul = True
        End If
    Else
        SelectMul = False
    End If
    Select Case Button
        Case vbLeftButton
            spinning = False
            anim = False
            moving = True
            beginx = x
            beginy = y
            If Shift = 1 Then
                scaling = True
                SelectMul = True
                Else
                scaling = False
                SelectMul = False
            End If
        Case vbRightButton
            Selection x, y
            StatusBar1.Panels(2).Text = ""
            '***********************************
            ' distancia punt al pla
            '***********************************
            If bDistPla1 Then
                texto = Format(Dist2Plane(atoms(atmsSelecds(1)), miPlano), formato)
                StatusBar1.Panels(2).Text = "Distància pla= " & texto
            End If
            '***********************************
            ' distancies
            '***********************************
            If UBound(atmsSelecds()) = 2 Then
                texto = Format(distanciaAB(atoms(atmsSelecds(1)), _
                                    atoms(atmsSelecds(2))), formato)
                txtDist = texto
            End If
            '***********************************
            'Angles
            '***********************************
            If UBound(atmsSelecds()) = 3 Then
                texto = Format(angleABC(atoms(atmsSelecds(1)), _
                                 atoms(atmsSelecds(2)), _
                                 atoms(atmsSelecds(3))), formato)
                txtAngle = texto
            End If
            '***********************************
            'Diedre
            '***********************************
            If UBound(atmsSelecds()) = 4 Then
                texto = Format(AngleDiedre(atoms(atmsSelecds(1)), _
                                 atoms(atmsSelecds(2)), _
                                 atoms(atmsSelecds(3)), _
                                 atoms(atmsSelecds(4))), formato)
                txtDiedre = texto
            End If
            '***********************************
            'plans
            '***********************************
            If UBound(atmsSelecds()) >= 3 And chkPla1 Then
                baricentroPlano = baricentro(atmsSelecds())
                miPlano = planoMinCuad2 'PLANO PEARSON
                Dim atmSelcc
                If bDibPla2 Then
                    List2.Clear
                Else
                    List1.Clear
                End If
                For Each atmSelcc In atmsSelecds()
                If bDibPla2 Then
                    List2.AddItem Format(atmSelcc & atoms(atmSelcc).simbol, "@@@@@") & "<-->" & _
                               Format(Format(Dist2Plane(atoms(atmSelcc), miPlano), formato), "@@@@@@@")
                Else
                    List1.AddItem Format(atmSelcc & atoms(atmSelcc).simbol, "@@@@@") & "<-->" & _
                               Format(Format(Dist2Plane(atoms(atmSelcc), miPlano), formato), "@@@@@@@")
                End If
                Next
            End If
            If bDibPla2 Then
                texto = Format(angleEntrePlans(plano1, miPlano), formato)
                
                If texto = "0,000" Then
                    StatusBar1.Panels(2).Text = "Distància entre plans= " & Format(distanciaABV(baricentroPlano, bcntPlano1), formato)
                Else
                    txtAnglePlans = texto
                End If
            End If
            Exit Sub
        End Select
End Sub

' +-------------------------------------------------------------------
' | MouseUp
' +-------------------------------------------------------------------
Private Sub pctOpenGL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        moving = False
        scaling = False
    Else
        If Not anim Then render pctOpenGL
    End If
End Sub

' +-------------------------------------------------------------------
' |  MouseMove
' +-------------------------------------------------------------------
Private Sub pctOpenGL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    h = pctOpenGL.ScaleHeight
    w = pctOpenGL.ScaleWidth
    
    If scaling Then
        scalefactor = scalefactor * (1 + ((beginy - y) / h))
        beginx = x
        beginy = y
        newModel = True
        render pctOpenGL
        spinning = True
        GoTo fin
    End If
     
    If moving Then
        TrackBall lastquat, ((2 * beginx - w) / w), ((h - 2 * beginy) / h), _
                ((2 * x - w) / w), ((h - 2 * y) / h)
        beginx = x
        beginy = y
        spinning = True
        anim = True
    End If
fin:
End Sub

' +-------------------------------------------------------------------
' |  Animate
' +-------------------------------------------------------------------
Private Sub Animate()
Dim i As Integer

            add_quats lastquat, curquat, curquat
            newModel = True
            render pctOpenGL
End Sub
' +-------------------------------------------------------------------
' |  Form_Load
' +-------------------------------------------------------------------
Private Sub Form_Load()
' +-------------------------------------------------------------------
' |  Inicializacion Quimica
' +-------------------------------------------------------------------
CommonDialog1.Filter = " PDB (*.PDB)|*.PDB|" & _
                       " CIF (*.CIF)|*.CIF|" & _
                       " XYZ (*.XYZ)|*.XYZ|" & _
                       " INT (*.INT)|*.INT|" & _
                       " FRC (*.FRC)|*.FRC"
Inicio
Call InitGL(pctOpenGL)
    ShowLights
    Show
    Do
            DoEvents                'reaction on mouse and other event enable
            If anim Then Animate
            If StopFlag Then Exit Do
    Loop
End Sub
' +-------------------------------------------------------------------
' |     pctOpenGL_paint
' +-------------------------------------------------------------------
Private Sub pctOpenGL_paint()
    render pctOpenGL
End Sub
' +-------------------------------------------------------------------
' | Form_Unload
' +-------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

'Necessary to avoid infinite Do Loop
    anim = False
    StopFlag = True
' +-------------------------------------------------------------------
' | Release OpenGL if we decide to quit.
' +-------------------------------------------------------------------
    If hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC
    End If
    
    If hPalette <> 0 Then
        DeleteObject hPalette
    End If
End Sub
' +-------------------------------------------------------------------
' |     Form_Resize
' +-------------------------------------------------------------------
Private Sub Form_Resize()
pctOpenGL.Top = 5
pctOpenGL.Left = 5
pctOpenGL.Height = frmMain.ScaleHeight - StatusBar1.Height - 4 'parell ?
pctOpenGL.Width = pctOpenGL.Height
frameSinNombre.Top = 10
frameSinNombre.Left = frmMain.ScaleWidth - frameSinNombre.Width - 10

    h = pctOpenGL.ScaleHeight
    w = pctOpenGL.ScaleWidth
    
    ' avoid colisions with math overflow
    If h = 0 Then h = 1
    If w = 0 Then w = 1
    
    MiProyeccion
    glMatrixMode GL_MODELVIEW
    render pctOpenGL
End Sub

Private Sub Inicio()
leerINI
BorrarSeleccion
End Sub
