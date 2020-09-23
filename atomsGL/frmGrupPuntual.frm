VERSION 5.00
Begin VB.Form frmGrupPuntual 
   Caption         =   "Grup Puntual"
   ClientHeight    =   4920
   ClientLeft      =   8475
   ClientTop       =   6120
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6780
   Begin VB.Frame frInversio 
      Caption         =   "Inversion"
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   855
      Begin VB.CheckBox Check1 
         Caption         =   "i"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame frPlans 
      Caption         =   "Planes"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
      Begin VB.CommandButton btnMostPlans 
         Caption         =   "Show All"
         Height          =   255
         Left            =   1320
         TabIndex        =   80
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton btnEsbPlans 
         Caption         =   "Erase All"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   32
         Top             =   1800
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   31
         Top             =   1800
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   30
         Top             =   1800
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   28
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   27
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   26
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   24
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   23
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   22
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame frEixos 
      Caption         =   "Axes"
      Height          =   2655
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
      Begin VB.CommandButton btnMostC 
         Caption         =   "Show All"
         Height          =   255
         Left            =   1920
         TabIndex        =   79
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton btnEsbC 
         Caption         =   "Erase All"
         Height          =   255
         Left            =   480
         TabIndex        =   75
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   34
         Left            =   3120
         TabIndex        =   67
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   33
         Left            =   2640
         TabIndex        =   66
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   32
         Left            =   2160
         TabIndex        =   65
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   31
         Left            =   1680
         TabIndex        =   64
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   30
         Left            =   1200
         TabIndex        =   63
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   29
         Left            =   720
         TabIndex        =   62
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   61
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   27
         Left            =   3120
         TabIndex        =   60
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   26
         Left            =   2640
         TabIndex        =   59
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   25
         Left            =   2160
         TabIndex        =   58
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   24
         Left            =   1680
         TabIndex        =   57
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   23
         Left            =   1200
         TabIndex        =   56
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   22
         Left            =   720
         TabIndex        =   55
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   21
         Left            =   240
         TabIndex        =   54
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   20
         Left            =   3120
         TabIndex        =   53
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   19
         Left            =   2640
         TabIndex        =   52
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   18
         Left            =   2160
         TabIndex        =   51
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   17
         Left            =   1680
         TabIndex        =   50
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   16
         Left            =   1200
         TabIndex        =   49
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   15
         Left            =   720
         TabIndex        =   48
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   13
         Left            =   3120
         TabIndex        =   46
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   45
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   44
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   43
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   41
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   40
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   39
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   36
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   35
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame frRotImp 
      Caption         =   "Imp Axes"
      Height          =   1695
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton btnBorrarS 
         Caption         =   "Erase All"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton btnMostraS 
         Caption         =   "Show All"
         Height          =   255
         Left            =   2040
         TabIndex        =   77
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   17
         Left            =   3120
         TabIndex        =   74
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   16
         Left            =   2520
         TabIndex        =   73
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   15
         Left            =   1920
         TabIndex        =   72
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   14
         Left            =   1320
         TabIndex        =   71
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   13
         Left            =   720
         TabIndex        =   70
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   16
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   13
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   68
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmGrupPuntual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cc() As Integer
Private ss() As Integer
Private sigm() As Integer
Private bS As Boolean
Private bC As Boolean
Private bSig As Boolean

'+---------------------------------------------------------------------------
'+  Seleccionar Eixos Impropis
'+---------------------------------------------------------------------------
Private Sub Check2_Click(Index As Integer)
Dim nE As Integer
nE = ss(Index + 1)
If bS Then
    bS = False
    Exit Sub
Else
    If nE = 0 Then Exit Sub
        If ElemSim(nE).visible = True Then
            ElemSim(nE).visible = False
        Else
            ElemSim(nE).visible = True
        End If
        bListaSimetria = GL_FALSE
        render frmMain!pctOpenGL
End If
End Sub

'+---------------------------------------------------------------------------
'+  btnBorrarS_Click()
'+---------------------------------------------------------------------------
Private Sub btnBorrarS_Click()
Dim i As Integer, n As Integer
n = UBound(ss())
    For i = 1 To n
        bS = True
        ElemSim(ss(i)).visible = False
        Check2(i - 1).Value = vbUnchecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub

'+---------------------------------------------------------------------------
'+  btnMostrarS_Click()
'+---------------------------------------------------------------------------
Private Sub btnMostraS_Click()
Dim i As Integer, n As Integer
n = UBound(ss())
    For i = 1 To n
    bS = True
        ElemSim(ss(i)).visible = True
        Check2(i - 1).Value = vbChecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub

'+---------------------------------------------------------------------------
'+  Seleccionar Plans
'+---------------------------------------------------------------------------
Private Sub Check3_Click(Index As Integer)
Dim nE As Integer
nE = sigm(Index + 1)
If bSig Then
    bSig = False
    Exit Sub
Else
    If nE = 0 Then Exit Sub
    If ElemSim(nE).visible = True Then
        ElemSim(nE).visible = False
    Else
        ElemSim(nE).visible = True
    End If
    bListaSimetria = GL_FALSE
    render frmMain!pctOpenGL
End If
End Sub
Private Sub btnEsbPlans_Click()
Dim i As Integer, n As Integer
n = UBound(sigm())
    For i = 1 To n
        bSig = True
        ElemSim(sigm(i)).visible = False
        Check3(i - 1).Value = vbUnchecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub
Private Sub btnMostPlans_Click()
Dim i As Integer, n As Integer
n = UBound(sigm())
    For i = 1 To n
    bSig = True
        ElemSim(sigm(i)).visible = True
        Check3(i - 1).Value = vbChecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub
'+---------------------------------------------------------------------------
'+  Seleccionar Eixos
'+---------------------------------------------------------------------------
Private Sub Check4_Click(Index As Integer)
Dim nE As Integer
nE = cc(Index + 1)
If bC Then
    bC = False
    Exit Sub
Else
    If nE = 0 Then Exit Sub
    If ElemSim(nE).visible = True Then
        ElemSim(nE).visible = False
    Else
        ElemSim(nE).visible = True
    End If
    bListaSimetria = GL_FALSE
    render frmMain!pctOpenGL
End If
End Sub
'+---------------------------------------------------------------------------
'+  btnEsbC_Click
'+---------------------------------------------------------------------------
Private Sub btnEsbC_Click()
Dim i As Integer, n As Integer
n = UBound(cc())
    For i = 1 To n
        bC = True
        ElemSim(cc(i)).visible = False
        Check4(i - 1).Value = vbUnchecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub
'+---------------------------------------------------------------------------
'+  btnMostC_Click
'+---------------------------------------------------------------------------
Private Sub btnMostC_Click()
Dim i As Integer, n As Integer
n = UBound(cc())
    For i = 1 To n
    bC = True
        ElemSim(cc(i)).visible = True
        Check4(i - 1).Value = vbChecked
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
End Sub

'+---------------------------------------------------------------------------
'+  FORM_LOAD()
'+---------------------------------------------------------------------------
Private Sub Form_Load()
Debug.Print "HOLA"
Label1.Caption = GrupPuntual
Dim i As Integer
Dim e As String
Dim c As Integer, s As Integer, sigma As Integer
For i = 1 To nElemSim
    e = ElemSim(i).tipo
    If e = "sigma" Then                 ' SIGMA
        sigma = sigma + 1
        ReDim Preserve sigm(sigma)
        Check3(sigma - 1).visible = True
        Check3(sigma - 1).Tag = i
        Check3(sigma - 1).Caption = i
        sigm(sigma) = i
        ElseIf Left(e, 1) = "C" Then    ' ejes C
            Select Case Mid(e, 2, 1)
                Case "2"
                    c = c + 1
                    checkC c, i, "c2"
                Case "3"
                    c = c + 1
                    checkC c, i, "c3"
                Case "4"
                    c = c + 1
                    checkC c, i, "c4"
                Case "5"
                    c = c + 1
                    checkC c, i, "c5"
                Case "6"
                    c = c + 1
                    checkC c, i, "c6"
            End Select
            ElseIf Left(e, 1) = "S" Then    'ejes S
                Select Case Mid(e, 2, 1)
                    Case "3"
                        s = s + 1
                        checkS s, i, "s3"
                    Case "4"
                        s = s + 1
                        checkS s, i, "s4"
                    Case "5"
                        s = s + 1
                        checkS s, i, "s5"
                    Case "6"
                        s = s + 1
                        checkS s, i, "s6"
                    Case "8"
                        s = s + 1
                        checkS s, i, "s8"
                    Case "1"
                        s = s + 1
                        checkS s, i, "s10"
                End Select
                ElseIf Left(e, 1) = "I" Then
                    Check1.visible = True
                    Check1.Value = 1
                    Check1.Tag = i
        End If
Next
End Sub

'+---------------------------------------------------------------------------
'+  CheckC
'+---------------------------------------------------------------------------
Private Sub checkC(c As Integer, n As Integer, s As String)
    ReDim Preserve cc(c)
    Check4(c - 1).Caption = s
    Check4(c - 1).visible = True
    cc(c) = n
End Sub

'+---------------------------------------------------------------------------
'+  CheckA
'+---------------------------------------------------------------------------
Private Sub checkS(c As Integer, n As Integer, s As String)
    ReDim Preserve ss(c)
    Check2(c - 1).Caption = s
    Check2(c - 1).visible = True
    ss(c) = n
End Sub

'+---------------------------------------------------------------------------
'+  FORM_TERMINATE()
'+---------------------------------------------------------------------------
Private Sub Form_Terminate()
Dim i As Integer
    For i = 1 To nElemSim
        ElemSim(i).visible = True
    Next
bListaSimetria = GL_FALSE
render frmMain!pctOpenGL
Set frmGrupPuntual = Nothing
End Sub

