Attribute VB_Name = "MDefiniciones"
Option Explicit
Option Base 1
'**********************************************************
Public Type TColorRGB
    r As Integer
    g As Integer
    b As Integer
End Type
'**********************************************************
Public Type TENLACES
    a As Integer
    ca As TColorRGB 'color atomo A
    b As Integer
    cb As TColorRGB 'color atomo B
    'definiciones para los enlacesGL
    d As Single 'distancia
    x0 As Single
    y0 As Single
    z0 As Single
    angle As Single
    vx As Single
    vy As Single
    vz As Single
End Type
'**********************************************************
Public Type TATOMO
    simbol As String
    label As String
    x As Double     'coordenadas cartesianas
    y As Double
    Z As Double
    fracX As Double '
    fracY As Double '
    fracZ As Double '
    radioCov As Double
    color As Integer
    numeroAtomico As Integer
    u(3, 3) As Double
End Type
'**********************************************************
Public Type TUATOMOANISO
    label As String
    u(3, 3) As Double
End Type
'**********************************************************
Public Type TVECTOR
    x As Double
    y As Double
    Z As Double
End Type
'**********************************************************
Public Type TATOMGL
    simbol As String * 2
    radioCov As Single
    color As Integer
End Type
'**********************************************************
Public Type TCELDA
    a As Double
    b As Double
    c As Double
    alfa As Double
    beta As Double
    gamma As Double
End Type
'**********************************************************
Public Type TPlano
    a As Double
    b As Double
    c As Double
    d As Double
End Type

'**************************************
Public atomGL(100) As TATOMGL
Public nParametros As Integer
'**************************************
Public atoms() As TATOMO
Public UatomsAniso() As TUATOMOANISO
Public enlace() As TENLACES
Public celda As TCELDA
Public xcm As Double, ycm As Double, zcm As Double
'**************************************
Public origenX As Single, origenY As Single
Public nAtoms As Integer
Public nEllipses As Integer

Public nEnlaces As Integer
Public fitxer As String
Public pointer As Integer ' tipus d'icona del mouse
Public escX As Single, escY As Single
Public vMomentoInercia1 As TVECTOR
Public vMomentoInercia2 As TVECTOR
Public vMomentoInercia3 As TVECTOR
'**************************************
Public Const PI = 3.1415926
Public Const cosAngulo = 0.999390827
Public Const sinAngulo = 0.034899496
Public Const formato = "#,###0.000"
Public texto As String

'**********************************************************
' variables definidas para la seleccion
'**********************************************************
Public seleccion As Integer
Public SelectMul As Boolean
Public atmsSelecds() As Integer
Public seleccionados As Integer

'**********************************************************
' variables definidas para dibujar el plano
'**********************************************************
Public dibujarPlano As Boolean
Public miPlano As TPlano
Public baricentroPlano As TVECTOR
Public plano1 As TPlano
Public plano2 As TPlano

'**********************************************************
' LISTAS
'**********************************************************
Public Const LISTA_MOLECULA = 1
Public bDibPla1 As Boolean
Public bDibPla2 As Boolean
Public bDistPla1 As Boolean
Public bcntPlano1 As TVECTOR
Public Const LISTA_ATOMOS_SELECCIONADOS = 4
Public Const LISTA_SIMETRIA = 3
'**********************************************************
' ESCALA MOLECULA
'**********************************************************
Public xmin As Double, xmax As Double
Public ymin As Double, ymax As Double
Public zmin As Double, zmax As Double
Public maximus As Double

'**********************************************************
' SIMETRIA
'**********************************************************
Public bsimetria As Boolean

