Attribute VB_Name = "mRreadINI"
Option Explicit
Option Base 1

' +-------------------------------------------------------------------
' | leerINI
' +-------------------------------------------------------------------
Public Sub leerINI()

Dim dummy As String
Dim stext() As String

Open App.Path & "\atomgl.ini" For Input As #1
nParametros = 0
Do While Not (EOF(1))
nParametros = nParametros + 1
    With atomGL(nParametros)
        Input #1, dummy
        stext = NeatSplit(dummy)
        .simbol = Trim(stext(0))
        .radioCov = Val(stext(1))
        .color = Val(stext(2))
    End With
Loop
Close #1
End Sub
