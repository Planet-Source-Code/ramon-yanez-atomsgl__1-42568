Attribute VB_Name = "mVarios"
Option Explicit

'**********************************************************
'
'**********************************************************
Public Sub borrarStatusBar()
    frmMain!StatusBar1.Panels(1).Text = ""
    frmMain!txtDist = ""
    frmMain!txtAngle = ""
    frmMain!txtDiedre = ""
End Sub

'**********************************************************
'
'**********************************************************

Public Function ConvertRGB(ByVal ColorSelect As Long) As TColorRGB
Dim blue As Long, green As Long, red As Long
    blue = Fix(ColorSelect / 65536)
    green = Fix(ColorSelect / 256) Mod 256
    red = Fix(ColorSelect) Mod 256
    ConvertRGB.b = blue
    ConvertRGB.g = green
    ConvertRGB.r = red
End Function


'**********************************************************
'
'**********************************************************

Public Function formatNumero(numero As Double) As String
formatNumero = Format(Format(numero, formato), "@@@@@@@@")
End Function

