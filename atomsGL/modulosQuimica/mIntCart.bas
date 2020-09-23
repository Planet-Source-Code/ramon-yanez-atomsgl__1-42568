Attribute VB_Name = "mIntCart"

'+--------------------------------------------------------
'+ Internas-cartesianas
'+--------------------------------------------------------
'Public Sub intCar(coorInt() As TInternas, atoms() As TATOMO)
'Dim i As Integer, j As Integer, k As Integer, m As Integer
'Dim ccos As Double, cosa As Double
'ReDim coord(3, nAtoms)
'ReDim na(nAtoms) As Integer
'ReDim nb(nAtoms) As Integer
'ReDim nc(nAtoms) As Integer
'Dim ma As Integer, mb As Integer, mc As Integer
'Dim xa As Double, ya As Double, za As Double, xb As Double, yb As Double, zb As Double
'Dim rbc As Double
'Dim xyb As Double
'Dim yza As Double
'Dim xpa As Double, xpb As Double, ypa As Double, xqa As Double, zqa As Double
'Dim cosph As Double, sinph As Double, costh As Double, sinth As Double, coskh As Double, sinkh As Double
'Dim cosd As Double, sina As Double, sind As Double
'Dim xd As Double, yd As Double, zd As Double
'Dim xpd As Double, ypd As Double, zpd As Double
'Dim xqd As Double, yqd As Double, zqd As Double
'Dim xrd As Double
'Const CONV = PI / 180
'ReDim geo(3, nAtoms) As Double
'ReDim geol(3, nAtoms) As Double
'For i = 1 To nAtoms
'    geo(1, i) = coorInt(i).r
'    geo(2, i) = coorInt(i).w
'    geo(3, i) = coorInt(i).t
'    na(i) = coorInt(i).na
'    nb(i) = coorInt(i).nb
'    nc(i) = coorInt(i).nc
'Next
'Debug.Print "CONV "; CONV
'For i = 1 To nAtoms
'geol(1, i) = geo(1, i)
'geol(2, i) = geo(2, i) * CONV
'geol(3, i) = geo(3, i) * CONV
'Next
'coord(1, 1) = 0#
'coord(2, 1) = 0#
'coord(3, 1) = 0#
'coord(1, 2) = geol(1, 2)
'coord(2, 2) = 0#
'coord(3, 2) = 0#
'
'ccos = Cos(geol(2, 3))
'If (na(3) = 1) Then
'    coord(1, 3) = coord(1, 1) + geol(1, 3) * ccos
'Else
'    coord(1, 3) = coord(1, 2) - geol(1, 3) * ccos
'    coord(2, 3) = geol(1, 3) * Sin(geol(2, 3))
'    coord(3, 3) = 0#
'    For i = 4 To nAtoms
'        cosa = Cos(geol(2, i))
'        mb = nb(i)
'        mc = na(i)
'        xb = coord(1, mb) - coord(1, mc)
'        yb = coord(2, mb) - coord(2, mc)
'        zb = coord(3, mb) - coord(3, mc)
'        rbc = 1# / Sqr(xb * xb + yb * yb + zb * zb)
'
'        If Abs(cosa) >= 0.999999991 Then
'            rbc = geol(1, i) * rbc * cosa
'            coord(1, i) = coord(1, mc) + xb * rbc
'            coord(2, i) = coord(2, mc) + yb * rbc
'            coord(3, 1) = coord(3, mc) + zb * rbc
'        Else
'            ma = nc(i)
'            xa = coord(1, ma) - coord(1, mc)
'            ya = coord(2, ma) - coord(2, mc)
'            za = coord(3, ma) - coord(3, mc)
'
'            xyb = Sqr(xb * xb + yb * yb)
'            k = -1
'            If xyb <= 0.1 Then
'                xpa = za
'                za = -xa
'                xa = xpa
'                xpb = zb
'                zb = -xb
'                xb = xpb
'                xyb = Sqr(xb * xb + yb * yb)
'                k = 1
'            End If
'            costh = xb / xyb
'            sinth = yb / xyb
'
'            xpa = xa * costh + ya * sinth
'            ypa = ya * costh - xa * sinth
'            sinph = zb * rbc
'            cosph = Sqr(Abs(1 - sinph * sinph))
'            xqa = xpa * cosph + za * sinph
'            zqa = za * cosph - xpa * sinph
'
'            yza = Sqr(ypa * ypa + zqa * zqa)
'            If ((yza < 0.1) And (yza > 0.0000000001)) Then
'                MsgBox "atoms " & mc & " AND " & _
'                                  mb & " AND " & _
'                                  ma & " Estan en una linea de " & yza & " angstroms"
'            End If
'            coskh = ypa / yza
'            sinkh = zqa / yza
'            If yza < 0.0000000001 Then
'                coskh = 1
'                sinkh = 0
'            End If
'            sina = Sin(geol(2, i))
'            sind = -Sin(geol(3, i))
'            cosd = Cos(geol(3, i))
'            xd = geol(1, i) * cosa
'            yd = geol(1, i) * sina * cosd
'            zd = geol(1, i) * sina * sind
'
'            ypd = yd * coskh - zd * sinkh
'            zpd = zd * coskh + yd * sinkh
'            xpd = xd * cosph - zpd * sinph
'            zqd = zpd * cosph + xd * sinph
'            zqd = xpd * costh - ypd * sinth
'            yqd = ypd * costh + xpd * sinth
'            If (k >= 1) Then
'                xrd = -zqd
'                zqd = xqd
'                zqd = xrd
'            End If
'            coord(1, i) = xqd + coord(1, mc)
'            coord(2, i) = yqd + coord(2, mc)
'            coord(3, i) = zqd + coord(3, mc)
'        End If
'    Next
'End If
'For i = 1 To nAtoms
'    With atoms(i)
'        .simbol = coorInt(i).simbol
'        .type = .simbol
'        .x = coord(1, i)
'        .y = coord(2, i)
'        .Z = coord(3, i)
'        Debug.Print .simbol, .x, .y, .Z
'    End With
'Next
'End Sub
'
