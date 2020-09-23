Attribute VB_Name = "mXYZATM"
Option Explicit
'c
'c
'c     ###################################################
'c     ##  COPYRIGHT (C)  1990  by  Jay William Ponder  ##
'c     ##              All Rights Reserved              ##
'c     ###################################################
Public Type TInternas
    simbol As String
    type As String
    r As Double
    w As Double
    t As Double
    na As Integer
    nb As Integer
    nc As Integer
End Type
Public coorInt() As TInternas
'c
'c     ################################################################
'c     ##                                                            ##
'c     ##  subroutine xyzatm  --  single atom internal to Cartesian  ##
'c     ##                                                            ##
'c     ################################################################
'c
'c
'c     "xyzatm" computes the Cartesian coordinates of a single
'c     atom from its defining internal coordinate values
'c
'c

Public Sub intCar()
Dim i As Integer
For i = 1 To nAtoms
xyzatm i, coorInt(i).na, coorInt(i).r, _
          coorInt(i).nb, coorInt(i).w, _
          coorInt(i).nc, coorInt(i).t, 0
Next
For i = 1 To nAtoms
    Debug.Print atoms(i).simbol, atoms(i).x, atoms(i).y, atoms(i).Z
Next
End Sub

Private Sub xyzatm(i As Integer, ia As Integer, bond As Double, _
                                 ib As Integer, angle1 As Double, _
                                 ic As Integer, angle2 As Double, chiral As Integer)

Dim eps, rad1, rad2
Dim sin1, cos1, sin2, cos2
Dim cosine, sine, sine2
Dim xab, yab, zab, rab
Dim xba, yba, zba, rba
Dim xbc, ybc, zbc, rbc
Dim xac, yac, zac, rac
Dim xt, yt, zt, xu, yu, zu
Dim cosb, sinb, cosg, sing
Dim xtmp, ztmp, a, b, c
Const CONV = PI / 180
'c
'c
'c     convert angles to radians, and get their sines and cosines
'c
      eps = 0.00000001
      rad1 = angle1 * CONV
      rad2 = angle2 * CONV
      sin1 = Sin(rad1)
      cos1 = Cos(rad1)
      sin2 = Sin(rad2)
      cos2 = Cos(rad2)
'c
'c     if no second site given, place the atom at the origin
'c
      If (ia = 0) Then
         atoms(i).x = 0#
         atoms(i).y = 0#
         atoms(i).Z = 0#
'c
'c     if no third site given, place the atom along the z-axis
'c
      Else
        If (ib = 0) Then
         atoms(i).x = atoms(ia).x
         atoms(i).y = atoms(ia).y
         atoms(i).Z = atoms(ia).Z + bond
'c
'c     if no fourth site given, place the atom in the x,z-plane
'c
        Else
            If (ic = 0) Then
                xab = atoms(ia).x - atoms(ib).x
                yab = atoms(ia).y - atoms(ib).y
                zab = atoms(ia).Z - atoms(ib).Z
                rab = Sqr(xab ^ 2 + yab ^ 2 + zab ^ 2)
                xab = xab / rab
                yab = yab / rab
                zab = zab / rab
                cosb = zab
                sinb = Sqr(xab ^ 2 + yab ^ 2)
                If (sinb = 0#) Then
                    cosg = 1#
                    sing = 0#
                Else
                    cosg = yab / sinb
                    sing = xab / sinb
                End If
                xtmp = bond * sin1
                ztmp = rab - bond * cos1
                atoms(i).x = atoms(ib).x + xtmp * cosg + ztmp * sing * sinb
                atoms(i).y = atoms(ib).y - xtmp * sing + ztmp * cosg * sinb
                atoms(i).Z = atoms(ib).Z + ztmp * cosb
'c
'c     general case where the second angle is a dihedral angle
'c
            Else
                If (chiral = 0) Then
                    xab = atoms(ia).x - atoms(ib).x
                    yab = atoms(ia).y - atoms(ib).y
                    zab = atoms(ia).Z - atoms(ib).Z
                    rab = Sqr(xab ^ 2 + yab ^ 2 + zab ^ 2)
                    xab = xab / rab
                    yab = yab / rab
                    zab = zab / rab
                    xbc = atoms(ib).x - atoms(ic).x
                    ybc = atoms(ib).y - atoms(ic).y
                    zbc = atoms(ib).Z - atoms(ic).Z
                    rbc = Sqr(xbc ^ 2 + ybc ^ 2 + zbc ^ 2)
                    xbc = xbc / rbc
                    ybc = ybc / rbc
                    zbc = zbc / rbc
                    xt = zab * ybc - yab * zbc
                    yt = xab * zbc - zab * xbc
                    zt = yab * xbc - xab * ybc
                    cosine = xab * xbc + yab * ybc + zab * zbc
                    sine = Sqr(1# - cosine ^ 2)
                    If (Abs(cosine) >= 1#) Then
                        Debug.Print "--  Undefined Dihedral"
                    End If
                    xt = xt / sine
                    yt = yt / sine
                    zt = zt / sine
                    xu = yt * zab - zt * yab
                    yu = zt * xab - xt * zab
                    zu = xt * yab - yt * xab
                    atoms(i).x = atoms(ia).x + bond * (xu * sin1 * cos2 + xt * sin1 * sin2 - xab * cos1)
                    atoms(i).y = atoms(ia).y + bond * (yu * sin1 * cos2 + yt * sin1 * sin2 - yab * cos1)
                    atoms(i).Z = atoms(ia).Z + bond * (zu * sin1 * cos2 + zt * sin1 * sin2 - zab * cos1)
'c
'c     general case where the second angle is a bond angle
'c
                Else
                    If (Abs(chiral) = 1) Then
                        xba = atoms(ib).x - atoms(ia).x
                        yba = atoms(ib).y - atoms(ia).y
                        zba = atoms(ib).Z - atoms(ia).Z
                        rba = Sqr(xba ^ 2 + yba ^ 2 + zba ^ 2)
                        xba = xba / rba
                        yba = yba / rba
                        zba = zba / rba
                        xac = atoms(ia).x - atoms(ic).x
                        yac = atoms(ia).y - atoms(ic).y
                        zac = atoms(ia).Z - atoms(ic).Z
                        rac = Sqr(xac ^ 2 + yac ^ 2 + zac ^ 2)
                        xac = xac / rac
                        yac = yac / rac
                        zac = zac / rac
                        xt = zba * yac - yba * zac
                        yt = xba * zac - zba * xac
                        zt = yba * xac - xba * yac
                        cosine = xba * xac + yba * yac + zba * zac
                        sine2 = 1# - cosine ^ 2
                        If (Abs(cosine) >= 1#) Then
                            Debug.Print "Defining Atoms Colinear"
                        End If
                        a = (-cos2 - cosine * cos1) / sine2
                        b = (cos1 + cosine * cos2) / sine2
                        c = (1# + a * cos2 - b * cos1) / sine2
                        If (c > eps) Then
                           c = chiral * Sqr(c)
                        Else
                            If (c < -eps) Then
                                c = Sqr((a * xac + b * xba) ^ 2 + (a * yac + b * yba) ^ 2 + (a * zac + b * zba) ^ 2)
                                a = a / c
                                b = b / c
                                c = 0#
                            End If
                        End If
                    Else
                        c = 0#
                    End If
                    atoms(i).x = atoms(ia).x + bond * (a * xac + b * xba + c * xt)
                    atoms(i).y = atoms(ia).y + bond * (a * yac + b * yba + c * yt)
                    atoms(i).Z = atoms(ia).Z + bond * (a * zac + b * zba + c * zt)
                End If
            End If
        End If
     End If
End Sub
