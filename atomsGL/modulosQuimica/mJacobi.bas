Attribute VB_Name = "mJacobi"
Option Explicit
Option Base 1


' +-------------------------------------------------------------------
' |         subrutina Jacobi
' |         diagonaliza una matriz mediante el metodo ciclico de jacobi
' +-------------------------------------------------------------------

Public Sub Jacobi(n As Integer, a() As Double, d() As Double, v() As Double)
Dim maxrot  As Integer, nrot As Integer
Dim ip As Integer, iq As Integer, np As Integer
Dim i As Integer, j As Integer, k As Integer
Dim sm As Double, tresh As Double, s As Double, c As Double, t As Double
Dim theta As Double, tau As Double, h As Double, g As Double, p As Double
Dim b(1 To 5)
Dim Z(1 To 5)
maxrot = 1000
nrot = 0
' matriz unidad
For ip = 1 To n
    For iq = 1 To n
        v(ip, iq) = 0#
    Next iq
    v(ip, ip) = 1#
Next ip
For ip = 1 To n
    b(ip) = a(ip, ip)
    d(ip) = b(ip)
    Z(ip) = 0#
Next ip
' perform the jacobi rotations
For i = 1 To maxrot
    sm = 0#
    For ip = 1 To n - 1
        For iq = ip + 1 To n
            sm = sm + Abs(a(ip, iq))
        Next iq
    Next ip
    If sm = 0# Then GoTo 10
    If i < 4 Then
        tresh = 0.2 * sm / n ^ 2
        Else
        tresh = 0#
    End If
    For ip = 1 To n - 1
        For iq = ip + 1 To n
            g = 100# * Abs(a(ip, iq))
            If i > 4 And (Abs(d(ip)) + g) = Abs(d(ip)) And (Abs(d(iq)) + g) = Abs(d(iq)) Then
                    a(ip, iq) = 0#
            ElseIf (Abs(a(ip, iq)) > tresh) Then
                  h = d(iq) - d(ip)
                  If (Abs(h) + g) = Abs(h) Then
                      t = a(ip, iq) / h
                  Else
                      theta = 0.5 * h / a(ip, iq)
                      t = 1# / (Abs(theta) + Sqr(1# + theta ^ 2))
                      If theta < 0# Then t = -t
                  End If
                  c = 1# / Sqr(1 + t ^ 2)
                  s = t * c
                  tau = s / (1# + c)
                  h = t * a(ip, iq)
                  Z(ip) = Z(ip) - h
                  Z(iq) = Z(iq) + h
                  d(ip) = d(ip) - h
                  d(iq) = d(iq) + h
                  a(ip, iq) = 0#
                  For j = 1 To ip - 1
                      g = a(j, ip)
                      h = a(j, iq)
                      a(j, ip) = g - s * (h + g * tau)
                      a(j, iq) = h + s * (g - h * tau)
                  Next j
                  For j = ip + 1 To iq - 1
                      g = a(ip, j)
                      h = a(j, iq)
                      a(ip, j) = g - s * (h + g * tau)
                      a(j, iq) = h + s * (g - h * tau)
                  Next j
                  For j = iq + 1 To n
                      g = a(ip, j)
                      h = a(iq, j)
                      a(ip, j) = g - s * (h + g * tau)
                      a(iq, j) = h + s * (g - h * tau)
                  Next j
                  For j = 1 To n
                      g = v(j, ip)
                      h = v(j, iq)
                      v(j, ip) = g - s * (h + g * tau)
                      v(j, iq) = h + s * (g - h * tau)
                  Next j
                  nrot = nrot + 1
            End If
        Next iq
    Next ip
    For ip = 1 To n
        b(ip) = b(ip) + Z(ip)
        d(ip) = b(ip)
        Z(ip) = 0#
    Next ip
Next i
       
10: If nrot = maxrot Then
            MsgBox ("Diagonalización no converge")
            Exit Sub
    End If
'***********************************************
'
'***********************************************

For i = 1 To n - 1
    k = i
    p = d(i)
    For j = i + 1 To n
        If (d(j) < p) Then
            k = j
            p = d(j)
        End If
    Next j
    If (k <> i) Then
        d(k) = d(i)
        d(i) = p
        For j = 1 To n
            p = v(j, i)
            v(j, i) = v(j, k)
            v(j, k) = p
        Next j
    End If
Next i
        
End Sub
