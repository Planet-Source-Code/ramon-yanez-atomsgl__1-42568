Attribute VB_Name = "GLTrackBall"
Option Explicit

' +-------------------------------------------------------------------
' | Trackball declarations
' +-------------------------------------------------------------------
Public spinning As Boolean, moving As Boolean
Public beginx As Integer, beginy As Integer
Public w As Integer, h As Integer
Public curquat(3) As GLfloat
Public lastquat(3) As GLfloat
Public newModel As Boolean
Public scaling As Boolean
Public scalefactor As Single
Public anim As Boolean
Public StopFlag As Boolean

' +-------------------------------------------------------------------
' | End of Trackball declarations
' +-------------------------------------------------------------------

Const TRACKBALLSIZE = 0.8
Const RENORMCOUNT = 97

' +-------------------------------------------------------------------
' | Local function prototypes (not defined in trackball.h)
' |
' | static float tb_project_to_sphere(float, float, float);
' +-------------------------------------------------------------------'
Public Function asin(t As Single) 'asin not defined in VB
If t = 1 Then
        asin = 3.14159265358979 / 2
    Else
        asin = Atn(t / (Sqr(1 - t * t)))
End If
End Function

' +-------------------------------------------------------------------
Public Sub vzero(v)
    
    v(0) = 0#
    v(1) = 0#
    v(2) = 0#

End Sub

' +-------------------------------------------------------------------
Public Sub vset(v, x As Single, y As Single, Z As Single)
    v(0) = x
    v(1) = y
    v(2) = Z
End Sub

' +-------------------------------------------------------------------
Public Sub vsub(src1, src2, dst)
    dst(0) = src1(0) - src2(0)
    dst(1) = src1(1) - src2(1)
    dst(2) = src1(2) - src2(2)
End Sub

' +-------------------------------------------------------------------
Public Sub vcopy(v1, v2)
Dim i As Integer
    For i = 0 To 3
        v2(i) = v1(i)
    Next i
End Sub

' +-------------------------------------------------------------------
Public Sub vcross(v1, v2, cross)
    Dim temp(3) As Single

    temp(0) = (v1(1) * v2(2)) - (v1(2) * v2(1))
    temp(1) = (v1(2) * v2(0)) - (v1(0) * v2(2))
    temp(2) = (v1(0) * v2(1)) - (v1(1) * v2(0))
    vcopy temp, cross
End Sub

' +-------------------------------------------------------------------
Public Function vlength(v)
    vlength = Sqr(v(0) * v(0) + v(1) * v(1) + v(2) * v(2))
End Function

' +-------------------------------------------------------------------
Public Sub vscale(v, div As Single)
    v(0) = v(0) * div
    v(1) = v(1) * div
    v(2) = v(2) * div
End Sub

' +-------------------------------------------------------------------
Public Sub vnormal(v)
    vscale v, 1# / vlength(v)
End Sub

' +-------------------------------------------------------------------
Public Function vdot(v1, v2)
vdot = v1(0) * v2(0) + v1(1) * v2(1) + v1(2) * v2(2)
End Function

' +-------------------------------------------------------------------
Public Sub vadd(src1, src2, dst)
    dst(0) = src1(0) + src2(0)
    dst(1) = src1(1) + src2(1)
    dst(2) = src1(2) + src2(2)
End Sub

' +-------------------------------------------------------------------
' | Ok, simulate a track-ball.  Project the points onto the virtual
' | trackball, then figure out the axis of rotation, which is the cross
' | product of P1 P2 and O P1 (O is the center of the ball, 0,0,0)
' | Note:  This is a deformed trackball-- is a trackball in the center,
' | but is deformed into a hyperbolic sheet of rotation away from the
' | center.  This particular function was chosen after trying out
' | several variations.
' |
' | It is assumed that the arguments to this routine are in the range
' | (-1.0 ... 1.0)
' +-------------------------------------------------------------------
Public Sub TrackBall(q, p1x As Single, p1y As Single, p2x As Single, p2y As Single)



Dim a(3) As Single '; /* Axis of rotation */
Dim phi As Single ' /* how much to rotate about axis */
Dim p1(3) As Single, p2(3) As Single, d(3) As Single
Dim t As Single

    If (p1x = p2x And p1y = p2y) Then
        '/* Zero rotation */
        vzero (q)
        q(3) = 1#
        GoTo fin
    Else
    

    '/*
    ' * First, figure out z-coordinates for projection of P1 and P2 to
    ' * deformed sphere
    ' */
    
    vset p1, p1x, p1y, tb_project_to_sphere(TRACKBALLSIZE, p1x, p1y)
    vset p2, p2x, p2y, tb_project_to_sphere(TRACKBALLSIZE, p2x, p2y)

    '/*
    ' *  Now, we want the cross product of P1 and P2
    ' */
    vcross p2, p1, a

    '/*
    ' *  Figure out how much to rotate around that axis.
    ' */
    vsub p1, p2, d
    t = vlength(d) / (2# * TRACKBALLSIZE)

    '/*
    ' * Avoid problems with out-of-control values...
    ' */
    If (t > 1#) Then t = 1#
    If (t < -1#) Then t = -1#
    phi = 2# * asin(t)

    axis_to_quat a, phi, q
End If
fin:
End Sub

' +-------------------------------------------------------------------
' |  Given an axis and angle, compute quaternion.
' +-------------------------------------------------------------------
Public Sub axis_to_quat(a, phi As Single, q)
    vnormal a
    vcopy a, q
    vscale q, Sin(phi / 2#)
    q(3) = Cos(phi / 2#)

End Sub

' +-------------------------------------------------------------------
' | Project an x,y pair onto a sphere of radius r OR a hyperbolic sheet
' | if we are away from the center of the sphere.
' +-------------------------------------------------------------------

Public Function tb_project_to_sphere(r As Single, x As Single, y As Single)
Dim d As Single, t As Single, Z As Single

    d = Sqr(x * x + y * y)
    'inside sphere
    If (d < r * 0.707106781186548) Then
        Z = Sqr(r * r - d * d)
     ' /* On hyperbola */
    Else
        t = r / 1.4142135623731
        Z = t * t / d
    End If
    
    tb_project_to_sphere = Z

End Function

' +-------------------------------------------------------------------
' |Given two rotations, e1 and e2, expressed as quaternion rotations,
' | figure out the equivalent single rotation and stuff it into dest.
' |
' | This routine also normalizes the result every RENORMCOUNT times it is
' | called, to keep error from creeping in.
' |
' | NOTE: This routine is written so that q1 or q2 may be the same
' | as dest (or each other).
' +-------------------------------------------------------------------


Public Sub add_quats(q1, q2, dest)
   Static count As Integer
    count = 0
    Dim t1(4) As Single, t2(4) As Single, t3(4) As Single
    Dim tf(4) As Single

    vcopy q1, t1
    vscale t1, (q2(3))

    vcopy q2, t2
    vscale t2, (q1(3))

    vcross q2, q1, t3
    vadd t1, t2, tf
    vadd t3, tf, tf
    tf(3) = q1(3) * q2(3) - vdot(q1, q2)

    dest(0) = tf(0)
    dest(1) = tf(1)
    dest(2) = tf(2)
    dest(3) = tf(3)
    normalize_quat dest
    

End Sub

' +-------------------------------------------------------------------
' |Quaternions always obey:  a^2 + b^2 + c^2 + d^2 = 1.0
' |If they don't add up to 1.0, dividing by their magnitued will
' |renormalize them.
' |
' |Note: See the following for more information on quaternions:
' |
' |- Shoemake, K., Animating rotation with quaternion curves, Computer
' |  Graphics 19, No 3 (Proc. SIGGRAPH'85), 245-254, 1985.
' |- Pletinckx, D., Quaternion calculus as a basic tool in computer
' |  graphics, The Visual Computer 5, 2-13, 1989.
' +-------------------------------------------------------------------
Public Sub normalize_quat(q)



   Dim i As Integer
   Dim mag As Single

    mag = Sqr(q(0) * q(0) + q(1) * q(1) + q(2) * q(2) + q(3) * q(3))
    For i = 0 To 3
        q(i) = q(i) / mag
    Next i
End Sub

' +-------------------------------------------------------------------
' * Build a rotation matrix, given a quaternion rotation.
' +-------------------------------------------------------------------

Public Sub build_rotmatrix(m, q)
    m(0) = 1 - 2 * (q(1) * q(1) + q(2) * q(2))
    m(1) = 2 * (q(0) * q(1) - q(2) * q(3))
    m(2) = 2 * (q(2) * q(0) + q(1) * q(3))
    m(3) = 0

    m(4) = 2 * (q(0) * q(1) + q(2) * q(3))
    m(5) = 1 - 2 * (q(2) * q(2) + q(0) * q(0))
    m(6) = 2 * (q(1) * q(2) - q(0) * q(3))
    m(7) = 0

    m(8) = 2 * (q(2) * q(0) - q(1) * q(3))
    m(9) = 2 * (q(1) * q(2) + q(0) * q(3))
    m(10) = 1 - 2 * (q(1) * q(1) + q(0) * q(0))
    m(11) = 0

    m(12) = 0
    m(13) = 0
    m(14) = 0
    m(15) = 1
End Sub

